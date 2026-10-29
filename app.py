import os
import json
import logging
import threading
from flask import Flask, render_template, request, Response, jsonify, stream_with_context
from flask_cors import CORS

from config import (
    PORT, SHAREPOINT_SITE_URL, SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH,
    AZURE_OPENAI_DEPLOYMENT, AZURE_OPENAI_ENDPOINT
)
from sharepoint_service import SharePointService
from bot_engine import SharePointBotEngine

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("sharepoint_app")

app = Flask(__name__)
CORS(app)

# Initialize SharePoint Service & Bot Engine
sp_service = SharePointService()
bot_engine = SharePointBotEngine(sp_service)

# Attempt to load from disk cache on startup, otherwise trigger background initial sync
def initialize_data():
    loaded = sp_service.load_cache()
    if not loaded:
        logger.info("No existing cache found. Initializing background sync...")
        t = threading.Thread(target=sp_service.sync_all, daemon=True)
        t.start()
    else:
        logger.info(f"Loaded existing cache with {len(sp_service.lists_metadata)} lists and {len(sp_service.drives_metadata)} libraries.")

# Trigger init on import
initialize_data()

@app.route("/")
def index():
    return render_template(
        "index.html",
        site_url=SHAREPOINT_SITE_URL,
        site_name=sp_service.site_info.get("displayName", "AISearchDev"),
        deployment=AZURE_OPENAI_DEPLOYMENT
    )

@app.route("/api/status")
def status():
    """Return health & synchronization status."""
    return jsonify({
        "status": "syncing" if sp_service.is_syncing else ("online" if sp_service.is_synced else "offline"),
        "sync_message": sp_service.sync_status_message,
        "last_sync": sp_service.last_sync_time,
        "site": {
            "name": sp_service.site_info.get("displayName", "AISearchDev"),
            "url": SHAREPOINT_SITE_URL,
            "hostname": SHAREPOINT_HOSTNAME,
            "path": SHAREPOINT_SITE_PATH
        },
        "stats": {
            "lists_count": len(sp_service.lists_metadata),
            "libraries_count": len(sp_service.drives_metadata),
            "total_list_items": sum(len(v) for v in sp_service.list_items_by_id.values()),
            "total_documents": sum(len(v) for v in sp_service.files_by_drive.values()),
            "indexed_chunks": len(sp_service.document_chunks)
        },
        "ai": {
            "endpoint": AZURE_OPENAI_ENDPOINT,
            "deployment": AZURE_OPENAI_DEPLOYMENT
        }
    })

@app.route("/api/sync", methods=["POST"])
def sync_data():
    """Trigger manual re-sync of SharePoint Lists & Document Libraries."""
    if sp_service.is_syncing:
        return jsonify({
            "success": True,
            "status": "syncing",
            "message": "Sync is already in progress..."
        })

    # Start sync in background thread so HTTP call returns immediately without timing out
    def run_sync_thread():
        sp_service.sync_all(force_refresh=False)

    t = threading.Thread(target=run_sync_thread, daemon=True)
    t.start()
    return jsonify({
        "success": True,
        "status": "started",
        "message": "Sync started in background"
    })

@app.route("/api/lists", methods=["GET"])
def get_lists():
    """List all available SharePoint lists with schema and preview counts."""
    lists = []
    for l in sp_service.lists_metadata:
        items = sp_service.list_items_by_id.get(l["id"], [])
        lists.append({
            "id": l["id"],
            "name": l["name"],
            "template": l["template"],
            "item_count": len(items),
            "webUrl": l.get("webUrl"),
            "columns": l.get("columns", [])
        })
    return jsonify({"lists": lists})

@app.route("/api/lists/<list_id>/items", methods=["GET"])
def get_list_items(list_id):
    """Get items of a specific list with optional search query."""
    q = request.args.get("q", "").strip().lower()
    page = int(request.args.get("page", 1))
    limit = int(request.args.get("limit", 20))
    
    items = sp_service.list_items_by_id.get(list_id, [])
    if q:
        filtered = []
        for it in items:
            fields_str = json.dumps(it.get("fields", {})).lower()
            if q in fields_str:
                filtered.append(it)
        items = filtered

    total = len(items)
    start = (page - 1) * limit
    paged = items[start:start + limit]

    return jsonify({
        "total": total,
        "page": page,
        "limit": limit,
        "items": paged
    })

@app.route("/api/libraries", methods=["GET"])
def get_libraries():
    """List all Document Libraries and their files."""
    libraries = []
    for d in sp_service.drives_metadata:
        drive_id = d["id"]
        files = sp_service.files_by_drive.get(drive_id, [])
        libraries.append({
            "id": drive_id,
            "name": d["name"],
            "description": d.get("description", ""),
            "driveType": d.get("driveType"),
            "file_count": len(files),
            "files": files,
            "webUrl": d.get("webUrl")
        })
    return jsonify({"libraries": libraries})

@app.route("/api/upload", methods=["POST"])
def upload_file():
    """Upload a document to a SharePoint Document Library and index it for RAG."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files["file"]
    if not file or not file.filename:
        return jsonify({"error": "Empty file name"}), 400

    drive_id = request.form.get("drive_id")
    if not drive_id and sp_service.drives_metadata:
        drive_id = sp_service.drives_metadata[0]["id"]

    try:
        file_bytes = file.read()
        res = sp_service.upload_document(drive_id, file.filename, file_bytes)
        return jsonify(res)
    except Exception as e:
        logger.exception("Upload failed")
        return jsonify({"error": str(e)}), 500

@app.route("/api/chat", methods=["POST"])
def chat():
    """Standard non-streaming chat endpoint."""
    try:
        data = request.get_json(force=True)
        messages = data.get("messages", [])
        if not messages:
            return jsonify({"error": "No messages provided"}), 400

        result = bot_engine.chat(messages)
        return jsonify(result)
    except Exception as e:
        logger.exception("Error in /api/chat")
        return jsonify({"error": str(e)}), 500

@app.route("/api/chat/stream", methods=["POST"])
def chat_stream():
    """Server-Sent Events streaming endpoint with real-time tool execution status."""
    try:
        data = request.get_json(force=True)
        messages = data.get("messages", [])
        if not messages:
            return jsonify({"error": "No messages provided"}), 400

        return Response(
            stream_with_context(bot_engine.chat_stream(messages)),
            content_type="text/event-stream",
            headers={
                "Cache-Control": "no-cache",
                "X-Accel-Buffering": "no",
                "Connection": "keep-alive"
            }
        )
    except Exception as e:
        logger.exception("Error in /api/chat/stream")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    logger.info(f"Starting SharePoint AI Bot on http://localhost:{PORT}")
    app.run(host="0.0.0.0", port=PORT, debug=False)
