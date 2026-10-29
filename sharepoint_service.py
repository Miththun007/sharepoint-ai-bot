import os
import time
import json
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional
import msal
import requests

from config import (
    AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET,
    SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH,
    GRAPH_API_BASE, AUTHORITY_URL, GRAPH_SCOPES, CACHE_DIR
)
from document_parser import parse_file
from search_index import SharePointIndexManager

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("sharepoint_service")

class SharePointService:
    def __init__(self):
        self.tenant_id = AZURE_TENANT_ID
        self.client_id = AZURE_CLIENT_ID
        self.client_secret = AZURE_CLIENT_SECRET
        self.hostname = SHAREPOINT_HOSTNAME
        self.site_path = SHAREPOINT_SITE_PATH
        
        self.site_id: Optional[str] = None
        self.site_info: Dict[str, Any] = {}
        
        # In-memory storage
        self.lists_metadata: List[Dict[str, Any]] = []
        self.list_items_by_id: Dict[str, List[Dict[str, Any]]] = {}
        self.drives_metadata: List[Dict[str, Any]] = []
        self.files_by_drive: Dict[str, List[Dict[str, Any]]] = {}
        self.document_chunks: List[Dict[str, Any]] = []
        
        # Search Index
        self.index_manager = SharePointIndexManager()
        
        self.msal_app = None
        self._access_token = None
        self._token_expires_at = 0
        self.is_synced = False
        self.is_syncing = False
        self.last_sync_time = None
        self.sync_status_message = "Not synchronized yet"

    def _get_app(self):
        if not self.msal_app:
            self.msal_app = msal.ConfidentialClientApplication(
                self.client_id,
                authority=AUTHORITY_URL,
                client_credential=self.client_secret
            )
        return self.msal_app

    def get_access_token(self) -> str:
        """Acquire or return valid Graph API access token."""
        now = time.time()
        if self._access_token and now < self._token_expires_at - 120:
            return self._access_token

        app = self._get_app()
        result = app.acquire_token_for_client(scopes=GRAPH_SCOPES)
        if "access_token" in result:
            self._access_token = result["access_token"]
            self._token_expires_at = now + result.get("expires_in", 3599)
            return self._access_token
        else:
            err = result.get("error_description", result.get("error", "Unknown MSAL error"))
            raise RuntimeError(f"Failed to acquire Microsoft Graph token: {err}")

    def _request(self, endpoint: str, params: Optional[Dict[str, Any]] = None, timeout: int = 45) -> requests.Response:
        """Helper to make authenticated Graph requests with auto-retry on 401."""
        token = self.get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        }
        url = endpoint if endpoint.startswith("http") else f"{GRAPH_API_BASE}/{endpoint.lstrip('/')}"
        
        resp = requests.get(url, headers=headers, params=params, timeout=timeout)
        if resp.status_code == 401:
            # Token expired or invalid, reset and retry once
            self._access_token = None
            token = self.get_access_token()
            headers["Authorization"] = f"Bearer {token}"
            resp = requests.get(url, headers=headers, params=params, timeout=timeout)
        return resp

    def resolve_site(self) -> Dict[str, Any]:
        """Resolve SharePoint site id from hostname and path."""
        if self.site_id and self.site_info:
            return self.site_info

        url = f"sites/{self.hostname}:{self.site_path}"
        resp = self._request(url)
        if resp.status_code != 200:
            raise RuntimeError(f"Error resolving SharePoint site '{url}': {resp.status_code} {resp.text}")

        self.site_info = resp.json()
        self.site_id = self.site_info["id"]
        logger.info(f"Connected to SharePoint site: {self.site_info.get('displayName')} ({self.site_id})")
        return self.site_info

    def fetch_all_lists(self) -> List[Dict[str, Any]]:
        """Fetch all SharePoint lists with their column definitions."""
        site_id = self.site_id or self.resolve_site()["id"]
        resp = self._request(f"sites/{site_id}/lists?$expand=columns")
        if resp.status_code != 200:
            raise RuntimeError(f"Error fetching lists: {resp.status_code} {resp.text}")

        raw_lists = resp.json().get("value", [])
        clean_lists = []
        for l in raw_lists:
            # We filter and format
            name = l.get("displayName", "")
            list_id = l.get("id", "")
            template = l.get("list", {}).get("template", "unknown")
            is_hidden = l.get("list", {}).get("hidden", False)
            web_url = l.get("webUrl", "")
            
            # Extract column definitions
            cols = []
            for c in l.get("columns", []):
                col_name = c.get("name", "")
                display_col = c.get("displayName", col_name)
                # Skip internal noisy columns
                if col_name.startswith("_") or col_name.startswith("@") or c.get("readOnly"):
                    if col_name not in ["Title", "id", "Created", "Modified"]:
                        continue
                cols.append({
                    "name": col_name,
                    "displayName": display_col,
                    "type": list(c.keys())[-1] if c else "text"
                })

            clean_lists.append({
                "id": list_id,
                "name": name,
                "template": template,
                "is_hidden": is_hidden,
                "webUrl": web_url,
                "columns": cols
            })

        self.lists_metadata = clean_lists
        return clean_lists

    def fetch_list_items(self, list_id: str, list_name: str = "") -> List[Dict[str, Any]]:
        """Fetch all items from a given list, handling Graph pagination."""
        site_id = self.site_id or self.resolve_site()["id"]
        items = []
        next_url = f"{GRAPH_API_BASE}/sites/{site_id}/lists/{list_id}/items?$expand=fields"

        while next_url:
            resp = self._request(next_url)
            if resp.status_code != 200:
                logger.error(f"Failed to fetch items for list {list_name} ({list_id}): {resp.status_code} {resp.text}")
                break

            data = resp.json()
            raw_items = data.get("value", [])
            for item in raw_items:
                raw_fields = item.get("fields", {})
                # Filter internal SharePoint noise
                clean_fields = {
                    k: v for k, v in raw_fields.items()
                    if not k.startswith("@") and not k.startswith("_")
                }
                
                # Make a human-readable searchable representation
                searchable_parts = [f"{k}: {v}" for k, v in clean_fields.items() if v is not None and str(v).strip()]
                searchable_text = f"List: {list_name}\n" + "\n".join(searchable_parts)

                items.append({
                    "id": item.get("id"),
                    "list_id": list_id,
                    "list_name": list_name,
                    "webUrl": item.get("webUrl", ""),
                    "createdDateTime": item.get("createdDateTime", ""),
                    "lastModifiedDateTime": item.get("lastModifiedDateTime", ""),
                    "fields": clean_fields,
                    "searchable_text": searchable_text
                })

            next_url = data.get("@odata.nextLink")

        self.list_items_by_id[list_id] = items
        logger.info(f"Loaded {len(items)} items for list '{list_name}'")
        return items

    def fetch_all_drives(self) -> List[Dict[str, Any]]:
        """Fetch all Document Libraries (Drives) for the site."""
        site_id = self.site_id or self.resolve_site()["id"]
        resp = self._request(f"sites/{site_id}/drives")
        if resp.status_code != 200:
            raise RuntimeError(f"Error fetching drives: {resp.status_code} {resp.text}")

        raw_drives = resp.json().get("value", [])
        clean_drives = []
        for d in raw_drives:
            clean_drives.append({
                "id": d.get("id"),
                "name": d.get("name"),
                "description": d.get("description", ""),
                "driveType": d.get("driveType"),
                "webUrl": d.get("webUrl", "")
            })

        self.drives_metadata = clean_drives
        return clean_drives

    def _traverse_folder(self, drive_id: str, drive_name: str, item_id: str = "root", current_path: str = "") -> List[Dict[str, Any]]:
        """Recursively fetch all files in a drive."""
        files = []
        endpoint = f"drives/{drive_id}/items/{item_id}/children" if item_id != "root" else f"drives/{drive_id}/root/children"
        resp = self._request(endpoint)
        if resp.status_code != 200:
            logger.error(f"Error traversing folder {item_id} in {drive_name}: {resp.status_code}")
            return files

        items = resp.json().get("value", [])
        for item in items:
            name = item.get("name", "")
            item_path = f"{current_path}/{name}" if current_path else name
            
            if "folder" in item:
                # Recurse into subfolder
                sub_files = self._traverse_folder(drive_id, drive_name, item.get("id"), item_path)
                files.extend(sub_files)
            elif "file" in item:
                files.append({
                    "id": item.get("id"),
                    "name": name,
                    "path": item_path,
                    "size": item.get("size", 0),
                    "drive_id": drive_id,
                    "drive_name": drive_name,
                    "webUrl": item.get("webUrl", ""),
                    "downloadUrl": item.get("@microsoft.graph.downloadUrl"),
                    "mimeType": item.get("file", {}).get("mimeType", "")
                })
        return files

    def fetch_drive_files(self, drive_id: str, drive_name: str) -> List[Dict[str, Any]]:
        """Fetch all files within a document library."""
        files = self._traverse_folder(drive_id, drive_name)
        self.files_by_drive[drive_id] = files
        logger.info(f"Found {len(files)} files in Document Library '{drive_name}'")
        return files

    def _download_file_bytes(self, drive_id: str, file_id: str, download_url: Optional[str] = None) -> Optional[bytes]:
        """Download file content using direct Graph API /content endpoint with fallback."""
        try:
            token = self.get_access_token()
            headers = {"Authorization": f"Bearer {token}"}
            url = f"{GRAPH_API_BASE}/drives/{drive_id}/items/{file_id}/content"
            resp = requests.get(url, headers=headers, timeout=25, allow_redirects=True)
            if resp.status_code == 200 and len(resp.content) > 0:
                return resp.content
        except Exception as e:
            logger.warning(f"Direct /content download attempt error for {file_id}: {e}")

        if download_url:
            try:
                resp = requests.get(download_url, timeout=25)
                if resp.status_code == 200 and len(resp.content) > 0:
                    return resp.content
            except Exception as e:
                logger.warning(f"downloadUrl attempt error for {file_id}: {e}")

        return None

    def parse_and_index_files(self, force_refresh: bool = False) -> List[Dict[str, Any]]:
        """Download document files from SharePoint, parse text chunks, and prepare for RAG.
        Uses smart incremental caching so unchanged files are reused instantly."""
        existing_by_file_id: Dict[str, List[Dict[str, Any]]] = {}
        if not force_refresh:
            for c in self.document_chunks:
                f_id = c.get("file_id")
                if f_id:
                    if f_id not in existing_by_file_id:
                        existing_by_file_id[f_id] = []
                    existing_by_file_id[f_id].append(c)

        all_chunks = []
        for drive in self.drives_metadata:
            drive_id = drive["id"]
            drive_name = drive["name"]
            files = self.files_by_drive.get(drive_id, [])

            for f in files:
                file_id = f["id"]
                file_name = f["name"]

                # Re-use existing cached chunks if already present
                if not force_refresh and file_id in existing_by_file_id and len(existing_by_file_id[file_id]) > 0:
                    logger.info(f"Reusing {len(existing_by_file_id[file_id])} cached chunks for '{file_name}'")
                    all_chunks.extend(existing_by_file_id[file_id])
                    continue

                logger.info(f"Downloading & indexing document: {drive_name} -> {file_name}")
                content = self._download_file_bytes(drive_id, file_id, f.get("downloadUrl"))
                if not content:
                    logger.warning(f"Could not download content for '{file_name}'")
                    if file_id in existing_by_file_id:
                        all_chunks.extend(existing_by_file_id[file_id])
                    continue

                try:
                    chunks = parse_file(content, file_name)
                    for c in chunks:
                        c["library_name"] = drive_name
                        c["drive_id"] = drive_id
                        c["file_id"] = file_id
                        c["file_path"] = f["path"]
                        c["webUrl"] = f.get("webUrl", "")
                        all_chunks.append(c)
                    logger.info(f"Extracted {len(chunks)} chunks from '{file_name}'")
                except Exception as e:
                    logger.error(f"Error parsing file {file_name}: {e}")

        self.document_chunks = all_chunks
        return all_chunks

    def upload_document(self, drive_id: str, file_name: str, file_bytes: bytes) -> Dict[str, Any]:
        """Upload a file to a SharePoint Document Library and index it immediately."""
        token = self.get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream"
        }
        url = f"{GRAPH_API_BASE}/drives/{drive_id}/root:/{file_name}:/content"
        resp = requests.put(url, headers=headers, data=file_bytes, timeout=60)
        if resp.status_code not in [200, 201]:
            raise RuntimeError(f"Failed to upload {file_name} to SharePoint: {resp.status_code} {resp.text}")

        item_data = resp.json()
        logger.info(f"Uploaded {file_name} to SharePoint: id={item_data.get('id')}")

        # Find drive name
        drive_name = next((d["name"] for d in self.drives_metadata if d["id"] == drive_id), "Knowledge Base")
        
        # Add to files_by_drive
        if drive_id not in self.files_by_drive:
            self.files_by_drive[drive_id] = []
        self.files_by_drive[drive_id].append({
            "id": item_data.get("id"),
            "name": file_name,
            "path": file_name,
            "size": len(file_bytes),
            "drive_id": drive_id,
            "drive_name": drive_name,
            "webUrl": item_data.get("webUrl", ""),
            "downloadUrl": item_data.get("@microsoft.graph.downloadUrl"),
            "mimeType": item_data.get("file", {}).get("mimeType", "")
        })

        # Parse and index chunks immediately
        new_chunks = parse_file(file_bytes, file_name)
        for c in new_chunks:
            c["library_name"] = drive_name
            c["drive_id"] = drive_id
            c["file_id"] = item_data.get("id")
            c["file_path"] = file_name
            c["webUrl"] = item_data.get("webUrl", "")
            self.document_chunks.append(c)

        # Update search index & persist cache
        self.index_manager.set_documents(self.document_chunks)
        self._save_cache()

        return {
            "success": True,
            "file_name": file_name,
            "drive_id": drive_id,
            "drive_name": drive_name,
            "webUrl": item_data.get("webUrl"),
            "chunks_added": len(new_chunks)
        }

    def delete_document(self, drive_id: str, item_id: str, file_name: str = "") -> Dict[str, Any]:
        """Delete a document from SharePoint drive and purge its chunks from RAG index."""
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        delete_url = f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"

        # 1. Attempt Graph API delete
        graph_deleted = False
        try:
            resp = requests.delete(delete_url, headers=headers, timeout=20)
            if resp.status_code in (200, 204):
                graph_deleted = True
                logger.info(f"Deleted item {item_id} from SharePoint drive {drive_id}")
            else:
                logger.warning(f"Graph API delete returned {resp.status_code}: {resp.text}")
        except Exception as e:
            logger.warning(f"Error calling Graph API delete: {e}")

        # 2. Remove file from memory
        removed_file_name = file_name
        if drive_id in self.files_by_drive:
            new_files = []
            for f in self.files_by_drive[drive_id]:
                if f.get("id") == item_id or (file_name and f.get("name") == file_name):
                    removed_file_name = f.get("name", file_name)
                else:
                    new_files.append(f)
            self.files_by_drive[drive_id] = new_files

        # 3. Remove chunks from memory
        before_chunks = len(self.document_chunks)
        self.document_chunks = [
            c for c in self.document_chunks
            if c.get("file_id") != item_id and (not removed_file_name or c.get("file_name") != removed_file_name)
        ]
        chunks_removed = before_chunks - len(self.document_chunks)

        # 4. Rebuild index & persist cache
        self.index_manager.set_documents(self.document_chunks)
        self._save_cache()

        logger.info(f"Successfully purged {chunks_removed} chunks for file '{removed_file_name}'")
        return {
            "success": True,
            "graph_deleted": graph_deleted,
            "file_name": removed_file_name,
            "chunks_removed": chunks_removed,
            "remaining_chunks": len(self.document_chunks)
        }

    def sync_all(self, force_refresh: bool = False) -> Dict[str, Any]:
        """Perform a full synchronization of SharePoint lists, libraries, and files."""
        if self.is_syncing:
            return {
                "success": True,
                "status": "syncing",
                "message": "Synchronization is already running in the background."
            }

        self.is_syncing = True
        start_time = time.time()
        self.sync_status_message = "Synchronizing SharePoint data..."
        logger.info("Starting SharePoint full sync...")
        
        try:
            # 1. Resolve Site
            self.resolve_site()
            
            # 2. Lists & Items
            self.fetch_all_lists()
            all_list_records = []
            for l in self.lists_metadata:
                items = self.fetch_list_items(l["id"], l["name"])
                all_list_records.extend(items)

            # 3. Document Libraries (Drives) & Files
            self.fetch_all_drives()
            for d in self.drives_metadata:
                self.fetch_drive_files(d["id"], d["name"])

            # 4. Download & Parse Documents incrementally
            chunks = self.parse_and_index_files(force_refresh=force_refresh)

            # 5. Build Indices
            self.index_manager.set_lists(all_list_records)
            self.index_manager.set_documents(chunks)

            self.is_synced = True
            self.last_sync_time = time.strftime("%Y-%m-%d %H:%M:%S")
            duration = round(time.time() - start_time, 2)
            self.sync_status_message = f"Synchronized in {duration}s ({len(all_list_records)} list items, {len(chunks)} document chunks)"

            # Save snapshot to cache
            self._save_cache()

            return {
                "success": True,
                "duration_seconds": duration,
                "site_name": self.site_info.get("displayName"),
                "total_lists": len(self.lists_metadata),
                "total_list_records": len(all_list_records),
                "total_libraries": len(self.drives_metadata),
                "total_document_chunks": len(chunks)
            }
        except Exception as e:
            self.sync_status_message = f"Sync failed: {str(e)}"
            logger.exception("SharePoint Sync failed")
            return {
                "success": False,
                "error": str(e)
            }
        finally:
            self.is_syncing = False

    def _save_cache(self):
        """Save indexed metadata to local disk cache for fast cold-start."""
        try:
            cache_file = CACHE_DIR / "sharepoint_cache.json"
            data = {
                "site_info": self.site_info,
                "site_id": self.site_id,
                "lists_metadata": self.lists_metadata,
                "list_items_by_id": self.list_items_by_id,
                "drives_metadata": self.drives_metadata,
                "files_by_drive": self.files_by_drive,
                "document_chunks": self.document_chunks,
                "last_sync_time": self.last_sync_time
            }
            with open(cache_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
            logger.info("Saved SharePoint snapshot to cache file")
        except Exception as e:
            logger.warning(f"Could not save cache: {e}")

    def load_cache(self) -> bool:
        """Load from local disk cache if available."""
        cache_file = CACHE_DIR / "sharepoint_cache.json"
        if not cache_file.exists():
            return False
        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.site_info = data.get("site_info", {})
            self.site_id = data.get("site_id")
            self.lists_metadata = data.get("lists_metadata", [])
            self.list_items_by_id = data.get("list_items_by_id", {})
            self.drives_metadata = data.get("drives_metadata", [])
            self.files_by_drive = data.get("files_by_drive", {})
            self.document_chunks = data.get("document_chunks", [])
            self.last_sync_time = data.get("last_sync_time")

            all_list_records = []
            for items in self.list_items_by_id.values():
                all_list_records.extend(items)

            self.index_manager.set_lists(all_list_records)
            self.index_manager.set_documents(self.document_chunks)
            self.is_synced = True
            self.sync_status_message = f"Loaded from cache ({self.last_sync_time})"
            logger.info(f"Loaded SharePoint data from cache successfully ({len(all_list_records)} items, {len(self.document_chunks)} chunks)")
            return True
        except Exception as e:
            logger.warning(f"Failed to load cache: {e}")
            return False

    def get_structure_summary(self) -> Dict[str, Any]:
        """Return high-level summary of all lists and document libraries."""
        lists_info = []
        for l in self.lists_metadata:
            list_id = l["id"]
            items = self.list_items_by_id.get(list_id, [])
            lists_info.append({
                "id": list_id,
                "name": l["name"],
                "template": l["template"],
                "total_items": len(items),
                "columns": [c["displayName"] for c in l.get("columns", [])[:8]],
                "webUrl": l.get("webUrl")
            })

        drives_info = []
        for d in self.drives_metadata:
            drive_id = d["id"]
            files = self.files_by_drive.get(drive_id, [])
            drives_info.append({
                "id": drive_id,
                "name": d["name"],
                "total_files": len(files),
                "files": [f["name"] for f in files],
                "webUrl": d.get("webUrl")
            })

        return {
            "site_name": self.site_info.get("displayName", "SharePoint Site"),
            "site_url": f"https://{self.hostname}{self.site_path}",
            "last_sync_time": self.last_sync_time,
            "lists": lists_info,
            "document_libraries": drives_info
        }

    # Query helpers for AI Agent
    def query_list_items(self, list_name_or_keyword: str, query: str = "", top_k: int = 15, column_filters: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
        """Query items within lists by name and optional column filters."""
        # Find matching list
        target_list_id = None
        target_list_name = None
        for l in self.lists_metadata:
            if list_name_or_keyword.lower() in l["name"].lower():
                target_list_id = l["id"]
                target_list_name = l["name"]
                break

        if target_list_id:
            items = self.list_items_by_id.get(target_list_id, [])
        else:
            # Search across all lists
            items = []
            for l_items in self.list_items_by_id.values():
                items.extend(l_items)

        # Apply column filters if provided
        filtered = []
        for it in items:
            fields = it.get("fields", {})
            match = True
            if column_filters:
                for col, val in column_filters.items():
                    actual = str(fields.get(col, "")).lower()
                    if str(val).lower() not in actual:
                        match = False
                        break
            if match:
                filtered.append(it)

        if not query.strip():
            return filtered[:top_k]

        # 1. Exact substring matches across fields
        q_lower = query.strip().lower()
        exact_matches = []
        for it in filtered:
            searchable = it.get("searchable_text", "").lower()
            if q_lower in searchable:
                it_copy = dict(it)
                it_copy["relevance_score"] = 10.0
                exact_matches.append(it_copy)

        # 2. Tokenized BM25 search
        scored = self.index_manager.search_lists(query, list_name=target_list_name, top_k=top_k)

        # Combine and deduplicate
        seen_ids = set()
        combined = []
        for it in exact_matches:
            it_id = it.get("id")
            if it_id not in seen_ids:
                seen_ids.add(it_id)
                combined.append(it)

        for it in scored:
            it_id = it.get("id")
            if it_id not in seen_ids:
                seen_ids.add(it_id)
                combined.append(it)

        return combined[:top_k] if combined else filtered[:5]

    def search_library_documents(self, query: str, library_name: Optional[str] = None, top_k: int = 5) -> List[Dict[str, Any]]:
        """Search across Document Libraries for relevant text passages."""
        return self.index_manager.search_documents(query, library_name=library_name, top_k=top_k)
