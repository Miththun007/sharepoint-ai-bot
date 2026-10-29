import sys
import logging

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
from config import SHAREPOINT_SITE_URL, AZURE_OPENAI_DEPLOYMENT
from sharepoint_service import SharePointService
from bot_engine import SharePointBotEngine

logging.basicConfig(level=logging.WARNING)

def main():
    print("=" * 70)
    print("  SharePoint AI Bot - Interactive CLI")
    print(f"  Target: {SHAREPOINT_SITE_URL}")
    print(f"  Model : {AZURE_OPENAI_DEPLOYMENT}")
    print("=" * 70)
    print("\n[1/2] Connecting to SharePoint & initializing knowledge bases...")

    sp_service = SharePointService()
    
    # Try cache first or sync
    if not sp_service.load_cache():
        print("  Synchronizing SharePoint Lists & Document Libraries (downloading & indexing)...")
        res = sp_service.sync_all()
        if not res.get("success"):
            print(f"[!] Sync failed: {res.get('error')}")
            sys.exit(1)
        print(f"  [DONE] Synced {res.get('total_list_records')} list items & {res.get('total_document_chunks')} document chunks.")
    else:
        print(f"  [DONE] Loaded cache with {len(sp_service.lists_metadata)} lists and {len(sp_service.drives_metadata)} libraries.")

    bot = SharePointBotEngine(sp_service)
    print("\n[2/2] SharePoint AI Bot Ready! Type your question below (or 'exit' to quit).")
    print("-" * 70)
    print("Examples:")
    print("  • Who is the Team Lead for employee Surya Kumar in Employee Details?")
    print("  • Summarize the performance appraisal rating system from the HR guide.")
    print("  • How do I resolve VPN and network connection errors in the IT guide?")
    print("  • What lists and document libraries exist in this SharePoint site?")
    print("-" * 70)

    history = []
    while True:
        try:
            prompt = input("\nYou: ").strip()
            if not prompt:
                continue
            if prompt.lower() in ["exit", "quit", "q"]:
                print("Goodbye!")
                break
            if prompt.lower() == "sync":
                print("Re-syncing SharePoint...")
                res = sp_service.sync_all(force_refresh=True)
                print(f"Sync complete: {res}")
                continue

            history.append({"role": "user", "content": prompt})
            print("\nBot: [Searching SharePoint data...]", end="\r")

            result = bot.chat(history)
            answer = result.get("response", "")
            citations = result.get("citations", [])

            print(" " * 50, end="\r")  # clear line
            print(f"\nBot:\n{answer}\n")

            if citations:
                print("Citations:")
                for c in citations[:4]:
                    print(f"  * {c.get('source')}")

            history.append({"role": "assistant", "content": answer})

        except KeyboardInterrupt:
            print("\nExiting...")
            break
        except Exception as e:
            print(f"\n[Error] {e}")

if __name__ == "__main__":
    main()
