import sys
import os
from pathlib import Path

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# Add current directory to path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from sharepoint_service import SharePointService
from bot_engine import SharePointBotEngine

def main():
    print("Testing SharePointService sync...")
    sp = SharePointService()
    sync_res = sp.sync_all()
    print("Sync Result:", sync_res)

    if not sync_res.get("success"):
        print("Sync failed!")
        return

    print(f"\nIndexed {sync_res['total_lists']} lists with {sync_res['total_list_records']} records.")
    print(f"Indexed {sync_res['total_libraries']} document libraries with {sync_res['total_document_chunks']} chunks.")

    bot = SharePointBotEngine(sp)

    # Test 1: Query SharePoint List
    print("\n--- Test 1: Querying SharePoint List (Employee Details) ---")
    q1 = "Who is Surya Kumar and what is his department, designation, and salary in Employee Details?"
    r1 = bot.chat([{"role": "user", "content": q1}])
    print("User:", q1)
    print("Bot Response:\n", r1["response"])
    print("Citations:", r1["citations"])

    # Test 2: Query SharePoint Document Library
    print("\n--- Test 2: Querying Document Library (HR/IT Knowledge Base) ---")
    q2 = "What are the key steps or goals mentioned in the Performance Appraisal Policy?"
    r2 = bot.chat([{"role": "user", "content": q2}])
    print("User:", q2)
    print("Bot Response:\n", r2["response"])
    print("Citations:", r2["citations"])

    print("\nAll verification tests completed successfully!")

if __name__ == "__main__":
    main()
