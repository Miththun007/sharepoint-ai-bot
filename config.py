import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env file from the current directory or parent directory
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

# Microsoft Entra ID (Azure AD) Credentials
AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID", "").strip()
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "").strip()
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "").strip()

# SharePoint Configuration
SHAREPOINT_HOSTNAME = os.getenv("SHAREPOINT_HOSTNAME", "veeleaddev.sharepoint.com").strip()
SHAREPOINT_SITE_PATH = os.getenv("SHAREPOINT_SITE_PATH", "/sites/AISearchDev").strip()
SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL", f"https://{SHAREPOINT_HOSTNAME}{SHAREPOINT_SITE_PATH}").strip()

# Microsoft Graph Endpoints
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
AUTHORITY_URL = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]

# Azure OpenAI Configuration
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT", "").rstrip("/")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY", "").strip()
AZURE_OPENAI_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o-mini").strip()
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION", "2024-08-01-preview").strip()

# App Settings
PORT = int(os.getenv("PORT", 5005))
CACHE_DIR = BASE_DIR / "cache"
CACHE_DIR.mkdir(exist_ok=True)
