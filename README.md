# SharePoint AI Assistant Bot

An intelligent enterprise AI Assistant connecting **Azure OpenAI (gpt-4o-mini)** with **SharePoint Online (`AISearchDev`)** via **Microsoft Graph API**.

---

## 🌟 Key Features

1. **SharePoint List Intelligence**:
   - Queries and filters across all SharePoint generic and custom lists (e.g., `Employee Details`, `PatientIntake`, `Employee Asset Subtable`, `Employee Official Letter Subtable`, `News Sources`, `Events`).
   - Retrieves records, aggregates statistics (counts, salaries, designations, managers, locations), and renders structured markdown tables.

2. **SharePoint Document Library RAG (Retrieval-Augmented Generation)**:
   - Deep search across Document Libraries (`HR Knowledge Base`, `IT Knowledge Base`, `Documents`).
   - Parses `.pdf`, `.docx`, `.xlsx`, `.txt`, `.csv` documents.
   - Extracts exact policy sections, troubleshooting instructions, and provides citations with document names and page numbers.

3. **Intelligent Query Routing with Function Calling**:
   - Uses OpenAI tool calling to dynamically choose whether to query lists, search document libraries, or synthesize data from both.

4. **Modern Responsive Web UI + Real-time Streaming**:
   - Tailwind CSS dark-mode interface with streaming token response via Server-Sent Events (SSE).
   - Interactive **Lists Explorer** to view tabular data directly.
   - Interactive **Document Libraries Explorer** with file cards.
   - Source citation badges linking to SharePoint references.
   - Live synchronization status with 1-click re-indexing button.

5. **Terminal CLI Mode**:
   - Includes an interactive terminal CLI (`cli.py`) for quick command-line queries.

---

## 🚀 Quick Start

### 1. Requirements
Make sure the required packages are installed:
```bash
pip install -r requirements.txt
```

### 2. Launch Web Application
Double-click `run.bat` or run:
```bash
python app.py
```
Open your browser and navigate to:
```
http://localhost:5005
```

### 3. Launch Interactive Terminal CLI
Double-click `run_cli.bat` or run:
```bash
python cli.py
```

---

## ⚙️ Configuration (.env)

The `.env` file contains your Microsoft Entra ID and Azure OpenAI configuration:

| Variable | Description |
|---|---|
| `AZURE_TENANT_ID` | Microsoft Entra (Azure AD) Directory ID |
| `AZURE_CLIENT_ID` | App Registration Client (Application) ID |
| `AZURE_CLIENT_SECRET` | App Registration Client Secret |
| `SHAREPOINT_HOSTNAME` | Hostname (e.g. `veeleaddev.sharepoint.com`) |
| `SHAREPOINT_SITE_PATH` | Site path (e.g. `/sites/AISearchDev`) |
| `AZURE_OPENAI_ENDPOINT`| Azure OpenAI Endpoint URL |
| `AZURE_OPENAI_API_KEY` | Azure OpenAI API Key |
| `AZURE_OPENAI_DEPLOYMENT` | Model deployment name (`gpt-4o-mini`) |
| `AZURE_OPENAI_API_VERSION` | API version (`2024-08-01-preview`) |
| `PORT` | Local web server port (default: `5000`) |

---

## 📁 SharePoint Data Sources Available

- **Lists**:
  - `Employee Details` (100 employee records with names, designations, salaries, managers, emails, locations)
  - `PatientIntake` (patient visits, symptoms, doctors, token numbers)
  - `Employee Asset Subtable` (hardware assets, laptops, monitors)
  - `Employee Official Letter Subtable` (offer letters, experience)
  - `News Sources`
- **Document Libraries**:
  - `HR Knowledge Base`:
    - `VLS-Onboarding-Offboarding-Guide.pdf` (10 pages)
    - `VLS-Performance-Appraisal-Policy.docx`
  - `IT Knowledge Base`:
    - `VLS-IT-020_Illustrated_Troubleshooting_Guide_v1.0.pdf` (15 pages)
    - `VLS-IT-Extended-Guides (1).pdf` (17 pages)
