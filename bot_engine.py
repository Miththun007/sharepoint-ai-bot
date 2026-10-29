import json
import logging
from typing import List, Dict, Any, Generator
import requests

from config import (
    AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_API_KEY,
    AZURE_OPENAI_DEPLOYMENT, AZURE_OPENAI_API_VERSION,
    SHAREPOINT_SITE_URL, SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH
)
from sharepoint_service import SharePointService

logger = logging.getLogger("bot_engine")

SYSTEM_PROMPT = f"""You are the official Veelead SharePoint Enterprise AI Policy & Knowledge Base Assistant.
You have real-time access to the Veelead SharePoint site:
Site URL: {SHAREPOINT_SITE_URL} (Hostname: {SHAREPOINT_HOSTNAME}, Path: {SHAREPOINT_SITE_PATH})

Your core purpose is to read, search, analyze, and answer questions from:
1. **Corporate Policy Documents & Knowledge Bases** (Document Libraries: "HR Knowledge Base", "IT Knowledge Base", "Documents"):
   - **HR Policies**: HR Policy Manual, Onboarding & Offboarding Guide, Performance Appraisal & Management Policy, Leave Entitlements, Code of Conduct, Grievance redressal, Probation rules.
   - **Financial Policies**: Financial & Budget Guidelines, Travel & Expense Reimbursement limits, Procurement approval tiers, Petty cash policies.
   - **Customer Support & SLAs**: Customer Support SLA Handbook, Severity definitions (P1-P4), Response/Resolution timeframes, Escalation matrices.
   - **IT & Technical Standards**: Software Engineering Standards (Python, TypeScript, Git), IT Architecture, Illustrated Troubleshooting Guide, Master Extended Support Manual.
2. **SharePoint Data Lists** (e.g. "Employee Details", "PatientIntake", "Employee Asset Subtable", "Employee Official Letter Subtable", "News Sources", "Events").

### Operating Principles for RAG & Policy Inquiries:
1. **Policy RAG Deep-Dive**:
   - Use `search_sharepoint_documents` to find relevant clauses, guidelines, and manuals.
   - Always cite the **Document Name**, **Section/Clause**, and **Page Number**.
   - When explaining policies, break down:
     * **Eligibility & Conditions**: Who does the policy apply to?
     * **Entitlements / Allowances / Limits**: Specific amounts, days, thresholds.
     * **Approval Workflow**: Who needs to approve (e.g. Manager, Department Head, Finance)?
     * **Timelines & SLAs**: Response times, deadlines, notice periods.
   - Use clean Markdown tables for quantitative rules (e.g. SLA matrices, travel allowances, leave types).

2. **SharePoint Lists Queries**:
   - Use `query_sharepoint_list` to find rows, employees, patients, assets, letters, or news.
   - Cross-reference lists with policies if relevant (e.g., cross-referencing an employee's designation or department with their applicable leave or travel approval tier).

3. **Tone & Formatting**:
   - Authoritative, precise, professional, and well-structured with Markdown headings, tables, and bullet points.
"""

TOOL_DEFINITIONS = [
    {
        "type": "function",
        "function": {
            "name": "query_sharepoint_list",
            "description": "Query or search records from a SharePoint list (e.g., 'Employee Details', 'PatientIntake', 'Employee Asset Subtable', 'Employee Official Letter Subtable', 'News Sources', 'Events'). Returns matching row fields.",
            "parameters": {
                "type": "object",
                "properties": {
                    "list_name": {
                        "type": "string",
                        "description": "The name of the SharePoint list to query (e.g. 'Employee Details', 'PatientIntake', 'Employee Asset Subtable', etc.)"
                    },
                    "query": {
                        "type": "string",
                        "description": "Optional search term to filter records (e.g., an employee name, designation, department, ID, patient symptom, doctor, email, city, etc.)"
                    },
                    "max_results": {
                        "type": "integer",
                        "description": "Maximum number of rows to return (default: 10, max: 50)"
                    }
                },
                "required": ["list_name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_sharepoint_documents",
            "description": "Search across SharePoint Document Libraries (e.g. 'HR Knowledge Base', 'IT Knowledge Base') for policies, guides, manuals, SOPs, and technical documentation.",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "The question or search query to find in document libraries (e.g. 'onboarding process steps', 'performance appraisal rating cycle', 'VPN connection failed error', 'password reset guide')"
                    },
                    "library_name": {
                        "type": "string",
                        "description": "Optional specific library name (e.g. 'HR Knowledge Base' or 'IT Knowledge Base') or document name (e.g. '01_HR_Policy_Manual.pdf', '06_Customer_Support_SLA_Handbook.pdf', '04_Financial_Policy_Guidelines.pdf', '02_Software_Engineering_Standards.pdf'). Leave empty to search all libraries."
                    },
                    "max_results": {
                        "type": "integer",
                        "description": "Maximum relevant passages to retrieve (default: 5)"
                    }
                },
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_sharepoint_overview",
            "description": "Get an overview of all available SharePoint Lists, Document Libraries, schemas, and files present on the SharePoint site.",
            "parameters": {
                "type": "object",
                "properties": {}
            }
        }
    }
]

class SharePointBotEngine:
    def __init__(self, sp_service: SharePointService):
        self.sp_service = sp_service
        self.endpoint = AZURE_OPENAI_ENDPOINT
        self.api_key = AZURE_OPENAI_API_KEY
        self.deployment = AZURE_OPENAI_DEPLOYMENT
        self.api_version = AZURE_OPENAI_API_VERSION
        self.completions_url = f"{self.endpoint}/openai/deployments/{self.deployment}/chat/completions?api-version={self.api_version}"

    def _call_openai(self, messages: List[Dict[str, Any]], tools=None, tool_choice="auto", stream: bool = False):
        headers = {
            "api-key": self.api_key,
            "Content-Type": "application/json"
        }
        payload = {
            "messages": messages,
            "temperature": 0.3,
            "max_tokens": 2500
        }
        if tools:
            payload["tools"] = tools
            payload["tool_choice"] = tool_choice
        if stream:
            payload["stream"] = True

        return requests.post(self.completions_url, headers=headers, json=payload, stream=stream, timeout=60)

    def execute_tool(self, name: str, args: Dict[str, Any]) -> Dict[str, Any]:
        """Execute a tool call against the SharePoint service and format the response."""
        logger.info(f"Executing tool: {name} with args: {args}")
        try:
            if name == "query_sharepoint_list":
                list_name = args.get("list_name", "")
                query = args.get("query", "")
                max_results = min(args.get("max_results", 10), 50)
                
                results = self.sp_service.query_list_items(list_name, query=query, top_k=max_results)
                
                # Format compact records for prompt efficiency
                formatted = []
                for r in results:
                    formatted.append({
                        "id": r.get("id"),
                        "list": r.get("list_name"),
                        "fields": r.get("fields", {}),
                        "relevance_score": r.get("relevance_score")
                    })
                    
                return {
                    "tool": name,
                    "list_queried": list_name,
                    "count": len(formatted),
                    "results": formatted
                }

            elif name == "search_sharepoint_documents":
                query = args.get("query", "")
                library_name = args.get("library_name")
                max_results = min(args.get("max_results", 5), 10)
                
                chunks = self.sp_service.search_library_documents(query, library_name=library_name, top_k=max_results)
                formatted = []
                for c in chunks:
                    formatted.append({
                        "library": c.get("library_name"),
                        "document": c.get("file_name"),
                        "page_or_section": c.get("page"),
                        "content_excerpt": c.get("text"),
                        "relevance_score": c.get("relevance_score"),
                        "webUrl": c.get("webUrl")
                    })
                    
                return {
                    "tool": name,
                    "query": query,
                    "library": library_name or "All Libraries",
                    "count": len(formatted),
                    "passages": formatted
                }

            elif name == "get_sharepoint_overview":
                overview = self.sp_service.get_structure_summary()
                return {
                    "tool": name,
                    "overview": overview
                }
            else:
                return {"error": f"Unknown tool: {name}"}

        except Exception as e:
            logger.exception(f"Error executing {name}")
            return {"error": str(e)}

    def chat(self, user_messages: List[Dict[str, str]]) -> Dict[str, Any]:
        """Non-streaming multi-turn chat handling tool execution loop."""
        messages = [{"role": "system", "content": SYSTEM_PROMPT}]
        messages.extend(user_messages)

        tool_activity = []
        citations = []

        # Tool calling loop (up to 4 steps)
        for _ in range(4):
            resp = self._call_openai(messages, tools=TOOL_DEFINITIONS)
            if resp.status_code != 200:
                raise RuntimeError(f"Azure OpenAI Error ({resp.status_code}): {resp.text}")

            res_json = resp.json()
            choice = res_json["choices"][0]
            msg = choice["message"]
            messages.append(msg)

            # Check if model wants to call tools
            tool_calls = msg.get("tool_calls")
            if not tool_calls:
                # No more tools, final response reached
                return {
                    "response": msg.get("content", ""),
                    "tool_activity": tool_activity,
                    "citations": citations
                }

            # Execute tool calls
            for tc in tool_calls:
                fn_name = tc["function"]["name"]
                raw_args = tc["function"]["arguments"]
                try:
                    fn_args = json.loads(raw_args)
                except Exception:
                    fn_args = {}

                tool_result = self.execute_tool(fn_name, fn_args)
                tool_activity.append({
                    "tool": fn_name,
                    "args": fn_args,
                    "result_summary": f"Returned {tool_result.get('count', 1)} items"
                })

                # Extract citations
                if fn_name == "search_sharepoint_documents":
                    for p in tool_result.get("passages", []):
                        citations.append({
                            "type": "document",
                            "source": f"{p['library']} / {p['document']} (Page {p['page_or_section']})",
                            "url": p.get("webUrl")
                        })
                elif fn_name == "query_sharepoint_list":
                    citations.append({
                        "type": "list",
                        "source": f"SharePoint List: {tool_result.get('list_queried')}",
                        "count": tool_result.get("count", 0)
                    })

                # Send tool response back to OpenAI
                messages.append({
                    "role": "tool",
                    "tool_call_id": tc["id"],
                    "content": json.dumps(tool_result, ensure_ascii=False)
                })

        # Fallback if loop finishes
        final_resp = self._call_openai(messages)
        content = final_resp.json()["choices"][0]["message"].get("content", "")
        return {
            "response": content,
            "tool_activity": tool_activity,
            "citations": citations
        }

    def chat_stream(self, user_messages: List[Dict[str, str]]) -> Generator[str, None, None]:
        """Streaming chat: executes tools synchronously and streams the final synthesis."""
        messages = [{"role": "system", "content": SYSTEM_PROMPT}]
        messages.extend(user_messages)

        tool_activity = []
        citations = []

        # Tool calling loop
        for step in range(3):
            resp = self._call_openai(messages, tools=TOOL_DEFINITIONS, stream=False)
            if resp.status_code != 200:
                err_data = json.dumps({"error": f"Azure OpenAI Error ({resp.status_code}): {resp.text}"})
                yield f"data: {err_data}\n\n"
                return

            res_json = resp.json()
            choice = res_json["choices"][0]
            msg = choice["message"]
            messages.append(msg)

            tool_calls = msg.get("tool_calls")
            if not tool_calls:
                # If no tools called on first turn or after previous tool calls, stream response
                content = msg.get("content", "")
                if content:
                    yield f"data: {json.dumps({'token': content})}\n\n"
                yield f"data: {json.dumps({'citations': citations, 'tool_activity': tool_activity, 'done': True})}\n\n"
                return

            # There are tools to execute
            for tc in tool_calls:
                fn_name = tc["function"]["name"]
                raw_args = tc["function"]["arguments"]
                try:
                    fn_args = json.loads(raw_args)
                except Exception:
                    fn_args = {}

                # Notify client about tool execution
                status_msg = f"Reading SharePoint { 'List: ' + fn_args.get('list_name', '') if fn_name == 'query_sharepoint_list' else 'Document Library: ' + (fn_args.get('library_name') or 'Knowledge Base') }..."
                yield f"data: {json.dumps({'status': status_msg, 'tool': fn_name, 'args': fn_args})}\n\n"

                tool_result = self.execute_tool(fn_name, fn_args)
                tool_activity.append({
                    "tool": fn_name,
                    "args": fn_args,
                    "result_summary": f"Returned {tool_result.get('count', 1)} items"
                })

                if fn_name == "search_sharepoint_documents":
                    for p in tool_result.get("passages", []):
                        citations.append({
                            "type": "document",
                            "source": f"{p['library']} / {p['document']} (Page {p['page_or_section']})",
                            "url": p.get("webUrl")
                        })
                elif fn_name == "query_sharepoint_list":
                    citations.append({
                        "type": "list",
                        "source": f"SharePoint List: {tool_result.get('list_queried')}",
                        "count": tool_result.get("count", 0)
                    })

                messages.append({
                    "role": "tool",
                    "tool_call_id": tc["id"],
                    "content": json.dumps(tool_result, ensure_ascii=False)
                })

        # Stream the final answer
        yield f"data: {json.dumps({'status': 'Generating response with SharePoint insights...'})}\n\n"
        stream_resp = self._call_openai(messages, stream=True)
        if stream_resp.status_code != 200:
            yield f"data: {json.dumps({'error': f'Stream error: {stream_resp.text}'})}\n\n"
            return

        for line in stream_resp.iter_lines():
            if not line:
                continue
            line_str = line.decode("utf-8")
            if line_str.startswith("data: "):
                payload_str = line_str[6:].strip()
                if payload_str == "[DONE]":
                    break
                try:
                    chunk_json = json.loads(payload_str)
                    delta = chunk_json["choices"][0].get("delta", {})
                    token = delta.get("content")
                    if token:
                        yield f"data: {json.dumps({'token': token})}\n\n"
                except Exception:
                    pass

        # Send final citations and done event
        yield f"data: {json.dumps({'citations': citations, 'tool_activity': tool_activity, 'done': True})}\n\n"
