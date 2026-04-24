# from dotenv import load_dotenv
# load_dotenv()

# import os
# import logging
# from typing import Optional, Dict, Any
# import time
# import requests
# from fastapi import APIRouter, HTTPException, Query
# from fastapi.responses import Response

# from openai import AzureOpenAI
# from azure.search.documents import SearchClient
# from azure.core.credentials import AzureKeyCredential
# from azure.search.documents.models import VectorizedQuery
# from azure.ai.formrecognizer import DocumentAnalysisClient
# import uuid
# import json

# logger = logging.getLogger("sharepoint_api")

# router = APIRouter(prefix="/api/sharepoint", tags=["sharepoint"])

# # =========================
# # ENV CONFIG
# # =========================
# TENANT_ID = os.getenv("MS_TENANT_ID")
# CLIENT_ID = os.getenv("MS_CLIENT_ID")
# CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")

# GRAPH_BASE = "https://graph.microsoft.com/v1.0"
# TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

# # =========================
# # TOKEN CACHE
# # =========================
# ACCESS_TOKEN = None
# TOKEN_EXPIRY = 0


# tools = [
#     {
#         "type": "function",
#         "function": {
#             "name": "search_sharepoint",
#             "description": "Search SharePoint documents and return relevant content",
#             "parameters": {
#                 "type": "object",
#                 "properties": {
#                     "query": {"type": "string"}
#                 },
#                 "required": ["query"]
#             }
#         }
#     },
     
#     {
#     "type": "function",
#     "function": {
#         "name": "query_data",
#         "description": "Use this tool to answer questions about numbers, statistics, claims data, totals, counts, or structured business data",
#         "parameters": {
#             "type": "object",
#             "properties": {
#                 "query": {"type": "string"}
#             },
#             "required": ["query"]
#         }
#     }
# }
# ]


# def get_access_token():
#     global ACCESS_TOKEN, TOKEN_EXPIRY

#     if ACCESS_TOKEN and time.time() < TOKEN_EXPIRY:
#         return ACCESS_TOKEN

#     data = {
#         "client_id": CLIENT_ID,
#         "client_secret": CLIENT_SECRET,
#         "scope": "https://graph.microsoft.com/.default",
#         "grant_type": "client_credentials"
#     }

#     res = requests.post(TOKEN_URL, data=data)

#     if res.status_code != 200:
#         raise HTTPException(status_code=500, detail="Failed to get access token")

#     token_json = res.json()
#     ACCESS_TOKEN = token_json["access_token"]
#     TOKEN_EXPIRY = time.time() + int(token_json.get("expires_in", 3599)) - 60

#     return ACCESS_TOKEN


# # =========================
# # GRAPH HELPER (WITH RETRY)
# # =========================
# def _graph_get_full(url: str):
#     while url:
#         token = get_access_token()

#         res = requests.get(url, headers={"Authorization": f"Bearer {token}"})

#         if res.status_code == 401:
#             token = get_access_token()
#             res = requests.get(url, headers={"Authorization": f"Bearer {token}"})

#         if res.status_code >= 400:
#             raise HTTPException(status_code=res.status_code, detail=res.text)

#         data = res.json()
#         yield data
#         url = data.get("@odata.nextLink")


# def _graph_get(path: str, params: Optional[Dict[str, str]] = None):
#     token = get_access_token()

#     res = requests.get(
#         f"{GRAPH_BASE}{path}",
#         headers={"Authorization": f"Bearer {token}"},
#         params=params
#     )

#     if res.status_code >= 400:
#         raise HTTPException(status_code=res.status_code, detail=res.text)

#     return res.json()


# # =========================
# # AI CLIENTS
# # =========================
# openai_client = AzureOpenAI(
#     api_key=os.environ["OPENAI_API_KEY"],
#     api_version="2024-02-01",
#     azure_endpoint=os.environ["OPENAI_ENDPOINT"]
# )

# search_client = SearchClient(
#     endpoint=os.environ["SEARCH_ENDPOINT"],
#     index_name="rag-index",
#     credential=AzureKeyCredential(os.environ["SEARCH_API_KEY"])
# )

# doc_client = DocumentAnalysisClient(
#     endpoint=os.environ["DOC_INT_ENDPOINT"],
#     credential=AzureKeyCredential(os.environ["DOC_INT_KEY"])
# )


# # =========================
# # UI APIs
# # =========================

# @router.get("/ui/sites")
# def ui_sites():
#     # directly return your known site
#     return {
#         "sites": [
#             {
#                 "id": "wearelucidgroup.sharepoint.com,708eb9a3-40f8-4f89-bff2-7c71f236edf7,1fa59446-f30d-4f1c-9623-d49b13cc8746",
#                 "name": "Team DeUS AI Development"
#             }
#         ]
#     }


# @router.get("/ui/drives")
# def ui_drives(site_id: str = Query(...)):
#     data = _graph_get(f"/sites/{site_id}/drives")

#     return {
#         "drives": [
#             {"id": d["id"], "name": d.get("name", "No Name")}
#             for d in data.get("value", [])
#         ]
#     }


# @router.get("/ui/files")
# def ui_files(drive_id: str = Query(...)):
#     files = _find_all_files(drive_id)
#     return {"files": files}


# # =========================
# # FILE HANDLING
# # =========================

# def _find_all_files(drive_id, folder_id=None):
#     results = []

#     if folder_id:
#         url = f"{GRAPH_BASE}/drives/{drive_id}/items/{folder_id}/children"
#     else:
#         url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"

#     for data in _graph_get_full(url):
#         for item in data.get("value", []):

#             # FILE
#             if "file" in item:
#                 results.append({
#                     "name": item["name"],
#                     "id": item["id"],
#                     "mimeType": item["file"].get("mimeType", "")
#                 })

#             # 🔁 FOLDER
#             elif "folder" in item:
#                 results.extend(_find_all_files(drive_id, item["id"]))

#     return results

# def chunk_text(text, chunk_size=500, overlap=100):
#     words = text.split()
#     chunks = []

#     i = 0
#     while i < len(words):
#         chunk = words[i:i + chunk_size]
#         chunks.append(" ".join(chunk))
#         i += chunk_size - overlap

#     return chunks


# @router.get("/preview")
# def preview_file(drive_id: str, item_id: str):
#     token = get_access_token()

#     url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"

#     res = requests.get(url, headers={"Authorization": f"Bearer {token}"})

#     return Response(content=res.content, media_type="application/pdf")


# @router.post("/process")
# def process_file(data: dict):
#     drive_id = data.get("drive_id")
#     item_id = data.get("item_id")

#     token = get_access_token()

#     url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"

#     res = requests.get(url, headers={"Authorization": f"Bearer {token}"})
#     file_bytes = res.content
#     content_type = res.headers.get("Content-Type", "")

#     text = ""

#     # =========================
#     # PDF / IMAGE
#     # =========================
#     if "pdf" in content_type or "image" in content_type:
#         poller = doc_client.begin_analyze_document(
#             model_id="prebuilt-layout",
#             document=file_bytes
#         )
#         result = poller.result()
#         text = result.content or ""

#     # =========================
#     # TEXT FILE
#     # =========================
#     elif "text" in content_type:
#         text = file_bytes.decode("utf-8", errors="ignore")

#     # =========================
#     # DOCX
#     # =========================
#     elif "wordprocessingml" in content_type:
#         from docx import Document
#         import io

#         doc = Document(io.BytesIO(file_bytes))
#         text = "\n".join([p.text for p in doc.paragraphs])

#     # =========================
#     # CSV
#     # =========================
#     elif "csv" in content_type:
#         import pandas as pd
#         import io

#         df = pd.read_csv(io.BytesIO(file_bytes))
#         text = df.to_string()

#     # =========================
#     # PPTX
#     # =========================
#     elif "presentationml" in content_type:
#         from pptx import Presentation
#         import io

#         prs = Presentation(io.BytesIO(file_bytes))
#         slides_text = []

#         for slide in prs.slides:
#             slide_content = []
#             for shape in slide.shapes:
#                 if hasattr(shape, "text"):
#                     slide_content.append(shape.text)
#             slides_text.append(" ".join(slide_content))

#         text = "\n".join(slides_text)

#     # =========================
#     # XLSX
#     # =========================
#     elif "spreadsheetml" in content_type:
#         import pandas as pd
#         import io

#         try:
#             excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
#             sheets_text = []

#             for sheet in excel_file.sheet_names:
#                 df = excel_file.parse(sheet)
#                 sheets_text.append(f"Sheet: {sheet}\n{df.to_string()}")

#             text = "\n\n".join(sheets_text)

#         except Exception as e:
#             return {"error": f"Excel parsing failed: {str(e)}"}

#     # =========================
#     # UNSUPPORTED
#     # =========================
#     else:
#         return {"error": f"Unsupported file type: {content_type}"}

#     if not text.strip():
#         return {"error": "No text extracted"}

#     # =========================
#     # CHUNKING
#     # =========================
#     chunks = chunk_text(text)

#     documents = []

#     for chunk in chunks:
#         embedding = openai_client.embeddings.create(
#             model="text-embedding-3-large",
#             input=chunk
#         ).data[0].embedding

#         documents.append({
#             "id": str(uuid.uuid4()),
#             "content": chunk,
#             "fileName": item_id,
#             "embedding": embedding
#         })

#     search_client.upload_documents(documents)

#     return {
#         "message": f"Processed {len(chunks)} chunks",
#         "type": content_type
#     }

# @router.post("/process-all")
# def process_all(drive_id: str):
#     pdfs = _find_all_pdfs(drive_id)

#     results = []

#     for pdf in pdfs:
#         try:
#             process_pdf({
#                 "drive_id": drive_id,
#                 "item_id": pdf["id"]
#             })
#             results.append(pdf["name"])
#         except Exception as e:
#             logger.error(f"Failed: {pdf['name']}")

#     return {"processed": results}


# def search_sharepoint_tool(query: str):
#     query_embedding = openai_client.embeddings.create(
#         model="text-embedding-3-large",
#         input=query
#     ).data[0].embedding

#     vector_query = VectorizedQuery(
#         vector=query_embedding,
#         k_nearest_neighbors=5,
#         fields="embedding"
#     )

#     results = search_client.search(
#         search_text=query,
#         vector_queries=[vector_query],
#         top=5,
#         query_type="semantic",
#         semantic_configuration_name="default"
#     )

#     context = ""

#     for r in results:
#         content = r.get("content", "")
#         source = r.get("fileName", "unknown")
#         score = r.get("@search.score", "")

#         context += f"\n[Source: {source} | Score: {score}]\n{content}\n"

#     return context

# def query_data_tool(query: str):
#     query = query.lower()

#     data = {
#         "total claims": "Total claims: 120",
#         "denied claims": "Denied claims: 15",
#         "approved claims": "Approved claims: 105",
#         "pending claims": "Pending claims: 10"
#     }

#     for key in data:
#         if key in query:
#             return data[key]

#     return "No structured data found"

# @router.post("/ask")
# def ask_ai(data: dict):
#     question = data.get("question")

#     if not question:
#         return {"error": "Missing question"}

#     messages = [
#         {
#   "role": "system",
#   "content": """
# You are an intelligent AI agent.

# You have access to tools:

# 1. search_sharepoint → for documents, policies, text
# 2. query_data → for numbers, counts, claims, statistics

# Rules:
# - If question asks about numbers, counts, totals → use query_data
# - If question asks about policy, documents → use search_sharepoint
# - If both needed → prefer best matching tool

# Always use tools when needed.
# """
# },
#         {"role": "user", "content": question}
#     ]

#     # Step 1: Agent decides
#     response = openai_client.chat.completions.create(
#         model="gpt-4o",
#         messages=messages,
#         tools=tools,
#         tool_choice="auto"
#     )

#     message = response.choices[0].message

#     # Step 2: Tool execution
#     if message.tool_calls:

#         messages.append(message)

#         for tool_call in message.tool_calls:
#             tool_name = tool_call.function.name
#             arguments = json.loads(tool_call.function.arguments)

#             if tool_name == "search_sharepoint":
#                 tool_result = search_sharepoint_tool(arguments["query"])

#             elif tool_name == "query_data":
#                 tool_result = query_data_tool(arguments["query"])

#             else:
#                 tool_result = "Unknown tool"

#             messages.append({
#                 "role": "tool",
#                 "tool_call_id": tool_call.id,
#                 "content": tool_result
#             })

#     # Final response after ALL tools
#         final_response = openai_client.chat.completions.create(
#             model="gpt-4o",
#             messages=messages
#         )

#         return {"answer": final_response.choices[0].message.content}