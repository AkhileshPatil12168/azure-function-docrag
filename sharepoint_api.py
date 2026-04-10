from dotenv import load_dotenv
load_dotenv()

import os
import logging
from typing import Optional, Dict, Any

import msal
import requests
from fastapi import APIRouter, HTTPException, Query
from fastapi.responses import RedirectResponse, Response
from fastapi import Request

from openai import AzureOpenAI
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
from azure.search.documents.models import VectorizedQuery
from azure.ai.formrecognizer import DocumentAnalysisClient
import uuid

logger = logging.getLogger("sharepoint_api")

router = APIRouter(prefix="/api/sharepoint", tags=["sharepoint"])

# In-memory token (POC purpose)
ACCESS_TOKEN = None

# Environment config
TENANT_ID = os.getenv("MS_TENANT_ID", "")
CLIENT_ID = os.getenv("MS_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET", "")
REDIRECT_URI = os.getenv("MS_REDIRECT_URI", "")

SCOPES = ["User.Read", "Sites.Read.All", "Files.Read.All"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"



# Initialize clients once
openai_client = AzureOpenAI(
    api_key=os.environ["OPENAI_API_KEY"],
    api_version="2024-02-01",
    azure_endpoint=os.environ["OPENAI_ENDPOINT"]
)

search_client = SearchClient(
    endpoint=os.environ["SEARCH_ENDPOINT"],
    index_name="rag-index",
    credential=AzureKeyCredential(os.environ["SEARCH_API_KEY"])
)

doc_client = DocumentAnalysisClient(
    endpoint=os.environ["DOC_INT_ENDPOINT"],
    credential=AzureKeyCredential(os.environ["DOC_INT_KEY"])
)


# Create MSAL app for authentication
def _msal_app():
    if not (TENANT_ID and CLIENT_ID and CLIENT_SECRET and REDIRECT_URI):
        raise RuntimeError("Missing MS_* environment variables")

    return msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=AUTHORITY,
    )


# Generic helper to call Microsoft Graph API
def _graph_get(token: str, path: str, params: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
    url = f"{GRAPH_BASE}{path}"

    res = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        params=params,
        timeout=60
    )

    if res.status_code >= 400:
        raise HTTPException(status_code=res.status_code, detail=res.text)

    return res.json()


# Ensure user is logged in
def _require_token(request: Request):
    token = request.session.get("access_token")

    if not token:
        raise HTTPException(status_code=401, detail="Not logged in")

    return token


# =========================
# AUTH APIs
# =========================

@router.get("/auth/start")
def auth_start():
    """
    Redirects user to Microsoft login page
    """
    app = _msal_app()

    auth_url = app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        prompt="select_account",
    )

    return RedirectResponse(auth_url)






@router.get("/auth/callback")
def auth_callback(request: Request, code: Optional[str] = None):

    if not code:
        raise HTTPException(status_code=400, detail="Missing code")

    app = _msal_app()

    token_result = app.acquire_token_by_authorization_code(
        code=code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )

    if "error" in token_result:
        raise HTTPException(status_code=400, detail="Token exchange failed")

    # Save token in session
    request.session["access_token"] = token_result.get("access_token")
    request.session["expires_in"] = token_result.get("expires_in")

    # 🔥 Redirect to homepage
    return RedirectResponse(url="https://doc-rag-intelligence.cognitiveservices.azure.com")

@router.get("/token")
def get_token():
    """
    Returns current access token (used by Azure Function)
    """
    return {"access_token": ACCESS_TOKEN}


# =========================
# UI APIs (Used by frontend)
# =========================

@router.get("/ui/sites")
def ui_sites(request: Request):
    """
    Fetch all SharePoint sites
    """
    token = _require_token(request)

    data = _graph_get(token, "/sites", params={"search": "*"})

    return {
        "sites": [
            {"id": s["id"], "name": s.get("displayName", "No Name")}
            for s in data.get("value", [])
        ]
    }


@router.get("/ui/drives")
def ui_drives(request: Request, site_id: str = Query(...)):
    """
    Fetch document libraries (drives) for selected site
    """
    token = _require_token(request)

    data = _graph_get(token, f"/sites/{site_id}/drives")

    return {
        "drives": [
            {"id": d["id"], "name": d.get("name", "No Name")}
            for d in data.get("value", [])
        ]
    }


@router.get("/ui/pdfs")
def ui_pdfs(request: Request, drive_id: str = Query(...) ):
    """
    Recursively fetch all PDF files inside a drive
    """
    token = _require_token(request)
    pdfs = _find_all_pdfs(token, drive_id)

    return {"files": pdfs}


# =========================
# FILE HANDLING
# =========================

def _find_all_pdfs(token, drive_id, folder_id=None):
    """
    Recursively traverse folders and collect all PDF files
    """
    results = []

    if folder_id:
        url = f"/drives/{drive_id}/items/{folder_id}/children"
    else:
        url = f"/drives/{drive_id}/root/children"

    data = _graph_get(token, url)

    for item in data.get("value", []):
        name = item.get("name", "").lower()

        # If file and PDF
        if "file" in item and name.endswith(".pdf"):
            results.append({
                "name": item["name"],
                "id": item["id"]
            })

        # If folder → go deeper
        elif "folder" in item:
            results.extend(_find_all_pdfs(token, drive_id, item["id"]))

    return results


@router.get("/preview")
def preview_file(drive_id: str, item_id: str, request: Request):
    """
    Returns PDF file stream for preview in browser
    """
    token = _require_token(request)

    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"

    res = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"}
    )

    return Response(content=res.content, media_type="application/pdf")

@router.post("/process")
def process_pdf(data: dict, request: Request):
    """
    Downloads PDF from SharePoint, extracts text,
    creates embeddings, and stores in AI Search
    """

    drive_id = data.get("drive_id")
    item_id = data.get("item_id")

    token = _require_token(request)

    # Step 1: Download file
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"

    res = requests.get(url, headers={
        "Authorization": f"Bearer {token}"
    })

    file_bytes = res.content

    # Step 2: Extract text
    poller = doc_client.begin_analyze_document(
        model_id="prebuilt-layout",
        document=file_bytes
    )
    result = poller.result()
    text = result.content or ""

    if not text:
        return {"error": "No text extracted"}

    # Step 3: Chunk
    words = text.split()
    chunks = []

    for i in range(0, len(words), 150):
        chunks.append(" ".join(words[i:i+150]))

    # Step 4: Create embeddings
    documents = []

    for chunk in chunks:
        embedding = openai_client.embeddings.create(
            model="text-embedding-3-large",
            input=chunk
        ).data[0].embedding

        documents.append({
            "id": str(uuid.uuid4()),
            "content": chunk,
            "fileName": item_id,
            "embedding": embedding
        })

    # Step 5: Store in AI Search
    search_client.upload_documents(documents)

    return {"message": f"Processed {len(chunks)} chunks"}

@router.post("/ask")
def ask_ai(data: dict, request: Request):
    """
    Performs RAG search + GPT answer
    """

    question = data.get("question")

    if not question:
        return {"error": "Missing question"}

    # Step 1: Embed question
    query_embedding = openai_client.embeddings.create(
        model="text-embedding-3-large",
        input=question
    ).data[0].embedding

    # Step 2: Search
    vector_query = VectorizedQuery(
        vector=query_embedding,
        k_nearest_neighbors=3,
        fields="embedding"
    )

    results = search_client.search(
        search_text=None,
        vector_queries=[vector_query]
    )

    context = "\n".join([r["content"] for r in results])

    # Step 3: GPT answer
    response = openai_client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "Answer only using context"},
            {"role": "user", "content": f"Context:\n{context}\n\nQuestion:{question}"}
        ]
    )

    return {"answer": response.choices[0].message.content}