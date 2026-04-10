from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.sessions import SessionMiddleware
from fastapi.responses import FileResponse

from sharepoint_api import router

app = FastAPI()

# Session (required for login)
app.add_middleware(SessionMiddleware, secret_key="poc-demo-for-rag-chatbot",session_cookie="session",)

# CORS (safe for POC)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(router)

# Serve UI
@app.get("/")
def home():
    return FileResponse("index.html")