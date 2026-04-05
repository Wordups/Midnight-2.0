"""
api.py — Midnight FastAPI backend v2.0
"""

import uuid
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path

from services import (
    get_uploaded_text,
    save_logo_bytes,
    run_llm_transform,
    run_framework_mapping,
    build_output_doc,
    build_grc_summary_doc,
)

app = FastAPI(title="Midnight API", version="2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory stores (resets on redeploy)
LOGO_STORE      = {}   # logo_token → file path
FRAMEWORK_STORE = {}   # policy_id  → {policy_data, framework_map}


# ── Serve tool UI ─────────────────────────────────────────────────────────────
@app.get("/")
def serve_tool():
    tool_path = Path("tool.html")
    if tool_path.exists():
        return FileResponse("tool.html", media_type="text/html")
    return HTMLResponse("<h2>Midnight API is running. Tool not found.</h2>", status_code=200)


# ── Health ────────────────────────────────────────────────────────────────────
@app.get("/api/health")
def health():
    return {"status": "ok", "version": "2.0"}


# ── Preview: extract + transform + framework map ──────────────────────────────
@app.post("/api/migrate/preview")
async def preview(
    source_file:   UploadFile = File(...),
    template_name: str        = Form(...),
    logo_file:     UploadFile = File(None),
):
    try:
        source_bytes = await source_file.read()
        source_text  = get_uploaded_text(source_file.filename, source_bytes)

        if len(source_text.strip()) < 5:
            raise HTTPException(400, "No readable text extracted from document.")

        # Step 1 — extract
        policy_data = run_llm_transform(source_text, template_name)

        # Step 2 — framework mapping
        framework_map = run_framework_mapping(policy_data)

        # Step 3 — store for GRC summary
        policy_id = str(uuid.uuid4())
        FRAMEWORK_STORE[policy_id] = {
            "policy_data":   policy_data,
            "framework_map": framework_map,
        }

        # Step 4 — handle logo
        logo_token = None
        if logo_file:
            logo_bytes = await logo_file.read()
            logo_path  = save_logo_bytes(logo_file.filename, logo_bytes)
            logo_token = str(uuid.uuid4())
            LOGO_STORE[logo_token] = logo_path

        return {
            "policy_data":   policy_data,
            "framework_map": framework_map,
            "logo_token":    logo_token,
            "policy_id":     policy_id,
        }

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


# ── Generate policy .docx ─────────────────────────────────────────────────────
@app.post("/api/migrate/generate")
async def generate(payload: dict):
    try:
        policy_data = payload.get("policy_data")
        logo_token  = payload.get("logo_token")

        if not policy_data:
            raise HTTPException(400, "Missing policy_data")

        logo_path = LOGO_STORE.get(logo_token)
        filename, file_bytes = build_output_doc(policy_data, logo_path)

        return StreamingResponse(
            iter([file_bytes]),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


# ── Generate GRC summary .docx ────────────────────────────────────────────────
@app.post("/api/migrate/grc-summary")
async def grc_summary(payload: dict):
    try:
        policy_id     = payload.get("policy_id")
        policy_data   = payload.get("policy_data")
        framework_map = payload.get("framework_map")

        # Try store first, fall back to payload
        if policy_id and policy_id in FRAMEWORK_STORE:
            stored        = FRAMEWORK_STORE[policy_id]
            policy_data   = stored["policy_data"]
            framework_map = stored["framework_map"]

        if not policy_data or not framework_map:
            raise HTTPException(400, "Missing policy_data or framework_map")

        filename, file_bytes = build_grc_summary_doc(policy_data, framework_map)

        return StreamingResponse(
            iter([file_bytes]),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


# ── Create policy ─────────────────────────────────────────────────────────────
@app.post("/api/create/generate")
async def create_generate(payload: dict):
    try:
        policy_data = payload.get("policy_data")
        logo_token  = payload.get("logo_token")

        if not policy_data:
            raise HTTPException(400, "Missing policy_data")

        # Run framework mapping on created policy too
        framework_map = run_framework_mapping(policy_data)
        policy_data["framework_map"] = framework_map

        logo_path = LOGO_STORE.get(logo_token)
        filename, file_bytes = build_output_doc(policy_data, logo_path)

        return StreamingResponse(
            iter([file_bytes]),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))
