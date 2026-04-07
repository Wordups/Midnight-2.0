"""
api.py — Midnight FastAPI backend v2.1
"""

from __future__ import annotations

import uuid
from pathlib import Path
from typing import Any

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, FileResponse, HTMLResponse

# ✅ FIXED IMPORTS (CRITICAL)
from .services import (
    get_uploaded_text,
    save_logo_bytes,
    run_llm_transform,
    run_framework_mapping,
    build_output_doc,
    build_grc_summary_doc,
)

from .supabase_client import supabase


app = FastAPI(title="Midnight API", version="2.1")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory stores
LOGO_STORE: dict[str, str] = {}
FRAMEWORK_STORE: dict[str, dict] = {}

TENANT_ID = "d1ed6950-8084-45f2-8ae3-0d5f1a15e442"


# ── Helpers ───────────────────────────────────────────────────────────────────

def _safe_supabase_insert(table: str, payload: dict[str, Any]) -> None:
    try:
        supabase.table(table).insert(payload).execute()
        print(f"[SUPABASE] INSERT {table}: OK")
    except Exception as e:
        print(f"[SUPABASE ERROR] INSERT {table}: {e}")


def _safe_supabase_update(table: str, match_field: str, match_value: Any, payload: dict[str, Any]) -> None:
    try:
        supabase.table(table).update(payload).eq(match_field, match_value).execute()
        print(f"[SUPABASE] UPDATE {table}: OK")
    except Exception as e:
        print(f"[SUPABASE ERROR] UPDATE {table}: {e}")


def _persist_preview_run(policy_data: dict[str, Any], framework_map: dict[str, Any], policy_id: str) -> None:
    print(f"[PERSIST] {policy_data.get('policy_name')}")

    _safe_supabase_insert("policies", {
        "id": policy_id,
        "org_id": TENANT_ID,
        "policy_name": policy_data.get("policy_name", ""),
        "policy_number": policy_data.get("policy_number", ""),
        "version": policy_data.get("version", ""),
        "status": "in_progress",
    })

    _safe_supabase_insert("policy_runs", {
        "policy_id": policy_id,
        "framework_map": framework_map,
        "overall_coverage": framework_map.get("overall_coverage", "unknown"),
        "total_mapped": framework_map.get("total_controls_mapped", 0),
        "total_gaps": framework_map.get("total_gaps", 0),
        "audit_summary": framework_map.get("audit_summary", ""),
    })


# ── Root ──────────────────────────────────────────────────────────────────────

@app.get("/")
def serve_tool():
    tool_path = Path("tool.html")
    if tool_path.exists():
        return FileResponse("tool.html", media_type="text/html")
    return HTMLResponse("<h2>Midnight API running</h2>")


@app.get("/api/health")
def health():
    return {"status": "ok", "version": "2.1"}


# ── Preview ───────────────────────────────────────────────────────────────────

@app.post("/api/migrate/preview")
async def preview(
    source_file: UploadFile = File(...),
    template_name: str = Form(...),
    logo_file: UploadFile | None = File(None),
):
    try:
        source_bytes = await source_file.read()
        source_text = get_uploaded_text(source_file.filename, source_bytes)

        policy_data = run_llm_transform(source_text, template_name)
        framework_map = run_framework_mapping(policy_data)

        policy_id = str(uuid.uuid4())

        FRAMEWORK_STORE[policy_id] = {
            "policy_data": policy_data,
            "framework_map": framework_map,
        }

        logo_token = None
        if logo_file:
            logo_bytes = await logo_file.read()
            logo_path = save_logo_bytes(logo_file.filename, logo_bytes)
            logo_token = str(uuid.uuid4())
            LOGO_STORE[logo_token] = logo_path

        _persist_preview_run(policy_data, framework_map, policy_id)

        return {
            "policy_data": policy_data,
            "framework_map": framework_map,
            "policy_id": policy_id,
            "logo_token": logo_token,
        }

    except Exception as e:
        raise HTTPException(400, str(e))


# ── DOCX ──────────────────────────────────────────────────────────────────────

@app.post("/api/migrate/generate")
async def generate(payload: dict[str, Any]):
    try:
        policy_data = payload.get("policy_data")
        logo_token = payload.get("logo_token")

        if not policy_data:
            raise HTTPException(400, "Missing policy_data")

        logo_path = LOGO_STORE.get(logo_token)

        filename, file_bytes = build_output_doc(policy_data, logo_path)

        return StreamingResponse(
            iter([file_bytes]),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    except Exception as e:
        raise HTTPException(400, str(e))


# ── GRC PDF ───────────────────────────────────────────────────────────────────

@app.post("/api/migrate/grc-summary")
async def grc_summary(payload: dict[str, Any]):
    try:
        policy_id = payload.get("policy_id")
        policy_data = payload.get("policy_data")
        framework_map = payload.get("framework_map")

        if policy_id and policy_id in FRAMEWORK_STORE:
            stored = FRAMEWORK_STORE[policy_id]
            policy_data = stored["policy_data"]
            framework_map = stored["framework_map"]

        if not policy_data or not framework_map:
            raise HTTPException(400, "Missing data")

        filename, file_bytes = build_grc_summary_doc(policy_data, framework_map)

        return StreamingResponse(
            iter([file_bytes]),
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    except Exception as e:
        raise HTTPException(400, str(e))
