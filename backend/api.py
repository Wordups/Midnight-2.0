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

from backend.services import (
    get_uploaded_text,
    save_logo_bytes,
    run_llm_transform,
    run_framework_mapping,
    build_output_doc,
    build_grc_summary_doc,
)
from supabase_client import supabase


app = FastAPI(title="Midnight API", version="2.1")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory stores (resets on redeploy)
LOGO_STORE: dict[str, str] = {}        # logo_token -> file path
FRAMEWORK_STORE: dict[str, dict] = {}  # policy_id  -> {policy_data, framework_map}

# Temporary demo tenant
TENANT_ID = "d1ed6950-8084-45f2-8ae3-0d5f1a15e442"


# ── Helpers ───────────────────────────────────────────────────────────────────

def _safe_supabase_insert(table: str, payload: dict[str, Any]) -> None:
    """Best-effort insert — database issues do not break the user-facing pipeline."""
    try:
        result = supabase.table(table).insert(payload).execute()
        print(f"[SUPABASE] INSERT {table}: OK")
    except Exception as e:
        print(f"[SUPABASE ERROR] INSERT {table}: {e}")


def _safe_supabase_update(table: str, match_field: str, match_value: Any, payload: dict[str, Any]) -> None:
    """Best-effort update — database issues do not break the user-facing pipeline."""
    try:
        supabase.table(table).update(payload).eq(match_field, match_value).execute()
        print(f"[SUPABASE] UPDATE {table}: OK")
    except Exception as e:
        print(f"[SUPABASE ERROR] UPDATE {table}: {e}")


def _persist_preview_run(
    policy_data: dict[str, Any],
    framework_map: dict[str, Any],
    policy_id: str,
) -> None:
    """
    Persists:
      - policy row
      - policy_run row
      - gap rows
      - activity_log row
    """
    print(f"[PERSIST] Starting persist for: {policy_data.get('policy_name')}")

    _safe_supabase_insert("policies", {
        "id":            policy_id,
        "org_id":        TENANT_ID,
        "policy_name":   policy_data.get("policy_name",   ""),
        "policy_number": policy_data.get("policy_number", ""),
        "version":       policy_data.get("version",       ""),
        "status":        "in_progress",
        "owner_name":    policy_data.get("owner_name",    ""),
        "owner_title":   policy_data.get("owner_title",   ""),
        "approver_name": policy_data.get("approver_name", ""),
        "effective_date":policy_data.get("effective_date",""),
        "last_reviewed": policy_data.get("last_reviewed", ""),
        "template_name": policy_data.get("template_name", ""),
    })

    _safe_supabase_insert("policy_runs", {
        "policy_id":       policy_id,
        "run_type":        "migrate",
        "template_name":   policy_data.get("template_name", ""),
        "framework_map":   framework_map,
        "overall_coverage":framework_map.get("overall_coverage", "unknown"),
        "total_mapped":    framework_map.get("total_controls_mapped", 0),
        "total_gaps":      framework_map.get("total_gaps", 0),
        "audit_summary":   framework_map.get("audit_summary", ""),
    })

    for gap in framework_map.get("gaps", []) or []:
        _safe_supabase_insert("policy_gaps", {
            "policy_id":      policy_id,
            "framework":      gap.get("framework",       ""),
            "control_id":     gap.get("control_id",      ""),
            "control_name":   gap.get("control_name",    ""),
            "gap_description":gap.get("gap_description", ""),
            "risk_level":     gap.get("risk_level",      "medium"),
            "suggestion":     gap.get("suggestion",      ""),
            "status":         "open",
        })

    _safe_supabase_insert("activity_log", {
        "org_id":    TENANT_ID,
        "policy_id": policy_id,
        "user_name": "System",
        "action":    "transform_policy",
        "result":    "audit-ready",
        "detail":    f"{policy_data.get('policy_number','')} {policy_data.get('policy_name','')}".strip(),
    })

    print(f"[PERSIST] Complete for: {policy_data.get('policy_name')}")


def _log_activity(
    action: str,
    policy_id: str | None = None,
    detail: str | None = None,
) -> None:
    payload: dict[str, Any] = {
        "org_id":    TENANT_ID,
        "action":    action,
        "user_name": "System",
        "result":    "audit-ready",
    }
    if policy_id:
        payload["policy_id"] = policy_id
    if detail:
        payload["detail"] = detail
    _safe_supabase_insert("activity_log", payload)


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
    return {"status": "ok", "version": "2.1"}


# ── Preview: extract + transform + framework map ──────────────────────────────

@app.post("/api/migrate/preview")
async def preview(
    source_file:   UploadFile = File(...),
    template_name: str        = Form(...),
    logo_file:     UploadFile | None = File(None),
):
    policy_id: str | None = None

    try:
        source_bytes = await source_file.read()
        source_text  = get_uploaded_text(source_file.filename, source_bytes)

        if len(source_text.strip()) < 5:
            raise HTTPException(400, "No readable text extracted from document.")

        # Step 1 — extract
        policy_data = run_llm_transform(source_text, template_name)

        # Step 2 — framework mapping
        framework_map = run_framework_mapping(policy_data)

        # Step 3 — store in memory for GRC summary
        policy_id = str(uuid.uuid4())
        FRAMEWORK_STORE[policy_id] = {
            "policy_data":  policy_data,
            "framework_map":framework_map,
        }

        # Step 4 — handle logo
        logo_token = None
        if logo_file:
            logo_bytes = await logo_file.read()
            logo_path  = save_logo_bytes(logo_file.filename, logo_bytes)
            logo_token = str(uuid.uuid4())
            LOGO_STORE[logo_token] = logo_path

        # Step 5 — persist to Supabase
        _persist_preview_run(policy_data, framework_map, policy_id)

        return {
            "policy_data":   policy_data,
            "framework_map": framework_map,
            "logo_token":    logo_token,
            "policy_id":     policy_id,
        }

    except HTTPException:
        raise
    except Exception as e:
        if policy_id:
            _log_activity("transform_policy_failed", policy_id=policy_id, detail=str(e))
        raise HTTPException(400, str(e))


# ── Generate policy .docx ─────────────────────────────────────────────────────

@app.post("/api/migrate/generate")
async def generate(payload: dict[str, Any]):
    try:
        policy_data = payload.get("policy_data")
        logo_token  = payload.get("logo_token")

        if not policy_data:
            raise HTTPException(400, "Missing policy_data")

        logo_path            = LOGO_STORE.get(logo_token)
        filename, file_bytes = build_output_doc(policy_data, logo_path)

        detail = f'{policy_data.get("policy_number","")} {policy_data.get("policy_name","")} {policy_data.get("version","")}'.strip()
        _log_activity("generate_document", detail=detail)

        return StreamingResponse(
            iter([file_bytes]),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


# ── Generate GRC summary — PDF ────────────────────────────────────────────────

@app.post("/api/migrate/grc-summary")
async def grc_summary(payload: dict[str, Any]):
    try:
        policy_id     = payload.get("policy_id")
        policy_data   = payload.get("policy_data")
        framework_map = payload.get("framework_map")

        # Prefer stored data over payload
        if policy_id and policy_id in FRAMEWORK_STORE:
            stored        = FRAMEWORK_STORE[policy_id]
            policy_data   = stored["policy_data"]
            framework_map = stored["framework_map"]

        if not policy_data or not framework_map:
            raise HTTPException(400, "Missing policy_data or framework_map")

        filename, file_bytes = build_grc_summary_doc(policy_data, framework_map)

        _log_activity("download_grc_summary", policy_id=policy_id)

        return StreamingResponse(
            iter([file_bytes]),
            media_type="application/pdf",  # ← PDF not docx
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


# ── Create policy ─────────────────────────────────────────────────────────────

@app.post("/api/create/generate")
async def create_generate(payload: dict[str, Any]):
    try:
        policy_data = payload.get("policy_data")
        logo_token  = payload.get("logo_token")

        if not policy_data:
            raise HTTPException(400, "Missing policy_data")

        # Run framework mapping on created policy too
        framework_map            = run_framework_mapping(policy_data)
        policy_data["framework_map"] = framework_map

        logo_path            = LOGO_STORE.get(logo_token)
        filename, file_bytes = build_output_doc(policy_data, logo_path)

        detail = f'{policy_data.get("policy_number","")} {policy_data.get("policy_name","")} {policy_data.get("version","")}'.strip()
        _log_activity("create_policy_generate", detail=detail)

        return StreamingResponse(
            iter([file_bytes]),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))
