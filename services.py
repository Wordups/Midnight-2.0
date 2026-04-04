import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from services import (
    get_uploaded_text,
    run_llm_transform,
    build_output_doc,
    save_logo_bytes,
)

app = FastAPI(title="Midnight API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/api/health")
def health():
    return {"status": "ok"}

@app.post("/api/migrate/preview")
async def migrate_preview(
    source_file: UploadFile = File(...),
    template_name: str = Form(...),
    logo_file: UploadFile | None = File(None),
):
    try:
        source_bytes = await source_file.read()
        source_text = get_uploaded_text(source_file.filename, source_bytes)

        if len(source_text.strip()) < 50:
            raise HTTPException(status_code=400, detail="Document too short.")

        policy_data = run_llm_transform(source_text, template_name)

        logo_token = None
        if logo_file:
            logo_bytes = await logo_file.read()
            logo_token = save_logo_bytes(logo_file.filename, logo_bytes)

        return {
            "ok": True,
            "policy_data": policy_data,
            "logo_token": logo_token,
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/migrate/generate")
async def migrate_generate(payload: dict):
    try:
        filename, docx_bytes = build_output_doc(
            payload["policy_data"],
            logo_path=payload.get("logo_token")
        )

        return StreamingResponse(
            iter([docx_bytes]),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
