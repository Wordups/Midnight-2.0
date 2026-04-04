import tempfile
import uuid

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

from services import (
    get_uploaded_text,
    save_logo_bytes,
    run_llm_transform,
    build_output_doc
)

app = FastAPI()

# allow your local HTML file to call API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# store temp logos between preview → generate
LOGO_STORE = {}


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/migrate/preview")
async def preview(
    source_file: UploadFile = File(...),
    template_name: str = Form(...),
    logo_file: UploadFile = File(None)
):
    try:
        source_bytes = await source_file.read()
        source_text = get_uploaded_text(source_file.filename, source_bytes)

        # 🔍 DEBUG LOGGING
        print("\n--- DEBUG EXTRACTION ---")
        print("Length:", len(source_text.strip()))
        print("Preview:", source_text[:500])
        print("------------------------\n")

        if len(source_text.strip()) < 5:
            raise HTTPException(
                status_code=400,
                detail="No readable text was extracted from the document."
            )

        policy_data = run_llm_transform(source_text, template_name)

        logo_token = None
        if logo_file:
            logo_bytes = await logo_file.read()
            logo_path = save_logo_bytes(logo_file.filename, logo_bytes)
            logo_token = str(uuid.uuid4())
            LOGO_STORE[logo_token] = logo_path

        return {
            "policy_data": policy_data,
            "logo_token": logo_token
        }

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/migrate/generate")
async def generate(payload: dict):
    try:
        policy_data = payload.get("policy_data")
        logo_token = payload.get("logo_token")

        if not policy_data:
            raise HTTPException(status_code=400, detail="Missing policy_data")

        logo_path = LOGO_STORE.get(logo_token)

        filename, file_bytes = build_output_doc(policy_data, logo_path)

        return StreamingResponse(
            iter([file_bytes]),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"'
            }
        )

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
