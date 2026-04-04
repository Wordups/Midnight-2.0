"""
services.py — Midnight backend processing layer
Handles: extraction, LLM transform, logo saving, doc building
"""

import os
import io
import tempfile
import traceback
from pathlib import Path

from docx import Document as DocxDocument
from groq import Groq
from hps_policy_migration_builder import build_policy_document


# ── Config ────────────────────────────────────────────────────────────────────
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")
GROQ_MODEL   = "llama-3.3-70b-versatile"
MIN_TEXT_LEN = 50

EXTRACTION_PROMPT = """
You are a policy migration specialist.

Read the attached legacy policy document and extract ALL content into the exact
Python dictionary structure below.

STRICT RULES:
- Do NOT summarize, rewrite, or remove content
- Preserve source wording as closely as possible
- Fix only minor spacing / punctuation / obvious grammar issues
- Map all content into the correct field
- For procedure items use exactly one type:
  "para"            = standalone paragraph
  "heading"         = bold underlined subsection title
  "bullet"          = first-level bullet
  "sub-bullet"      = second-level bullet
  "bold_intro"      = paragraph starting with bold label; keys: "bold" and "rest"
  "bold_intro_semi" = same as bold_intro but "rest" contains semicolons
  "empty"           = blank spacer line

Return ONLY a valid Python dictionary assignment.
No explanation. No markdown fences. No preamble.
Start your response with:
POLICY_DATA = {
End with the closing brace.

POLICY_DATA = {
    "policy_name": "",
    "policy_number": "",
    "version": "",
    "grc_id": "",
    "supersedes": "",
    "effective_date": "",
    "last_reviewed": "",
    "last_revised": "",
    "custodians": "",
    "owner_name": "",
    "owner_title": "",
    "approver_name": "",
    "approver_title": "",
    "date_signed": "",
    "date_approved": "",
    "applicable_to": {
        "hps_inc": False,
        "agency": True,
        "corporate": True,
        "govt_affairs": False,
        "legal_review": False
    },
    "policy_types": {
        "carrier_specific": False,
        "cross_carrier": False,
        "global": False,
        "on_off_hix": False
    },
    "line_of_business": {
        "all_lobs": True,
        "specific_lob": "",
        "specific_lob_checked": False
    },
    "purpose": "",
    "definitions": {},
    "policy_statement": "",
    "procedures": [],
    "related_policies": [],
    "citations": [],
    "revision_history": []
}

HERE IS THE LEGACY POLICY DOCUMENT:
"""


# ══════════════════════════════════════════════════════════════════════════════
# EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

def _extract_docx_bytes(file_bytes: bytes) -> str:
    """
    Robust .docx extraction — captures paragraphs AND table cells.
    Handles documents where most content lives in tables (common in HPS templates).
    """
    lines = []

    try:
        doc = DocxDocument(io.BytesIO(file_bytes))
    except Exception as e:
        raise ValueError(f"Could not open .docx file: {e}")

    # ── Body paragraphs ───────────────────────────────────────────────────────
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            lines.append(text)

    # ── Table cells ───────────────────────────────────────────────────────────
    # Many enterprise policy docs store metadata and content in tables.
    # Extract every non-empty cell, joining row cells with " | ".
    for table in doc.tables:
        for row in table.rows:
            row_parts = []
            for cell in row.cells:
                # Each cell can have multiple paragraphs
                cell_text = "\n".join(
                    p.text.strip() for p in cell.paragraphs if p.text.strip()
                ).strip()
                if cell_text:
                    row_parts.append(cell_text)
            if row_parts:
                # Deduplicate — merged cells repeat the same text
                seen = []
                for part in row_parts:
                    if part not in seen:
                        seen.append(part)
                lines.append(" | ".join(seen))

    # ── Headers and footers ───────────────────────────────────────────────────
    for section in doc.sections:
        for hdr_para in section.header.paragraphs:
            text = hdr_para.text.strip()
            if text:
                lines.append(f"[HEADER] {text}")
        for ftr_para in section.footer.paragraphs:
            text = ftr_para.text.strip()
            if text:
                lines.append(f"[FOOTER] {text}")

    result = "\n".join(lines)
    return result


def _extract_txt_bytes(file_bytes: bytes) -> str:
    for encoding in ("utf-8", "latin-1", "cp1252"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return file_bytes.decode("utf-8", errors="replace")


def get_uploaded_text(filename: str, file_bytes: bytes) -> str:
    """
    Dispatch extraction by file type.
    Returns the full text content of the uploaded file.
    """
    name = filename.lower().strip()

    if name.endswith(".docx"):
        text = _extract_docx_bytes(file_bytes)
    elif name.endswith((".txt", ".md")):
        text = _extract_txt_bytes(file_bytes)
    else:
        # Try docx first, fall back to plain text
        try:
            text = _extract_docx_bytes(file_bytes)
        except Exception:
            text = _extract_txt_bytes(file_bytes)

    # Debug log — always visible in uvicorn console
    print("\n" + "="*60)
    print(f"EXTRACTION  |  file: {filename}")
    print(f"            |  chars extracted: {len(text.strip())}")
    print(f"            |  first 400 chars:")
    print(text[:400])
    print("="*60 + "\n")

    return text


# ══════════════════════════════════════════════════════════════════════════════
# LLM TRANSFORM
# ══════════════════════════════════════════════════════════════════════════════

def _parse_policy_data(raw: str) -> dict:
    """Parse LLM output into a POLICY_DATA dict."""
    if "POLICY_DATA = {" in raw:
        raw = raw[raw.index("POLICY_DATA = {"):]
    elif "POLICY_DATA={" in raw:
        raw = raw[raw.index("POLICY_DATA={"):]

    # Sanitize smart quotes
    raw = (raw
           .replace("\u201c", '"').replace("\u201d", '"')
           .replace("\u2018", "'").replace("\u2019", "'"))

    namespace = {}
    exec(raw, {}, namespace)
    result = namespace.get("POLICY_DATA")

    if not result or not isinstance(result, dict):
        raise ValueError("LLM did not return a valid POLICY_DATA dictionary.")

    return result


def run_llm_transform(source_text: str, template_name: str) -> dict:
    """Send extracted text to Groq LLM and return parsed POLICY_DATA."""
    api_key = GROQ_API_KEY
    if not api_key:
        raise ValueError(
            "GROQ_API_KEY not set. Add it to your environment or .env file."
        )

    if len(source_text.strip()) < MIN_TEXT_LEN:
        raise ValueError(
            f"Extracted text is too short ({len(source_text.strip())} chars). "
            f"Minimum is {MIN_TEXT_LEN}. Check that the document contains readable text."
        )

    # Truncate to ~24k chars to stay within context limits
    # LLaMA 3.3 70B has 128k context but Groq has throughput limits
    MAX_INPUT = 24000
    if len(source_text) > MAX_INPUT:
        print(f"[WARN] Source text truncated from {len(source_text)} to {MAX_INPUT} chars")
        source_text = source_text[:MAX_INPUT]

    client = Groq(api_key=api_key)

    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{
                "role": "user",
                "content": EXTRACTION_PROMPT + "\n\n" + source_text
            }],
            temperature=0.1,
            max_tokens=8000,
        )
    except Exception as e:
        raise ValueError(f"Groq API call failed: {e}")

    raw = response.choices[0].message.content.strip()

    print("\n" + "="*60)
    print(f"LLM OUTPUT  |  raw length: {len(raw)} chars")
    print(f"            |  first 300 chars:")
    print(raw[:300])
    print("="*60 + "\n")

    policy_data = _parse_policy_data(raw)
    policy_data["template_name"] = template_name
    return policy_data


# ══════════════════════════════════════════════════════════════════════════════
# LOGO
# ══════════════════════════════════════════════════════════════════════════════

def save_logo_bytes(filename: str, file_bytes: bytes) -> str:
    """Save logo to a temp file and return the path."""
    tmp_dir = os.path.join(tempfile.gettempdir(), "midnight_logos")
    os.makedirs(tmp_dir, exist_ok=True)

    ext  = Path(filename).suffix.lower() or ".png"
    stem = "".join(c for c in Path(filename).stem if c.isalnum() or c in "-_") or "logo"
    path = os.path.join(tmp_dir, f"{stem}_{uuid.uuid4().hex[:8]}{ext}")

    with open(path, "wb") as f:
        f.write(file_bytes)

    return path


# ══════════════════════════════════════════════════════════════════════════════
# DOC BUILDER
# ══════════════════════════════════════════════════════════════════════════════

def build_output_doc(policy_data: dict, logo_path: str | None = None) -> tuple[str, bytes]:
    """Build the final .docx and return (filename, bytes)."""
    name   = policy_data.get("policy_name",   "Policy")
    number = policy_data.get("policy_number", "SEC-P")
    ver    = policy_data.get("version",       "V1.0")
    fname  = f"{number} {name} {ver}-NEW.docx"

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp_path = tmp.name

    try:
        build_policy_document(policy_data, tmp_path, logo_path=logo_path)
        with open(tmp_path, "rb") as f:
            doc_bytes = f.read()
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass

    return fname, doc_bytes
