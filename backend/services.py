"""
services.py — Midnight backend processing layer

Pipeline:
    get_uploaded_text()
        ↓
    run_llm_transform()       → POLICY_DATA
        ↓
    run_framework_mapping()   → framework_map{}
        ↓
    build_output_doc()        → policy.docx
    build_grc_summary_doc()   → grc_summary.docx
"""

import os
import io
import json
import uuid
import tempfile
from pathlib import Path

from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from groq import Groq
from hps_policy_migration_builder import build_policy_document

# ── Template registry ─────────────────────────────────────────────────────────
# Each template is a standalone module in the templates/ directory.
# Add new templates here — nothing else needs to change.
try:
    from templates.template_generic import build_document as _build_generic
    _TEMPLATES_AVAILABLE = True
except ImportError:
    _TEMPLATES_AVAILABLE = False
    _build_generic = None

TEMPLATE_REGISTRY = {
    # key (template_name in POLICY_DATA)  →  builder function
    "Generic Policy Template":   "_generic",
    "Generic":                   "_generic",
    "Enterprise Policy Template":"_generic",
    # Future templates:
    # "Healthcare":  _build_healthcare,
    # "Finance":     _build_finance,
    # "Tech":        _build_tech,
    # "Legal":       _build_legal,
}


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

FRAMEWORK_MAPPING_PROMPT = """
You are a compliance framework specialist with deep expertise in:
- HIPAA (45 CFR 164)
- HiTrust CSF v9.3 and v11
- PCI DSS v3.2 and v4.0
- ISO/IEC 27001:2022
- CoBIT 5.0 and 2019
- NIST Cybersecurity Framework v1.1 and v2.0
- NIST SP 800-53
- SOC 2 Type II

Analyze this policy and produce a complete framework mapping.

Do FOUR things:

1. MAPPED CITATIONS — for each policy section that satisfies a control,
   cite the exact control number. Be specific.

2. GAP ANALYSIS — identify controls this policy SHOULD address based on
   its type and subject matter but does not. Only flag genuine relevant gaps.

3. SUGGESTIONS — for each gap, write 1-3 sentences of specific policy language
   the organization could add to close it. Write in formal policy language.

4. AUDIT SUMMARY — 2-3 plain-language sentences a non-technical person can read
   to understand: what frameworks are covered, how many gaps exist, overall posture.

Return ONLY valid JSON. No explanation. No markdown. No preamble.
Start with { and end with }

{
  "policy_name": "",
  "policy_type": "",
  "overall_coverage": "strong|moderate|weak",
  "mapped_citations": [
    {
      "framework": "",
      "control_id": "",
      "control_name": "",
      "policy_section": "",
      "coverage_note": ""
    }
  ],
  "gaps": [
    {
      "framework": "",
      "control_id": "",
      "control_name": "",
      "gap_description": "",
      "risk_level": "high|medium|low",
      "suggestion": ""
    }
  ],
  "audit_summary": "",
  "frameworks_covered": [],
  "total_controls_mapped": 0,
  "total_gaps": 0
}

HERE IS THE POLICY TO ANALYZE:
"""


# ── Extraction ────────────────────────────────────────────────────────────────

def _extract_docx_bytes(file_bytes: bytes) -> str:
    lines = []
    try:
        doc = DocxDocument(io.BytesIO(file_bytes))
    except Exception as e:
        raise ValueError(f"Could not open .docx file: {e}")

    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            lines.append(text)

    for table in doc.tables:
        for row in table.rows:
            row_parts = []
            for cell in row.cells:
                cell_text = "\n".join(
                    p.text.strip() for p in cell.paragraphs if p.text.strip()
                ).strip()
                if cell_text:
                    row_parts.append(cell_text)
            if row_parts:
                seen = []
                for part in row_parts:
                    if part not in seen:
                        seen.append(part)
                lines.append(" | ".join(seen))

    for section in doc.sections:
        for hdr_para in section.header.paragraphs:
            text = hdr_para.text.strip()
            if text:
                lines.append(f"[HEADER] {text}")
        for ftr_para in section.footer.paragraphs:
            text = ftr_para.text.strip()
            if text:
                lines.append(f"[FOOTER] {text}")

    return "\n".join(lines)


def _extract_txt_bytes(file_bytes: bytes) -> str:
    for encoding in ("utf-8", "latin-1", "cp1252"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return file_bytes.decode("utf-8", errors="replace")


def get_uploaded_text(filename: str, file_bytes: bytes) -> str:
    name = filename.lower().strip()
    if name.endswith(".docx"):
        text = _extract_docx_bytes(file_bytes)
    elif name.endswith((".txt", ".md")):
        text = _extract_txt_bytes(file_bytes)
    else:
        try:
            text = _extract_docx_bytes(file_bytes)
        except Exception:
            text = _extract_txt_bytes(file_bytes)

    print(f"\n{'='*60}\nEXTRACTION  |  {filename}  |  {len(text.strip())} chars\n{'='*60}\n")
    return text


# ── LLM helpers ───────────────────────────────────────────────────────────────

def _groq_call(prompt: str, content: str, max_tokens: int = 8000) -> str:
    api_key = GROQ_API_KEY
    if not api_key:
        raise ValueError("GROQ_API_KEY not set.")
    client = Groq(api_key=api_key)
    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": prompt + "\n\n" + content}],
        temperature=0.1,
        max_tokens=max_tokens,
    )
    return response.choices[0].message.content.strip()


def _parse_policy_data(raw: str) -> dict:
    if "POLICY_DATA = {" in raw:
        raw = raw[raw.index("POLICY_DATA = {"):]
    elif "POLICY_DATA={" in raw:
        raw = raw[raw.index("POLICY_DATA={"):]

    raw = (raw
           .replace("\u201c", '"').replace("\u201d", '"')
           .replace("\u2018", "'").replace("\u2019", "'"))

    # 🔥 ADD THIS — shows you EXACTLY what is breaking
    print("\n\n===== RAW MODEL OUTPUT START =====\n")
    print(raw)
    print("\n===== RAW MODEL OUTPUT END =====\n\n")

    namespace = {}
    exec(raw, {}, namespace)

    result = namespace.get("POLICY_DATA")
    if not result or not isinstance(result, dict):
        raise ValueError("LLM did not return a valid POLICY_DATA dictionary.")
    return result


def run_llm_transform(source_text: str, template_name: str) -> dict:
    if len(source_text.strip()) < MIN_TEXT_LEN:
        raise ValueError(f"Extracted text too short ({len(source_text.strip())} chars).")

    if len(source_text) > 24000:
        source_text = source_text[:24000]

    try:
        raw = _groq_call(EXTRACTION_PROMPT, source_text, max_tokens=8000)
    except Exception as e:
        raise ValueError(f"Groq extraction failed: {e}")

    policy_data = _parse_policy_data(raw)
    policy_data["template_name"] = template_name
    return policy_data


# ── Framework Mapping ─────────────────────────────────────────────────────────

def _policy_to_text(policy_data: dict) -> str:
    lines = []
    lines.append(f"POLICY: {policy_data.get('policy_name','')} ({policy_data.get('policy_number','')})")
    lines.append(f"VERSION: {policy_data.get('version','')}")
    lines.append("")

    if policy_data.get("purpose"):
        lines.append(f"PURPOSE:\n{policy_data['purpose']}\n")
    if policy_data.get("policy_statement"):
        lines.append(f"POLICY STATEMENT:\n{policy_data['policy_statement']}\n")

    defs = policy_data.get("definitions", {})
    if defs:
        lines.append("DEFINITIONS:")
        for k, v in defs.items():
            lines.append(f"  {k}: {v}")
        lines.append("")

    procs = policy_data.get("procedures", [])
    if procs:
        lines.append("PROCEDURES:")
        for item in procs:
            kind = item.get("type", "")
            if kind == "heading":
                lines.append(f"\n  [{item.get('text','')}]")
            elif kind == "bullet":
                lines.append(f"  • {item.get('text','')}")
            elif kind == "sub-bullet":
                lines.append(f"    ◦ {item.get('text','')}")
            elif kind in ("bold_intro", "bold_intro_semi"):
                lines.append(f"  {item.get('bold','')} {item.get('rest','')}")
            elif kind == "para":
                lines.append(f"  {item.get('text','')}")
        lines.append("")

    existing = policy_data.get("citations", [])
    if existing:
        lines.append("EXISTING CITATIONS:")
        for c in existing:
            lines.append(f"  {c}")

    return "\n".join(lines)


def _empty_map(policy_data: dict) -> dict:
    return {
        "policy_name":          policy_data.get("policy_name", ""),
        "policy_type":          "",
        "overall_coverage":     "unknown",
        "mapped_citations":     [],
        "gaps":                 [],
        "audit_summary":        "Framework mapping unavailable for this policy.",
        "frameworks_covered":   [],
        "total_controls_mapped": 0,
        "total_gaps":           0,
    }


def run_framework_mapping(policy_data: dict) -> dict:
    """
    Layer 2 of the pipeline.
    Maps policy content to compliance framework controls.
    Identifies gaps and generates suggested policy language.
    Returns a framework_map dict.
    """
    policy_text = _policy_to_text(policy_data)
    if len(policy_text) > 16000:
        policy_text = policy_text[:16000]

    print(f"\n{'='*60}\nFRAMEWORK   |  {policy_data.get('policy_name','')}  |  {len(policy_text)} chars\n{'='*60}\n")

    try:
        raw = _groq_call(FRAMEWORK_MAPPING_PROMPT, policy_text, max_tokens=6000)
    except Exception as e:
        print(f"[WARN] Framework mapping failed: {e}")
        return _empty_map(policy_data)

    raw = raw.strip()
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else raw
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    try:
        framework_map = json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"[WARN] JSON parse failed: {e} — raw: {raw[:300]}")
        return _empty_map(policy_data)

    print(f"FRAMEWORK   |  mapped: {framework_map.get('total_controls_mapped',0)}  gaps: {framework_map.get('total_gaps',0)}  coverage: {framework_map.get('overall_coverage','?')}\n")
    return framework_map


# ── GRC Summary Doc ───────────────────────────────────────────────────────────

def _add_para(doc, text, bold=False, size=10, rgb=None, align=WD_ALIGN_PARAGRAPH.LEFT):
    para = doc.add_paragraph()
    para.alignment = align
    run = para.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Arial"
    if rgb:
        run.font.color.rgb = RGBColor(*rgb)
    return para


def build_grc_summary_doc(policy_data: dict, framework_map: dict) -> tuple[str, bytes]:
    """Build the GRC summary document for handoff to GRC tools (Ostendio, Archer, etc.)"""
    name   = policy_data.get("policy_name",   "Policy")
    number = policy_data.get("policy_number", "SEC-P")
    ver    = policy_data.get("version",       "V1.0")
    fname  = f"{number} {name} {ver}-GRC-Summary.docx"

    doc = DocxDocument()
    for sec in doc.sections:
        sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Pt(72)

    # Header
    _add_para(doc, "MIDNIGHT — GRC COMPLIANCE SUMMARY",
              bold=True, size=9, rgb=(120,120,120), align=WD_ALIGN_PARAGRAPH.CENTER)
    _add_para(doc, "")
    _add_para(doc, name, bold=True, size=18)
    _add_para(doc, f"{number}  ·  {ver}  ·  Framework Compliance Report",
              size=10, rgb=(80,110,150))
    _add_para(doc, "")

    # Audit summary
    summary = framework_map.get("audit_summary", "")
    if summary:
        _add_para(doc, "AUDIT SUMMARY", bold=True, size=10, rgb=(0,140,170))
        _add_para(doc, summary, size=10)
        _add_para(doc, "")

    # Coverage overview table
    _add_para(doc, "COVERAGE OVERVIEW", bold=True, size=10, rgb=(0,140,170))
    t = doc.add_table(rows=4, cols=2)
    t.style = "Table Grid"
    overview = [
        ("Overall Coverage",    framework_map.get("overall_coverage","—").upper()),
        ("Controls Mapped",     str(framework_map.get("total_controls_mapped", 0))),
        ("Gaps Identified",     str(framework_map.get("total_gaps", 0))),
        ("Frameworks Assessed", ", ".join(framework_map.get("frameworks_covered",[]))
                                 or "HIPAA, HiTrust, PCI DSS, ISO 27001, NIST CSF, CoBIT"),
    ]
    for i, (lbl, val) in enumerate(overview):
        t.rows[i].cells[0].text = lbl
        t.rows[i].cells[1].text = val
        for cell in t.rows[i].cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(9)
                    run.font.name = "Arial"
        t.rows[i].cells[0].paragraphs[0].runs[0].bold = True
    _add_para(doc, "")

    # Mapped citations
    citations = framework_map.get("mapped_citations", [])
    if citations:
        _add_para(doc, "CONTROLS COVERED", bold=True, size=10, rgb=(0,140,170))
        t2 = doc.add_table(rows=1 + len(citations), cols=4)
        t2.style = "Table Grid"
        for i, h in enumerate(["Framework", "Control ID", "Control Name", "Policy Section"]):
            t2.rows[0].cells[i].text = h
            t2.rows[0].cells[i].paragraphs[0].runs[0].bold = True
            t2.rows[0].cells[i].paragraphs[0].runs[0].font.size = Pt(9)
        for ri, c in enumerate(citations, 1):
            vals = [c.get("framework",""), c.get("control_id",""),
                    c.get("control_name",""), c.get("policy_section","")]
            for ci, v in enumerate(vals):
                cell = t2.rows[ri].cells[ci]
                cell.text = v
                cell.paragraphs[0].runs[0].font.size = Pt(9)
                cell.paragraphs[0].runs[0].font.name = "Arial"
        _add_para(doc, "")

    # Gaps + suggestions
    gaps = framework_map.get("gaps", [])
    if gaps:
        _add_para(doc, "COMPLIANCE GAPS — ACTION REQUIRED", bold=True, size=10, rgb=(190,50,50))
        for i, gap in enumerate(gaps, 1):
            risk = gap.get("risk_level","medium").upper()
            _add_para(doc, f"Gap {i} — {gap.get('framework','')} {gap.get('control_id','')}",
                      bold=True, size=10)
            t3 = doc.add_table(rows=4, cols=2)
            t3.style = "Table Grid"
            gap_rows = [
                ("Control",    f"{gap.get('control_id','')} — {gap.get('control_name','')}"),
                ("Risk",       risk),
                ("Gap",        gap.get("gap_description","")),
                ("Suggestion", gap.get("suggestion","")),
            ]
            for ri, (lbl, val) in enumerate(gap_rows):
                t3.rows[ri].cells[0].text = lbl
                t3.rows[ri].cells[1].text = val
                for cell in t3.rows[ri].cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.font.size = Pt(9)
                            run.font.name = "Arial"
                t3.rows[ri].cells[0].paragraphs[0].runs[0].bold = True
            _add_para(doc, "")

    # Footer
    _add_para(doc, "Generated by Midnight · Takeoff · For internal use only.",
              size=8, rgb=(150,150,150), align=WD_ALIGN_PARAGRAPH.CENTER)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp_path = tmp.name
    try:
        doc.save(tmp_path)
        with open(tmp_path, "rb") as f:
            doc_bytes = f.read()
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass

    return fname, doc_bytes


# ── Logo ──────────────────────────────────────────────────────────────────────

def save_logo_bytes(filename: str, file_bytes: bytes) -> str:
    tmp_dir = os.path.join(tempfile.gettempdir(), "midnight_logos")
    os.makedirs(tmp_dir, exist_ok=True)
    ext  = Path(filename).suffix.lower() or ".png"
    stem = "".join(c for c in Path(filename).stem if c.isalnum() or c in "-_") or "logo"
    path = os.path.join(tmp_dir, f"{stem}_{uuid.uuid4().hex[:8]}{ext}")
    with open(path, "wb") as f:
        f.write(file_bytes)
    return path


# ── Policy Doc Builder ────────────────────────────────────────────────────────

def build_output_doc(policy_data: dict, logo_path: str | None = None) -> tuple[str, bytes]:
    """
    Route to the correct template builder based on policy_data["template_name"].

    Routing logic:
      "Wipro HealthPlan Services"  → hps_policy_migration_builder (legacy, client-specific)
      Everything else              → templates/template_generic.py (generic minimal)
      Future: "Healthcare" etc.   → templates/template_healthcare.py etc.
    """
    name          = policy_data.get("policy_name",    "Policy")
    number        = policy_data.get("policy_number",  "POL")
    ver           = policy_data.get("version",        "V1.0")
    template_name = policy_data.get("template_name",  "Generic Policy Template")
    logo_pos      = policy_data.get("logo_position",  "left")
    fname         = f"{number} {name} {ver}-NEW.docx"

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp_path = tmp.name

    try:
        registry_key = TEMPLATE_REGISTRY.get(template_name, "_generic")

        if template_name == "Wipro HealthPlan Services":
            # Legacy HPS builder — kept for existing client
            build_policy_document(policy_data, tmp_path, logo_path=logo_path)

        elif registry_key == "_generic" and _TEMPLATES_AVAILABLE and _build_generic:
            # Generic minimal template
            _build_generic(
                policy_data,
                tmp_path,
                logo_path=logo_path,
                logo_position=logo_pos,
            )

        else:
            # Fallback to HPS builder if template module not found
            print(f"[WARN] Template '{template_name}' not found in registry — using HPS builder")
            build_policy_document(policy_data, tmp_path, logo_path=logo_path)

        with open(tmp_path, "rb") as f:
            doc_bytes = f.read()

    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass

    print(f"BUILD  |  template: {template_name}  |  file: {fname}")
    return fname, doc_bytes
