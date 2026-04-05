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

from __future__ import annotations

import os
import io
import re
import json
import uuid
import tempfile
from pathlib import Path
from typing import Any

from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from groq import Groq

from hps_policy_migration_builder import build_policy_document

# ── Template registry ─────────────────────────────────────────────────────────

try:
    from templates.template_generic import build_document as _build_generic
    _TEMPLATES_AVAILABLE = True
except ImportError:
    _TEMPLATES_AVAILABLE = False
    _build_generic = None

TEMPLATE_REGISTRY = {
    "Generic Policy Template": "_generic",
    "Generic": "_generic",
    "Enterprise Policy Template": "_generic",
}

# ── Runtime config ────────────────────────────────────────────────────────────

GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")
GROQ_MODEL = "llama-3.3-70b-versatile"
MIN_TEXT_LEN = 50
MAX_SOURCE_LEN = 24000
MAX_POLICY_ANALYSIS_LEN = 16000

# ── Prompts ──────────────────────────────────────────────────────────────────

EXTRACTION_PROMPT = """
You are a policy migration specialist.

Read the attached legacy policy document and extract ALL content into the exact
JSON structure below.

STRICT RULES:
- Do NOT summarize, rewrite, or remove content
- Preserve source wording as closely as possible
- Fix only minor spacing / punctuation / obvious grammar issues
- Map all content into the correct field
- If a field is missing in the source, return an empty string, empty object, or empty list
- Preserve revision history if present
- Preserve related policies and citations if present
- Preserve definitions if present
- Preserve procedures in the correct order
- For procedure items use exactly one type:
  "para"            = standalone paragraph
  "heading"         = bold underlined subsection title
  "bullet"          = first-level bullet
  "sub-bullet"      = second-level bullet
  "bold_intro"      = paragraph starting with bold label; keys: "bold" and "rest"
  "bold_intro_semi" = same as bold_intro but "rest" contains semicolons
  "empty"           = blank spacer line

OUTPUT RULES:
- Return ONLY valid JSON
- No explanation
- No markdown fences
- No comments
- No preamble
- No trailing commas
- Use double quotes for all keys and string values
- Escape quotes inside strings
- Start with {
- End with }

JSON SCHEMA:
{
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
    "hps_inc": false,
    "agency": true,
    "corporate": true,
    "govt_affairs": false,
    "legal_review": false
  },
  "policy_types": {
    "carrier_specific": false,
    "cross_carrier": false,
    "global": false,
    "on_off_hix": false
  },
  "line_of_business": {
    "all_lobs": true,
    "specific_lob": "",
    "specific_lob_checked": false
  },
  "purpose": "",
  "definitions": {},
  "policy_statement": "",
  "procedures": [],
  "related_policies": [],
  "citations": [],
  "revision_history": []
}

REVISION HISTORY FORMAT:
[
  {
    "date": "",
    "version": "",
    "updated_by": "",
    "description": ""
  }
]

PROCEDURES FORMAT:
[
  {
    "type": "para",
    "text": ""
  },
  {
    "type": "heading",
    "text": ""
  },
  {
    "type": "bullet",
    "text": ""
  },
  {
    "type": "sub-bullet",
    "text": ""
  },
  {
    "type": "bold_intro",
    "bold": "",
    "rest": ""
  },
  {
    "type": "bold_intro_semi",
    "bold": "",
    "rest": ""
  },
  {
    "type": "empty"
  }
]

HERE IS THE LEGACY POLICY DOCUMENT:
"""

JSON_REPAIR_PROMPT = """
You are repairing malformed JSON from a policy extraction pipeline.

Your job:
- Convert the input into valid JSON only
- Preserve content exactly
- Do NOT summarize
- Do NOT remove fields
- Do NOT add explanation
- Escape inner quotes correctly
- Fix commas, brackets, and broken strings only
- Return only one valid JSON object
- Start with {
- End with }

BROKEN JSON:
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

Return ONLY valid JSON.
- No explanation
- No markdown
- No preamble
- No comments
- No trailing commas

Start with { and end with }

{
  "policy_name": "",
  "policy_type": "",
  "overall_coverage": "strong|moderate|weak|unknown",
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

# ── Extraction helpers ───────────────────────────────────────────────────────

def _extract_docx_bytes(file_bytes: bytes) -> str:
    lines: list[str] = []
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
            row_parts: list[str] = []
            for cell in row.cells:
                cell_text = "\n".join(
                    p.text.strip() for p in cell.paragraphs if p.text.strip()
                ).strip()
                if cell_text:
                    row_parts.append(cell_text)
            if row_parts:
                seen: list[str] = []
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

    text = text.strip()
    print(f"\n{'=' * 60}\nEXTRACTION  |  {filename}  |  {len(text)} chars\n{'=' * 60}\n")
    return text

# ── LLM helpers ───────────────────────────────────────────────────────────────

def _groq_call(prompt: str, content: str, max_tokens: int = 8000) -> str:
    if not GROQ_API_KEY:
        raise ValueError("GROQ_API_KEY not set.")

    client = Groq(api_key=GROQ_API_KEY)
    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": prompt + "\n\n" + content}],
        temperature=0.1,
        max_tokens=max_tokens,
    )
    return (response.choices[0].message.content or "").strip()


def _strip_code_fences(raw: str) -> str:
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json|python)?\s*", "", raw, flags=re.IGNORECASE)
        raw = re.sub(r"\s*```$", "", raw)
    return raw.strip()


def _extract_json_object(raw: str) -> str:
    raw = _strip_code_fences(raw)

    if raw.startswith("POLICY_DATA ="):
        raw = raw.split("=", 1)[1].strip()
    elif raw.startswith("POLICY_DATA="):
        raw = raw.split("=", 1)[1].strip()

    raw = (raw
           .replace("\u201c", '"').replace("\u201d", '"')
           .replace("\u2018", "'").replace("\u2019", "'"))

    first = raw.find("{")
    last = raw.rfind("}")
    if first == -1 or last == -1 or last < first:
        raise ValueError("Model response did not contain a JSON object.")

    return raw[first:last + 1]


def _json_error_context(raw: str, lineno: int | None, window: int = 6) -> str:
    lines = raw.splitlines()
    if not lines:
        return raw[:600]

    line_no = lineno or 1
    start = max(0, line_no - window - 1)
    end = min(len(lines), line_no + window)

    return "\n".join(f"{i + 1}: {lines[i]}" for i in range(start, end))


def _repair_json_via_llm(raw_json: str) -> str:
    repaired = _groq_call(JSON_REPAIR_PROMPT, raw_json, max_tokens=8000)
    return _extract_json_object(repaired)


def _normalize_revision_history(value: Any) -> list[dict[str, str]]:
    if not value:
        return []

    out: list[dict[str, str]] = []

    if isinstance(value, list):
        for entry in value:
            if isinstance(entry, dict):
                out.append({
                    "date": str(entry.get("date", "")).strip(),
                    "version": str(entry.get("version", "")).strip(),
                    "updated_by": str(entry.get("updated_by", "")).strip(),
                    "description": str(entry.get("description", "")).strip(),
                })
            elif isinstance(entry, (list, tuple)):
                padded = list(entry) + ["", "", "", ""]
                out.append({
                    "date": str(padded[0]).strip(),
                    "version": str(padded[1]).strip(),
                    "updated_by": str(padded[2]).strip(),
                    "description": str(padded[3]).strip(),
                })
            else:
                out.append({
                    "date": "",
                    "version": "",
                    "updated_by": "",
                    "description": str(entry).strip(),
                })

    return out


def _normalize_procedures(value: Any) -> list[dict[str, Any]]:
    allowed = {"para", "heading", "bullet", "sub-bullet", "bold_intro", "bold_intro_semi", "empty"}

    if not isinstance(value, list):
        return []

    normalized: list[dict[str, Any]] = []
    for item in value:
        if isinstance(item, str):
            normalized.append({"type": "para", "text": item.strip()})
            continue

        if not isinstance(item, dict):
            continue

        kind = str(item.get("type", "para")).strip()
        if kind not in allowed:
            kind = "para"

        if kind in {"bold_intro", "bold_intro_semi"}:
            normalized.append({
                "type": kind,
                "bold": str(item.get("bold", "")).strip(),
                "rest": str(item.get("rest", "")).strip(),
            })
        elif kind == "empty":
            normalized.append({"type": "empty"})
        else:
            normalized.append({
                "type": kind,
                "text": str(item.get("text", "")).strip(),
            })

    return normalized


def _coerce_bool(value: Any, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        v = value.strip().lower()
        if v in {"true", "yes", "y", "1", "checked", "x"}:
            return True
        if v in {"false", "no", "n", "0", "unchecked"}:
            return False
    return default


def _validate_policy_data(data: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(data, dict):
        raise ValueError("Parsed POLICY_DATA is not a JSON object.")

    scalar_fields = [
        "policy_name", "policy_number", "version", "grc_id", "supersedes",
        "effective_date", "last_reviewed", "last_revised", "custodians",
        "owner_name", "owner_title", "approver_name", "approver_title",
        "date_signed", "date_approved", "purpose", "policy_statement",
    ]
    for key in scalar_fields:
        data[key] = str(data.get(key, "") or "").strip()

    app = data.get("applicable_to", {}) if isinstance(data.get("applicable_to"), dict) else {}
    data["applicable_to"] = {
        "hps_inc": _coerce_bool(app.get("hps_inc"), False),
        "agency": _coerce_bool(app.get("agency"), True),
        "corporate": _coerce_bool(app.get("corporate"), True),
        "govt_affairs": _coerce_bool(app.get("govt_affairs"), False),
        "legal_review": _coerce_bool(app.get("legal_review"), False),
    }

    pol = data.get("policy_types", {}) if isinstance(data.get("policy_types"), dict) else {}
    data["policy_types"] = {
        "carrier_specific": _coerce_bool(pol.get("carrier_specific"), False),
        "cross_carrier": _coerce_bool(pol.get("cross_carrier"), False),
        "global": _coerce_bool(pol.get("global"), False),
        "on_off_hix": _coerce_bool(pol.get("on_off_hix"), False),
    }

    lob = data.get("line_of_business", {}) if isinstance(data.get("line_of_business"), dict) else {}
    data["line_of_business"] = {
        "all_lobs": _coerce_bool(lob.get("all_lobs"), True),
        "specific_lob": str(lob.get("specific_lob", "") or "").strip(),
        "specific_lob_checked": _coerce_bool(lob.get("specific_lob_checked"), False),
    }

    defs = data.get("definitions", {})
    if isinstance(defs, dict):
        data["definitions"] = {
            str(k).strip(): str(v).strip()
            for k, v in defs.items()
            if str(k).strip()
        }
    else:
        data["definitions"] = {}

    for key in ("related_policies", "citations"):
        value = data.get(key, [])
        if isinstance(value, list):
            data[key] = [str(x).strip() for x in value if str(x).strip()]
        else:
            data[key] = []

    data["procedures"] = _normalize_procedures(data.get("procedures", []))
    data["revision_history"] = _normalize_revision_history(data.get("revision_history", []))

    required = ["policy_name", "policy_number", "version"]
    missing = [k for k in required if not data.get(k)]
    if missing:
        raise ValueError(f"Missing required fields after extraction: {', '.join(missing)}")

    return data


def _parse_policy_data(raw: str) -> dict[str, Any]:
    raw_json = _extract_json_object(raw)

    try:
        parsed = json.loads(raw_json)
    except json.JSONDecodeError as e:
        print(f"[WARN] Initial POLICY_DATA JSON parse failed at line {e.lineno}, col {e.colno}")
        print(_json_error_context(raw_json, e.lineno))

        try:
            repaired_json = _repair_json_via_llm(raw_json)
            parsed = json.loads(repaired_json)
            print("[INFO] POLICY_DATA repaired successfully by secondary JSON repair pass.")
        except Exception as repair_error:
            context = _json_error_context(raw_json, e.lineno)
            raise ValueError(
                f"JSON parsing failed at line {e.lineno}, column {e.colno}.\n\n"
                f"Context:\n{context}\n\n"
                f"Original error: {e}\n\n"
                f"Repair attempt failed: {repair_error}"
            )

    return _validate_policy_data(parsed)


def run_llm_transform(source_text: str, template_name: str) -> dict[str, Any]:
    source_text = source_text.strip()
    if len(source_text) < MIN_TEXT_LEN:
        raise ValueError(f"Extracted text too short ({len(source_text)} chars).")

    if len(source_text) > MAX_SOURCE_LEN:
        source_text = source_text[:MAX_SOURCE_LEN]

    try:
        raw = _groq_call(EXTRACTION_PROMPT, source_text, max_tokens=8000)
    except Exception as e:
        raise ValueError(f"Groq extraction failed: {e}")

    policy_data = _parse_policy_data(raw)
    policy_data["template_name"] = template_name
    return policy_data

# ── Framework Mapping ─────────────────────────────────────────────────────────

def _policy_to_text(policy_data: dict[str, Any]) -> str:
    lines: list[str] = []
    lines.append(f"POLICY: {policy_data.get('policy_name', '')} ({policy_data.get('policy_number', '')})")
    lines.append(f"VERSION: {policy_data.get('version', '')}")
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
                lines.append(f"\n  [{item.get('text', '')}]")
            elif kind == "bullet":
                lines.append(f"  • {item.get('text', '')}")
            elif kind == "sub-bullet":
                lines.append(f"    ◦ {item.get('text', '')}")
            elif kind in ("bold_intro", "bold_intro_semi"):
                lines.append(f"  {item.get('bold', '')} {item.get('rest', '')}")
            elif kind == "para":
                lines.append(f"  {item.get('text', '')}")
        lines.append("")

    existing = policy_data.get("citations", [])
    if existing:
        lines.append("EXISTING CITATIONS:")
        for c in existing:
            lines.append(f"  {c}")

    return "\n".join(lines)


def _empty_map(policy_data: dict[str, Any]) -> dict[str, Any]:
    return {
        "policy_name": policy_data.get("policy_name", ""),
        "policy_type": "",
        "overall_coverage": "unknown",
        "mapped_citations": [],
        "gaps": [],
        "audit_summary": "Framework mapping unavailable for this policy.",
        "frameworks_covered": [],
        "total_controls_mapped": 0,
        "total_gaps": 0,
    }


def _parse_framework_map(raw: str, policy_data: dict[str, Any]) -> dict[str, Any]:
    raw = _strip_code_fences(raw)

    first = raw.find("{")
    last = raw.rfind("}")
    if first == -1 or last == -1 or last < first:
        return _empty_map(policy_data)

    raw_json = raw[first:last + 1]

    try:
        framework_map = json.loads(raw_json)
    except json.JSONDecodeError as e:
        context = _json_error_context(raw_json, e.lineno)
        print(f"[WARN] Framework JSON parse failed at line {e.lineno}, col {e.colno}\n{context}")
        return _empty_map(policy_data)

    if not isinstance(framework_map, dict):
        return _empty_map(policy_data)

    framework_map["policy_name"] = str(framework_map.get("policy_name", policy_data.get("policy_name", "")))
    framework_map["policy_type"] = str(framework_map.get("policy_type", ""))
    framework_map["overall_coverage"] = str(framework_map.get("overall_coverage", "unknown")).lower()

    if framework_map["overall_coverage"] not in {"strong", "moderate", "weak", "unknown"}:
        framework_map["overall_coverage"] = "unknown"

    mapped = framework_map.get("mapped_citations", [])
    framework_map["mapped_citations"] = mapped if isinstance(mapped, list) else []

    gaps = framework_map.get("gaps", [])
    framework_map["gaps"] = gaps if isinstance(gaps, list) else []

    covered = framework_map.get("frameworks_covered", [])
    framework_map["frameworks_covered"] = covered if isinstance(covered, list) else []

    framework_map["audit_summary"] = str(framework_map.get("audit_summary", "") or "")
    framework_map["total_controls_mapped"] = int(framework_map.get("total_controls_mapped", len(framework_map["mapped_citations"])) or 0)
    framework_map["total_gaps"] = int(framework_map.get("total_gaps", len(framework_map["gaps"])) or 0)

    return framework_map


def run_framework_mapping(policy_data: dict[str, Any]) -> dict[str, Any]:
    policy_text = _policy_to_text(policy_data)
    if len(policy_text) > MAX_POLICY_ANALYSIS_LEN:
        policy_text = policy_text[:MAX_POLICY_ANALYSIS_LEN]

    print(
        f"\n{'=' * 60}\nFRAMEWORK   |  "
        f"{policy_data.get('policy_name', '')}  |  {len(policy_text)} chars\n{'=' * 60}\n"
    )

    try:
        raw = _groq_call(FRAMEWORK_MAPPING_PROMPT, policy_text, max_tokens=6000)
    except Exception as e:
        print(f"[WARN] Framework mapping failed: {e}")
        return _empty_map(policy_data)

    framework_map = _parse_framework_map(raw, policy_data)

    print(
        "FRAMEWORK   |  "
        f"mapped: {framework_map.get('total_controls_mapped', 0)}  "
        f"gaps: {framework_map.get('total_gaps', 0)}  "
        f"coverage: {framework_map.get('overall_coverage', '?')}\n"
    )
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


def build_grc_summary_doc(policy_data: dict[str, Any], framework_map: dict[str, Any]) -> tuple[str, bytes]:
    name = policy_data.get("policy_name", "Policy")
    number = policy_data.get("policy_number", "SEC-P")
    ver = policy_data.get("version", "V1.0")
    fname = f"{number} {name} {ver}-GRC-Summary.docx"

    doc = DocxDocument()
    for sec in doc.sections:
        sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Pt(72)

    _add_para(
        doc,
        "MIDNIGHT — GRC COMPLIANCE SUMMARY",
        bold=True,
        size=9,
        rgb=(120, 120, 120),
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    _add_para(doc, "")
    _add_para(doc, name, bold=True, size=18)
    _add_para(doc, f"{number}  ·  {ver}  ·  Framework Compliance Report", size=10, rgb=(80, 110, 150))
    _add_para(doc, "")

    summary = framework_map.get("audit_summary", "")
    if summary:
        _add_para(doc, "AUDIT SUMMARY", bold=True, size=10, rgb=(0, 140, 170))
        _add_para(doc, summary, size=10)
        _add_para(doc, "")

    _add_para(doc, "COVERAGE OVERVIEW", bold=True, size=10, rgb=(0, 140, 170))
    t = doc.add_table(rows=4, cols=2)
    t.style = "Table Grid"
    overview = [
        ("Overall Coverage", framework_map.get("overall_coverage", "—").upper()),
        ("Controls Mapped", str(framework_map.get("total_controls_mapped", 0))),
        ("Gaps Identified", str(framework_map.get("total_gaps", 0))),
        (
            "Frameworks Assessed",
            ", ".join(framework_map.get("frameworks_covered", []))
            or "HIPAA, HiTrust, PCI DSS, ISO 27001, NIST CSF, CoBIT",
        ),
    ]
    for i, (lbl, val) in enumerate(overview):
        t.rows[i].cells[0].text = lbl
        t.rows[i].cells[1].text = val
        for cell in t.rows[i].cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(9)
                    run.font.name = "Arial"
        if t.rows[i].cells[0].paragraphs[0].runs:
            t.rows[i].cells[0].paragraphs[0].runs[0].bold = True
    _add_para(doc, "")

    citations = framework_map.get("mapped_citations", [])
    if citations:
        _add_para(doc, "CONTROLS COVERED", bold=True, size=10, rgb=(0, 140, 170))
        t2 = doc.add_table(rows=1 + len(citations), cols=4)
        t2.style = "Table Grid"
        headers = ["Framework", "Control ID", "Control Name", "Policy Section"]
        for i, h in enumerate(headers):
            t2.rows[0].cells[i].text = h
            if t2.rows[0].cells[i].paragraphs[0].runs:
                t2.rows[0].cells[i].paragraphs[0].runs[0].bold = True
                t2.rows[0].cells[i].paragraphs[0].runs[0].font.size = Pt(9)

        for ri, c in enumerate(citations, 1):
            vals = [
                str(c.get("framework", "")),
                str(c.get("control_id", "")),
                str(c.get("control_name", "")),
                str(c.get("policy_section", "")),
            ]
            for ci, v in enumerate(vals):
                cell = t2.rows[ri].cells[ci]
                cell.text = v
                if cell.paragraphs[0].runs:
                    cell.paragraphs[0].runs[0].font.size = Pt(9)
                    cell.paragraphs[0].runs[0].font.name = "Arial"
        _add_para(doc, "")

    gaps = framework_map.get("gaps", [])
    if gaps:
        _add_para(doc, "COMPLIANCE GAPS — ACTION REQUIRED", bold=True, size=10, rgb=(190, 50, 50))
        for i, gap in enumerate(gaps, 1):
            risk = str(gap.get("risk_level", "medium")).upper()
            _add_para(
                doc,
                f"Gap {i} — {gap.get('framework', '')} {gap.get('control_id', '')}",
                bold=True,
                size=10,
            )
            t3 = doc.add_table(rows=4, cols=2)
            t3.style = "Table Grid"
            gap_rows = [
                ("Control", f"{gap.get('control_id', '')} — {gap.get('control_name', '')}"),
                ("Risk", risk),
                ("Gap", str(gap.get("gap_description", ""))),
                ("Suggestion", str(gap.get("suggestion", ""))),
            ]
            for ri, (lbl, val) in enumerate(gap_rows):
                t3.rows[ri].cells[0].text = lbl
                t3.rows[ri].cells[1].text = val
                for cell in t3.rows[ri].cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.font.size = Pt(9)
                            run.font.name = "Arial"
                if t3.rows[ri].cells[0].paragraphs[0].runs:
                    t3.rows[ri].cells[0].paragraphs[0].runs[0].bold = True
            _add_para(doc, "")

    _add_para(
        doc,
        "Generated by Midnight · Takeoff · For internal use only.",
        size=8,
        rgb=(150, 150, 150),
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )

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
    ext = Path(filename).suffix.lower() or ".png"
    stem = "".join(c for c in Path(filename).stem if c.isalnum() or c in "-_") or "logo"
    path = os.path.join(tmp_dir, f"{stem}_{uuid.uuid4().hex[:8]}{ext}")
    with open(path, "wb") as f:
        f.write(file_bytes)
    return path

# ── Policy Doc Builder ────────────────────────────────────────────────────────

def build_output_doc(policy_data: dict[str, Any], logo_path: str | None = None) -> tuple[str, bytes]:
    name = policy_data.get("policy_name", "Policy")
    number = policy_data.get("policy_number", "POL")
    ver = policy_data.get("version", "V1.0")
    template_name = policy_data.get("template_name", "Generic Policy Template")
    logo_pos = policy_data.get("logo_position", "left")
    fname = f"{number} {name} {ver}-NEW.docx"

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp_path = tmp.name

    try:
        registry_key = TEMPLATE_REGISTRY.get(template_name, "_generic")

        if template_name == "Wipro HealthPlan Services":
            build_policy_document(policy_data, tmp_path, logo_path=logo_path)

        elif registry_key == "_generic" and _TEMPLATES_AVAILABLE and _build_generic:
            _build_generic(
                policy_data,
                tmp_path,
                logo_path=logo_path,
                logo_position=logo_pos,
            )

        else:
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
