# ── GRC Summary Doc — PDF ────────────────────────────────────────────────────

def build_grc_summary_doc(
    policy_data: dict[str, Any],
    framework_map: dict[str, Any],
) -> tuple[str, bytes]:
    name = _clean_scalar(policy_data.get("policy_name", "Policy")) or "Policy"
    number = _clean_scalar(policy_data.get("policy_number", "SEC-P")) or "SEC-P"
    ver = _clean_scalar(policy_data.get("version", "V1.0")) or "V1.0"
    fname = f"{number} {name} {ver}-GRC-Summary.pdf"

    print(f"GRC PDF  |  {name}  |  {number}  |  {ver}")

    if not isinstance(framework_map, dict):
        framework_map = {}

    safe_framework_map = {
        "policy_name": _clean_scalar(framework_map.get("policy_name", policy_data.get("policy_name", ""))),
        "policy_type": _clean_scalar(framework_map.get("policy_type", "")),
        "overall_coverage": _clean_scalar(framework_map.get("overall_coverage", "unknown")).lower(),
        "mapped_citations": [],
        "gaps": [],
        "audit_summary": _clean_scalar(framework_map.get("audit_summary", "")),
        "frameworks_covered": _normalize_string_list(framework_map.get("frameworks_covered", [])),
        "total_controls_mapped": 0,
        "total_gaps": 0,
        "_mapping_failed": bool(framework_map.get("_mapping_failed", False)),
        "_mapping_failure_reason": _clean_scalar(framework_map.get("_mapping_failure_reason", "")),
    }

    if safe_framework_map["overall_coverage"] not in {"strong", "moderate", "weak", "unknown"}:
        safe_framework_map["overall_coverage"] = "unknown"

    for entry in framework_map.get("mapped_citations", []):
        if not isinstance(entry, dict):
            continue
        safe_framework_map["mapped_citations"].append({
            "framework": _clean_scalar(entry.get("framework", "")),
            "control_id": _clean_scalar(entry.get("control_id", "")),
            "control_name": _clean_scalar(entry.get("control_name", "")),
            "policy_section": _clean_scalar(entry.get("policy_section", "")),
            "coverage_note": _clean_scalar(entry.get("coverage_note", "")),
        })

    for entry in framework_map.get("gaps", []):
        if not isinstance(entry, dict):
            continue

        risk_level = _clean_scalar(entry.get("risk_level", "medium")).lower()
        if risk_level not in {"high", "medium", "low"}:
            risk_level = "medium"

        safe_framework_map["gaps"].append({
            "framework": _clean_scalar(entry.get("framework", "")),
            "control_id": _clean_scalar(entry.get("control_id", "")),
            "control_name": _clean_scalar(entry.get("control_name", "")),
            "gap_description": _clean_scalar(entry.get("gap_description", "")),
            "risk_level": risk_level,
            "suggestion": _clean_scalar(entry.get("suggestion", "")),
        })

    safe_framework_map["total_controls_mapped"] = _safe_int(
        framework_map.get("total_controls_mapped"),
        default=len(safe_framework_map["mapped_citations"]),
    )
    safe_framework_map["total_gaps"] = _safe_int(
        framework_map.get("total_gaps"),
        default=len(safe_framework_map["gaps"]),
    )

    if not safe_framework_map["audit_summary"]:
        safe_framework_map["audit_summary"] = (
            f"{name} was analyzed against the selected frameworks. "
            f"{safe_framework_map['total_controls_mapped']} mapped controls and "
            f"{safe_framework_map['total_gaps']} gaps were identified. "
            "This summary reflects only the reviewed document set."
        )

    print(
        "GRC PDF  |  "
        f"mapped: {safe_framework_map['total_controls_mapped']}  "
        f"gaps: {safe_framework_map['total_gaps']}  "
        f"coverage: {safe_framework_map['overall_coverage']}"
    )

    try:
        from grc_summary_pdf import build_grc_pdf
    except Exception as e:
        raise RuntimeError(f"[IMPORT ERROR] grc_summary_pdf failed: {e}") from e

    try:
        pdf_bytes = build_grc_pdf(policy_data, safe_framework_map)
    except Exception as e:
        raise RuntimeError(f"[PDF ERROR] GRC generation failed: {e}") from e

    if not isinstance(pdf_bytes, (bytes, bytearray)):
        raise RuntimeError(f"[PDF ERROR] Expected bytes, got {type(pdf_bytes).__name__}")

    pdf_bytes = bytes(pdf_bytes)

    if not pdf_bytes:
        raise RuntimeError("[PDF ERROR] GRC generation returned empty bytes.")

    if not pdf_bytes.startswith(b"%PDF"):
        raise RuntimeError("[PDF ERROR] Returned content is not a PDF.")

    return fname, pdf_bytes
