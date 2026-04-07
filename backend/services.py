# ── GRC Summary Doc — PDF ────────────────────────────────────────────────────

def _normalize_grc_pdf_payload(
    policy_data: dict[str, Any],
    framework_map: dict[str, Any],
) -> dict[str, Any]:
    """
    Normalize framework_map into a ReportLab-safe payload.
    Prevents PDF crashes caused by None values, non-string fields,
    malformed gap entries, or missing summary text.
    """
    if not isinstance(framework_map, dict):
        framework_map = {}

    mapped_citations: list[dict[str, str]] = []
    for entry in framework_map.get("mapped_citations", []):
        if not isinstance(entry, dict):
            continue
        mapped_citations.append({
            "framework": _clean_scalar(entry.get("framework", "")),
            "control_id": _clean_scalar(entry.get("control_id", "")),
            "control_name": _clean_scalar(entry.get("control_name", "")),
            "policy_section": _clean_scalar(entry.get("policy_section", "")),
            "coverage_note": _clean_scalar(entry.get("coverage_note", "")),
        })

    gaps: list[dict[str, str]] = []
    for entry in framework_map.get("gaps", []):
        if not isinstance(entry, dict):
            continue
        risk_level = _clean_scalar(entry.get("risk_level", "medium")).lower()
        if risk_level not in {"high", "medium", "low"}:
            risk_level = "medium"

        gaps.append({
            "framework": _clean_scalar(entry.get("framework", "")),
            "control_id": _clean_scalar(entry.get("control_id", "")),
            "control_name": _clean_scalar(entry.get("control_name", "")),
            "gap_description": _clean_scalar(entry.get("gap_description", "")),
            "risk_level": risk_level,
            "suggestion": _clean_scalar(entry.get("suggestion", "")),
        })

    frameworks_covered = _normalize_string_list(framework_map.get("frameworks_covered", []))

    overall_coverage = _clean_scalar(framework_map.get("overall_coverage", "unknown")).lower()
    if overall_coverage not in {"strong", "moderate", "weak", "unknown"}:
        overall_coverage = "unknown"

    audit_summary = _clean_scalar(framework_map.get("audit_summary", ""))
    if not audit_summary:
        policy_name = _clean_scalar(policy_data.get("policy_name", "This policy")) or "This policy"
        mapped_count = len(mapped_citations)
        gap_count = len(gaps)
        audit_summary = (
            f"{policy_name} was analyzed against the selected frameworks. "
            f"{mapped_count} mapped controls and {gap_count} gaps were identified. "
            "This summary reflects only the reviewed document set."
        )

    normalized = {
        "policy_name": _clean_scalar(
            framework_map.get("policy_name", policy_data.get("policy_name", ""))
        ),
        "policy_type": _clean_scalar(framework_map.get("policy_type", "")),
        "overall_coverage": overall_coverage,
        "mapped_citations": mapped_citations,
        "gaps": gaps,
        "audit_summary": audit_summary,
        "frameworks_covered": frameworks_covered,
        "total_controls_mapped": _safe_int(
            framework_map.get("total_controls_mapped"),
            default=len(mapped_citations),
        ),
        "total_gaps": _safe_int(
            framework_map.get("total_gaps"),
            default=len(gaps),
        ),
        "_mapping_failed": bool(framework_map.get("_mapping_failed", False)),
        "_mapping_failure_reason": _clean_scalar(
            framework_map.get("_mapping_failure_reason", "")
        ),
    }

    return normalized


def build_grc_summary_doc(
    policy_data: dict[str, Any],
    framework_map: dict[str, Any],
) -> tuple[str, bytes]:
    name = _clean_scalar(policy_data.get("policy_name", "Policy")) or "Policy"
    number = _clean_scalar(policy_data.get("policy_number", "SEC-P")) or "SEC-P"
    ver = _clean_scalar(policy_data.get("version", "V1.0")) or "V1.0"
    fname = f"{number} {name} {ver}-GRC-Summary.pdf"

    print(f"GRC PDF  |  {name}  |  {number}  |  {ver}")

    normalized_framework_map = _normalize_grc_pdf_payload(policy_data, framework_map)

    print(
        "GRC PDF  |  "
        f"mapped: {normalized_framework_map['total_controls_mapped']}  "
        f"gaps: {normalized_framework_map['total_gaps']}  "
        f"coverage: {normalized_framework_map['overall_coverage']}"
    )

    try:
        from grc_summary_pdf import build_grc_pdf
    except Exception as e:
        raise RuntimeError(f"[IMPORT ERROR] grc_summary_pdf failed: {e}") from e

    try:
        pdf_bytes = build_grc_pdf(policy_data, normalized_framework_map)
    except Exception as e:
        raise RuntimeError(
            "[PDF ERROR] GRC generation failed. "
            f"summary_len={len(normalized_framework_map.get('audit_summary', ''))} "
            f"mapped={len(normalized_framework_map.get('mapped_citations', []))} "
            f"gaps={len(normalized_framework_map.get('gaps', []))} "
            f"reason={e}"
        ) from e

    if not isinstance(pdf_bytes, (bytes, bytearray)):
        raise RuntimeError(
            f"[PDF ERROR] GRC generation returned {type(pdf_bytes).__name__}, expected bytes."
        )

    pdf_bytes = bytes(pdf_bytes)
    if not pdf_bytes:
        raise RuntimeError("[PDF ERROR] GRC generation returned empty PDF bytes.")

    # Basic PDF signature check
    if not pdf_bytes.startswith(b"%PDF"):
        raise RuntimeError("[PDF ERROR] Returned bytes are not a valid PDF payload.")

    return fname, pdf_bytes
