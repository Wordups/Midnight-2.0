"""
grc_summary_pdf.py
Midnight GRC Compliance Summary PDF Builder
Pure renderer only. No sample data. No local test fixture.
"""

import io
from typing import Any, Dict, List

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    HRFlowable,
    Flowable,
)

# =============================================================================
# BRAND COLORS
# =============================================================================

DARK = HexColor("#050a12")
DARK2 = HexColor("#080e18")
CYAN = HexColor("#00d4f5")
CYAN_DIM = HexColor("#00a8c8")
WHITE = HexColor("#ffffff")

T1 = HexColor("#dde8f7")
T2 = HexColor("#7d95b5")
T3 = HexColor("#3d5470")

GREEN = HexColor("#00e89a")
RED = HexColor("#c00020")
ORANGE = HexColor("#ff6b35")

LGRAY = HexColor("#f5f7fa")
MGRAY = HexColor("#e2e8f0")
BGRAY = HexColor("#1a2640")

BODY_TEXT = HexColor("#334155")
BODY_TEXT_DARK = HexColor("#1e293b")
SUGGESTION_TEXT = HexColor("#1e3a5f")
ROW_ALT = HexColor("#f8fafc")
SUMMARY_BG = HexColor("#f0f9ff")
RISK_HIGH_BG = HexColor("#fff1f2")
RISK_MEDIUM_BG = HexColor("#fff7ed")
RISK_LOW_BG = HexColor("#f0f9ff")


# =============================================================================
# STYLES
# =============================================================================

def make_styles() -> Dict[str, ParagraphStyle]:
    return {
        "section_label": ParagraphStyle(
            "section_label",
            fontName="Helvetica-Bold",
            fontSize=8,
            textColor=CYAN,
            leading=10,
            spaceBefore=18,
            spaceAfter=6,
        ),
        "section_title": ParagraphStyle(
            "section_title",
            fontName="Helvetica-Bold",
            fontSize=16,
            textColor=DARK,
            leading=20,
            spaceAfter=6,
        ),
        "body": ParagraphStyle(
            "body",
            fontName="Helvetica",
            fontSize=9,
            textColor=BODY_TEXT,
            leading=14,
            spaceAfter=4,
        ),
        "body_small": ParagraphStyle(
            "body_small",
            fontName="Helvetica",
            fontSize=8,
            textColor=T2,
            leading=12,
            spaceAfter=4,
        ),
        "body_bold": ParagraphStyle(
            "body_bold",
            fontName="Helvetica-Bold",
            fontSize=9,
            textColor=DARK,
            leading=14,
        ),
        "table_header": ParagraphStyle(
            "table_header",
            fontName="Helvetica-Bold",
            fontSize=8,
            textColor=WHITE,
            leading=10,
        ),
        "table_cell": ParagraphStyle(
            "table_cell",
            fontName="Helvetica",
            fontSize=8,
            textColor=BODY_TEXT,
            leading=11,
        ),
        "table_cell_bold": ParagraphStyle(
            "table_cell_bold",
            fontName="Helvetica-Bold",
            fontSize=8,
            textColor=DARK,
            leading=11,
        ),
        "footer": ParagraphStyle(
            "footer",
            fontName="Helvetica",
            fontSize=7,
            textColor=T3,
            leading=9,
            alignment=TA_CENTER,
        ),
    }


# =============================================================================
# HELPERS
# =============================================================================

def _safe_dict(value: Any) -> Dict[str, Any]:
    return value if isinstance(value, dict) else {}


def _safe_list(value: Any) -> List[Any]:
    return value if isinstance(value, list) else []


def _safe_str(value: Any, default: str = "") -> str:
    if value is None:
        return default
    return str(value).strip()


def _safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except Exception:
        return default


def _truncate(text: Any, limit: int = 3000) -> str:
    value = _safe_str(text)
    if len(value) <= limit:
        return value
    return value[: limit - 3].rstrip() + "..."


def _join_frameworks(items: Any) -> str:
    values = _safe_list(items)
    cleaned = [_safe_str(x) for x in values if _safe_str(x)]
    return ", ".join(cleaned) if cleaned else "—"


def _coverage_color(coverage: str):
    normalized = coverage.upper()
    if normalized == "STRONG":
        return GREEN
    if normalized == "MODERATE":
        return ORANGE
    return RED


def _risk_palette(risk_level: str):
    risk = _safe_str(risk_level, "medium").upper()
    if risk == "HIGH":
        return risk, RED, RISK_HIGH_BG
    if risk == "LOW":
        return risk, HexColor("#0ea5e9"), RISK_LOW_BG
    return "MEDIUM", ORANGE, RISK_MEDIUM_BG


# =============================================================================
# OPTIONAL FLOWABLE
# =============================================================================

class ColorRect(Flowable):
    def __init__(self, width, height, color, radius=0):
        super().__init__()
        self.width = width
        self.height = height
        self.color = color
        self.radius = radius

    def draw(self):
        self.canv.setFillColor(self.color)
        if self.radius:
            self.canv.roundRect(0, 0, self.width, self.height, self.radius, fill=1, stroke=0)
        else:
            self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


# =============================================================================
# HEADER / FOOTER
# =============================================================================

def header_footer(canvas, doc, policy_name: str, policy_number: str, version: str):
    width, height = letter
    canvas.saveState()

    canvas.setFillColor(DARK)
    canvas.rect(0, height - 0.55 * inch, width, 0.55 * inch, fill=1, stroke=0)

    canvas.setFillColor(CYAN)
    canvas.rect(0, height - 0.55 * inch, 4, 0.55 * inch, fill=1, stroke=0)

    canvas.setFont("Helvetica-Bold", 7)
    canvas.setFillColor(WHITE)
    canvas.drawString(0.35 * inch, height - 0.32 * inch, "MIDNIGHT")

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(T2)
    canvas.drawString(
        1.1 * inch,
        height - 0.32 * inch,
        f"GRC COMPLIANCE SUMMARY  ·  {policy_number}  ·  {version}",
    )

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(T3)
    canvas.drawRightString(width - 0.35 * inch, height - 0.32 * inch, "CONFIDENTIAL")

    canvas.setFillColor(LGRAY)
    canvas.rect(0, 0, width, 0.4 * inch, fill=1, stroke=0)

    canvas.setFillColor(CYAN)
    canvas.rect(0, 0.38 * inch, width, 1.5, fill=1, stroke=0)

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(T2)
    canvas.drawString(
        0.35 * inch,
        0.15 * inch,
        "Generated by Midnight · Takeoff LLC · For internal use only.",
    )

    canvas.setFont("Helvetica-Bold", 7)
    canvas.setFillColor(T2)
    canvas.drawRightString(width - 0.35 * inch, 0.15 * inch, f"Page {doc.page}")

    canvas.restoreState()


# =============================================================================
# MAIN BUILDER
# =============================================================================

def build_grc_pdf(policy_data: dict, framework_map: dict) -> bytes:
    policy_data = _safe_dict(policy_data)
    framework_map = _safe_dict(framework_map)

    styles = make_styles()
    page_width, _ = letter
    content_width = page_width - 1.4 * inch

    buffer = io.BytesIO()

    policy_name = _safe_str(policy_data.get("policy_name"), "Policy")
    policy_number = _safe_str(policy_data.get("policy_number"), "SEC-P")
    version = _safe_str(policy_data.get("version"), "V1.0")
    owner = _safe_str(policy_data.get("owner_name"))
    effective = _safe_str(policy_data.get("effective_date"))

    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=0.7 * inch,
        rightMargin=0.7 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.6 * inch,
        title=f"{policy_name} — GRC Compliance Summary",
        author="Midnight · Takeoff LLC",
        subject="GRC Compliance Summary",
        creator="Midnight Policy Intelligence Platform",
    )

    story = []

    cover_data = [[
        Paragraph(
            "MIDNIGHT",
            ParagraphStyle(
                "cover_brand",
                fontName="Helvetica-Bold",
                fontSize=9,
                textColor=CYAN,
                leading=11,
            ),
        ),
        Paragraph(
            "GRC COMPLIANCE SUMMARY",
            ParagraphStyle(
                "cover_label",
                fontName="Helvetica",
                fontSize=8,
                textColor=T3,
                leading=10,
                alignment=TA_RIGHT,
            ),
        ),
    ]]

    cover_table = Table(cover_data, colWidths=[content_width * 0.5, content_width * 0.5])
    cover_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), DARK),
        ("TOPPADDING", (0, 0), (-1, -1), 14),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
        ("LEFTPADDING", (0, 0), (-1, -1), 14),
        ("RIGHTPADDING", (0, 0), (-1, -1), 14),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 0), (-1, -1), 2, CYAN),
    ]))
    story.append(cover_table)
    story.append(Spacer(1, 14))

    story.append(Paragraph(
        policy_name,
        ParagraphStyle(
            "policy_name",
            fontName="Helvetica-Bold",
            fontSize=22,
            textColor=DARK,
            leading=26,
            spaceAfter=4,
        ),
    ))

    meta_parts = [policy_number, version]
    if effective:
        meta_parts.append(f"Effective {effective}")
    meta_parts.append("Framework Compliance Report")

    story.append(Paragraph(
        "  ·  ".join(meta_parts),
        ParagraphStyle(
            "policy_meta",
            fontName="Helvetica",
            fontSize=9,
            textColor=CYAN_DIM,
            leading=12,
            spaceAfter=2,
        ),
    ))

    if owner:
        story.append(Paragraph(
            f"Policy Owner: {owner}",
            ParagraphStyle(
                "policy_owner",
                fontName="Helvetica",
                fontSize=8,
                textColor=T2,
                leading=11,
                spaceAfter=12,
            ),
        ))

    story.append(HRFlowable(
        width=content_width,
        thickness=1,
        color=MGRAY,
        spaceAfter=16,
    ))

    audit_summary = _truncate(framework_map.get("audit_summary"), 4000)
    if audit_summary:
        summary_table = Table(
            [
                [Paragraph(
                    "AUDIT SUMMARY",
                    ParagraphStyle(
                        "summary_label",
                        fontName="Helvetica-Bold",
                        fontSize=7,
                        textColor=CYAN,
                        leading=9,
                    ),
                )],
                [Paragraph(
                    audit_summary,
                    ParagraphStyle(
                        "summary_body",
                        fontName="Helvetica",
                        fontSize=9,
                        textColor=BODY_TEXT_DARK,
                        leading=14,
                    ),
                )],
            ],
            colWidths=[content_width],
        )

        summary_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), SUMMARY_BG),
            ("LINEBEFORE", (0, 0), (0, -1), 3, CYAN),
            ("LEFTPADDING", (0, 0), (-1, -1), 14),
            ("RIGHTPADDING", (0, 0), (-1, -1), 14),
            ("TOPPADDING", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ]))
        story.append(summary_table)
        story.append(Spacer(1, 16))

    story.append(Paragraph("COVERAGE OVERVIEW", styles["section_label"]))

    coverage = _safe_str(framework_map.get("overall_coverage"), "unknown").upper()
    coverage_color = _coverage_color(coverage)
    total_controls_mapped = _safe_int(framework_map.get("total_controls_mapped"), 0)
    total_gaps = _safe_int(framework_map.get("total_gaps"), 0)
    frameworks = _join_frameworks(framework_map.get("frameworks_covered"))

    overview_rows = [
        [
            Paragraph("METRIC", styles["table_header"]),
            Paragraph("VALUE", styles["table_header"]),
            Paragraph("DETAIL", styles["table_header"]),
        ],
        [
            Paragraph("Overall Coverage", styles["table_cell_bold"]),
            Paragraph(
                coverage,
                ParagraphStyle(
                    "coverage_value",
                    fontName="Helvetica-Bold",
                    fontSize=9,
                    textColor=coverage_color,
                    leading=11,
                ),
            ),
            Paragraph("Based on mapped controls vs expected controls", styles["table_cell"]),
        ],
        [
            Paragraph("Controls Mapped", styles["table_cell_bold"]),
            Paragraph(str(total_controls_mapped), styles["table_cell"]),
            Paragraph("Framework controls satisfied by this policy", styles["table_cell"]),
        ],
        [
            Paragraph("Gaps Identified", styles["table_cell_bold"]),
            Paragraph(
                str(total_gaps),
                ParagraphStyle(
                    "gaps_value",
                    fontName="Helvetica-Bold",
                    fontSize=9,
                    textColor=RED if total_gaps > 0 else GREEN,
                    leading=11,
                ),
            ),
            Paragraph("Controls required but not addressed", styles["table_cell"]),
        ],
        [
            Paragraph("Frameworks Assessed", styles["table_cell_bold"]),
            Paragraph("", styles["table_cell"]),
            Paragraph(frameworks, styles["table_cell"]),
        ],
    ]

    overview_table = Table(
        overview_rows,
        colWidths=[content_width * 0.28, content_width * 0.17, content_width * 0.55],
    )
    overview_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), DARK),
        ("BACKGROUND", (0, 1), (-1, 1), ROW_ALT),
        ("BACKGROUND", (0, 2), (-1, 2), WHITE),
        ("BACKGROUND", (0, 3), (-1, 3), ROW_ALT),
        ("BACKGROUND", (0, 4), (-1, 4), WHITE),
        ("GRID", (0, 0), (-1, -1), 0.5, MGRAY),
        ("LINEBELOW", (0, 0), (-1, 0), 1.5, CYAN),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(overview_table)
    story.append(Spacer(1, 20))

    citations = _safe_list(framework_map.get("mapped_citations"))
    if citations:
        story.append(Paragraph("CONTROLS COVERED", styles["section_label"]))

        citation_rows = [[
            Paragraph("FRAMEWORK", styles["table_header"]),
            Paragraph("CONTROL ID", styles["table_header"]),
            Paragraph("CONTROL NAME", styles["table_header"]),
            Paragraph("POLICY SECTION", styles["table_header"]),
        ]]

        for item in citations:
            item = _safe_dict(item)
            citation_rows.append([
                Paragraph(_truncate(item.get("framework"), 120), styles["table_cell"]),
                Paragraph(_truncate(item.get("control_id"), 60), styles["table_cell_bold"]),
                Paragraph(_truncate(item.get("control_name"), 250), styles["table_cell"]),
                Paragraph(_truncate(item.get("policy_section"), 180), styles["table_cell"]),
            ])

        citation_table = Table(
            citation_rows,
            colWidths=[
                content_width * 0.22,
                content_width * 0.14,
                content_width * 0.34,
                content_width * 0.30,
            ],
            repeatRows=1,
        )

        citation_style = [
            ("BACKGROUND", (0, 0), (-1, 0), DARK),
            ("LINEBELOW", (0, 0), (-1, 0), 1.5, CYAN),
            ("GRID", (0, 0), (-1, -1), 0.4, MGRAY),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]

        for row_index in range(1, len(citation_rows)):
            if row_index % 2 == 0:
                citation_style.append(("BACKGROUND", (0, row_index), (-1, row_index), ROW_ALT))

        citation_table.setStyle(TableStyle(citation_style))
        story.append(citation_table)
        story.append(Spacer(1, 20))

    gaps = _safe_list(framework_map.get("gaps"))
    if gaps:
        story.append(Paragraph("COMPLIANCE GAPS — ACTION REQUIRED", styles["section_label"]))
        story.append(Paragraph(
            f"{len(gaps)} gap{'s' if len(gaps) != 1 else ''} identified. "
            "Each gap includes suggested policy language to close the control.",
            ParagraphStyle(
                "gaps_intro",
                fontName="Helvetica",
                fontSize=8,
                textColor=T2,
                leading=12,
                spaceAfter=10,
            ),
        ))

        for idx, gap in enumerate(gaps, start=1):
            gap = _safe_dict(gap)

            risk_label, risk_color, risk_bg = _risk_palette(gap.get("risk_level"))
            framework = _truncate(gap.get("framework"), 100)
            control_id = _truncate(gap.get("control_id"), 80)
            control_name = _truncate(gap.get("control_name"), 300)
            gap_description = _truncate(gap.get("gap_description"), 2200)
            suggestion = _truncate(gap.get("suggestion"), 2200)

            gap_rows = [
                [
                    Paragraph(
                        f"Gap {idx} &nbsp;—&nbsp; {framework} {control_id}",
                        ParagraphStyle(
                            "gap_header_left",
                            fontName="Helvetica-Bold",
                            fontSize=10,
                            textColor=DARK,
                            leading=13,
                        ),
                    ),
                    Paragraph(
                        risk_label,
                        ParagraphStyle(
                            "gap_header_right",
                            fontName="Helvetica-Bold",
                            fontSize=8,
                            textColor=risk_color,
                            leading=10,
                            alignment=TA_RIGHT,
                        ),
                    ),
                ],
                [
                    Paragraph(
                        control_name,
                        ParagraphStyle(
                            "gap_control_name",
                            fontName="Helvetica",
                            fontSize=8,
                            textColor=T2,
                            leading=11,
                        ),
                    ),
                    Paragraph("", styles["table_cell"]),
                ],
                [
                    Paragraph(
                        f"<b>Gap:</b> {gap_description}",
                        ParagraphStyle(
                            "gap_desc",
                            fontName="Helvetica",
                            fontSize=8,
                            textColor=BODY_TEXT,
                            leading=12,
                        ),
                    ),
                    Paragraph("", styles["table_cell"]),
                ],
                [
                    Paragraph(
                        f"<b>Suggested Language:</b> {suggestion}",
                        ParagraphStyle(
                            "gap_suggestion",
                            fontName="Helvetica-Oblique",
                            fontSize=8,
                            textColor=SUGGESTION_TEXT,
                            leading=12,
                        ),
                    ),
                    Paragraph("", styles["table_cell"]),
                ],
            ]

            gap_table = Table(
                gap_rows,
                colWidths=[content_width * 0.84, content_width * 0.16],
            )
            gap_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), risk_bg),
                ("BACKGROUND", (0, 1), (-1, 1), WHITE),
                ("BACKGROUND", (0, 2), (-1, 2), ROW_ALT),
                ("BACKGROUND", (0, 3), (-1, 3), SUMMARY_BG),
                ("BOX", (0, 0), (-1, -1), 1, MGRAY),
                ("LINEBEFORE", (0, 0), (0, -1), 3, risk_color),
                ("LINEBELOW", (0, 0), (-1, 0), 0.5, MGRAY),
                ("LINEBELOW", (0, 1), (-1, 1), 0.5, MGRAY),
                ("LINEBELOW", (0, 2), (-1, 2), 0.5, MGRAY),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]))

            story.append(gap_table)
            story.append(Spacer(1, 10))

    story.append(Spacer(1, 20))

    disclaimer_table = Table(
        [[Paragraph(
            "This report was generated by Midnight, an enterprise policy intelligence platform "
            "by Takeoff LLC. Content is derived from AI-assisted framework analysis and should "
            "be reviewed by a qualified compliance professional before use in formal audit proceedings.",
            ParagraphStyle(
                "disclaimer",
                fontName="Helvetica",
                fontSize=7,
                textColor=T2,
                leading=11,
            ),
        )]],
        colWidths=[content_width],
    )
    disclaimer_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LGRAY),
        ("LINEABOVE", (0, 0), (-1, 0), 1, MGRAY),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))
    story.append(disclaimer_table)

    try:
        doc.build(
            story,
            onFirstPage=lambda canvas, d: header_footer(canvas, d, policy_name, policy_number, version),
            onLaterPages=lambda canvas, d: header_footer(canvas, d, policy_name, policy_number, version),
        )
        pdf_bytes = buffer.getvalue()
        if not pdf_bytes:
            raise RuntimeError("PDF build completed but returned empty content.")
        return pdf_bytes

    except Exception as exc:
        raise RuntimeError(f"GRC PDF build failed: {exc}") from exc

    finally:
        buffer.close()
