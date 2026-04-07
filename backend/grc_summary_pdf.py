"""
grc_summary_pdf.py — Midnight GRC Compliance Summary PDF Builder
Branded, locked, tamper-proof. Dark header, cyan accents.
"""

import io
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, white, black
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether, PageBreak
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import Flowable

# ── Brand colors ──────────────────────────────────────────────────────────────
DARK    = HexColor("#050a12")
DARK2   = HexColor("#080e18")
CYAN    = HexColor("#00d4f5")
CYAN_DIM= HexColor("#00a8c8")
WHITE   = HexColor("#ffffff")
T1      = HexColor("#dde8f7")
T2      = HexColor("#7d95b5")
T3      = HexColor("#3d5470")
GREEN   = HexColor("#00e89a")
RED     = HexColor("#c00020")
ORANGE  = HexColor("#ff6b35")
LGRAY   = HexColor("#f5f7fa")
MGRAY   = HexColor("#e2e8f0")
BGRAY   = HexColor("#1a2640")

# ── Styles ────────────────────────────────────────────────────────────────────
def make_styles():
    return {
        "cover_title": ParagraphStyle(
            "cover_title", fontName="Helvetica-Bold",
            fontSize=32, textColor=WHITE, leading=36,
            spaceAfter=6
        ),
        "cover_sub": ParagraphStyle(
            "cover_sub", fontName="Helvetica",
            fontSize=11, textColor=CYAN, leading=14,
            spaceAfter=4
        ),
        "cover_meta": ParagraphStyle(
            "cover_meta", fontName="Helvetica",
            fontSize=9, textColor=T2, leading=12
        ),
        "section_label": ParagraphStyle(
            "section_label", fontName="Helvetica-Bold",
            fontSize=8, textColor=CYAN, leading=10,
            spaceBefore=18, spaceAfter=6,
            letterSpacing=2
        ),
        "section_title": ParagraphStyle(
            "section_title", fontName="Helvetica-Bold",
            fontSize=16, textColor=DARK, leading=20,
            spaceAfter=6
        ),
        "body": ParagraphStyle(
            "body", fontName="Helvetica",
            fontSize=9, textColor=HexColor("#334155"),
            leading=14, spaceAfter=4
        ),
        "body_bold": ParagraphStyle(
            "body_bold", fontName="Helvetica-Bold",
            fontSize=9, textColor=DARK, leading=14
        ),
        "table_header": ParagraphStyle(
            "table_header", fontName="Helvetica-Bold",
            fontSize=8, textColor=WHITE, leading=10
        ),
        "table_cell": ParagraphStyle(
            "table_cell", fontName="Helvetica",
            fontSize=8, textColor=HexColor("#334155"), leading=11
        ),
        "table_cell_bold": ParagraphStyle(
            "table_cell_bold", fontName="Helvetica-Bold",
            fontSize=8, textColor=DARK, leading=11
        ),
        "gap_title": ParagraphStyle(
            "gap_title", fontName="Helvetica-Bold",
            fontSize=10, textColor=DARK, leading=13,
            spaceBefore=10, spaceAfter=4
        ),
        "suggestion": ParagraphStyle(
            "suggestion", fontName="Helvetica-Oblique",
            fontSize=8, textColor=HexColor("#334155"),
            leading=12, leftIndent=8
        ),
        "footer": ParagraphStyle(
            "footer", fontName="Helvetica",
            fontSize=7, textColor=T3, leading=9,
            alignment=TA_CENTER
        ),
        "mono": ParagraphStyle(
            "mono", fontName="Courier",
            fontSize=8, textColor=T2, leading=11
        ),
        "risk_high": ParagraphStyle(
            "risk_high", fontName="Helvetica-Bold",
            fontSize=7, textColor=RED, leading=9
        ),
        "risk_medium": ParagraphStyle(
            "risk_medium", fontName="Helvetica-Bold",
            fontSize=7, textColor=ORANGE, leading=9
        ),
        "risk_low": ParagraphStyle(
            "risk_low", fontName="Helvetica-Bold",
            fontSize=7, textColor=HexColor("#0ea5e9"), leading=9
        ),
    }


class ColorRect(Flowable):
    """A solid filled rectangle used for banners and dividers."""
    def __init__(self, width, height, color, radius=0):
        super().__init__()
        self.width  = width
        self.height = height
        self.color  = color
        self.radius = radius

    def draw(self):
        self.canv.setFillColor(self.color)
        if self.radius:
            self.canv.roundRect(0, 0, self.width, self.height,
                                self.radius, fill=1, stroke=0)
        else:
            self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


def header_footer(canvas, doc, policy_name, policy_number, version):
    """Draws header and footer on every page."""
    W, H = letter
    canvas.saveState()

    # ── Header bar ──
    canvas.setFillColor(DARK)
    canvas.rect(0, H - 0.55 * inch, W, 0.55 * inch, fill=1, stroke=0)

    canvas.setFillColor(CYAN)
    canvas.rect(0, H - 0.55 * inch, 4, 0.55 * inch, fill=1, stroke=0)

    canvas.setFont("Helvetica-Bold", 7)
    canvas.setFillColor(WHITE)
    canvas.drawString(0.35 * inch, H - 0.32 * inch, "MIDNIGHT")

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(T2)
    canvas.drawString(1.1 * inch, H - 0.32 * inch,
                      f"GRC COMPLIANCE SUMMARY  ·  {policy_number}  ·  {version}")

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(T3)
    canvas.drawRightString(W - 0.35 * inch, H - 0.32 * inch, "CONFIDENTIAL")

    # ── Footer bar ──
    canvas.setFillColor(LGRAY)
    canvas.rect(0, 0, W, 0.4 * inch, fill=1, stroke=0)

    canvas.setFillColor(CYAN)
    canvas.rect(0, 0.38 * inch, W, 1.5, fill=1, stroke=0)

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(T2)
    canvas.drawString(0.35 * inch, 0.15 * inch,
                      f"Generated by Midnight · Takeoff LLC · For internal use only.")

    canvas.setFont("Helvetica-Bold", 7)
    canvas.setFillColor(T2)
    canvas.drawRightString(W - 0.35 * inch, 0.15 * inch,
                           f"Page {doc.page}")

    canvas.restoreState()


def build_grc_pdf(policy_data: dict, framework_map: dict) -> bytes:
    """
    Build a premium branded GRC Summary PDF.
    Returns raw bytes.
    """
    S = make_styles()
    W, H = letter
    content_w = W - 1.4 * inch  # margins 0.7" each side

    buf = io.BytesIO()

    policy_name   = policy_data.get("policy_name",   "Policy")
    policy_number = policy_data.get("policy_number", "SEC-P")
    version       = policy_data.get("version",       "V1.0")
    owner         = policy_data.get("owner_name",    "")
    effective     = policy_data.get("effective_date","")

    doc = SimpleDocTemplate(
        buf,
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

    # ── COVER BLOCK ───────────────────────────────────────────────────────────
    # Dark banner with policy name
    cover_data = [[
        Paragraph("MIDNIGHT", ParagraphStyle(
            "mn", fontName="Helvetica-Bold", fontSize=9,
            textColor=CYAN, letterSpacing=3
        )),
        Paragraph("GRC COMPLIANCE SUMMARY", ParagraphStyle(
            "grc", fontName="Helvetica", fontSize=8,
            textColor=T3, letterSpacing=2, alignment=TA_RIGHT
        )),
    ]]
    cover_table = Table(cover_data, colWidths=[content_w * 0.5, content_w * 0.5])
    cover_table.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), DARK),
        ("TOPPADDING",   (0, 0), (-1, -1), 14),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 14),
        ("LEFTPADDING",  (0, 0), (-1, -1), 14),
        ("RIGHTPADDING", (0, 0), (-1, -1), 14),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW",    (0, 0), (-1, -1), 2, CYAN),
    ]))
    story.append(cover_table)
    story.append(Spacer(1, 14))

    # Policy title block
    story.append(Paragraph(policy_name, ParagraphStyle(
        "pn", fontName="Helvetica-Bold", fontSize=22,
        textColor=DARK, leading=26, spaceAfter=4
    )))

    meta_parts = [policy_number, version]
    if effective:
        meta_parts.append(f"Effective {effective}")
    meta_parts.append("Framework Compliance Report")

    story.append(Paragraph(
        "  ·  ".join(meta_parts),
        ParagraphStyle("meta", fontName="Helvetica", fontSize=9,
                       textColor=CYAN_DIM, leading=12, spaceAfter=2)
    ))

    if owner:
        story.append(Paragraph(
            f"Policy Owner: {owner}",
            ParagraphStyle("own", fontName="Helvetica", fontSize=8,
                           textColor=T2, leading=11, spaceAfter=12)
        ))

    story.append(HRFlowable(width=content_w, thickness=1,
                             color=MGRAY, spaceAfter=16))

    # ── AUDIT SUMMARY ─────────────────────────────────────────────────────────
    audit_text = framework_map.get("audit_summary", "")
    if audit_text:
        summary_data = [[
            Paragraph("AUDIT SUMMARY", ParagraphStyle(
                "as_lbl", fontName="Helvetica-Bold", fontSize=7,
                textColor=CYAN, letterSpacing=2, spaceAfter=4
            )),
        ], [
            Paragraph(audit_text, ParagraphStyle(
                "as_body", fontName="Helvetica", fontSize=9,
                textColor=HexColor("#1e293b"), leading=14
            )),
        ]]
        summary_table = Table(summary_data, colWidths=[content_w])
        summary_table.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), HexColor("#f0f9ff")),
            ("LINERIGHT",     (0, 0), (0, -1),  3, CYAN),
            ("LEFTPADDING",   (0, 0), (-1, -1), 14),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 14),
            ("TOPPADDING",    (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ("ROUNDEDCORNERS",(0, 0), (-1, -1), [4]),
        ]))
        story.append(summary_table)
        story.append(Spacer(1, 16))

    # ── COVERAGE OVERVIEW ─────────────────────────────────────────────────────
    story.append(Paragraph("COVERAGE OVERVIEW", S["section_label"]))

    coverage    = framework_map.get("overall_coverage", "unknown").upper()
    cov_color   = GREEN if coverage == "STRONG" else (
                  ORANGE if coverage == "MODERATE" else RED)
    total_mapped= framework_map.get("total_controls_mapped", 0)
    total_gaps  = framework_map.get("total_gaps", 0)
    frameworks  = ", ".join(framework_map.get("frameworks_covered", [])) or "—"

    overview_rows = [
        # Header
        [
            Paragraph("METRIC",   S["table_header"]),
            Paragraph("VALUE",    S["table_header"]),
            Paragraph("DETAIL",   S["table_header"]),
        ],
        [
            Paragraph("Overall Coverage",    S["table_cell_bold"]),
            Paragraph(coverage, ParagraphStyle(
                "cov", fontName="Helvetica-Bold", fontSize=9,
                textColor=cov_color
            )),
            Paragraph("Based on mapped controls vs expected controls", S["table_cell"]),
        ],
        [
            Paragraph("Controls Mapped",     S["table_cell_bold"]),
            Paragraph(str(total_mapped),     S["table_cell"]),
            Paragraph("Framework controls satisfied by this policy", S["table_cell"]),
        ],
        [
            Paragraph("Gaps Identified",     S["table_cell_bold"]),
            Paragraph(str(total_gaps), ParagraphStyle(
                "gaps", fontName="Helvetica-Bold", fontSize=9,
                textColor=RED if total_gaps > 0 else GREEN
            )),
            Paragraph("Controls required but not addressed", S["table_cell"]),
        ],
        [
            Paragraph("Frameworks Assessed", S["table_cell_bold"]),
            Paragraph("",                    S["table_cell"]),
            Paragraph(frameworks,            S["table_cell"]),
        ],
    ]

    col1 = content_w * 0.28
    col2 = content_w * 0.17
    col3 = content_w * 0.55

    ov_table = Table(overview_rows, colWidths=[col1, col2, col3])
    ov_table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  DARK),
        ("BACKGROUND",    (0, 1), (-1, 1),  HexColor("#f8fafc")),
        ("BACKGROUND",    (0, 2), (-1, 2),  WHITE),
        ("BACKGROUND",    (0, 3), (-1, 3),  HexColor("#f8fafc")),
        ("BACKGROUND",    (0, 4), (-1, 4),  WHITE),
        ("GRID",          (0, 0), (-1, -1), 0.5, MGRAY),
        ("LINEBELOW",     (0, 0), (-1, 0),  1.5, CYAN),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(ov_table)
    story.append(Spacer(1, 20))

    # ── CONTROLS COVERED ──────────────────────────────────────────────────────
    citations = framework_map.get("mapped_citations", [])
    if citations:
        story.append(Paragraph("CONTROLS COVERED", S["section_label"]))

        cit_rows = [[
            Paragraph("FRAMEWORK",    S["table_header"]),
            Paragraph("CONTROL ID",   S["table_header"]),
            Paragraph("CONTROL NAME", S["table_header"]),
            Paragraph("POLICY SECTION",S["table_header"]),
        ]]

        for c in citations:
            cit_rows.append([
                Paragraph(str(c.get("framework",     "")), S["table_cell"]),
                Paragraph(str(c.get("control_id",    "")), S["table_cell_bold"]),
                Paragraph(str(c.get("control_name",  "")), S["table_cell"]),
                Paragraph(str(c.get("policy_section","")), S["table_cell"]),
            ])

        c1 = content_w * 0.22
        c2 = content_w * 0.14
        c3 = content_w * 0.34
        c4 = content_w * 0.30

        cit_table = Table(cit_rows, colWidths=[c1, c2, c3, c4],
                          repeatRows=1)
        row_styles = [
            ("BACKGROUND",    (0, 0), (-1, 0),  DARK),
            ("LINEBELOW",     (0, 0), (-1, 0),  1.5, CYAN),
            ("GRID",          (0, 0), (-1, -1), 0.4, MGRAY),
            ("TOPPADDING",    (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING",   (0, 0), (-1, -1), 8),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ]
        for i in range(1, len(cit_rows)):
            if i % 2 == 0:
                row_styles.append(("BACKGROUND", (0, i), (-1, i), HexColor("#f8fafc")))

        cit_table.setStyle(TableStyle(row_styles))
        story.append(cit_table)
        story.append(Spacer(1, 20))

    # ── GAPS ──────────────────────────────────────────────────────────────────
    gaps = framework_map.get("gaps", [])
    if gaps:
        story.append(Paragraph("COMPLIANCE GAPS — ACTION REQUIRED", S["section_label"]))
        story.append(Paragraph(
            f"{len(gaps)} gap{'s' if len(gaps) != 1 else ''} identified. "
            "Each gap includes suggested policy language to close the control.",
            ParagraphStyle("gap_intro", fontName="Helvetica", fontSize=8,
                           textColor=T2, leading=12, spaceAfter=10)
        ))

        for i, gap in enumerate(gaps, 1):
            risk = str(gap.get("risk_level", "medium")).upper()
            risk_color = RED if risk == "HIGH" else (
                         ORANGE if risk == "MEDIUM" else HexColor("#0ea5e9"))
            risk_bg    = HexColor("#fff1f2") if risk == "HIGH" else (
                         HexColor("#fff7ed") if risk == "MEDIUM" else HexColor("#f0f9ff"))

            framework  = str(gap.get("framework",       ""))
            control_id = str(gap.get("control_id",      ""))
            ctrl_name  = str(gap.get("control_name",    ""))
            gap_desc   = str(gap.get("gap_description", ""))
            suggestion = str(gap.get("suggestion",      ""))

            gap_block = [
                # Title row
                [
                    Paragraph(
                        f"Gap {i} &nbsp;—&nbsp; {framework} {control_id}",
                        ParagraphStyle("gt", fontName="Helvetica-Bold",
                                       fontSize=10, textColor=DARK, leading=13)
                    ),
                    Paragraph(risk, ParagraphStyle(
                        "rsk", fontName="Helvetica-Bold", fontSize=8,
                        textColor=risk_color, leading=10, alignment=TA_RIGHT
                    )),
                ],
                # Control name
                [
                    Paragraph(ctrl_name, ParagraphStyle(
                        "cn", fontName="Helvetica", fontSize=8,
                        textColor=T2, leading=11
                    )),
                    Paragraph("", S["table_cell"]),
                ],
                # Gap description
                [
                    Paragraph(f"<b>Gap:</b> {gap_desc}", ParagraphStyle(
                        "gd", fontName="Helvetica", fontSize=8,
                        textColor=HexColor("#334155"), leading=12
                    )),
                    Paragraph("", S["table_cell"]),
                ],
                # Suggestion
                [
                    Paragraph(f"<b>Suggested Language:</b> {suggestion}",
                              ParagraphStyle(
                                  "sg", fontName="Helvetica-Oblique",
                                  fontSize=8,
                                  textColor=HexColor("#1e3a5f"), leading=12
                              )),
                    Paragraph("", S["table_cell"]),
                ],
            ]

            g1 = content_w * 0.84
            g2 = content_w * 0.16

            gap_table = Table(gap_block, colWidths=[g1, g2])
            gap_table.setStyle(TableStyle([
                ("BACKGROUND",    (0, 0), (-1, 0),  risk_bg),
                ("BACKGROUND",    (0, 1), (-1, 1),  WHITE),
                ("BACKGROUND",    (0, 2), (-1, 2),  HexColor("#f8fafc")),
                ("BACKGROUND",    (0, 3), (-1, 3),  HexColor("#f0f9ff")),
                ("BOX",           (0, 0), (-1, -1), 1,   MGRAY),
                ("LINEBEFORE",    (0, 0), (0, -1),  3,   risk_color),
                ("LINEBELOW",     (0, 0), (-1, 0),  0.5, MGRAY),
                ("LINEBELOW",     (0, 1), (-1, 1),  0.5, MGRAY),
                ("LINEBELOW",     (0, 2), (-1, 2),  0.5, MGRAY),
                ("TOPPADDING",    (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("LEFTPADDING",   (0, 0), (-1, -1), 12),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
                ("VALIGN",        (0, 0), (-1, -1), "TOP"),
                ("SPAN",          (0, 1), (0, 1)),
                ("SPAN",          (0, 2), (0, 2)),
                ("SPAN",          (0, 3), (0, 3)),
            ]))

            story.append(KeepTogether([gap_table, Spacer(1, 10)]))

    # ── CLOSING BLOCK ─────────────────────────────────────────────────────────
    story.append(Spacer(1, 20))
    close_data = [[
        Paragraph(
            "This report was generated by Midnight, an enterprise policy intelligence platform "
            "by Takeoff LLC. Content is derived from AI-assisted framework analysis and should "
            "be reviewed by a qualified compliance professional before use in formal audit proceedings.",
            ParagraphStyle("disc", fontName="Helvetica", fontSize=7,
                           textColor=T2, leading=11)
        )
    ]]
    close_table = Table(close_data, colWidths=[content_w])
    close_table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), LGRAY),
        ("LINEABOVE",     (0, 0), (-1, 0),  1, MGRAY),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING",   (0, 0), (-1, -1), 12),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
    ]))
    story.append(close_table)

    # ── BUILD ─────────────────────────────────────────────────────────────────
    doc.build(
        story,
        onFirstPage=lambda c, d: header_footer(c, d, policy_name, policy_number, version),
        onLaterPages=lambda c, d: header_footer(c, d, policy_name, policy_number, version),
    )

    return buf.getvalue()


# ── Test with sample data ─────────────────────────────────────────────────────
if __name__ == "__main__":
    sample_policy = {
        "policy_name":   "Access Control Policy",
        "policy_number": "SEC-P023",
        "version":       "V28.0",
        "owner_name":    "Brian Word",
        "owner_title":   "CISO",
        "effective_date":"04/05/2026",
    }

    sample_map = {
        "overall_coverage":     "moderate",
        "total_controls_mapped": 5,
        "total_gaps":            3,
        "frameworks_covered":   ["NIST CSF v1.1", "HIPAA (45 CFR 164)", "PCI DSS v3.2", "ISO/IEC 27001:2022"],
        "audit_summary": (
            "The Access Control Policy (SEC-P023) provides moderate coverage of access control "
            "requirements, with some gaps in remote access management, security awareness and "
            "training, and password management. The policy aligns with several frameworks, "
            "including NIST Cybersecurity Framework v1.1, HIPAA, and PCI DSS v3.2."
        ),
        "mapped_citations": [
            {"framework": "NIST CSF v1.1",       "control_id": "PR.AC-1", "control_name": "Identities and credentials are issued, managed, and verified", "policy_section": "Account Administration"},
            {"framework": "NIST CSF v1.1",       "control_id": "PR.AC-2", "control_name": "Access to assets is limited to authorized users",              "policy_section": "Policy Statement"},
            {"framework": "HIPAA (45 CFR 164)",  "control_id": "164.308(a)(4)", "control_name": "Information Access Control",                            "policy_section": "Policy Statement"},
            {"framework": "PCI DSS v3.2",         "control_id": "8.1",    "control_name": "Assign all users a unique ID",                                "policy_section": "Account Administration"},
            {"framework": "ISO/IEC 27001:2022",   "control_id": "A.6.1.2","control_name": "Management of access rights",                                 "policy_section": "Account Administration"},
        ],
        "gaps": [
            {
                "framework":       "NIST CSF v1.1",
                "control_id":      "PR.AC-3",
                "control_name":    "Remote access is managed",
                "gap_description": "The policy does not address remote access management.",
                "risk_level":      "medium",
                "suggestion":      "The organization should add a section to the policy that outlines procedures for managing remote access, including the use of multifactor authentication and secure protocols for remote connections.",
            },
            {
                "framework":       "HIPAA (45 CFR 164)",
                "control_id":      "164.308(a)(5)",
                "control_name":    "Security Awareness and Training",
                "gap_description": "The policy does not address security awareness and training for all users.",
                "risk_level":      "medium",
                "suggestion":      "The organization should add a section to the policy that outlines procedures for providing security awareness and training to all users, including employees, contractors, and vendors.",
            },
            {
                "framework":       "PCI DSS v3.2",
                "control_id":      "8.2",
                "control_name":    "Proper user authentication and password management",
                "gap_description": "The policy does not address password management.",
                "risk_level":      "medium",
                "suggestion":      "The organization should add a section to the policy that outlines procedures for password management, including password complexity, expiration, and reset procedures.",
            },
        ],
    }

    pdf_bytes = build_grc_pdf(sample_policy, sample_map)

    with open("/home/claude/grc_summary_sample.pdf", "wb") as f:
        f.write(pdf_bytes)

    print(f"PDF built: {len(pdf_bytes):,} bytes → grc_summary_sample.pdf")
