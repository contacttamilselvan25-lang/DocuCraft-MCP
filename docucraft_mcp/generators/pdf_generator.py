"""PDF Report Generator for DocuCraft MCP.

Generates publication-quality, print-ready PDF documents (.pdf) using ReportLab with
custom color themes, page headers/footers, callout quote boxes, and formatted data tables.
"""

import os
from typing import Dict, List, Optional, Any
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

from themes import get_theme_config


class NumberedCanvas(canvas.Canvas):
    """Two-pass canvas to dynamically compute and draw total page counts."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 9)
        self.setFillColor(colors.HexColor("#64748B"))

        # Footer line and text
        self.setStrokeColor(colors.HexColor("#CBD5E1"))
        self.setLineWidth(0.5)
        self.line(54, 45, 612 - 54, 45)

        footer_text = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(612 - 54, 30, footer_text)
        self.drawString(54, 30, "Generated via DocuCraft MCP  |  Confidential")

        self.restoreState()


def build_pdf(
    title: str,
    sections: List[Dict[str, Any]],
    output_path: str,
    subtitle: Optional[str] = None,
    author: str = "DocuCraft MCP",
    theme: str = "corporate_blue",
    primary_color: Optional[str] = None,
    secondary_color: Optional[str] = None,
    accent_color: Optional[str] = None,
    font_family: Optional[str] = None
) -> Dict[str, Any]:
    """Generate a themed, publication-grade PDF report (.pdf).

    Args:
        title: Main document title.
        sections: List of section dicts (heading, paragraphs, callouts, tables, bullets).
        output_path: Target .pdf file path.
        subtitle: Subtitle or summary statement.
        author: Author or corporate name.
        theme: Theme preset name.
        primary_color: Custom primary hex.
        secondary_color: Custom secondary hex.
        accent_color: Custom accent hex.
        font_family: Custom font name.

    Returns:
        Dict with document metrics and output path.
    """
    theme_cfg = get_theme_config(
        theme_name=theme,
        primary_color=primary_color,
        secondary_color=secondary_color,
        accent_color=accent_color,
        font_family=font_family
    )

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    # Standard letter: 8.5 x 11 inches, 0.75 in margins (54 pt)
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=60
    )

    p_col = colors.HexColor(theme_cfg["primary"])
    s_col = colors.HexColor(theme_cfg["secondary"])
    bg_col = colors.HexColor(theme_cfg["bg_light"])
    td_col = colors.HexColor(theme_cfg["text_dark"])

    # Base stylesheet
    styles = getSampleStyleSheet()

    # Custom paragraph styles
    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=26,
        leading=32,
        textColor=p_col,
        spaceAfter=6
    )

    sub_style = ParagraphStyle(
        "DocSubtitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=13,
        leading=18,
        textColor=s_col,
        spaceAfter=14
    )

    h1_style = ParagraphStyle(
        "Heading1_Custom",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=22,
        textColor=p_col,
        spaceBefore=14,
        spaceAfter=6,
        keepWithNext=True
    )

    body_style = ParagraphStyle(
        "Body_Custom",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=15,
        textColor=td_col,
        spaceAfter=8
    )

    bullet_style = ParagraphStyle(
        "Bullet_Custom",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        textColor=td_col,
        leftIndent=15,
        spaceAfter=4
    )

    callout_style = ParagraphStyle(
        "Callout_Custom",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=10,
        leading=14,
        textColor=td_col
    )

    callout_hdr_style = ParagraphStyle(
        "Callout_Hdr",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10.5,
        leading=14,
        textColor=p_col,
        spaceAfter=3
    )

    story = []

    # Title Block
    story.append(Paragraph(title, title_style))
    if subtitle:
        story.append(Paragraph(subtitle, sub_style))

    # Divider bar
    story.append(HRFlowable(width="100%", thickness=1.5, color=s_col, spaceBefore=4, spaceAfter=14))

    # Process Sections
    for sec in sections:
        h_text = sec.get("heading")
        if h_text:
            story.append(Paragraph(h_text, h1_style))

        # Paragraphs
        for p_text in sec.get("paragraphs", []):
            story.append(Paragraph(p_text, body_style))

        # Bullets
        for b_text in sec.get("bullets", []):
            story.append(Paragraph(f"•  {b_text}", bullet_style))

        # Callout box
        callout = sec.get("callout")
        if callout:
            c_text = callout.get("text", "")
            c_title = callout.get("title", "KEY TAKEAWAY").upper()
            c_content = [
                Paragraph(f"<b>📌 {c_title}</b>", callout_hdr_style),
                Paragraph(c_text, callout_style)
            ]
            c_table = Table([[c_content]], colWidths=[504])
            c_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), bg_col),
                ("BOX", (0, 0), (-1, -1), 1, s_col),
                ("LEFTPADDING", (0, 0), (-1, -1), 14),
                ("RIGHTPADDING", (0, 0), (-1, -1), 14),
                ("TOPPADDING", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ]))
            story.append(Spacer(1, 6))
            story.append(c_table)
            story.append(Spacer(1, 10))

        # Tables
        tbl_data = sec.get("table")
        if tbl_data:
            headers = tbl_data.get("headers", [])
            rows = tbl_data.get("rows", [])
            if headers:
                col_count = len(headers)
                col_w = 504 / col_count

                tbl_matrix = []
                # Header cells wrapped in Paragraphs
                th_style = ParagraphStyle(
                    "TH", fontName="Helvetica-Bold", fontSize=9.5, leading=12, textColor=colors.white, alignment=1
                )
                tbl_matrix.append([Paragraph(str(h), th_style) for h in headers])

                # Data cells wrapped in Paragraphs
                td_cell_style = ParagraphStyle(
                    "TD", fontName="Helvetica", fontSize=9, leading=12, textColor=td_col
                )
                for r in rows:
                    row_cells = []
                    for c_idx in range(col_count):
                        val = r[c_idx] if c_idx < len(r) else ""
                        row_cells.append(Paragraph(str(val), td_cell_style))
                    tbl_matrix.append(row_cells)

                t = Table(tbl_matrix, colWidths=[col_w] * col_count)
                t_style = [
                    ("BACKGROUND", (0, 0), (-1, 0), p_col),
                    ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ]
                # Alternating row colors
                for r_i in range(1, len(tbl_matrix)):
                    if r_i % 2 == 1:
                        t_style.append(("BACKGROUND", (0, r_i), (-1, r_i), bg_col))

                t.setStyle(TableStyle(t_style))
                story.append(Spacer(1, 6))
                story.append(t)
                story.append(Spacer(1, 12))

    # Build PDF with dynamic multi-pass NumberedCanvas
    doc.build(story, canvasmaker=NumberedCanvas)

    return {
        "format": "pdf",
        "title": title,
        "sections_count": len(sections),
        "theme_used": theme_cfg["name"],
        "output_path": os.path.abspath(output_path)
    }
