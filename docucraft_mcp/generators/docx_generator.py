"""Word Document Generator for DocuCraft MCP.

Generates structured, executive-grade Word documents (.docx) with custom color themes,
callout alert boxes, styled headers/footers, and formatted data tables.
"""

import os
from typing import Dict, List, Optional, Any
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from themes import get_theme_config, hex_to_rgb


def _set_cell_background(cell, hex_color: str):
    """Set shading color for a table cell."""
    clean_hex = hex_color.lstrip("#")
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{clean_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def _set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Set cell internal padding in dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def build_document(
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
    """Generate a themed Word document (.docx).

    Args:
        title: Main document title.
        sections: List of section dicts (heading, paragraphs, callouts, tables, bullets).
        output_path: Target .docx file path.
        subtitle: Document subtitle or executive summary summary line.
        author: Author or corporate name.
        theme: Theme preset name.
        primary_color: Custom primary hex color.
        secondary_color: Custom secondary hex color.
        accent_color: Custom accent hex color.
        font_family: Custom font family name.

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

    doc = Document()

    # Set page margins (1 inch standard)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        # Header / Footer
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.text = f"{title}  |  {author}"
        fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        fp.runs[0].font.name = theme_cfg["font_body"]
        fp.runs[0].font.size = Pt(9)
        fp.runs[0].font.color.rgb = RGBColor(140, 150, 165)

    p_rgb = RGBColor(*theme_cfg["primary_rgb"])
    s_rgb = RGBColor(*theme_cfg["secondary_rgb"])
    td_rgb = RGBColor(*theme_cfg["text_dark_rgb"])
    font_h = theme_cfg["font_heading"]
    font_b = theme_cfg["font_body"]

    # --- COVER / HEADER TITLE BLOCK ---
    title_p = doc.add_paragraph()
    title_run = title_p.add_run(title)
    title_run.font.name = font_h
    title_run.font.size = Pt(28)
    title_run.font.bold = True
    title_run.font.color.rgb = p_rgb
    title_p.paragraph_format.space_after = Pt(4)

    if subtitle:
        sub_p = doc.add_paragraph()
        sub_run = sub_p.add_run(subtitle)
        sub_run.font.name = font_b
        sub_run.font.size = Pt(14)
        sub_run.font.color.rgb = s_rgb
        sub_p.paragraph_format.space_after = Pt(18)

    # Accent divider bar
    div_p = doc.add_paragraph()
    div_run = div_p.add_run("―" * 35)
    div_run.font.size = Pt(14)
    div_run.font.color.rgb = s_rgb
    div_p.paragraph_format.space_after = Pt(20)

    # --- SECTIONS LOOP ---
    for sec in sections:
        # 1. Section Heading
        h_text = sec.get("heading")
        if h_text:
            hp = doc.add_paragraph()
            hrun = hp.add_run(h_text)
            hrun.font.name = font_h
            hrun.font.size = Pt(18)
            hrun.font.bold = True
            hrun.font.color.rgb = p_rgb
            hp.paragraph_format.space_before = Pt(16)
            hp.paragraph_format.space_after = Pt(6)

        # 2. Paragraphs
        paragraphs = sec.get("paragraphs", [])
        for para_text in paragraphs:
            pp = doc.add_paragraph()
            prun = pp.add_run(para_text)
            prun.font.name = font_b
            prun.font.size = Pt(11)
            prun.font.color.rgb = td_rgb
            pp.paragraph_format.space_after = Pt(8)
            pp.paragraph_format.line_spacing = 1.15

        # 3. Bullets
        bullets = sec.get("bullets", [])
        for b_text in bullets:
            bp = doc.add_paragraph(style="List Bullet")
            brun = bp.add_run(b_text)
            brun.font.name = font_b
            brun.font.size = Pt(11)
            brun.font.color.rgb = td_rgb
            bp.paragraph_format.space_after = Pt(4)

        # 4. Callout Box (Note / Important / Highlight)
        callout = sec.get("callout")
        if callout:
            c_text = callout.get("text", "")
            c_title = callout.get("title", "KEY TAKEAWAY")

            # 1x1 Table as a styled callout box
            c_table = doc.add_table(rows=1, cols=1)
            c_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            c_cell = c_table.cell(0, 0)
            c_cell.width = Inches(6.5)
            _set_cell_background(c_cell, theme_cfg["bg_light"])
            _set_cell_margins(c_cell, top=140, bottom=140, left=200, right=200)

            cp = c_cell.paragraphs[0]
            cp_hdr = cp.add_run(f"📌 {c_title}\n")
            cp_hdr.font.name = font_h
            cp_hdr.font.size = Pt(11)
            cp_hdr.font.bold = True
            cp_hdr.font.color.rgb = p_rgb

            cp_body = cp.add_run(c_text)
            cp_body.font.name = font_b
            cp_body.font.size = Pt(10.5)
            cp_body.font.color.rgb = td_rgb
            doc.add_paragraph().paragraph_format.space_after = Pt(8)

        # 5. Formatted Table
        tbl_data = sec.get("table")
        if tbl_data:
            headers = tbl_data.get("headers", [])
            rows = tbl_data.get("rows", [])
            if headers:
                t = doc.add_table(rows=len(rows) + 1, cols=len(headers))
                t.alignment = WD_TABLE_ALIGNMENT.CENTER

                # Header Row
                for c_idx, h in enumerate(headers):
                    cell = t.cell(0, c_idx)
                    _set_cell_background(cell, theme_cfg["primary"])
                    _set_cell_margins(cell, top=120, bottom=120, left=150, right=150)
                    p = cell.paragraphs[0]
                    r = p.add_run(str(h))
                    r.font.name = font_h
                    r.font.size = Pt(10.5)
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(255, 255, 255)

                # Data Rows
                for r_idx, row_vals in enumerate(rows):
                    is_even = (r_idx % 2 == 0)
                    bg_col = theme_cfg["bg_light"] if is_even else "#FFFFFF"
                    for c_idx, cell_val in enumerate(row_vals):
                        if c_idx < len(headers):
                            cell = t.cell(r_idx + 1, c_idx)
                            _set_cell_background(cell, bg_col)
                            _set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
                            p = cell.paragraphs[0]
                            r = p.add_run(str(cell_val))
                            r.font.name = font_b
                            r.font.size = Pt(10)
                            r.font.color.rgb = td_rgb

                doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Save document
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)

    return {
        "format": "docx",
        "title": title,
        "sections_count": len(sections),
        "theme_used": theme_cfg["name"],
        "output_path": os.path.abspath(output_path)
    }
