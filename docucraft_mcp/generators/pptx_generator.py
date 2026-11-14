"""PowerPoint Presentation Generator for DocuCraft MCP.

Generates professional, widescreen (16:9) presentations with customizable themes,
multiple slide layout archetypes, KPI callout cards, comparison columns, and tables.
"""

import os
from typing import Dict, List, Optional, Any
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

from themes import get_theme_config


def build_presentation(
    title: str,
    slides: List[Dict[str, Any]],
    output_path: str,
    subtitle: Optional[str] = None,
    theme: str = "corporate_blue",
    primary_color: Optional[str] = None,
    secondary_color: Optional[str] = None,
    accent_color: Optional[str] = None,
    font_family: Optional[str] = None,
    author: str = "DocuCraft MCP"
) -> Dict[str, Any]:
    """Generate a themed PowerPoint presentation (.pptx).

    Args:
        title: Deck main title.
        slides: List of slide dicts, each with 'layout', 'title', and layout-specific content.
        output_path: Target .pptx file path.
        subtitle: Deck subtitle or presenter credentials.
        theme: Theme preset name.
        primary_color: Custom hex color.
        secondary_color: Custom hex color.
        accent_color: Custom hex color.
        font_family: Custom font family name.
        author: Presentation author or company name.

    Returns:
        Dict with status, output path, slide count, and theme used.
    """
    theme_cfg = get_theme_config(
        theme_name=theme,
        primary_color=primary_color,
        secondary_color=secondary_color,
        accent_color=accent_color,
        font_family=font_family
    )

    prs = Presentation()
    # 16:9 Widescreen dimensions
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Convert RGB
    p_rgb = RGBColor(*theme_cfg["primary_rgb"])
    s_rgb = RGBColor(*theme_cfg["secondary_rgb"])
    a_rgb = RGBColor(*theme_cfg["accent_rgb"])
    bg_rgb = RGBColor(*theme_cfg["bg_light_rgb"])
    td_rgb = RGBColor(*theme_cfg["text_dark_rgb"])
    tl_rgb = RGBColor(*theme_cfg["text_light_rgb"])

    font_h = theme_cfg["font_heading"]
    font_b = theme_cfg["font_body"]

    # --- 1. TITLE SLIDE ---
    title_slide = prs.slides.add_slide(blank_layout)

    # Background shape
    bg_shape = title_slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5)
    )
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = p_rgb
    bg_shape.line.fill.background()

    # Top accent line
    accent_bar = title_slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(1.2), Inches(1.5), Inches(0.1)
    )
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = a_rgb
    accent_bar.line.fill.background()

    # Title box
    tx_box = title_slide.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.3), Inches(2.5))
    tf = tx_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = font_h
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = tl_rgb

    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.name = font_b
        p2.font.size = Pt(22)
        p2.font.color.rgb = RGBColor(220, 230, 245)
        p2.space_before = Pt(20)

    # Author / Footer Info
    author_box = title_slide.shapes.add_textbox(Inches(1.0), Inches(6.0), Inches(11.3), Inches(0.8))
    ap = author_box.text_frame.paragraphs[0]
    ap.text = f"{author}  |  Executive Presentation"
    ap.font.name = font_b
    ap.font.size = Pt(14)
    ap.font.color.rgb = RGBColor(180, 200, 225)

    # --- HELPER: HEADER BANNER FOR CONTENT SLIDES ---
    def add_slide_header(slide, slide_title, slide_category=None):
        # Subtle header bottom rule
        rule = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.35), Inches(11.7), Inches(0.04)
        )
        rule.fill.solid()
        rule.fill.fore_color.rgb = s_rgb
        rule.line.fill.background()

        # Category pill if present
        if slide_category:
            pill = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(6.0), Inches(0.4))
            pp = pill.text_frame.paragraphs[0]
            pp.text = slide_category.upper()
            pp.font.name = font_b
            pp.font.size = Pt(11)
            pp.font.bold = True
            pp.font.color.rgb = s_rgb

        # Header title
        hdr_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(0.65) if slide_category else Inches(0.5), Inches(11.7), Inches(0.7)
        )
        hp = hdr_box.text_frame.paragraphs[0]
        hp.text = slide_title
        hp.font.name = font_h
        hp.font.size = Pt(28)
        hp.font.bold = True
        hp.font.color.rgb = p_rgb

    # --- 2. CONTENT SLIDES LOOP ---
    for item in slides:
        slide = prs.slides.add_slide(blank_layout)
        layout = item.get("layout", "bullets").lower()
        s_title = item.get("title", "Untitled Slide")
        s_cat = item.get("category", "")

        add_slide_header(slide, s_title, s_cat)

        # ----------------------------------------------------
        # LAYOUT: STATS / KPI CARDS
        # ----------------------------------------------------
        if layout in ["stats", "kpi"]:
            stats_list = item.get("stats", [])
            count = max(1, min(len(stats_list), 4))
            card_w = (11.7 - (count - 1) * 0.4) / count
            for i, st in enumerate(stats_list[:count]):
                left = Inches(0.8 + i * (card_w + 0.4))
                # Card container
                card = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(2.2), Inches(card_w), Inches(3.8)
                )
                card.fill.solid()
                card.fill.fore_color.rgb = bg_rgb
                card.line.color.rgb = s_rgb
                card.line.width = Pt(1.5)

                # Stat Value
                val_box = slide.shapes.add_textbox(left, Inches(2.6), Inches(card_w), Inches(1.2))
                vp = val_box.text_frame.paragraphs[0]
                vp.text = str(st.get("value", ""))
                vp.alignment = PP_ALIGN.CENTER
                vp.font.name = font_h
                vp.font.size = Pt(40)
                vp.font.bold = True
                vp.font.color.rgb = p_rgb

                # Stat Label
                lbl_box = slide.shapes.add_textbox(left + Inches(0.2), Inches(3.8), Inches(card_w - 0.4), Inches(1.8))
                lp = lbl_box.text_frame.paragraphs[0]
                lp.text = str(st.get("label", ""))
                lp.alignment = PP_ALIGN.CENTER
                lp.font.name = font_b
                lp.font.size = Pt(16)
                lp.font.bold = True
                lp.font.color.rgb = td_rgb

                # Subtext if present
                if st.get("description"):
                    lp2 = lbl_box.text_frame.add_paragraph()
                    lp2.text = str(st.get("description"))
                    lp2.alignment = PP_ALIGN.CENTER
                    lp2.font.name = font_b
                    lp2.font.size = Pt(12)
                    lp2.font.color.rgb = RGBColor(100, 115, 130)
                    lp2.space_before = Pt(8)

        # ----------------------------------------------------
        # LAYOUT: COMPARISON / SPLIT COLUMN
        # ----------------------------------------------------
        elif layout in ["comparison", "split_column"]:
            col1 = item.get("column_1", {"title": "Option A", "bullets": []})
            col2 = item.get("column_2", {"title": "Option B", "bullets": []})

            for idx, col in enumerate([col1, col2]):
                left = Inches(0.8 + idx * 6.0)
                card = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(1.8), Inches(5.6), Inches(4.8)
                )
                card.fill.solid()
                card.fill.fore_color.rgb = bg_rgb
                card.line.color.rgb = p_rgb if idx == 0 else s_rgb
                card.line.width = Pt(1.5)

                # Column header
                ch_box = slide.shapes.add_textbox(left + Inches(0.3), Inches(2.0), Inches(5.0), Inches(0.6))
                cp = ch_box.text_frame.paragraphs[0]
                cp.text = col.get("title", f"Section {idx+1}")
                cp.font.name = font_h
                cp.font.size = Pt(20)
                cp.font.bold = True
                cp.font.color.rgb = p_rgb if idx == 0 else s_rgb

                # Bullets
                cb_box = slide.shapes.add_textbox(left + Inches(0.3), Inches(2.7), Inches(5.0), Inches(3.6))
                ctf = cb_box.text_frame
                ctf.word_wrap = True
                bullets = col.get("bullets", [])
                for b_idx, b_text in enumerate(bullets):
                    bp = ctf.paragraphs[0] if b_idx == 0 else ctf.add_paragraph()
                    bp.text = f"•  {b_text}"
                    bp.font.name = font_b
                    bp.font.size = Pt(15)
                    bp.font.color.rgb = td_rgb
                    bp.space_before = Pt(10)

        # ----------------------------------------------------
        # LAYOUT: TABLE
        # ----------------------------------------------------
        elif layout == "table":
            headers = item.get("headers", ["Column 1", "Column 2", "Column 3"])
            rows = item.get("rows", [["Data 1", "Data 2", "Data 3"]])
            n_rows = len(rows) + 1
            n_cols = len(headers)

            t_shape = slide.shapes.add_table(
                n_rows, n_cols, Inches(0.8), Inches(1.8), Inches(11.7), Inches(0.6 * n_rows)
            )
            table = t_shape.table

            # Header row styling
            for c_idx, h_text in enumerate(headers):
                cell = table.cell(0, c_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = p_rgb
                cell.text_frame.word_wrap = True
                p = cell.text_frame.paragraphs[0]
                p.text = str(h_text)
                p.font.name = font_h
                p.font.size = Pt(15)
                p.font.bold = True
                p.font.color.rgb = tl_rgb
                p.alignment = PP_ALIGN.CENTER

            # Data rows
            for r_idx, row in enumerate(rows):
                is_even = (r_idx % 2 == 0)
                for c_idx, cell_val in enumerate(row):
                    if c_idx < n_cols:
                        cell = table.cell(r_idx + 1, c_idx)
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = bg_rgb if is_even else RGBColor(255, 255, 255)
                        cell.text_frame.word_wrap = True
                        p = cell.text_frame.paragraphs[0]
                        p.text = str(cell_val)
                        p.font.name = font_b
                        p.font.size = Pt(14)
                        p.font.color.rgb = td_rgb

        # ----------------------------------------------------
        # LAYOUT: DEFAULT / BULLETS & CARDS
        # ----------------------------------------------------
        else:
            bullets = item.get("bullets", [])
            card_count = len(bullets)
            if card_count <= 4 and card_count > 0:
                # Render as visual cards
                c_height = min(1.1, 4.8 / card_count - 0.2)
                for b_idx, b_text in enumerate(bullets):
                    top = Inches(1.8 + b_idx * (c_height + 0.2))
                    c_box = slide.shapes.add_shape(
                        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top, Inches(11.7), Inches(c_height)
                    )
                    c_box.fill.solid()
                    c_box.fill.fore_color.rgb = bg_rgb
                    c_box.line.color.rgb = s_rgb
                    c_box.line.width = Pt(1.0)

                    # Accent icon tag
                    tag = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE, Inches(0.8), top, Inches(0.15), Inches(c_height)
                    )
                    tag.fill.solid()
                    tag.fill.fore_color.rgb = p_rgb
                    tag.line.fill.background()

                    tbox = slide.shapes.add_textbox(Inches(1.2), top + Inches(0.1), Inches(11.1), Inches(c_height - 0.2))
                    tp = tbox.text_frame.paragraphs[0]
                    tp.text = b_text
                    tp.font.name = font_b
                    tp.font.size = Pt(16)
                    tp.font.color.rgb = td_rgb
            else:
                # Standard clean bullet frame
                tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
                tf = tx_box.text_frame
                tf.word_wrap = True
                for b_idx, b_text in enumerate(bullets):
                    bp = tf.paragraphs[0] if b_idx == 0 else tf.add_paragraph()
                    bp.text = f"•   {b_text}"
                    bp.font.name = font_b
                    bp.font.size = Pt(18)
                    bp.font.color.rgb = td_rgb
                    bp.space_before = Pt(14)

    # Save presentation
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)

    return {
        "format": "pptx",
        "title": title,
        "slide_count": len(prs.slides),
        "theme_used": theme_cfg["name"],
        "output_path": os.path.abspath(output_path)
    }
