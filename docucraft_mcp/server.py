"""DocuCraft MCP: Universal AI Document, Presentation, Spreadsheet & PDF Studio.

Generates beautifully formatted PowerPoint (.pptx), Word (.docx), Excel (.xlsx),
and PDF (.pdf) documents directly from natural language agent prompts.
"""

import sys
import os
from typing import Dict, List, Optional, Any

# Ensure directory is on sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mcp.server.fastmcp import FastMCP
from themes import PRESET_THEMES, get_theme_config
from generators.pptx_generator import build_presentation
from generators.docx_generator import build_document
from generators.excel_generator import build_spreadsheet
from generators.pdf_generator import build_pdf

# Initialize FastMCP Server
mcp = FastMCP("docucraft_mcp")


@mcp.tool(
    name="docucraft_list_themes",
    annotations={"readOnlyHint": True, "idempotentHint": True}
)
def tool_list_themes() -> Dict[str, Any]:
    """List all available pre-designed visual themes and color palettes for documents.

    Themes include:
    - corporate_blue (Navy/Cyan for business, finance, and corporate proposals)
    - emerald_growth (Forest/Mint for healthcare, pharma, biotech, ESG)
    - modern_dark (Midnight Slate/Electric Violet for modern tech & AI pitch decks)
    - warm_executive (Burgundy/Amber Gold for legal, executive, and annual reviews)
    - minimal_slate (Cool Gray/Sky Blue for clean documentation and academic work)
    - crimson_bold (Ruby/Coral for high-energy sales and marketing)
    - sunset_orange (Rust/Sunburst Gold for creative and consumer brands)
    - royal_indigo (Indigo/Purple for SaaS, analytics, and software suites)
    """
    return {
        "total_presets": len(PRESET_THEMES),
        "presets": PRESET_THEMES,
        "custom_support": (
            "You can also pass custom hex values to 'primary_color', 'secondary_color', "
            "'accent_color', and 'font_family' on any tool."
        )
    }


@mcp.tool(
    name="docucraft_create_presentation",
    annotations={"readOnlyHint": False, "destructiveHint": False}
)
def tool_create_presentation(
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
    """Create a professional widescreen (16:9) PowerPoint presentation (.pptx).

    Slide layout options for the 'slides' list:
    1. 'bullets': {'title': '...', 'category': '...', 'bullets': ['Point 1', 'Point 2']}
    2. 'kpi' / 'stats': {'title': '...', 'stats': [{'value': '$10M+', 'label': 'Revenue', 'description': '+40% YoY'}]}
    3. 'comparison' / 'split_column': {'title': '...', 'column_1': {'title': '...', 'bullets': [...]}, 'column_2': {'title': '...', 'bullets': [...]}}
    4. 'table': {'title': '...', 'headers': ['Col 1', 'Col 2'], 'rows': [['Val 1', 'Val 2']]}

    Args:
        title: Main presentation title.
        slides: List of slide objects with layout, title, and content.
        output_path: Destination filepath (e.g. 'output/pitch_deck.pptx').
        subtitle: Optional presentation subtitle.
        theme: Theme preset name ('corporate_blue', 'emerald_growth', 'modern_dark', etc.).
        primary_color: Optional custom hex color (e.g. '#1E3A8A').
        secondary_color: Optional custom hex color.
        accent_color: Optional custom hex color.
        font_family: Optional font name (e.g. 'Calibri', 'Segoe UI', 'Arial').
        author: Author or organization name.
    """
    try:
        return build_presentation(
            title=title,
            slides=slides,
            output_path=output_path,
            subtitle=subtitle,
            theme=theme,
            primary_color=primary_color,
            secondary_color=secondary_color,
            accent_color=accent_color,
            font_family=font_family,
            author=author
        )
    except Exception as e:
        return {"error": f"Failed to build presentation: {str(e)}"}


@mcp.tool(
    name="docucraft_create_document",
    annotations={"readOnlyHint": False, "destructiveHint": False}
)
def tool_create_document(
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
    """Create a structured executive Word document (.docx) with custom color themes and callout boxes.

    Section object structure in 'sections':
    - 'heading': Section title
    - 'paragraphs': ['Paragraph 1 text...', 'Paragraph 2 text...']
    - 'bullets': ['Bullet point 1', 'Bullet point 2']
    - 'callout': {'title': 'KEY NOTE', 'text': 'Important highlight text...'}
    - 'table': {'headers': ['Col A', 'Col B'], 'rows': [['Val 1', 'Val 2']]}

    Args:
        title: Main document title.
        sections: List of section dicts containing headings, text, callouts, or tables.
        output_path: Destination filepath (e.g. 'output/project_proposal.docx').
        subtitle: Document subtitle or executive summary summary line.
        author: Author or corporate name.
        theme: Theme preset name.
        primary_color: Optional custom hex color.
        secondary_color: Optional custom hex color.
        accent_color: Optional custom hex color.
        font_family: Optional font name.
    """
    try:
        return build_document(
            title=title,
            sections=sections,
            output_path=output_path,
            subtitle=subtitle,
            author=author,
            theme=theme,
            primary_color=primary_color,
            secondary_color=secondary_color,
            accent_color=accent_color,
            font_family=font_family
        )
    except Exception as e:
        return {"error": f"Failed to build Word document: {str(e)}"}


@mcp.tool(
    name="docucraft_create_spreadsheet",
    annotations={"readOnlyHint": False, "destructiveHint": False}
)
def tool_create_spreadsheet(
    title: str,
    sheets: List[Dict[str, Any]],
    output_path: str,
    theme: str = "corporate_blue",
    primary_color: Optional[str] = None,
    secondary_color: Optional[str] = None,
    accent_color: Optional[str] = None,
    font_family: Optional[str] = None
) -> Dict[str, Any]:
    """Create a professional multi-sheet Excel workbook (.xlsx) with KPI summary cards and auto-formulas.

    Sheet object structure in 'sheets':
    - 'sheet_name': Tab name (e.g. 'Q1 Budget', 'Inventory')
    - 'kpi_cards': [{'label': 'Total Budget', 'value': '$125,000'}, {'label': 'Utilized', 'value': '78%'}]
    - 'headers': ['Item', 'Category', 'Unit Cost', 'Quantity', 'Total']
    - 'rows': [['Lab Reagent A', 'Chemicals', 450.0, 10, 4500.0]]
    - 'column_formats': {'Unit Cost': 'currency', 'Total': 'currency'}
    - 'include_totals': True (auto-injects =SUM and =AVERAGE formulas on numeric columns)

    Args:
        title: Main workbook title.
        sheets: List of sheet specification dicts.
        output_path: Destination filepath (e.g. 'output/financial_model.xlsx').
        theme: Theme preset name.
        primary_color: Optional custom hex color.
        secondary_color: Optional custom hex color.
        accent_color: Optional custom hex color.
        font_family: Optional font name.
    """
    try:
        return build_spreadsheet(
            title=title,
            sheets=sheets,
            output_path=output_path,
            theme=theme,
            primary_color=primary_color,
            secondary_color=secondary_color,
            accent_color=accent_color,
            font_family=font_family
        )
    except Exception as e:
        return {"error": f"Failed to build spreadsheet: {str(e)}"}


@mcp.tool(
    name="docucraft_create_pdf",
    annotations={"readOnlyHint": False, "destructiveHint": False}
)
def tool_create_pdf(
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
    """Create a publication-grade, vector PDF report (.pdf) with multi-pass running headers and footers.

    Section object structure in 'sections':
    - 'heading': Section heading
    - 'paragraphs': ['Paragraph text...']
    - 'bullets': ['Key takeaway 1', 'Key takeaway 2']
    - 'callout': {'title': 'REGULATORY NOTE', 'text': 'Callout highlight text...'}
    - 'table': {'headers': ['Parameter', 'Specification', 'Result'], 'rows': [['Assay', '98-102%', '99.8%']]}

    Args:
        title: Main report title.
        sections: List of section dicts.
        output_path: Destination filepath (e.g. 'output/executive_summary.pdf').
        subtitle: Subtitle or executive summary tagline.
        author: Author or corporate name.
        theme: Theme preset name.
        primary_color: Optional custom hex color.
        secondary_color: Optional custom hex color.
        accent_color: Optional custom hex color.
        font_family: Optional font name.
    """
    try:
        return build_pdf(
            title=title,
            sections=sections,
            output_path=output_path,
            subtitle=subtitle,
            author=author,
            theme=theme,
            primary_color=primary_color,
            secondary_color=secondary_color,
            accent_color=accent_color,
            font_family=font_family
        )
    except Exception as e:
        return {"error": f"Failed to build PDF: {str(e)}"}


@mcp.tool(
    name="docucraft_create_suite",
    annotations={"readOnlyHint": False, "destructiveHint": False}
)
def tool_create_suite(
    project_title: str,
    output_dir: str,
    presentation_slides: List[Dict[str, Any]],
    document_sections: List[Dict[str, Any]],
    spreadsheet_sheets: List[Dict[str, Any]],
    subtitle: Optional[str] = None,
    theme: str = "corporate_blue",
    primary_color: Optional[str] = None,
    secondary_color: Optional[str] = None,
    accent_color: Optional[str] = None,
    font_family: Optional[str] = None,
    author: str = "DocuCraft MCP"
) -> Dict[str, Any]:
    """Generate all four documents (.pptx, .docx, .xlsx, .pdf) in one command with unified branding and theme.

    Args:
        project_title: Overall project or company title.
        output_dir: Directory where the four files will be saved.
        presentation_slides: Slide outline for the .pptx presentation.
        document_sections: Sections for the .docx report and .pdf summary.
        spreadsheet_sheets: Sheets for the .xlsx financial / data workbook.
        subtitle: Optional project subtitle.
        theme: Visual theme preset.
        primary_color: Optional custom brand hex.
        secondary_color: Optional custom brand accent hex.
        accent_color: Optional custom badge hex.
        font_family: Optional typography font.
        author: Author or organization name.
    """
    os.makedirs(output_dir, exist_ok=True)
    slug = "".join([c if c.isalnum() else "_" for c in project_title.lower()]).strip("_")

    pptx_path = os.path.join(output_dir, f"{slug}_deck.pptx")
    docx_path = os.path.join(output_dir, f"{slug}_report.docx")
    xlsx_path = os.path.join(output_dir, f"{slug}_model.xlsx")
    pdf_path = os.path.join(output_dir, f"{slug}_summary.pdf")

    res_pptx = build_presentation(
        title=project_title,
        slides=presentation_slides,
        output_path=pptx_path,
        subtitle=subtitle,
        theme=theme,
        primary_color=primary_color,
        secondary_color=secondary_color,
        accent_color=accent_color,
        font_family=font_family,
        author=author
    )

    res_docx = build_document(
        title=project_title,
        sections=document_sections,
        output_path=docx_path,
        subtitle=subtitle,
        author=author,
        theme=theme,
        primary_color=primary_color,
        secondary_color=secondary_color,
        accent_color=accent_color,
        font_family=font_family
    )

    res_xlsx = build_spreadsheet(
        title=project_title,
        sheets=spreadsheet_sheets,
        output_path=xlsx_path,
        theme=theme,
        primary_color=primary_color,
        secondary_color=secondary_color,
        accent_color=accent_color,
        font_family=font_family
    )

    res_pdf = build_pdf(
        title=project_title,
        sections=document_sections,
        output_path=pdf_path,
        subtitle=subtitle,
        author=author,
        theme=theme,
        primary_color=primary_color,
        secondary_color=secondary_color,
        accent_color=accent_color,
        font_family=font_family
    )

    return {
        "status": "success",
        "project_title": project_title,
        "theme_applied": theme,
        "generated_files": {
            "presentation_pptx": res_pptx.get("output_path"),
            "document_docx": res_docx.get("output_path"),
            "spreadsheet_xlsx": res_xlsx.get("output_path"),
            "report_pdf": res_pdf.get("output_path")
        }
    }


if __name__ == "__main__":
    mcp.run(transport="stdio")
