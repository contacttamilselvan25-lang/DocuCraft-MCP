"""Excel Workbook Generator for DocuCraft MCP.

Generates beautiful, executive-formatted Excel spreadsheets (.xlsx) with custom color themes,
KPI cards, auto-calculated summary formulas, number formatting, and auto-adjusted column widths.
"""

import os
from typing import Dict, List, Optional, Any
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from themes import get_theme_config


def build_spreadsheet(
    title: str,
    sheets: List[Dict[str, Any]],
    output_path: str,
    theme: str = "corporate_blue",
    primary_color: Optional[str] = None,
    secondary_color: Optional[str] = None,
    accent_color: Optional[str] = None,
    font_family: Optional[str] = None
) -> Dict[str, Any]:
    """Generate a themed Excel workbook (.xlsx).

    Args:
        title: Overall workbook title.
        sheets: List of sheet dicts (sheet_name, kpi_cards, headers, rows, include_totals).
        output_path: Target .xlsx file path.
        theme: Theme preset name.
        primary_color: Custom primary hex.
        secondary_color: Custom secondary hex.
        accent_color: Custom accent hex.
        font_family: Custom font family.

    Returns:
        Dict with workbook summary and output path.
    """
    theme_cfg = get_theme_config(
        theme_name=theme,
        primary_color=primary_color,
        secondary_color=secondary_color,
        accent_color=accent_color,
        font_family=font_family
    )

    wb = Workbook()
    # Remove default sheet
    default_sheet = wb.active

    font_name = theme_cfg["font_heading"]
    p_hex = theme_cfg["primary"].lstrip("#")
    s_hex = theme_cfg["secondary"].lstrip("#")
    bg_hex = theme_cfg["bg_light"].lstrip("#")
    td_hex = theme_cfg["text_dark"].lstrip("#")

    # Fills
    primary_fill = PatternFill(start_color=p_hex, end_color=p_hex, fill_type="solid")
    secondary_fill = PatternFill(start_color=s_hex, end_color=s_hex, fill_type="solid")
    zebra_fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    # Fonts
    title_font = Font(name=font_name, size=16, bold=True, color="FFFFFF")
    kpi_val_font = Font(name=font_name, size=18, bold=True, color=p_hex)
    kpi_lbl_font = Font(name=font_name, size=9, bold=True, color="64748B")
    hdr_font = Font(name=font_name, size=11, bold=True, color="FFFFFF")
    data_font = Font(name=font_name, size=10, color=td_hex)
    total_font = Font(name=font_name, size=11, bold=True, color=p_hex)

    # Borders
    thin_border_side = Side(border_style="thin", color="CBD5E1")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    double_bottom_border = Border(
        top=Side(border_style="thin", color=p_hex),
        bottom=Side(border_style="double", color=p_hex)
    )

    for s_idx, s_data in enumerate(sheets):
        sheet_title = s_data.get("sheet_name", f"Sheet{s_idx + 1}")[:30]
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True

        current_row = 1

        # 1. Main Title Banner (merged across 6 columns)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 1, end_column=8)
        top_cell = ws.cell(row=current_row, column=1)
        top_cell.value = f"  {title.upper()}  ―  {sheet_title}"
        top_cell.font = title_font
        top_cell.alignment = Alignment(vertical="center")

        for r in range(current_row, current_row + 2):
            for c in range(1, 9):
                ws.cell(row=r, column=c).fill = primary_fill
        current_row += 3

        # 2. KPI Summary Cards (if provided)
        kpis = s_data.get("kpi_cards", [])
        if kpis:
            col_pos = 1
            for kpi in kpis[:4]:
                # 2x2 cell card
                ws.merge_cells(start_row=current_row, start_column=col_pos, end_row=current_row, end_column=col_pos + 1)
                lbl_cell = ws.cell(row=current_row, column=col_pos)
                lbl_cell.value = str(kpi.get("label", "")).upper()
                lbl_cell.font = kpi_lbl_font
                lbl_cell.fill = zebra_fill
                lbl_cell.alignment = Alignment(horizontal="center", vertical="center")

                ws.merge_cells(start_row=current_row + 1, start_column=col_pos, end_row=current_row + 1, end_column=col_pos + 1)
                val_cell = ws.cell(row=current_row + 1, column=col_pos)
                val_cell.value = str(kpi.get("value", ""))
                val_cell.font = kpi_val_font
                val_cell.fill = zebra_fill
                val_cell.alignment = Alignment(horizontal="center", vertical="center")

                # Borders for card
                for r in range(current_row, current_row + 2):
                    for c in range(col_pos, col_pos + 2):
                        ws.cell(row=r, column=c).border = cell_border

                col_pos += 2
            current_row += 3

        # 3. Data Table Headers
        headers = s_data.get("headers", [])
        rows = s_data.get("rows", [])
        formats = s_data.get("column_formats", {})  # e.g. {"Amount": "currency", "Margin": "percent"}

        if headers:
            tbl_start_row = current_row
            ws.freeze_panes = ws.cell(row=current_row + 1, column=1)

            for c_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=current_row, column=c_idx)
                cell.value = h
                cell.font = hdr_font
                cell.fill = secondary_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = cell_border
            ws.row_dimensions[current_row].height = 26
            current_row += 1

            # 4. Data Rows
            data_start_row = current_row
            for r_idx, row_vals in enumerate(rows):
                is_even = (r_idx % 2 == 0)
                row_fill = zebra_fill if is_even else white_fill
                for c_idx, val in enumerate(row_vals, start=1):
                    cell = ws.cell(row=current_row, column=c_idx)
                    cell.value = val
                    cell.font = data_font
                    cell.fill = row_fill
                    cell.border = cell_border
                    cell.alignment = Alignment(vertical="center")

                    # Apply number formatting
                    h_name = headers[c_idx - 1] if c_idx - 1 < len(headers) else ""
                    fmt = formats.get(h_name, "").lower()
                    if fmt == "currency" or any(curr in h_name.lower() for curr in ["price", "cost", "revenue", "amount", "budget", "$"]):
                        cell.number_format = '"$"#,##0.00'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif fmt == "percent" or "%" in h_name or "rate" in h_name.lower():
                        cell.number_format = "0.0%"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif isinstance(val, (int, float)):
                        cell.number_format = "#,##0"
                        cell.alignment = Alignment(horizontal="right", vertical="center")

                ws.row_dimensions[current_row].height = 20
                current_row += 1

            data_end_row = current_row - 1

            # 5. Summary / Totals Row
            if s_data.get("include_totals", True) and len(rows) > 0:
                ws.cell(row=current_row, column=1).value = "TOTAL / AVERAGE"
                ws.cell(row=current_row, column=1).font = total_font
                ws.cell(row=current_row, column=1).border = double_bottom_border

                for c_idx, h in enumerate(headers, start=1):
                    cell = ws.cell(row=current_row, column=c_idx)
                    cell.border = double_bottom_border
                    if c_idx > 1:
                        # Check if column is numeric
                        first_val = rows[0][c_idx - 1] if len(rows[0]) >= c_idx else None
                        if isinstance(first_val, (int, float)):
                            col_letter = get_column_letter(c_idx)
                            h_lower = h.lower()
                            if "avg" in h_lower or "average" in h_lower or "rate" in h_lower or "pct" in h_lower:
                                cell.value = f"=AVERAGE({col_letter}{data_start_row}:{col_letter}{data_end_row})"
                            else:
                                cell.value = f"=SUM({col_letter}{data_start_row}:{col_letter}{data_end_row})"
                            cell.font = total_font

                            fmt = formats.get(h, "").lower()
                            if fmt == "currency" or any(curr in h_lower for curr in ["price", "cost", "revenue", "amount", "$"]):
                                cell.number_format = '"$"#,##0.00'
                            elif fmt == "percent" or "%" in h:
                                cell.number_format = "0.0%"
                            else:
                                cell.number_format = "#,##0"
                current_row += 1

        # 6. Auto-fit column widths
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                # Ignore merged title row length
                if cell.row > 2 and len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Remove the default empty sheet created by openpyxl
    if default_sheet in wb.worksheets:
        wb.remove(default_sheet)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)

    return {
        "format": "xlsx",
        "title": title,
        "sheets_count": len(wb.worksheets),
        "theme_used": theme_cfg["name"],
        "output_path": os.path.abspath(output_path)
    }
