"""FastMCP server for Excel operations"""

from typing import Any

from fastmcp import FastMCP

from . import operations
from .models import (
    AlignmentFormatRequest,
    BorderFormatRequest,
    CellReadRequest,
    CellWriteRequest,
    FillFormatRequest,
    FontFormatRequest,
    NumberFormatRequest,
    RangeReadRequest,
    RangeWriteRequest,
    SheetCreateRequest,
    SheetRenameRequest,
)

# Initialize FastMCP server
mcp = FastMCP("Excel MCP Server")


# ==================== WORKBOOK OPERATIONS ====================

@mcp.tool()
def create_workbook(file_path: str) -> dict:
    """
    Create a new Excel workbook.

    Args:
        file_path: Path where the workbook will be created (must end with .xlsx)

    Returns:
        Dictionary with success status and message
    """
    result = operations.workbook.create(file_path)
    return result.model_dump()


@mcp.tool()
def get_workbook_info(file_path: str) -> dict:
    """
    Get information about an Excel workbook (sheets, size, etc.).

    Args:
        file_path: Path to the Excel workbook

    Returns:
        Dictionary with workbook metadata
    """
    return operations.workbook.get_info(file_path)


@mcp.tool()
def list_sheets(file_path: str) -> dict:
    """
    List all worksheet names in a workbook.

    Args:
        file_path: Path to the Excel workbook

    Returns:
        Dictionary with list of sheet names
    """
    try:
        sheets = operations.workbook.list_sheets(file_path)
        return {"success": True, "sheets": sheets, "count": len(sheets)}
    except Exception as e:
        return {"success": False, "error": str(e)}


# ==================== SHEET OPERATIONS ====================

@mcp.tool()
def create_sheet(workbook_path: str, sheet_name: str, index: int | None = None) -> dict:
    """
    Create a new worksheet in the workbook.

    Args:
        workbook_path: Path to the Excel workbook
        sheet_name: Name for the new worksheet
        index: Optional position to insert the sheet (0-based)

    Returns:
        Dictionary with success status and message
    """
    request = SheetCreateRequest(workbook_path=workbook_path, sheet_name=sheet_name, index=index)
    result = operations.sheet.create(request)
    return result.model_dump()


@mcp.tool()
def delete_sheet(workbook_path: str, sheet_name: str) -> dict:
    """
    Delete a worksheet from the workbook.

    Args:
        workbook_path: Path to the Excel workbook
        sheet_name: Name of the worksheet to delete

    Returns:
        Dictionary with success status and message
    """
    result = operations.sheet.delete(workbook_path, sheet_name)
    return result.model_dump()


@mcp.tool()
def rename_sheet(workbook_path: str, old_name: str, new_name: str) -> dict:
    """
    Rename a worksheet.

    Args:
        workbook_path: Path to the Excel workbook
        old_name: Current name of the worksheet
        new_name: New name for the worksheet

    Returns:
        Dictionary with success status and message
    """
    request = SheetRenameRequest(workbook_path=workbook_path, old_name=old_name, new_name=new_name)
    result = operations.sheet.rename(request)
    return result.model_dump()


@mcp.tool()
def copy_sheet(workbook_path: str, source_sheet: str, new_name: str) -> dict:
    """
    Copy a worksheet within the workbook.

    Args:
        workbook_path: Path to the Excel workbook
        source_sheet: Name of the worksheet to copy
        new_name: Name for the copied worksheet

    Returns:
        Dictionary with success status and message
    """
    result = operations.sheet.copy_sheet(workbook_path, source_sheet, new_name)
    return result.model_dump()


# ==================== CELL OPERATIONS ====================

@mcp.tool()
def write_cell(workbook_path: str, sheet_name: str, cell: str, value: Any) -> dict:
    """
    Write a value to a specific cell.

    Args:
        workbook_path: Path to the Excel workbook
        sheet_name: Name of the worksheet
        cell: Cell reference (e.g., 'A1', 'B10')
        value: Value to write (string, number, boolean, etc.)

    Returns:
        Dictionary with success status and the value written
    """
    request = CellWriteRequest(workbook_path=workbook_path, sheet_name=sheet_name, cell=cell, value=value)
    result = operations.cell.write_cell_value(request)
    return result.model_dump()


@mcp.tool()
def read_cell(workbook_path: str, sheet_name: str, cell: str) -> dict:
    """
    Read a value from a specific cell.

    Args:
        workbook_path: Path to the Excel workbook
        sheet_name: Name of the worksheet
        cell: Cell reference (e.g., 'A1', 'B10')

    Returns:
        Dictionary with the cell value
    """
    request = CellReadRequest(workbook_path=workbook_path, sheet_name=sheet_name, cell=cell)
    result = operations.cell.read_cell_value(request)
    return result.model_dump()


@mcp.tool()
def write_range(workbook_path: str, sheet_name: str, start_cell: str, data: list[list[Any]]) -> dict:
    """
    Write data to a range of cells.

    Args:
        workbook_path: Path to the Excel workbook
        sheet_name: Name of the worksheet
        start_cell: Top-left cell of the range (e.g., 'A1')
        data: 2D list of values to write [[row1], [row2], ...]

    Returns:
        Dictionary with success status and range info
    """
    request = RangeWriteRequest(workbook_path=workbook_path, sheet_name=sheet_name, start_cell=start_cell, data=data)
    result = operations.cell.write_range_values(request)
    return result.model_dump()


@mcp.tool()
def read_range(workbook_path: str, sheet_name: str, range_ref: str) -> dict:
    """
    Read data from a range of cells.

    Args:
        workbook_path: Path to the Excel workbook
        sheet_name: Name of the worksheet
        range_ref: Range reference (e.g., 'A1:D10')

    Returns:
        Dictionary with the data from the range
    """
    request = RangeReadRequest(workbook_path=workbook_path, sheet_name=sheet_name, range_ref=range_ref)
    result = operations.cell.read_range_values(request)
    return result.model_dump()


@mcp.tool()
def write_formula(workbook_path: str, sheet_name: str, cell: str, formula: str) -> dict:
    """
    Write a formula to a cell.

    Args:
        workbook_path: Path to the Excel workbook
        sheet_name: Name of the worksheet
        cell: Cell reference (e.g., 'A1')
        formula: Excel formula (e.g., '=SUM(A1:A10)', '=B2*C2')

    Returns:
        Dictionary with success status
    """
    result = operations.cell.write_formula(workbook_path, sheet_name, cell, formula)
    return result.model_dump()


# ==================== FORMATTING OPERATIONS ====================

@mcp.tool()
def format_font(
    workbook_path: str,
    sheet_name: str,
    range_ref: str,
    font_name: str | None = None,
    font_size: int | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: str | None = None,
    color: str | None = None,
) -> dict:
    """
    Apply font formatting to a range of cells.

    Args:
        workbook_path: Path to the Excel workbook
        sheet_name: Name of the worksheet
        range_ref: Range to format (e.g., 'A1:B10')
        font_name: Font name (e.g., 'Arial', 'Calibri')
        font_size: Font size (8-72)
        bold: Bold text
        italic: Italic text
        underline: Underline style ('single', 'double', or None)
        color: Hex color code (e.g., 'FF0000' for red)

    Returns:
        Dictionary with success status
    """
    request = FontFormatRequest(
        workbook_path=workbook_path,
        sheet_name=sheet_name,
        range_ref=range_ref,
        font_name=font_name,
        font_size=font_size,
        bold=bold,
        italic=italic,
        underline=underline,
        color=color,
    )
    result = operations.formatting.format_font(request)
    return result.model_dump()


@mcp.tool()
def format_fill(workbook_path: str, sheet_name: str, range_ref: str, color: str, fill_type: str = "solid") -> dict:
    """
    Apply background color to a range of cells.

    Args:
        workbook_path: Path to the Excel workbook
        sheet_name: Name of the worksheet
        range_ref: Range to format (e.g., 'A1:B10')
        color: Hex color code (e.g., 'FFFF00' for yellow)
        fill_type: Fill type ('solid' or 'pattern')

    Returns:
        Dictionary with success status
    """
    request = FillFormatRequest(
        workbook_path=workbook_path, sheet_name=sheet_name, range_ref=range_ref, color=color, fill_type=fill_type
    )
    result = operations.formatting.format_fill(request)
    return result.model_dump()


@mcp.tool()
def format_border(
    workbook_path: str,
    sheet_name: str,
    range_ref: str,
    style: str = "thin",
    color: str | None = None,
    sides: list[str] | None = None,
) -> dict:
    """
    Apply border formatting to a range of cells.

    Args:
        workbook_path: Path to the Excel workbook
        sheet_name: Name of the worksheet
        range_ref: Range to format (e.g., 'A1:B10')
        style: Border style ('thin', 'medium', 'thick', 'double')
        color: Hex color code for border
        sides: Which sides to apply border to (['top', 'bottom', 'left', 'right'])

    Returns:
        Dictionary with success status
    """
    if sides is None:
        sides = ["top", "bottom", "left", "right"]

    request = BorderFormatRequest(
        workbook_path=workbook_path, sheet_name=sheet_name, range_ref=range_ref, style=style, color=color, sides=sides
    )
    result = operations.formatting.format_border(request)
    return result.model_dump()


@mcp.tool()
def format_alignment(
    workbook_path: str,
    sheet_name: str,
    range_ref: str,
    horizontal: str | None = None,
    vertical: str | None = None,
    wrap_text: bool | None = None,
    text_rotation: int | None = None,
) -> dict:
    """
    Apply alignment formatting to a range of cells.

    Args:
        workbook_path: Path to the Excel workbook
        sheet_name: Name of the worksheet
        range_ref: Range to format (e.g., 'A1:B10')
        horizontal: Horizontal alignment ('left', 'center', 'right')
        vertical: Vertical alignment ('top', 'center', 'bottom')
        wrap_text: Enable text wrapping
        text_rotation: Text rotation angle (0-180)

    Returns:
        Dictionary with success status
    """
    request = AlignmentFormatRequest(
        workbook_path=workbook_path,
        sheet_name=sheet_name,
        range_ref=range_ref,
        horizontal=horizontal,
        vertical=vertical,
        wrap_text=wrap_text,
        text_rotation=text_rotation,
    )
    result = operations.formatting.format_alignment(request)
    return result.model_dump()


@mcp.tool()
def format_number(workbook_path: str, sheet_name: str, range_ref: str, format_string: str) -> dict:
    """
    Apply number formatting to a range of cells.

    Args:
        workbook_path: Path to the Excel workbook
        sheet_name: Name of the worksheet
        range_ref: Range to format (e.g., 'A1:B10')
        format_string: Excel number format string
            Examples:
            - '0.00' = Two decimal places
            - '#,##0' = Thousands separator
            - '0%' = Percentage
            - '$#,##0.00' = Currency
            - 'mm/dd/yyyy' = Date format

    Returns:
        Dictionary with success status
    """
    request = NumberFormatRequest(
        workbook_path=workbook_path, sheet_name=sheet_name, range_ref=range_ref, format_string=format_string
    )
    result = operations.formatting.format_number(request)
    return result.model_dump()


# ==================== ENTRY POINT ====================

if __name__ == "__main__":
    mcp.run()
