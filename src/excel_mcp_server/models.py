"""Pydantic models for Excel operations"""

import re
from typing import Any

from pydantic import BaseModel, Field, field_validator

# ========== Base Result Models ==========

class OperationResult(BaseModel):
    """Base result model for all operations"""

    success: bool
    message: str | None = None


class WorkbookResult(OperationResult):
    """Result for workbook operations"""

    file_path: str | None = None


class SheetResult(OperationResult):
    """Result for sheet operations"""

    sheet_name: str | None = None


class CellResult(OperationResult):
    """Result for cell operations"""

    cell: str | None = None
    value: Any | None = None


class RangeResult(OperationResult):
    """Result for range operations"""

    range: str | None = None
    rows: int | None = None
    cols: int | None = None
    data: list[list[Any]] | None = None


# ========== Request Models ==========

class WorkbookInfo(BaseModel):
    """Workbook metadata"""

    file_path: str
    sheets: list[str]
    sheet_count: int
    file_size: int | None = None


class CellWriteRequest(BaseModel):
    """Request to write a value to a cell"""

    workbook_path: str = Field(..., description="Path to the Excel workbook")
    sheet_name: str = Field(..., description="Name of the worksheet")
    cell: str = Field(..., description="Cell reference (e.g., 'A1')")
    value: Any = Field(..., description="Value to write to the cell")

    @field_validator("cell")
    @classmethod
    def validate_cell(cls, v: str) -> str:
        """Validate cell reference format"""
        if not re.match(r"^[A-Z]{1,3}[1-9]\d*$", v.upper()):
            raise ValueError(f"Invalid cell reference: {v}")
        return v.upper()


class CellReadRequest(BaseModel):
    """Request to read a cell value"""

    workbook_path: str
    sheet_name: str
    cell: str

    @field_validator("cell")
    @classmethod
    def validate_cell(cls, v: str) -> str:
        if not re.match(r"^[A-Z]{1,3}[1-9]\d*$", v.upper()):
            raise ValueError(f"Invalid cell reference: {v}")
        return v.upper()


class RangeWriteRequest(BaseModel):
    """Request to write data to a range"""

    workbook_path: str
    sheet_name: str
    start_cell: str = Field(..., description="Top-left cell of the range (e.g., 'A1')")
    data: list[list[Any]] = Field(..., description="2D list of values to write")

    @field_validator("start_cell")
    @classmethod
    def validate_cell(cls, v: str) -> str:
        if not re.match(r"^[A-Z]{1,3}[1-9]\d*$", v.upper()):
            raise ValueError(f"Invalid cell reference: {v}")
        return v.upper()


class RangeReadRequest(BaseModel):
    """Request to read a range of cells"""

    workbook_path: str
    sheet_name: str
    range_ref: str = Field(..., description="Range reference (e.g., 'A1:D10')")

    @field_validator("range_ref")
    @classmethod
    def validate_range(cls, v: str) -> str:
        """Validate range reference format"""
        if not re.match(r"^[A-Z]{1,3}[1-9]\d*:[A-Z]{1,3}[1-9]\d*$", v.upper()):
            raise ValueError(f"Invalid range reference: {v}")
        return v.upper()


class SheetCreateRequest(BaseModel):
    """Request to create a new sheet"""

    workbook_path: str
    sheet_name: str
    index: int | None = Field(None, description="Position to insert the sheet")


class SheetRenameRequest(BaseModel):
    """Request to rename a sheet"""

    workbook_path: str
    old_name: str
    new_name: str


class FormulaWriteRequest(BaseModel):
    """Request to write a formula to a cell"""

    workbook_path: str
    sheet_name: str
    cell: str
    formula: str = Field(..., description="Excel formula (should start with '=')")

    @field_validator("cell")
    @classmethod
    def validate_cell(cls, v: str) -> str:
        if not re.match(r"^[A-Z]{1,3}[1-9]\d*$", v.upper()):
            raise ValueError(f"Invalid cell reference: {v}")
        return v.upper()

    @field_validator("formula")
    @classmethod
    def validate_formula(cls, v: str) -> str:
        """Ensure formula starts with ="""
        if not v.startswith("="):
            return f"={v}"
        return v


# ========== Formatting Models ==========

class FontFormatRequest(BaseModel):
    """Request to format font"""

    workbook_path: str
    sheet_name: str
    range_ref: str = Field(..., description="Range to format (e.g., 'A1:B10')")
    font_name: str | None = Field(None, description="Font name (e.g., 'Arial', 'Calibri')")
    font_size: int | None = Field(None, ge=8, le=72, description="Font size (8-72)")
    bold: bool | None = None
    italic: bool | None = None
    underline: str | None = Field(None, description="'single', 'double', or None")
    color: str | None = Field(None, description="Hex color code (e.g., 'FF0000')")

    @field_validator("color")
    @classmethod
    def validate_color(cls, v: str | None) -> str | None:
        """Validate hex color format"""
        if v is None:
            return v
        # Remove # if present
        v = v.lstrip("#")
        if not re.match(r"^[0-9A-Fa-f]{6}$", v):
            raise ValueError(f"Invalid hex color: {v}")
        return v.upper()


class FillFormatRequest(BaseModel):
    """Request to format cell fill (background color)"""

    workbook_path: str
    sheet_name: str
    range_ref: str
    fill_type: str = Field("solid", description="Fill type: 'solid' or 'pattern'")
    color: str = Field(..., description="Hex color code (e.g., 'FFFF00')")

    @field_validator("color")
    @classmethod
    def validate_color(cls, v: str) -> str:
        v = v.lstrip("#")
        if not re.match(r"^[0-9A-Fa-f]{6}$", v):
            raise ValueError(f"Invalid hex color: {v}")
        return v.upper()


class BorderFormatRequest(BaseModel):
    """Request to format cell borders"""

    workbook_path: str
    sheet_name: str
    range_ref: str
    style: str = Field("thin", description="Border style: 'thin', 'medium', 'thick', 'double'")
    color: str | None = Field(None, description="Hex color code")
    sides: list[str] = Field(
        ["top", "bottom", "left", "right"], description="Which sides to apply border to"
    )

    @field_validator("color")
    @classmethod
    def validate_color(cls, v: str | None) -> str | None:
        if v is None:
            return v
        v = v.lstrip("#")
        if not re.match(r"^[0-9A-Fa-f]{6}$", v):
            raise ValueError(f"Invalid hex color: {v}")
        return v.upper()

    @field_validator("sides")
    @classmethod
    def validate_sides(cls, v: list[str]) -> list[str]:
        valid_sides = {"top", "bottom", "left", "right"}
        for side in v:
            if side not in valid_sides:
                raise ValueError(f"Invalid side: {side}. Must be one of {valid_sides}")
        return v


class AlignmentFormatRequest(BaseModel):
    """Request to format cell alignment"""

    workbook_path: str
    sheet_name: str
    range_ref: str
    horizontal: str | None = Field(
        None, description="Horizontal alignment: 'left', 'center', 'right'"
    )
    vertical: str | None = Field(None, description="Vertical alignment: 'top', 'center', 'bottom'")
    wrap_text: bool | None = None
    text_rotation: int | None = Field(None, ge=0, le=180, description="Text rotation (0-180)")


class NumberFormatRequest(BaseModel):
    """Request to format numbers"""

    workbook_path: str
    sheet_name: str
    range_ref: str
    format_string: str = Field(
        ...,
        description="Excel number format string (e.g., '0.00', '#,##0', '0%', '$#,##0.00', 'mm/dd/yyyy')",
    )
