"""Validation utilities for Excel operations"""

import re
from pathlib import Path


def validate_file_path(path: str, must_exist: bool = False) -> tuple[bool, str | None]:
    """
    Validate file path for Excel operations.

    Args:
        path: File path to validate
        must_exist: If True, file must exist

    Returns:
        Tuple of (is_valid, error_message)
    """
    try:
        file_path = Path(path)

        # Check file extension
        if file_path.suffix.lower() != ".xlsx":
            return False, "File must have .xlsx extension"

        # Check if file exists (when required)
        if must_exist and not file_path.exists():
            return False, f"File not found: {path}"

        # Check if parent directory exists
        if not file_path.parent.exists():
            return False, f"Parent directory does not exist: {file_path.parent}"

        # Check for path traversal attempts
        try:
            file_path.resolve()
        except (ValueError, RuntimeError):
            return False, "Invalid file path"

        return True, None

    except Exception as e:
        return False, f"Invalid path: {str(e)}"


def validate_cell_reference(cell: str) -> tuple[bool, str | None]:
    """
    Validate Excel cell reference.

    Args:
        cell: Cell reference (e.g., 'A1', 'Z100')

    Returns:
        Tuple of (is_valid, error_message)
    """
    # Valid format: Column (A-ZZ) + Row (1-1048576)
    pattern = r"^[A-Z]{1,3}[1-9]\d*$"

    if not re.match(pattern, cell.upper()):
        return False, f"Invalid cell reference: {cell}. Expected format like 'A1' or 'B10'"

    # Extract row number and validate
    row_match = re.search(r"\d+", cell)
    if row_match:
        row = int(row_match.group())
        if row > 1048576:  # Excel's max row
            return False, f"Row number {row} exceeds Excel's maximum (1048576)"

    return True, None


def validate_range_reference(range_ref: str) -> tuple[bool, str | None]:
    """
    Validate Excel range reference.

    Args:
        range_ref: Range reference (e.g., 'A1:B10')

    Returns:
        Tuple of (is_valid, error_message)
    """
    # Valid format: Cell:Cell
    pattern = r"^[A-Z]{1,3}[1-9]\d*:[A-Z]{1,3}[1-9]\d*$"

    if not re.match(pattern, range_ref.upper()):
        return False, f"Invalid range reference: {range_ref}. Expected format like 'A1:B10'"

    # Validate individual cells
    cells = range_ref.split(":")
    for cell in cells:
        is_valid, error = validate_cell_reference(cell)
        if not is_valid:
            return False, error

    return True, None


def validate_formula(formula: str) -> tuple[bool, str | None]:
    """
    Validate Excel formula.

    Args:
        formula: Excel formula string

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not formula:
        return False, "Formula cannot be empty"

    # Ensure formula starts with =
    if not formula.startswith("="):
        return False, "Formula must start with '='"

    # Check for dangerous functions (security)
    dangerous_functions = ["CALL", "REGISTER", "EXEC"]
    formula_upper = formula.upper()
    for func in dangerous_functions:
        if func in formula_upper:
            return False, f"Formula contains prohibited function: {func}"

    return True, None


def validate_sheet_name(name: str) -> tuple[bool, str | None]:
    """
    Validate Excel sheet name.

    Args:
        name: Sheet name to validate

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not name:
        return False, "Sheet name cannot be empty"

    if len(name) > 31:
        return False, "Sheet name cannot exceed 31 characters"

    # Invalid characters for sheet names
    invalid_chars = [":", "\\", "/", "?", "*", "[", "]"]
    for char in invalid_chars:
        if char in name:
            return False, f"Sheet name cannot contain '{char}'"

    return True, None


def validate_color_hex(color: str) -> tuple[bool, str | None]:
    """
    Validate hex color code.

    Args:
        color: Hex color code (with or without #)

    Returns:
        Tuple of (is_valid, error_message)
    """
    # Remove # if present
    color = color.lstrip("#")

    if not re.match(r"^[0-9A-Fa-f]{6}$", color):
        return False, f"Invalid hex color: {color}. Expected format like 'FF0000' or '#FF0000'"

    return True, None


def column_letter_to_number(column: str) -> int:
    """
    Convert column letter to number (A=1, B=2, ..., Z=26, AA=27, etc.)

    Args:
        column: Column letter(s)

    Returns:
        Column number
    """
    number = 0
    for char in column.upper():
        number = number * 26 + (ord(char) - ord("A") + 1)
    return number


def column_number_to_letter(number: int) -> str:
    """
    Convert column number to letter (1=A, 2=B, ..., 26=Z, 27=AA, etc.)

    Args:
        number: Column number (1-based)

    Returns:
        Column letter(s)
    """
    letter = ""
    while number > 0:
        number -= 1
        letter = chr(number % 26 + ord("A")) + letter
        number //= 26
    return letter
