"""Tests for cell operations"""

import pytest
from excel_mcp_server.operations import cell
from excel_mcp_server.models import CellWriteRequest, CellReadRequest, RangeWriteRequest, RangeReadRequest


def test_write_cell(sample_workbook):
    """Test writing a value to a cell"""
    request = CellWriteRequest(workbook_path=sample_workbook, sheet_name="Sheet1", cell="D1", value="Test")
    result = cell.write_cell_value(request)

    assert result.success is True
    assert result.value == "Test"
    assert result.cell == "D1"


def test_read_cell(sample_workbook):
    """Test reading a value from a cell"""
    request = CellReadRequest(workbook_path=sample_workbook, sheet_name="Sheet1", cell="A1")
    result = cell.read_cell_value(request)

    assert result.success is True
    assert result.value == "Name"


def test_write_read_cell_number(sample_workbook):
    """Test writing and reading a number"""
    write_request = CellWriteRequest(workbook_path=sample_workbook, sheet_name="Sheet1", cell="E1", value=42)
    write_result = cell.write_cell_value(write_request)
    assert write_result.success is True

    read_request = CellReadRequest(workbook_path=sample_workbook, sheet_name="Sheet1", cell="E1")
    read_result = cell.read_cell_value(read_request)
    assert read_result.success is True
    assert read_result.value == 42


def test_write_cell_invalid_reference(sample_workbook):
    """Test writing to an invalid cell reference"""
    with pytest.raises(ValueError, match="Invalid cell reference"):
        CellWriteRequest(workbook_path=sample_workbook, sheet_name="Sheet1", cell="INVALID", value="Test")


def test_write_cell_sheet_not_found(sample_workbook):
    """Test writing to a non-existent sheet"""
    request = CellWriteRequest(workbook_path=sample_workbook, sheet_name="NonExistent", cell="A1", value="Test")
    result = cell.write_cell_value(request)

    assert result.success is False
    assert "not found" in result.message.lower()


def test_write_range(sample_workbook):
    """Test writing data to a range"""
    data = [["Header1", "Header2"], ["Value1", "Value2"], ["Value3", "Value4"]]
    request = RangeWriteRequest(workbook_path=sample_workbook, sheet_name="Sheet1", start_cell="F1", data=data)
    result = cell.write_range_values(request)

    assert result.success is True
    assert result.rows == 3
    assert result.cols == 2


def test_read_range(sample_workbook):
    """Test reading a range of cells"""
    request = RangeReadRequest(workbook_path=sample_workbook, sheet_name="Sheet1", range_ref="A1:C1")
    result = cell.read_range_values(request)

    assert result.success is True
    assert result.rows == 1
    assert result.cols == 3
    assert result.data[0] == ["Name", "Age", "City"]


def test_write_formula(sample_workbook):
    """Test writing a formula"""
    result = cell.write_formula(sample_workbook, "Sheet1", "D2", "=B2*2")

    assert result.success is True
    assert result.cell == "D2"


def test_write_formula_without_equals(sample_workbook):
    """Test writing a formula without = prefix"""
    result = cell.write_formula(sample_workbook, "Sheet1", "D3", "SUM(B2:B4)")

    assert result.success is True
    assert "=" in str(result.value)
