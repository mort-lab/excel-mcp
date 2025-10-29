"""Tests for workbook operations"""

import pytest
from pathlib import Path
from excel_mcp_server.operations import workbook


def test_create_workbook(temp_dir):
    """Test creating a new workbook"""
    file_path = temp_dir / "new_workbook.xlsx"
    result = workbook.create(str(file_path))

    assert result.success is True
    assert Path(file_path).exists()
    assert result.file_path == str(file_path)


def test_create_workbook_already_exists(sample_workbook):
    """Test that creating an existing workbook fails"""
    result = workbook.create(sample_workbook)

    assert result.success is False
    assert "already exists" in result.message.lower()


def test_create_workbook_invalid_extension(temp_dir):
    """Test that creating a workbook with invalid extension fails"""
    file_path = temp_dir / "test.txt"
    result = workbook.create(str(file_path))

    assert result.success is False
    assert "xlsx" in result.message.lower()


def test_open_file(sample_workbook):
    """Test opening an existing workbook"""
    info = workbook.open_file(sample_workbook)

    assert info.file_path == sample_workbook
    assert len(info.sheets) > 0
    assert info.sheet_count == len(info.sheets)
    assert info.file_size > 0


def test_open_file_not_found(temp_dir):
    """Test opening a non-existent file"""
    file_path = temp_dir / "nonexistent.xlsx"

    with pytest.raises(ValueError, match="not found"):
        workbook.open_file(str(file_path))


def test_list_sheets(sample_workbook):
    """Test listing sheets in a workbook"""
    sheets = workbook.list_sheets(sample_workbook)

    assert isinstance(sheets, list)
    assert len(sheets) > 0
    assert "Sheet1" in sheets


def test_save_workbook(sample_workbook):
    """Test saving a workbook"""
    result = workbook.save(sample_workbook)

    assert result.success is True
    assert result.file_path == sample_workbook


def test_get_info(sample_workbook):
    """Test getting workbook info"""
    info = workbook.get_info(sample_workbook)

    assert "file_path" in info
    assert "sheets" in info
    assert "sheet_count" in info
    assert info["sheet_count"] == len(info["sheets"])
