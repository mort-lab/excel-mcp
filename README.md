# Excel MCP Server

A comprehensive [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that enables AI assistants to perform Excel file operations without requiring Microsoft Excel installation.

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

### Core Operations

- **📊 Workbook Management**: Create, open, save Excel workbooks
- **📄 Sheet Operations**: Create, delete, rename, copy worksheets
- **📝 Cell Operations**: Read and write individual cells or ranges
- **🔢 Formula Support**: Write and evaluate Excel formulas
- **🎨 Rich Formatting**: Fonts, colors, borders, alignment, number formats

### Key Capabilities

- ✅ No Excel installation required (uses openpyxl)
- ✅ Type-safe operations with Pydantic validation
- ✅ Comprehensive error handling
- ✅ Support for stdio transport (Claude Desktop, etc.)
- ✅ Well-tested with pytest
- ✅ Full Python type hints

## Installation

### Using uvx (Recommended)

```bash
uvx excel-mcp-server
```

### Using pip

```bash
pip install excel-mcp-server
```

### From Source

```bash
git clone https://github.com/mort-lab/excel-mcp
cd excel-mcp-server
uv sync
```

## Quick Start

### Claude Desktop Configuration

Add to your Claude Desktop config file:

**MacOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows**: `%APPDATA%/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["excel-mcp-server"]
    }
  }
}
```

Restart Claude Desktop and you can start using Excel operations!

### Example Usage

Once configured, you can ask Claude to:

```
"Create a new Excel workbook called sales_report.xlsx"
"Write 'Product' to cell A1 in Sheet1"
"Write a table of sales data starting at A1"
"Format the header row with bold text and blue background"
"Add a formula in D2 that multiplies B2 and C2"
"Read the data from range A1:D10"
```

## Available Tools

### Workbook Operations

- `create_workbook(file_path)` - Create a new Excel workbook
- `get_workbook_info(file_path)` - Get workbook metadata
- `list_sheets(file_path)` - List all worksheet names

### Sheet Operations

- `create_sheet(workbook_path, sheet_name, index?)` - Create a new worksheet
- `delete_sheet(workbook_path, sheet_name)` - Delete a worksheet
- `rename_sheet(workbook_path, old_name, new_name)` - Rename a worksheet
- `copy_sheet(workbook_path, source_sheet, new_name)` - Copy a worksheet

### Cell Operations

- `write_cell(workbook_path, sheet_name, cell, value)` - Write to a cell
- `read_cell(workbook_path, sheet_name, cell)` - Read from a cell
- `write_range(workbook_path, sheet_name, start_cell, data)` - Write data to a range
- `read_range(workbook_path, sheet_name, range_ref)` - Read data from a range
- `write_formula(workbook_path, sheet_name, cell, formula)` - Write a formula

### Formatting Operations

- `format_font(...)` - Apply font formatting (bold, italic, color, size)
- `format_fill(...)` - Apply background color
- `format_border(...)` - Apply cell borders
- `format_alignment(...)` - Apply text alignment
- `format_number(...)` - Apply number formatting

For complete API documentation, see [TOOLS.md](TOOLS.md).

## Development

### Setup

```bash
# Clone the repository
git clone https://github.com/yourusername/excel-mcp-server
cd excel-mcp-server

# Install dependencies
uv sync

# Run tests
uv run pytest

# Run with coverage
uv run pytest --cov=excel_mcp_server --cov-report=term-missing

# Format code
uv run ruff format

# Lint
uv run ruff check
```

### Project Structure

```
excel-mcp-server/
├── src/excel_mcp_server/
│   ├── server.py          # FastMCP server with tool definitions
│   ├── models.py          # Pydantic models for validation
│   ├── operations/        # Business logic
│   │   ├── workbook.py
│   │   ├── cell.py
│   │   ├── sheet.py
│   │   └── formatting.py
│   └── utils/
│       └── validators.py  # Input validation utilities
├── tests/                 # Test suite
└── pyproject.toml        # Project configuration
```

## Requirements

- Python 3.10 or higher
- Dependencies:
  - `fastmcp` >= 0.1.0
  - `openpyxl` >= 3.1.0
  - `pydantic` >= 2.0.0

## Examples

### Create a Report

```python
# Ask Claude:
"Create a new workbook called monthly_report.xlsx with these steps:
1. Create a sheet called 'Sales'
2. Write headers: Product, Quantity, Price, Total in row 1
3. Add 5 rows of sample data
4. Make headers bold with blue background
5. Format the Total column as currency
6. Add a formula in each Total cell (Quantity * Price)"
```

### Data Analysis

```python
# Ask Claude:
"Read the data from sales.xlsx, range A1:D100,
then calculate the sum of the Total column and
write it to cell D101 with the label 'Grand Total:' in C101"
```

## Troubleshooting

### Import Errors

If you see import errors, make sure dependencies are installed:

```bash
uv sync
```

### Permission Errors

Ensure the directory where you're creating/modifying files has write permissions.

### File Already Exists

When creating a workbook, if the file already exists, the operation will fail. Delete the existing file or use a different name.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with [FastMCP](https://github.com/jlowin/fastmcp)
- Excel operations powered by [openpyxl](https://openpyxl.readthedocs.io/)
- Inspired by the [Model Context Protocol](https://modelcontextprotocol.io/)

## Roadmap

- [ ] Support for charts and graphs
- [ ] Pivot table operations
- [ ] Data validation rules
- [ ] Conditional formatting
- [ ] Image insertion
- [ ] HTTP transport for remote access
- [ ] CSV import/export
- [ ] PDF export

## Support

If you encounter any issues or have questions:

- Open an issue on [GitHub](https://github.com/yourusername/excel-mcp-server/issues)
- Check the [documentation](TOOLS.md)
- Review existing issues for solutions

---

Made with ❤️ by Martin Irurozki
