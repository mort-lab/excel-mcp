# Excel MCP Server - Tools Reference

Complete reference for all available tools in the Excel MCP Server.

## Table of Contents

- [Workbook Operations](#workbook-operations)
- [Sheet Operations](#sheet-operations)
- [Cell Operations](#cell-operations)
- [Formatting Operations](#formatting-operations)

---

## Workbook Operations

### create_workbook

Create a new Excel workbook.

**Parameters:**

- `file_path` (string, required): Path where the workbook will be created. Must end with `.xlsx`

**Returns:**

```json
{
  "success": true,
  "message": "Workbook created successfully",
  "file_path": "/path/to/workbook.xlsx"
}
```

**Example:**

```python
create_workbook("/Users/name/Documents/report.xlsx")
```

**Errors:**

- File already exists
- Invalid file extension
- Parent directory doesn't exist
- Permission denied

---

### get_workbook_info

Get comprehensive information about an Excel workbook.

**Parameters:**

- `file_path` (string, required): Path to the Excel workbook

**Returns:**

```json
{
  "file_path": "/path/to/workbook.xlsx",
  "sheets": ["Sheet1", "Sheet2"],
  "sheet_count": 2,
  "file_size": 8456
}
```

**Example:**

```python
get_workbook_info("/Users/name/Documents/report.xlsx")
```

---

### list_sheets

List all worksheet names in a workbook.

**Parameters:**

- `file_path` (string, required): Path to the Excel workbook

**Returns:**

```json
{
  "success": true,
  "sheets": ["Sheet1", "Sales", "Summary"],
  "count": 3
}
```

**Example:**

```python
list_sheets("/Users/name/Documents/report.xlsx")
```

---

## Sheet Operations

### create_sheet

Create a new worksheet in the workbook.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `sheet_name` (string, required): Name for the new worksheet
- `index` (integer, optional): Position to insert the sheet (0-based). If not provided, sheet is added at the end

**Returns:**

```json
{
  "success": true,
  "message": "Sheet 'Sales' created successfully",
  "sheet_name": "Sales"
}
```

**Example:**

```python
create_sheet("/path/to/workbook.xlsx", "Sales", index=0)
```

**Validation:**

- Sheet name cannot be empty
- Sheet name max 31 characters
- Cannot contain: `:`, `\`, `/`, `?`, `*`, `[`, `]`
- Sheet name must be unique

---

### delete_sheet

Delete a worksheet from the workbook.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `sheet_name` (string, required): Name of the worksheet to delete

**Returns:**

```json
{
  "success": true,
  "message": "Sheet 'OldData' deleted successfully",
  "sheet_name": "OldData"
}
```

**Example:**

```python
delete_sheet("/path/to/workbook.xlsx", "OldData")
```

**Errors:**

- Cannot delete the last sheet in a workbook
- Sheet not found

---

### rename_sheet

Rename a worksheet.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `old_name` (string, required): Current name of the worksheet
- `new_name` (string, required): New name for the worksheet

**Returns:**

```json
{
  "success": true,
  "message": "Sheet renamed from 'Sheet1' to 'Sales'",
  "sheet_name": "Sales"
}
```

**Example:**

```python
rename_sheet("/path/to/workbook.xlsx", "Sheet1", "Sales")
```

---

### copy_sheet

Copy a worksheet within the workbook.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `source_sheet` (string, required): Name of the worksheet to copy
- `new_name` (string, required): Name for the copied worksheet

**Returns:**

```json
{
  "success": true,
  "message": "Sheet 'Template' copied to 'January'",
  "sheet_name": "January"
}
```

**Example:**

```python
copy_sheet("/path/to/workbook.xlsx", "Template", "January")
```

---

## Cell Operations

### write_cell

Write a value to a specific cell.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `sheet_name` (string, required): Name of the worksheet
- `cell` (string, required): Cell reference (e.g., 'A1', 'B10')
- `value` (any, required): Value to write (string, number, boolean, etc.)

**Returns:**

```json
{
  "success": true,
  "message": "Value written to A1",
  "cell": "A1",
  "value": "Hello World"
}
```

**Example:**

```python
write_cell("/path/to/workbook.xlsx", "Sheet1", "A1", "Product Name")
write_cell("/path/to/workbook.xlsx", "Sheet1", "B1", 42)
write_cell("/path/to/workbook.xlsx", "Sheet1", "C1", True)
```

**Cell Reference Format:**

- Valid: `A1`, `B10`, `AA100`, `ZZ999`
- Invalid: `A`, `1`, `1A`, `AAA1000000`

---

### read_cell

Read a value from a specific cell.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `sheet_name` (string, required): Name of the worksheet
- `cell` (string, required): Cell reference (e.g., 'A1')

**Returns:**

```json
{
  "success": true,
  "message": "Value read from A1",
  "cell": "A1",
  "value": "Hello World"
}
```

**Example:**

```python
read_cell("/path/to/workbook.xlsx", "Sheet1", "A1")
```

---

### write_range

Write data to a range of cells.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `sheet_name` (string, required): Name of the worksheet
- `start_cell` (string, required): Top-left cell of the range (e.g., 'A1')
- `data` (array of arrays, required): 2D list of values `[[row1], [row2], ...]`

**Returns:**

```json
{
  "success": true,
  "message": "Data written to range starting at A1",
  "range": "A1",
  "rows": 3,
  "cols": 2
}
```

**Example:**

```python
data = [
    ["Product", "Price"],
    ["Widget", 10.99],
    ["Gadget", 24.99]
]
write_range("/path/to/workbook.xlsx", "Sheet1", "A1", data)
```

---

### read_range

Read data from a range of cells.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `sheet_name` (string, required): Name of the worksheet
- `range_ref` (string, required): Range reference (e.g., 'A1:D10')

**Returns:**

```json
{
  "success": true,
  "message": "Data read from range A1:C3",
  "range": "A1:C3",
  "rows": 3,
  "cols": 3,
  "data": [
    ["Name", "Age", "City"],
    ["Alice", 30, "NYC"],
    ["Bob", 25, "LA"]
  ]
}
```

**Example:**

```python
read_range("/path/to/workbook.xlsx", "Sheet1", "A1:D10")
```

**Range Format:**

- Valid: `A1:B10`, `A1:Z100`
- Invalid: `A:B`, `1:10`, `B10:A1`

---

### write_formula

Write a formula to a cell.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `sheet_name` (string, required): Name of the worksheet
- `cell` (string, required): Cell reference (e.g., 'D1')
- `formula` (string, required): Excel formula (can start with or without '=')

**Returns:**

```json
{
  "success": true,
  "message": "Formula written to D2",
  "cell": "D2",
  "value": "=B2*C2"
}
```

**Example:**

```python
write_formula("/path/to/workbook.xlsx", "Sheet1", "D2", "=B2*C2")
write_formula("/path/to/workbook.xlsx", "Sheet1", "E10", "=SUM(E2:E9)")
write_formula("/path/to/workbook.xlsx", "Sheet1", "F2", "=IF(D2>100, 'High', 'Low')")
```

**Common Formulas:**

- `=SUM(A1:A10)` - Sum a range
- `=AVERAGE(B1:B10)` - Average
- `=COUNT(C1:C10)` - Count numbers
- `=IF(D1>50, "Pass", "Fail")` - Conditional
- `=VLOOKUP(E2, A1:B10, 2, FALSE)` - Lookup
- `=CONCATENATE(A1, " ", B1)` - Combine text
- `=TODAY()` - Current date
- `=NOW()` - Current date and time

---

## Formatting Operations

### format_font

Apply font formatting to a range of cells.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `sheet_name` (string, required): Name of the worksheet
- `range_ref` (string, required): Range to format (e.g., 'A1:B10')
- `font_name` (string, optional): Font name (e.g., 'Arial', 'Calibri', 'Times New Roman')
- `font_size` (integer, optional): Font size (8-72)
- `bold` (boolean, optional): Bold text
- `italic` (boolean, optional): Italic text
- `underline` (string, optional): Underline style ('single', 'double', or None)
- `color` (string, optional): Hex color code (e.g., 'FF0000' for red, 'FFFF00' for yellow)

**Returns:**

```json
{
  "success": true,
  "message": "Font formatting applied to A1:B1"
}
```

**Example:**

```python
# Make header row bold and blue
format_font("/path/to/workbook.xlsx", "Sheet1", "A1:D1",
            bold=True, color="0000FF", font_size=14)

# Italic text
format_font("/path/to/workbook.xlsx", "Sheet1", "A10",
            italic=True)
```

**Color Examples:**

- `FF0000` - Red
- `00FF00` - Green
- `0000FF` - Blue
- `FFFF00` - Yellow
- `FF00FF` - Magenta
- `00FFFF` - Cyan
- `000000` - Black
- `FFFFFF` - White

---

### format_fill

Apply background color to a range of cells.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `sheet_name` (string, required): Name of the worksheet
- `range_ref` (string, required): Range to format
- `color` (string, required): Hex color code
- `fill_type` (string, optional): Fill type ('solid' or 'pattern'). Default: 'solid'

**Returns:**

```json
{
  "success": true,
  "message": "Fill formatting applied to A1:D1"
}
```

**Example:**

```python
# Yellow background for headers
format_fill("/path/to/workbook.xlsx", "Sheet1", "A1:D1", color="FFFF00")

# Light gray background
format_fill("/path/to/workbook.xlsx", "Sheet1", "A2:D10", color="F2F2F2")
```

---

### format_border

Apply border formatting to a range of cells.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `sheet_name` (string, required): Name of the worksheet
- `range_ref` (string, required): Range to format
- `style` (string, optional): Border style. Options: 'thin', 'medium', 'thick', 'double'. Default: 'thin'
- `color` (string, optional): Hex color code for border
- `sides` (array, optional): Which sides to apply border. Options: ['top', 'bottom', 'left', 'right']. Default: all sides

**Returns:**

```json
{
  "success": true,
  "message": "Border formatting applied to A1:D10"
}
```

**Example:**

```python
# Border around entire range
format_border("/path/to/workbook.xlsx", "Sheet1", "A1:D10", style="thin")

# Only top and bottom borders
format_border("/path/to/workbook.xlsx", "Sheet1", "A1:D1",
              style="medium", sides=["top", "bottom"])

# Thick red border
format_border("/path/to/workbook.xlsx", "Sheet1", "A1:D1",
              style="thick", color="FF0000")
```

---

### format_alignment

Apply alignment formatting to a range of cells.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `sheet_name` (string, required): Name of the worksheet
- `range_ref` (string, required): Range to format
- `horizontal` (string, optional): Horizontal alignment. Options: 'left', 'center', 'right'
- `vertical` (string, optional): Vertical alignment. Options: 'top', 'center', 'bottom'
- `wrap_text` (boolean, optional): Enable text wrapping
- `text_rotation` (integer, optional): Text rotation angle (0-180)

**Returns:**

```json
{
  "success": true,
  "message": "Alignment formatting applied to A1:D1"
}
```

**Example:**

```python
# Center align headers
format_alignment("/path/to/workbook.xlsx", "Sheet1", "A1:D1",
                 horizontal="center", vertical="center")

# Right align numbers
format_alignment("/path/to/workbook.xlsx", "Sheet1", "D2:D10",
                 horizontal="right")

# Wrap text
format_alignment("/path/to/workbook.xlsx", "Sheet1", "A1",
                 wrap_text=True)
```

---

### format_number

Apply number formatting to a range of cells.

**Parameters:**

- `workbook_path` (string, required): Path to the Excel workbook
- `sheet_name` (string, required): Name of the worksheet
- `range_ref` (string, required): Range to format
- `format_string` (string, required): Excel number format string

**Returns:**

```json
{
  "success": true,
  "message": "Number formatting applied to D2:D10"
}
```

**Example:**

```python
# Currency format
format_number("/path/to/workbook.xlsx", "Sheet1", "D2:D10",
              format_string="$#,##0.00")

# Percentage
format_number("/path/to/workbook.xlsx", "Sheet1", "E2:E10",
              format_string="0.0%")

# Date format
format_number("/path/to/workbook.xlsx", "Sheet1", "F2:F10",
              format_string="mm/dd/yyyy")

# Two decimal places
format_number("/path/to/workbook.xlsx", "Sheet1", "G2:G10",
              format_string="0.00")
```

**Common Format Strings:**

- `0.00` - Two decimal places
- `0.000` - Three decimal places
- `#,##0` - Thousands separator
- `#,##0.00` - Thousands separator with decimals
- `0%` - Percentage (no decimals)
- `0.0%` - Percentage (one decimal)
- `$#,##0.00` - Currency (US Dollar)
- `€#,##0.00` - Currency (Euro)
- `mm/dd/yyyy` - Date (US format)
- `dd/mm/yyyy` - Date (European format)
- `yyyy-mm-dd` - Date (ISO format)
- `h:mm AM/PM` - Time (12-hour)
- `h:mm:ss` - Time (24-hour)
- `@` - Text format

---

## Error Handling

All tools return a `success` field indicating whether the operation succeeded. If `success` is `false`, a `message` field will contain the error description.

**Common Errors:**

1. **File Not Found**: The specified workbook doesn't exist
2. **Sheet Not Found**: The specified worksheet doesn't exist
3. **Invalid Cell Reference**: Cell reference format is incorrect
4. **Invalid Range**: Range reference format is incorrect
5. **Permission Denied**: Cannot read/write file
6. **Invalid Formula**: Formula syntax is incorrect

**Error Response Format:**

```json
{
  "success": false,
  "message": "Sheet 'Data' not found. Available sheets: ['Sheet1', 'Summary']"
}
```

---

## Tips & Best Practices

### Working with Ranges

- Always use uppercase for cell references (though the API accepts lowercase)
- Ensure start cell is before end cell in range (e.g., `A1:B10`, not `B10:A1`)

### File Paths

- Use absolute paths when possible
- Ensure parent directories exist before creating files
- Use `.xlsx` extension (not `.xls`)

### Formulas

- Test formulas in Excel first if unsure of syntax
- Use cell references relative to formula location
- Formulas are evaluated when the file is opened in Excel

### Formatting

- Apply formatting after writing data
- Group similar formatting operations
- Color codes are in RGB hex format (RRGGBB)

### Performance

- Use `write_range` instead of multiple `write_cell` calls
- Minimize the number of operations
- Read/write entire ranges when possible

---

## Complete Workflow Example

```python
# 1. Create a new workbook
create_workbook("/path/to/sales_report.xlsx")

# 2. Create sheets
create_sheet("/path/to/sales_report.xlsx", "Sales Data")
create_sheet("/path/to/sales_report.xlsx", "Summary")

# 3. Write headers
write_range("/path/to/sales_report.xlsx", "Sales Data", "A1",
            [["Date", "Product", "Quantity", "Price", "Total"]])

# 4. Format headers
format_font("/path/to/sales_report.xlsx", "Sales Data", "A1:E1",
            bold=True, font_size=12)
format_fill("/path/to/sales_report.xlsx", "Sales Data", "A1:E1",
            color="4472C4")
format_font("/path/to/sales_report.xlsx", "Sales Data", "A1:E1",
            color="FFFFFF")  # White text

# 5. Write data
data = [
    ["2025-01-01", "Widget", 10, 15.99, None],
    ["2025-01-02", "Gadget", 5, 29.99, None],
    ["2025-01-03", "Widget", 8, 15.99, None]
]
write_range("/path/to/sales_report.xlsx", "Sales Data", "A2", data)

# 6. Add formulas for totals
write_formula("/path/to/sales_report.xlsx", "Sales Data", "E2", "=C2*D2")
write_formula("/path/to/sales_report.xlsx", "Sales Data", "E3", "=C3*D3")
write_formula("/path/to/sales_report.xlsx", "Sales Data", "E4", "=C4*D4")

# 7. Format numbers
format_number("/path/to/sales_report.xlsx", "Sales Data", "D2:D4",
              format_string="$#,##0.00")
format_number("/path/to/sales_report.xlsx", "Sales Data", "E2:E4",
              format_string="$#,##0.00")

# 8. Add borders
format_border("/path/to/sales_report.xlsx", "Sales Data", "A1:E4",
              style="thin")

# 9. Adjust alignment
format_alignment("/path/to/sales_report.xlsx", "Sales Data", "A1:E1",
                 horizontal="center")
format_alignment("/path/to/sales_report.xlsx", "Sales Data", "C2:E4",
                 horizontal="right")
```

---

For more examples and use cases, see the [README](README.md).
