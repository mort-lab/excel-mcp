# Excel MCP Server - Product Requirements Document (PRD)

## 1. Executive Summary

### Vision
Create a professional, enterprise-grade Model Context Protocol (MCP) server that enables AI assistants to perform comprehensive Excel file manipulation without requiring Microsoft Excel installation.

### Goals
- Provide a complete suite of Excel operations (create, read, update, format)
- Support multiple transport protocols (stdio, Streamable HTTP)
- Ensure robust error handling and validation
- Enable both local and remote usage
- Maintain high code quality and test coverage
- Publish to PyPI for easy installation via `uvx`

### Target Users
- AI developers building assistants that need Excel capabilities
- Enterprise users requiring automated Excel workflows
- Data analysts using AI tools for spreadsheet manipulation
- Developers integrating Excel functionality into LLM applications

---

## 2. Technical Stack

### Core Technologies

#### Programming Language
- **Python 3.10+** (Required by MCP SDK)
  - Modern type hints support
  - Robust error handling
  - Excellent library ecosystem

#### MCP Framework
- **mcp[cli]** (Official Python SDK)
  - FastMCP for rapid development
  - Built-in transport support
  - Automatic tool schema generation from type hints

#### Excel Library
- **openpyxl** (Primary library)
  - Pure Python implementation
  - No Excel installation required
  - Comprehensive Excel 2010+ support
  - Features:
    - Read/write .xlsx files
    - Cell formatting (fonts, colors, borders, alignment)
    - Formulas and calculations
    - Charts (line, bar, pie, scatter, etc.)
    - Data validation
    - Conditional formatting
    - Pivot tables
    - Excel tables
    - Images and shapes
    - Named ranges
    - Sheet protection

#### Additional Libraries
- **Pydantic** - Data validation and schema definitions
- **typing** - Type hints for robust code
- **pytest** - Testing framework
- **black** - Code formatting
- **ruff** - Linting and code quality

### Package Management
- **uv** - Fast Python package installer
- **pyproject.toml** - Modern Python project configuration

---

## 3. MCP Architecture

### Transport Protocols

#### 1. Stdio Transport (Local Use)
- **Use Case**: Direct integration with local AI clients
- **Communication**: Standard input/output streams
- **Session**: Single persistent connection
- **Best For**: Claude Desktop, local development

**Configuration Example:**
```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["excel-mcp-server", "stdio"]
    }
  }
}
```

#### 2. Streamable HTTP Transport (Remote Use)
- **Use Case**: Remote server deployments, multi-client support
- **Communication**: HTTP POST/GET with optional SSE streaming
- **Session**: Stateful via `Mcp-Session-Id` header
- **Best For**: Cloud deployments, enterprise environments

**Configuration Example:**
```json
{
  "mcpServers": {
    "excel": {
      "url": "http://localhost:8017/mcp"
    }
  }
}
```

**Environment Variables:**
- `EXCEL_FILES_PATH` - Base directory for Excel files (required for HTTP)
- `FASTMCP_PORT` - Server port (default: 8017)
- `MCP_LOG_LEVEL` - Logging verbosity (debug, info, warning, error)

### MCP Components

#### 1. Tools (Actions)
Functions that perform operations and side effects:
- Create/modify Excel files
- Write data and formulas
- Apply formatting
- Generate charts and pivot tables

**Example:**
```python
@mcp.tool()
def write_cell(workbook_path: str, sheet_name: str, cell: str, value: str) -> dict:
    """Write a value to a specific cell"""
    # Implementation
    return {"status": "success", "cell": cell, "value": value}
```

#### 2. Resources (Data)
Read-only data exposure:
- List workbooks in directory
- Get workbook metadata
- Read sheet structure
- Access cell values

**Example:**
```python
@mcp.resource("excel://{workbook}/sheets")
def list_sheets(workbook: str) -> dict:
    """List all sheets in a workbook"""
    # Implementation
    return {"sheets": ["Sheet1", "Sheet2"]}
```

#### 3. Prompts (Templates)
Reusable AI interaction patterns:
- Data analysis templates
- Report generation workflows
- Formatting suggestions

### JSON-RPC Protocol
All MCP communication uses JSON-RPC 2.0:
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call",
  "params": {
    "name": "write_cell",
    "arguments": {
      "workbook_path": "/path/to/file.xlsx",
      "sheet_name": "Sheet1",
      "cell": "A1",
      "value": "Hello"
    }
  }
}
```

---

## 4. Feature Requirements

### 4.1 Core Excel Operations

#### Workbook Management
- **create_workbook** - Create new Excel file
  - Input: `file_path: str`
  - Output: `{status, path}`
  - Validation: Path must be writable

- **open_workbook** - Open existing Excel file
  - Input: `file_path: str`
  - Output: `{status, sheets: List[str]}`
  - Validation: File must exist and be valid .xlsx

- **save_workbook** - Save changes to workbook
  - Input: `file_path: str`
  - Output: `{status}`

- **get_workbook_info** - Get workbook metadata
  - Input: `file_path: str`
  - Output: `{sheets, properties, size}`

#### Sheet Management
- **create_sheet** - Add new worksheet
  - Input: `workbook_path: str, sheet_name: str, index?: int`
  - Output: `{status, sheet_name}`

- **delete_sheet** - Remove worksheet
  - Input: `workbook_path: str, sheet_name: str`
  - Output: `{status}`

- **rename_sheet** - Rename worksheet
  - Input: `workbook_path: str, old_name: str, new_name: str`
  - Output: `{status, new_name}`

- **copy_sheet** - Duplicate worksheet
  - Input: `workbook_path: str, source_sheet: str, new_name: str`
  - Output: `{status, new_sheet}`

- **list_sheets** - Get all sheet names
  - Input: `workbook_path: str`
  - Output: `{sheets: List[str]}`

### 4.2 Data Operations

#### Cell Operations
- **write_cell** - Write value to cell
  - Input: `workbook_path, sheet_name, cell, value`
  - Output: `{status, cell, value}`
  - Supports: strings, numbers, dates, booleans

- **read_cell** - Read cell value
  - Input: `workbook_path, sheet_name, cell`
  - Output: `{value, type, formula?}`

- **write_range** - Write data to range
  - Input: `workbook_path, sheet_name, start_cell, data: List[List[Any]]`
  - Output: `{status, rows_written, cols_written}`

- **read_range** - Read range of cells
  - Input: `workbook_path, sheet_name, range (e.g., "A1:D10")`
  - Output: `{data: List[List[Any]]}`

#### Formula Operations
- **write_formula** - Write formula to cell
  - Input: `workbook_path, sheet_name, cell, formula`
  - Output: `{status, cell, formula, calculated_value?}`
  - Validation: Formula syntax

- **bulk_write_formulas** - Write multiple formulas
  - Input: `workbook_path, sheet_name, formulas: Dict[cell, formula]`
  - Output: `{status, count}`

### 4.3 Formatting Operations

#### Font Styling
- **format_font** - Apply font formatting
  - Input: `workbook_path, sheet_name, range, font_config`
  - Font Config:
    - `name: str` (e.g., "Arial", "Calibri")
    - `size: int` (8-72)
    - `bold: bool`
    - `italic: bool`
    - `underline: str` ("single", "double", "none")
    - `color: str` (hex color, e.g., "FF0000")
  - Output: `{status, range}`

#### Cell Styling
- **format_fill** - Apply background color
  - Input: `workbook_path, sheet_name, range, fill_config`
  - Fill Config:
    - `type: str` ("solid", "pattern")
    - `color: str` (hex color)
    - `pattern_type?: str`
  - Output: `{status, range}`

- **format_border** - Apply borders
  - Input: `workbook_path, sheet_name, range, border_config`
  - Border Config:
    - `style: str` ("thin", "medium", "thick", "double")
    - `color: str`
    - `sides: List[str]` (["top", "bottom", "left", "right"])
  - Output: `{status, range}`

- **format_alignment** - Apply alignment
  - Input: `workbook_path, sheet_name, range, alignment_config`
  - Alignment Config:
    - `horizontal: str` ("left", "center", "right")
    - `vertical: str` ("top", "center", "bottom")
    - `wrap_text: bool`
    - `text_rotation: int` (0-180)
  - Output: `{status, range}`

#### Number Formatting
- **format_number** - Apply number format
  - Input: `workbook_path, sheet_name, range, format_string`
  - Format Examples:
    - `"0.00"` - Two decimal places
    - `"#,##0"` - Thousands separator
    - `"0%"` - Percentage
    - `"$#,##0.00"` - Currency
    - `"mm/dd/yyyy"` - Date
  - Output: `{status, range}`

#### Conditional Formatting
- **add_conditional_format** - Add conditional formatting rule
  - Input: `workbook_path, sheet_name, range, rule_config`
  - Rule Types:
    - Color scale (2-color, 3-color)
    - Data bars
    - Icon sets
    - Cell value conditions (greater than, less than, between, etc.)
    - Formula-based
  - Output: `{status, rule_id}`

### 4.4 Advanced Features

#### Excel Tables
- **create_table** - Create Excel table
  - Input: `workbook_path, sheet_name, range, table_name, style?`
  - Output: `{status, table_name, range}`
  - Features: Auto-filtering, structured references

- **add_table_column** - Add column to table
  - Input: `workbook_path, sheet_name, table_name, column_name, formula?`
  - Output: `{status, column_name}`

#### Chart Creation
- **create_chart** - Generate chart
  - Input: `workbook_path, sheet_name, chart_config`
  - Chart Config:
    - `type: str` ("line", "bar", "column", "pie", "scatter", "area")
    - `title: str`
    - `data_range: str`
    - `position: str` (cell reference)
    - `x_axis_title?: str`
    - `y_axis_title?: str`
    - `legend: bool`
  - Output: `{status, chart_id}`

- **create_combo_chart** - Create combination chart
  - Input: Multiple chart types on same axes
  - Output: `{status, chart_id}`

#### Pivot Tables
- **create_pivot_table** - Create pivot table
  - Input: `workbook_path, source_sheet, source_range, dest_sheet, dest_cell, pivot_config`
  - Pivot Config:
    - `rows: List[str]` - Row fields
    - `columns: List[str]` - Column fields
    - `values: List[Dict]` - Value fields with aggregation (sum, avg, count, etc.)
    - `filters: List[str]` - Filter fields
  - Output: `{status, pivot_table_id}`

#### Data Validation
- **add_data_validation** - Add validation rule
  - Input: `workbook_path, sheet_name, range, validation_config`
  - Validation Types:
    - List (dropdown)
    - Whole number
    - Decimal
    - Date
    - Time
    - Text length
    - Custom formula
  - Config:
    - `type: str`
    - `criteria: Dict` (operator, value1, value2)
    - `error_message?: str`
    - `prompt_message?: str`
  - Output: `{status, range}`

#### Images and Shapes
- **insert_image** - Insert image into sheet
  - Input: `workbook_path, sheet_name, image_path, cell, width?, height?`
  - Output: `{status, image_id}`

- **insert_shape** - Insert shape
  - Input: `workbook_path, sheet_name, shape_type, cell, text?`
  - Output: `{status, shape_id}`

### 4.5 Utility Operations

#### Sheet Operations
- **get_used_range** - Get range of used cells
  - Input: `workbook_path, sheet_name`
  - Output: `{range, rows, cols}`

- **clear_range** - Clear cell contents
  - Input: `workbook_path, sheet_name, range`
  - Output: `{status, cells_cleared}`

- **merge_cells** - Merge cell range
  - Input: `workbook_path, sheet_name, range`
  - Output: `{status, merged_range}`

- **unmerge_cells** - Unmerge cells
  - Input: `workbook_path, sheet_name, range`
  - Output: `{status}`

#### Column/Row Operations
- **set_column_width** - Set column width
  - Input: `workbook_path, sheet_name, column, width`
  - Output: `{status}`

- **set_row_height** - Set row height
  - Input: `workbook_path, sheet_name, row, height`
  - Output: `{status}`

- **auto_fit_columns** - Auto-fit column widths
  - Input: `workbook_path, sheet_name, columns`
  - Output: `{status, widths: Dict[column, width]}`

- **insert_rows** - Insert empty rows
  - Input: `workbook_path, sheet_name, row, count`
  - Output: `{status, rows_inserted}`

- **delete_rows** - Delete rows
  - Input: `workbook_path, sheet_name, row, count`
  - Output: `{status, rows_deleted}`

- **insert_columns** - Insert empty columns
  - Input: `workbook_path, sheet_name, column, count`
  - Output: `{status, columns_inserted}`

- **delete_columns** - Delete columns
  - Input: `workbook_path, sheet_name, column, count`
  - Output: `{status, columns_deleted}`

---

## 5. Data Validation & Error Handling

### Input Validation

#### File Path Validation
```python
def validate_file_path(path: str, must_exist: bool = False) -> bool:
    """
    - Must be absolute or relative path
    - Must end with .xlsx
    - If must_exist=True, file must exist
    - Parent directory must exist and be writable
    - No path traversal attacks (../../)
    """
```

#### Cell Reference Validation
```python
def validate_cell_reference(cell: str) -> bool:
    """
    - Valid formats: A1, B10, AA100
    - Column: A-ZZ (max 16384)
    - Row: 1-1048576
    """
```

#### Range Validation
```python
def validate_range(range_str: str) -> bool:
    """
    - Valid formats: A1:B10, A:A, 1:1
    - Start cell must be before end cell
    """
```

#### Formula Validation
```python
def validate_formula(formula: str) -> bool:
    """
    - Must start with =
    - No dangerous functions (CALL, REGISTER)
    - Valid Excel formula syntax
    """
```

### Error Handling Strategy

#### Error Types
1. **FileNotFoundError** - Workbook doesn't exist
2. **PermissionError** - Cannot read/write file
3. **ValueError** - Invalid input parameters
4. **FormulaError** - Invalid formula syntax
5. **SheetNotFoundError** - Sheet doesn't exist
6. **CellRangeError** - Invalid cell/range reference

#### JSON-RPC Error Codes
```python
ERROR_CODES = {
    -32700: "Parse error",           # Invalid JSON
    -32600: "Invalid request",       # Invalid JSON-RPC
    -32601: "Method not found",      # Unknown tool
    -32602: "Invalid params",        # Invalid parameters
    -32603: "Internal error",        # Server error
    -32001: "File not found",        # Custom: File doesn't exist
    -32002: "Permission denied",     # Custom: Cannot access file
    -32003: "Invalid cell reference",# Custom: Bad cell/range
    -32004: "Invalid formula",       # Custom: Formula error
    -32005: "Sheet not found",       # Custom: Sheet doesn't exist
}
```

#### Error Response Format
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "error": {
    "code": -32001,
    "message": "File not found",
    "data": {
      "file_path": "/path/to/missing.xlsx",
      "suggestion": "Please check the file path and try again"
    }
  }
}
```

### Logging Strategy

#### Log Levels (RFC 5424)
- **DEBUG**: Function entry/exit, variable values
- **INFO**: Operation success, progress updates
- **NOTICE**: Configuration changes
- **WARNING**: Deprecated features, non-critical issues
- **ERROR**: Operation failures, caught exceptions
- **CRITICAL**: Component failures
- **ALERT**: Data corruption detected
- **EMERGENCY**: Complete system failure

#### Log Message Format
```python
{
  "level": "error",
  "logger": "excel_operations",
  "timestamp": "2025-01-29T12:00:00Z",
  "data": {
    "operation": "write_cell",
    "workbook": "/path/to/file.xlsx",
    "error": "Sheet 'Data' not found",
    "available_sheets": ["Sheet1", "Sheet2"]
  }
}
```

---

## 6. Security Considerations

### File System Security
1. **Path Sanitization**
   - No path traversal (../)
   - Whitelist allowed directories
   - Validate file extensions

2. **File Access Control**
   - Read/write permissions checks
   - File size limits (max 100MB)
   - Rate limiting for operations

3. **Formula Security**
   - Blacklist dangerous functions
   - No external data connections
   - Sandbox formula evaluation

### HTTP Transport Security
1. **Origin Validation**
   - MUST validate `Origin` header
   - CORS configuration

2. **Authentication**
   - Bearer token support (RFC 6750)
   - OAuth 2.0 integration
   - API key validation

3. **Network Security**
   - Bind to localhost for local deployments
   - HTTPS required for remote access
   - TLS 1.3 minimum

4. **Session Management**
   - Session ID validation (`Mcp-Session-Id`)
   - Session timeout (30 minutes)
   - Session cleanup

### Data Protection
1. **Input Sanitization**
   - Escape special characters
   - Validate data types
   - Size limits on inputs

2. **Output Encoding**
   - Proper JSON encoding
   - Base64 for binary data
   - No sensitive data in logs

---

## 7. Performance Requirements

### Response Time Targets
- **Simple operations** (read cell, write cell): < 100ms
- **Medium operations** (read range, format cells): < 500ms
- **Complex operations** (create chart, pivot table): < 2s
- **Large file operations** (>10MB): < 5s

### Resource Limits
- **Max file size**: 100MB
- **Max range size**: 1M cells (1000x1000)
- **Max concurrent operations**: 10
- **Memory limit**: 512MB per operation

### Optimization Strategies
1. **Lazy Loading**: Only load required sheets
2. **Caching**: Cache workbook objects for repeated operations
3. **Batch Operations**: Support bulk operations to reduce overhead
4. **Streaming**: Stream large data ranges
5. **Connection Pooling**: Reuse connections in HTTP mode

---

## 8. Testing Strategy

### Unit Tests
- Test each tool function independently
- Mock file system operations
- Test error handling
- Coverage target: 90%

### Integration Tests
- Test complete workflows (create → write → format → save)
- Test with real Excel files
- Test transport protocols
- Test concurrent operations

### Test Categories
1. **Basic Operations**
   - Create/open/save workbooks
   - Read/write cells
   - Sheet management

2. **Formatting**
   - Font, fill, border, alignment
   - Number formats
   - Conditional formatting

3. **Advanced Features**
   - Charts, pivot tables
   - Excel tables
   - Data validation

4. **Error Handling**
   - Invalid inputs
   - File not found
   - Permission errors
   - Formula errors

5. **Performance**
   - Large file handling
   - Bulk operations
   - Memory usage

### Test Files
- Sample Excel files with various features
- Corrupted files for error testing
- Large files for performance testing

---

## 9. Documentation Requirements

### User Documentation
1. **README.md**
   - Quick start guide
   - Installation instructions
   - Basic examples
   - Feature overview

2. **TOOLS.md**
   - Complete tool reference
   - Parameters and return values
   - Examples for each tool
   - Common workflows

3. **CONFIGURATION.md**
   - Environment variables
   - Transport setup
   - Security configuration

4. **EXAMPLES.md**
   - Real-world use cases
   - Code snippets
   - Common patterns

### Developer Documentation
1. **ARCHITECTURE.md**
   - System design
   - Component overview
   - Data flow diagrams

2. **CONTRIBUTING.md**
   - Development setup
   - Code style guide
   - Pull request process

3. **API Reference**
   - Auto-generated from docstrings
   - Type annotations
   - Error codes

---

## 10. Deployment & Distribution

### PyPI Package
- **Package name**: `excel-mcp-server`
- **Version**: Semantic versioning (v1.0.0)
- **Dependencies**: Minimal, well-maintained
- **License**: MIT

### Installation Methods
```bash
# Via uvx (recommended)
uvx excel-mcp-server stdio

# Via pip
pip install excel-mcp-server
excel-mcp-server stdio

# From source
git clone https://github.com/username/excel-mcp-server
cd excel-mcp-server
uv sync
uv run excel-mcp-server stdio
```

### Docker Support
```dockerfile
FROM python:3.11-slim
WORKDIR /app
RUN pip install excel-mcp-server
EXPOSE 8017
CMD ["excel-mcp-server", "streamable-http"]
```

### Configuration Files
1. **pyproject.toml** - Project metadata and dependencies
2. **uv.lock** - Locked dependencies
3. **.env.example** - Environment variable template

---

## 11. Implementation Phases

### Phase 1: Core Foundation (Week 1-2)
**Goals**: Basic infrastructure and simple operations

**Deliverables**:
- Project structure and configuration
- MCP server setup with FastMCP
- Stdio transport implementation
- Basic workbook operations:
  - create_workbook
  - open_workbook
  - save_workbook
  - list_sheets
- Basic cell operations:
  - write_cell
  - read_cell
  - write_range
  - read_range
- Unit tests for core functions
- Basic documentation

**Success Criteria**:
- Can create a workbook via AI assistant
- Can write and read cell values
- Tests pass with 80% coverage

### Phase 2: Formatting & Styling (Week 3)
**Goals**: Complete formatting capabilities

**Deliverables**:
- Font formatting (bold, italic, color, size)
- Cell styling (fill, borders, alignment)
- Number formatting
- Sheet management (create, delete, rename, copy)
- Column/row operations (width, height, insert, delete)
- Integration tests for formatting
- TOOLS.md documentation

**Success Criteria**:
- Can create a formatted report via AI
- All formatting options work correctly
- Examples in documentation

### Phase 3: Advanced Features (Week 4)
**Goals**: Charts, tables, and advanced functionality

**Deliverables**:
- Excel tables support
- Chart creation (line, bar, pie, scatter)
- Pivot tables
- Data validation
- Conditional formatting
- Formula support
- Advanced tests
- Example workflows

**Success Criteria**:
- Can create a dashboard with charts
- Can create a pivot table report
- Complex workflows documented

### Phase 4: HTTP Transport & Security (Week 5)
**Goals**: Remote access and production-ready features

**Deliverables**:
- Streamable HTTP transport
- Session management
- Authentication support
- Origin validation
- Security hardening
- Performance optimization
- Docker support
- Deployment guide

**Success Criteria**:
- Can run as remote server
- Secure by default
- Performance targets met

### Phase 5: Polish & Release (Week 6)
**Goals**: Production release

**Deliverables**:
- Complete documentation
- Tutorial videos/guides
- PyPI publication
- GitHub repository setup
- CI/CD pipeline
- Logo and branding
- Announcement blog post

**Success Criteria**:
- Published to PyPI
- Documentation complete
- Community feedback positive

---

## 12. Success Metrics

### Technical Metrics
- **Test Coverage**: >90%
- **Response Time**: 95th percentile < 500ms
- **Error Rate**: <0.1%
- **Uptime**: 99.9% (for HTTP mode)

### User Metrics
- **Downloads**: 1000+ in first month
- **GitHub Stars**: 500+ in first 3 months
- **Issues Resolved**: >90% in 7 days
- **User Satisfaction**: 4.5/5 stars

### Community Metrics
- **Contributors**: 5+ in first 6 months
- **Forks**: 100+ in first year
- **Blog Mentions**: 10+ articles
- **Integration Examples**: 20+ community projects

---

## 13. Risks & Mitigation

### Technical Risks
1. **Risk**: openpyxl limitations
   - **Mitigation**: Test extensively, document limitations, consider xlsxwriter for write-only operations

2. **Risk**: Large file performance
   - **Mitigation**: Implement streaming, set size limits, optimize algorithms

3. **Risk**: Formula compatibility
   - **Mitigation**: Document supported formulas, validate before writing

### Security Risks
1. **Risk**: Path traversal attacks
   - **Mitigation**: Strict path validation, whitelist directories

2. **Risk**: Formula injection
   - **Mitigation**: Blacklist dangerous functions, sanitize inputs

3. **Risk**: Resource exhaustion
   - **Mitigation**: Rate limiting, size limits, timeouts

### Operational Risks
1. **Risk**: Breaking changes in MCP protocol
   - **Mitigation**: Follow spec closely, version pinning, test with multiple clients

2. **Risk**: Dependency vulnerabilities
   - **Mitigation**: Regular updates, security scanning, minimal dependencies

---

## 14. Future Enhancements (Post-v1.0)

### Planned Features
1. **Multi-format Support**
   - CSV import/export
   - PDF export
   - HTML export

2. **Advanced Analytics**
   - Statistical functions
   - Data transformation
   - Machine learning integration

3. **Collaboration Features**
   - Multi-user editing
   - Change tracking
   - Comments and annotations

4. **Performance**
   - Parallel processing
   - Incremental saves
   - Distributed operations

5. **Integration**
   - Database connectors
   - API integrations
   - Cloud storage support

---

## 15. Conclusion

This PRD defines a comprehensive Excel MCP server that will:
- Provide full Excel manipulation capabilities
- Support multiple deployment scenarios
- Maintain high security and performance standards
- Enable AI assistants to work seamlessly with spreadsheets

The phased approach ensures steady progress while maintaining quality. The focus on documentation and testing will create a reliable, maintainable codebase that the community can build upon.

**Next Steps**:
1. Review and approve this PRD
2. Set up project repository and infrastructure
3. Begin Phase 1 implementation
4. Regular progress reviews and adjustments

---

**Document Version**: 1.0
**Last Updated**: 2025-01-29
**Status**: Draft - Pending Approval
