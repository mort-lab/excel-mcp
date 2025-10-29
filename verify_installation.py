#!/usr/bin/env python3
"""Quick verification script to test Excel MCP Server installation"""

import sys
from pathlib import Path

print("🔍 Verificando instalación de Excel MCP Server...\n")

# Test 1: Import the package
try:
    from excel_mcp_server import mcp
    print("✅ Paquete importado correctamente")
except ImportError as e:
    print(f"❌ Error al importar paquete: {e}")
    sys.exit(1)

# Test 2: Check FastMCP instance
try:
    assert mcp is not None
    print(f"✅ FastMCP server inicializado: {mcp.name}")
except AssertionError:
    print("❌ FastMCP server no inicializado")
    sys.exit(1)

# Test 3: Check operations modules
try:
    from excel_mcp_server import operations
    assert hasattr(operations, 'workbook')
    assert hasattr(operations, 'cell')
    assert hasattr(operations, 'sheet')
    assert hasattr(operations, 'formatting')
    print("✅ Módulos de operaciones disponibles")
except (ImportError, AssertionError) as e:
    print(f"❌ Error en módulos de operaciones: {e}")
    sys.exit(1)

# Test 4: Create a test workbook
try:
    import tempfile
    from excel_mcp_server.operations import workbook

    # Create temp file path (don't create the file yet)
    tmpdir = Path(tempfile.gettempdir())
    test_file = tmpdir / "test_excel_mcp.xlsx"

    # Remove if exists
    if test_file.exists():
        test_file.unlink()

    result = workbook.create(str(test_file))
    assert result.success is True, f"Expected success, got: {result.message}"
    assert test_file.exists(), f"File not created at {test_file}"

    # Cleanup
    test_file.unlink()

    print("✅ Operaciones de workbook funcionan correctamente")
except Exception as e:
    print(f"❌ Error en operaciones: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Test 5: Check validators
try:
    from excel_mcp_server.utils import validators

    is_valid, _ = validators.validate_cell_reference("A1")
    assert is_valid is True

    is_valid, _ = validators.validate_cell_reference("INVALID")
    assert is_valid is False

    print("✅ Validadores funcionan correctamente")
except Exception as e:
    print(f"❌ Error en validadores: {e}")
    sys.exit(1)

# Test 6: Check models
try:
    from excel_mcp_server.models import CellWriteRequest

    request = CellWriteRequest(
        workbook_path="/tmp/test.xlsx",
        sheet_name="Sheet1",
        cell="A1",
        value="test"
    )
    assert request.cell == "A1"

    print("✅ Modelos Pydantic funcionan correctamente")
except Exception as e:
    print(f"❌ Error en modelos: {e}")
    sys.exit(1)

print("\n" + "="*60)
print("🎉 ¡Todas las verificaciones pasaron exitosamente!")
print("="*60)
print("\n📝 Próximos pasos:")
print("   1. Ejecutar: uv run python -m excel_mcp_server")
print("   2. Configurar en Claude Desktop")
print("   3. ¡Empezar a usar!")
