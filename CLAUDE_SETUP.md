# Configuración para Claude Desktop

Esta guía te ayudará a configurar el Excel MCP Server para usarlo con Claude Desktop.

## Paso 1: Verificar la Instalación

Ejecuta el script de verificación para asegurarte de que todo funciona:

```bash
uv run python verify_installation.py
```

Deberías ver:

```
✅ Paquete importado correctamente
✅ FastMCP server inicializado: Excel MCP Server
✅ Módulos de operaciones disponibles
✅ Operaciones de workbook funcionan correctamente
✅ Validadores funcionan correctamente
✅ Modelos Pydantic funcionan correctamente

🎉 ¡Todas las verificaciones pasaron exitosamente!
```

## Paso 2: Ubicar el Archivo de Configuración

### En macOS:

```bash
open ~/Library/Application\ Support/Claude/
```

El archivo se llama: `claude_desktop_config.json`

### En Windows:

```
%APPDATA%\Claude\claude_desktop_config.json
```

## Paso 3: Editar la Configuración

Abre `claude_desktop_config.json` y agrega la configuración del servidor:

```json
{
  "mcpServers": {
    "excel": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "/Users/m.irurozki/Developer/martin/proyectos/excel-mcp",
        "python",
        "-m",
        "excel_mcp_server"
      ]
    }
  }
}
```

**⚠️ IMPORTANTE:** Reemplaza la ruta `/Users/m.irurozki/Developer/martin/proyectos/excel-mcp` con la ruta completa a tu directorio del proyecto.

### Alternativa: Si ya instalaste el paquete globalmente

Si planeas publicar a PyPI o instalar globalmente:

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

## Paso 4: Reiniciar Claude Desktop

1. Cierra completamente Claude Desktop
2. Abre Claude Desktop de nuevo
3. El servidor MCP se iniciará automáticamente

## Paso 5: Verificar la Conexión

En Claude Desktop, intenta estos comandos:

```
"Crea un nuevo archivo Excel llamado test.xlsx en mi escritorio"

"Lista las hojas del archivo test.xlsx"

"Escribe 'Hola Mundo' en la celda A1 de Sheet1"

"Lee el valor de la celda A1"
```

## Herramientas Disponibles

Una vez configurado, tendrás acceso a 20 herramientas:

### Workbook

- `create_workbook` - Crear nuevos archivos Excel
- `get_workbook_info` - Obtener información del archivo
- `list_sheets` - Listar hojas

### Sheets

- `create_sheet` - Crear hojas
- `delete_sheet` - Eliminar hojas
- `rename_sheet` - Renombrar hojas
- `copy_sheet` - Copiar hojas

### Cells

- `write_cell` - Escribir en celdas
- `read_cell` - Leer celdas
- `write_range` - Escribir rangos
- `read_range` - Leer rangos
- `write_formula` - Escribir fórmulas

### Formatting

- `format_font` - Formatear fuente
- `format_fill` - Color de fondo
- `format_border` - Bordes
- `format_alignment` - Alineación
- `format_number` - Formato numérico

## Ejemplos de Uso

### Crear un Reporte de Ventas

```
"Crea un reporte de ventas:
1. Crea un archivo sales_report.xlsx en mi escritorio
2. En Sheet1, escribe los encabezados: Producto, Cantidad, Precio, Total en A1:D1
3. Haz los encabezados en negrita y con fondo azul
4. Escribe estos datos en A2:D4:
   - Widget, 10, 15.99, (fórmula: =B2*C2)
   - Gadget, 5, 29.99, (fórmula: =B3*C3)
   - Doohickey, 8, 12.50, (fórmula: =B4*C4)
5. Formatea la columna D como moneda
6. Centra los encabezados"
```

### Analizar Datos Existentes

```
"Lee los datos del archivo sales_data.xlsx, rango A1:E100,
y dime el total de la columna E"
```

### Crear una Hoja Formateada

```
"En el archivo report.xlsx:
1. Crea una nueva hoja llamada 'Resumen'
2. Escribe un título en A1 con fuente tamaño 16 y negrita
3. Agrega bordes a todas las celdas de A1:D10
4. Aplica alineación centrada a los encabezados"
```

## Solución de Problemas

### Error: "command not found: uv"

Instala uv:

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### Error: "No module named 'excel_mcp_server'"

Asegúrate de estar en el directorio correcto:

```bash
cd /Users/m.irurozki/Developer/martin/proyectos/excel-mcp
uv sync
```

### El servidor no aparece en Claude

1. Verifica que el archivo de configuración esté en el lugar correcto
2. Revisa que la sintaxis JSON sea correcta
3. Asegúrate de haber reiniciado Claude Desktop completamente

### Ver logs del servidor

Los logs aparecerán en la consola de Claude Desktop. En macOS:

```bash
# Ver logs en tiempo real
tail -f ~/Library/Logs/Claude/mcp*.log
```

## Recursos Adicionales

- **Documentación completa:** Ver `TOOLS.md` para referencia de todas las herramientas
- **Ejemplos:** Ver `README.md` para más ejemplos
- **Tests:** Ejecutar `uv run pytest` para ver tests de ejemplo

## Feedback y Soporte

Si encuentras problemas o tienes sugerencias:

1. Revisa los logs
2. Ejecuta `verify_installation.py` para diagnóstico
3. Revisa la documentación en `TOOLS.md`

¡Disfruta usando Excel MCP Server con Claude! 🎉
