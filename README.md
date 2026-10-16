# Docx MCP Server

A Model Context Protocol (MCP) server for reading and editing Microsoft Word (`.docx`) documents.

## Features

- **`read_docx`**: Extracts all paragraph text and tables from a `.docx` file.
- **`create_or_append_docx`**: Creates a new `.docx` file or appends headings and paragraphs to an existing file.
- **`replace_text_in_docx`**: Finds and replaces specific text across paragraphs and tables.
- **`create_table_docx`**: Creates a new table from 2D data rows with an optional heading.
- **`edit_table_docx`**: Edits specific cell content by table index, row, and column.
- **`delete_table_docx`**: Deletes a table by its table index.
- **`clear_docx`**: Clears all content (paragraphs and tables) from a document.
- **`add_column_docx`**: Appends or inserts a new column to a table with optional cell values, width, and position (`insert_index`).
- **`resize_table_docx`**: Resizes column widths (in inches) across all columns or for a single column.
- **`set_table_visibility_docx`**: Toggles table border visibility (visible grid or transparent/borderless).
- **`delete_column_docx`**: Deletes a specific column from a table by column index.
- **`add_row_docx`**: Appends or inserts a new row to a table with optional cell values, height, and position (`insert_index`).
- **`delete_row_docx`**: Deletes a specific row from a table by row index.
- **`merge_cells_docx`**: Merges a rectangular range of cells within a table.
- **`set_cell_shading_docx`**: Applies background color (HEX) to a cell, row, column, or entire table.
- **`set_table_style_docx`**: Sets the Word table template style (e.g. Table Grid, Light Shading).
- **`set_table_alignment_docx`**: Sets table alignment on the document page (left, center, right).
- **`set_cell_alignment_docx`**: Sets cell text alignment horizontally and/or vertically.
- **`set_table_pagination_docx`**: Configures repeat header rows (`tblHeader`) and prevents row split (`cantSplit`).
- **`set_cell_margins_docx`**: Sets cell padding/margins (in pt) on a specific cell or the entire table.
- **`set_table_borders_docx`**: Configures custom borders per side (top, bottom, left, right, inside) with custom color and size.
- **`set_cell_text_direction_docx`**: Rotates cell text direction (horizontal, 90° up, 270° down).
- **`split_table_docx`**: Splits a table into two separate tables at a specified row index.
- **`format_text_docx`**: Formats font name, size (Pt), color (RGB/HEX), bold, italic, underline, strikethrough, and highlight on paragraphs, table cells, or targeted phrases.
- **`add_styled_paragraph_docx`**: Appends a new paragraph or heading with comprehensive font, styling, and alignment options.

## Requirements

- Python >= 3.12
- [uv](https://docs.astral.sh/uv/) (recommended)

## Installation

Clone the repository and install dependencies:

```bash
cd docx-mcp
uv sync
```

## Running the Server

Run directly with `uv`:

```bash
uv run docx-mcp
```

## MCP Client Configuration

Add this server to your MCP client configuration (e.g. `mcp_config.json`):

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/docx-mcp",
        "run",
        "docx-mcp"
      ]
    }
  }
}
```
