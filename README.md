# Docx MCP Server

A Model Context Protocol (MCP) server for reading and editing Microsoft Word (`.docx`) documents.

## Features

- **`read_docx`**: Extracts all paragraph text and tables from a `.docx` file.
- **`create_or_append_docx`**: Creates a new `.docx` file or appends headings and paragraphs to an existing file.
- **`replace_text_in_docx`**: Finds and replaces specific text across paragraphs and tables.

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
uv run server.py
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
        "server.py"
      ]
    }
  }
}
```
