# PowerPoint-MCP-Server

A compact MCP (Model Context Protocol) server for practical PowerPoint editing with `python-pptx`.

This repository is the core editing edition of the server and currently focuses on file lifecycle, slide operations, placeholder editing, and direct shape insertion.

## Current Scope

- Template-based presentation creation
- Open and save existing `.pptx` files
- Slide inspection and management
- Placeholder text and image insertion
- Add text boxes and images at explicit coordinates
- In-memory multi-presentation state via `presentation_id`

## Important: Template Requirement

New presentations are created by copying a template file.

- Default template path: `Template.pptx`
- Override via CLI: `--template-file-path /path/to/template.pptx`
- If the template does not exist, creation fails with `FileNotFoundError`

This is required by the `create_presentation_file` tool in the current implementation.

## Tool Modules and Tools

The server currently exposes 15 tools in total.

### Server

1. `get_server_info`

### File Management (`tools/file_management.py`)

1. `create_presentation_file`
2. `open_presentation_file`
3. `save_presentation_file`
4. `cleanup_presentation_file`

### Slide Tools (`tools/slide_tools.py`)

1. `get_slide_layouts`
2. `get_slides`
3. `add_slide`
4. `remove_slide`
5. `rename_slide`
6. `get_slide_content`

### Placeholder Tools (`tools/placeholder_tools.py`)

1. `insert_picture_into_placeholder`
2. `insert_text_into_placeholder`

### Shape Tools (`tools/shape_tools.py`)

1. `add_image_to_slide`
2. `add_text_to_slide`

## Installation

### Prerequisites

- Python 3.6+
- A valid PowerPoint template file (default: `Template.pptx`)

### Install dependencies

```bash
pip install -r requirements.txt
```

## Running the Server

### Stdio transport

```bash
python ppt_mcp_server.py
```

### HTTP transport

```bash
python ppt_mcp_server.py --transport http --port 8000
```

### Use a custom default template

```bash
python ppt_mcp_server.py --template-file-path "C:/path/to/Template.pptx"
```

## Minimal MCP Config Example

```json
{
  "mcpServers": {
    "ppt": {
      "command": "python",
      "args": [
        "/absolute/path/to/ppt_mcp_server.py",
        "--template-file-path",
        "/absolute/path/to/Template.pptx"
      ],
      "env": {}
    }
  }
}
```

## Project Structure

```text
ppt-mcp/
|- ppt_mcp_server.py
|- tools/
|  |- __init__.py
|  |- file_management.py
|  |- placeholder_tools.py
|  |- shape_tools.py
|  |- slide_tools.py
|- utils/
|  |- __init__.py
|  |- models.py
|- Template.pptx
|- requirements.txt
|- pyproject.toml
```

## Notes

- Position and size values for shape insertion are in EMUs.
- Image insertion supports local file paths and `http/https` URLs.
- Temporary presentations are stored under `temp_presentations/`.

## License

MIT
