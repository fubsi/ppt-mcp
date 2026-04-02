#!/usr/bin/env python
"""
MCP Server for PowerPoint manipulation using python-pptx.
"""
import argparse
from typing import Dict
from mcp.server.fastmcp import FastMCP

# import utils  # Currently unused
from tools import (
    register_file_management_tools,
    register_placeholder_tools,
    register_shape_tools,
    register_slide_tools,
)

# Initialize the FastMCP server
app = FastMCP(
    name="ppt-mcp-server"
)

# Global state to store presentations in memory
presentations = {}

def register_tools(default_template_file_path: str = "Template.pptx"):
    """Register all tool modules."""
    register_file_management_tools(app, presentations, default_template_file_path)
    register_placeholder_tools(app, presentations)
    register_shape_tools(app, presentations)
    register_slide_tools(app, presentations)


@app.tool()
def get_server_info() -> Dict:
    """Get information about the MCP server."""
    return {
        "name": "PowerPoint MCP Server",
        "version": "1.0.0"
    }

# ---- Main Function ----
def main(
    transport: str = "stdio",
    port: int = 8000,
    template_file_path: str = "Template.pptx",
):
    register_tools(template_file_path)

    if transport == "http":
        import asyncio
        # Set the port for HTTP transport
        app.settings.port = port
        # Start the FastMCP server with HTTP transport
        try:
            app.run(transport='streamable-http')
        except asyncio.exceptions.CancelledError:
            print("Server stopped by user.")
        except KeyboardInterrupt:
            print("Server stopped by user.")
        except Exception as e:
            print(f"Error starting server: {e}")
    else:
        # Run the FastMCP server
        app.run(transport='stdio')

if __name__ == "__main__":
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="MCP Server for PowerPoint manipulation using python-pptx")

    parser.add_argument(
        "-t",
        "--transport",
        type=str,
        default="stdio",
        choices=["stdio", "http", "sse"],
        help="Transport method for the MCP server (default: stdio)"
    )

    parser.add_argument(
        "-p",
        "--port",
        type=int,
        default=8000,
        help="Port to run the MCP server on (default: 8000)"
    )

    parser.add_argument(
        "--template-file-path",
        type=str,
        default="Template.pptx",
        help="Path to the default PowerPoint template file used when creating new presentations (default: Template.pptx)"
    )

    args = parser.parse_args()
    main(args.transport, args.port, args.template_file_path)
