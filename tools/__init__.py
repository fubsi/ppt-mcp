"""
Tools package for PowerPoint MCP Server.
Organizes tools into logical modules for better maintainability.
"""
from .file_management import register_file_management_tools
from .placeholder_tools import register_placeholder_tools
from .shape_tools import register_shape_tools
from .slide_tools import register_slide_tools

__all__ = [
    "register_file_management_tools",
    "register_placeholder_tools",
    "register_shape_tools",
    "register_slide_tools",
]