"""Simple file management tools for PowerPoint files."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Dict
from uuid import uuid4

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from utils.models import PresentationFile


BASE_DIR = Path(__file__).resolve().parent.parent
TEMP_PATH = BASE_DIR / "temp_presentations"
TEMP_PATH.mkdir(exist_ok=True)



def register_file_management_tools(app: FastMCP, presentations: Dict[str, PresentationFile], default_template_file_path: str):
    """Register simple PowerPoint file management tools."""

    def _resolve_template_path(template_path_input: str) -> Path:
        template_path = Path(template_path_input).expanduser()
        if template_path.is_absolute():
            return template_path
        return (BASE_DIR / template_path).resolve()

    @app.tool(
        annotations=ToolAnnotations(
            title="Create Presentation File",
            readOnlyHint=False,
        ),
        description="Creates a new Powerpoint presentation file and returns its information. \
        The presentation_id returned must be used for subsequent operations on this file. \
        Placeholder-first policy: when adding content to slides later, prefer placeholder tools before manual shape tools."
    )
    def create_presentation_file() -> Dict:
        """Create a new PowerPoint presentation file."""
        individual_id = uuid4().hex[:12]
        file_name = f"{individual_id}.pptx"
        file_path = TEMP_PATH / file_name
        
        # Create an empty presentation from the default template
        template_path = _resolve_template_path(default_template_file_path)
        if not template_path.exists():
            return {
                "error": "Default template file not found.",
                "template_path": str(template_path),
            }
        
        # Copy the template to create a new presentation
        with template_path.open("rb") as src, file_path.open("wb") as dst:
            dst.write(src.read())

        if not file_path.exists() or file_path.stat().st_size == 0:
            return {
                "error": "Failed to initialize presentation file from template.",
                "template_path": str(template_path),
                "target_path": str(file_path),
            }
        
        # Load the presentation into memory and store it in the global state
        try:
            presentations[individual_id] = PresentationFile(str(file_path))
        except Exception as e:
            return {
                "error": f"Failed to load presentation package: {str(e)}",
                "template_path": str(template_path),
                "target_path": str(file_path),
            }
        
        return {
            "message": "Presentation file created successfully",
            "presentation_id": individual_id,
            "file_info": presentations[individual_id].get_file_info()
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Open Presentation File",
            readOnlyHint=False,
        ),
        description="Opens an existing presentation file from disk and returns its information. \
        The presentation_id returned must be used for subsequent operations on this file. \
        Placeholder-first policy: when adding content to slides later, prefer placeholder tools before manual shape tools."
    )
    def open_presentation_file(file_path: str) -> Dict:
        """Open an existing presentation file."""
        file_path = Path(file_path).expanduser()
        if not file_path.is_absolute():
            file_path = (BASE_DIR / file_path).resolve()

        if not file_path.exists():
            return {"error": "File not found."}

        individual_id = uuid4().hex[:12]
        try:
            presentations[individual_id] = PresentationFile(str(file_path))
        except Exception as e:
            return {
                "error": f"Failed to open presentation package: {str(e)}",
                "file_path": str(file_path),
            }
        
        return {
            "message": "Presentation file opened successfully",
            "presentation_id": individual_id,
            "file_info": presentations[individual_id].get_file_info()
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Save Presentation File",
            readOnlyHint=False,
        ),
        description="Saves the presentation file to disk and returns its information. \
        The presentation_id must be provided for this operation. \
        This tool must be run each time a modifying iteration is completed. \
        Placeholder-first policy: when adding content to slides later, prefer placeholder tools before manual shape tools."
    )
    def save_presentation_file(presentation_id: str, file_path: str = None, file_name: str = None) -> Dict:
        """Save the presentation file to disk."""
        if presentation_id not in presentations:
            return {"error": "Presentation ID not found."}
        
        presentation_file = presentations[presentation_id]
        requested_destination = None

        if file_path:
            expanded_path = Path(os.path.expandvars(file_path)).expanduser()
            # Backward compatibility: when file_name is provided, treat file_path as a directory.
            requested_destination = expanded_path / file_name if file_name else expanded_path

        try:
            if requested_destination is not None:
                requested_destination.parent.mkdir(parents=True, exist_ok=True)
                presentation_file.save(requested_destination)
                return {
                    "message": "Presentation file saved successfully",
                    "presentation_id": presentation_id,
                    "requested_path": str(requested_destination),
                    "effective_path": str(Path(presentation_file.file_path).resolve()),
                    "used_fallback": False,
                    "file_info": presentation_file.get_file_info(),
                }

            presentation_file.save()
            return {
                "message": "Presentation file saved successfully",
                "presentation_id": presentation_id,
                "requested_path": None,
                "effective_path": str(Path(presentation_file.file_path).resolve()),
                "used_fallback": False,
                "file_info": presentation_file.get_file_info(),
            }
        except Exception as requested_save_error:
            if requested_destination is not None:
                # Fallback to the current temp/original location if requested destination fails.
                try:
                    presentation_file.save()
                    return {
                        "message": "Presentation file saved using fallback path",
                        "presentation_id": presentation_id,
                        "requested_path": str(requested_destination),
                        "effective_path": str(Path(presentation_file.file_path).resolve()),
                        "used_fallback": True,
                        "fallback_reason": str(requested_save_error),
                        "file_info": presentation_file.get_file_info(),
                    }
                except Exception as fallback_save_error:
                    return {
                        "error": f"Failed to save presentation file: requested path error: {requested_save_error}; fallback error: {fallback_save_error}",
                        "presentation_id": presentation_id,
                        "requested_path": str(requested_destination),
                    }

            return {
                "error": f"Failed to save presentation file: {str(requested_save_error)}",
                "presentation_id": presentation_id,
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="Cleanup Presentation File",
            readOnlyHint=False,
        ),
        description="Removes a presentation file from the global state and deletes the file from disk. \
        The presentation_id must be provided for this operation. \
        Use this tool to clean up temporary files created during testing or if a file is no longer needed. \
        Placeholder-first policy: when adding content to slides later, prefer placeholder tools before manual shape tools."
    )
    def cleanup_presentation_file(presentation_id: str) -> Dict:
        """Remove a presentation file from global state and delete the file from disk."""
        if presentation_id not in presentations:
            return {"error": "Presentation ID not found."}
        
        presentation_file = presentations.pop(presentation_id)
        file_path = Path(presentation_file.file_path)
        
        try:
            if file_path.exists():
                file_path.unlink()
                return {
                    "message": "Presentation file cleaned up successfully",
                    "presentation_id": presentation_id,
                    "file_path": str(file_path)
                }
            else:
                return {
                    "message": "Presentation file removed from state, but file not found on disk",
                    "presentation_id": presentation_id,
                    "file_path": str(file_path)
                }
        except Exception as e:
            return {
                "error": f"Failed to delete file: {str(e)}",
                "presentation_id": presentation_id,
                "file_path": str(file_path)
            }