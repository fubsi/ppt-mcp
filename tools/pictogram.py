import importlib
from pathlib import Path
from typing import Dict, Optional
from tempfile import NamedTemporaryFile

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from utils.helper_methods import get_slide_by_id
from utils.models import PictogramLibrary, PresentationFile

def register_pictogram_tools(app: FastMCP, presentations: Dict[str, PresentationFile]):
    """Register pictogram tools for PowerPoint file management."""

    pictogram_library = PictogramLibrary()

    @app.tool(
        annotations=ToolAnnotations(
            title="Get Pictogram List",
            readOnlyHint=True,
        ),
        description="Returns a list of all available pictograms in the library. \
        Placeholder-first policy: when adding content to slides, prefer placeholder tools before manual shape tools."
    )
    def get_pictogram_list() -> Dict:
        """Return a list of all available pictograms."""
        return {
            "message": "Available pictograms:",
            "pictograms": list(pictogram_library.pictograms.keys())
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Add Pictogram to Slide",
            readOnlyHint=False,
        ),
        description="Adds a pictogram image from the pictogram library to a slide at an exact position. \
        The presentation_id, slide_id, pictogram_name, left, and top must be provided. \
        Optional width and height can be provided to resize the pictogram. \
        Position and size values are expected in EMUs. \
        Placeholder-first policy: prefer placeholder tools before adding shapes manually."
    )
    def add_pictogram_to_slide(
        presentation_id: str,
        slide_id: int,
        pictogram_name: str,
        top: int = 0,
        left: int = 0,
        width: Optional[int] = None,
        height: Optional[int] = None,
    ) -> Dict:
        """Add a pictogram from the local library to a specific slide."""
        if presentation_id not in presentations:
            return {"error": "Presentation ID not found."}

        presentation_file = presentations[presentation_id]
        pptx_object = presentation_file.get_pptx_object()

        slide_to_update = get_slide_by_id(pptx_object, slide_id)
        if slide_to_update is None:
            return {"error": "Slide ID not found."}

        if pictogram_name not in pictogram_library.pictograms:
            return {
                "error": "Pictogram not found in library.",
                "pictogram_name": pictogram_name,
            }

        pictogram = pictogram_library.pictograms[pictogram_name]

        pictogram_path = Path(pictogram.image_path)
        if not pictogram_path.exists() or not pictogram_path.is_file():
            return {
                "error": "Pictogram image file not found.",
                "pictogram_name": pictogram_name,
                "image_path": str(pictogram_path),
            }

        insert_path = pictogram_path
        temp_converted_path = None
        supported_extensions = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".wmf"}

        if pictogram_path.suffix.lower() == ".svg":
            try:
                cairosvg = importlib.import_module("cairosvg")
            except ImportError:
                return {
                    "error": "SVG pictograms are not directly supported by python-pptx. Install cairosvg to enable SVG conversion.",
                    "pictogram_name": pictogram_name,
                    "image_path": str(pictogram_path),
                }

            try:
                png_bytes = cairosvg.svg2png(url=str(pictogram_path))
                with NamedTemporaryFile(delete=False, suffix=".png") as tmp_file:
                    tmp_file.write(png_bytes)
                    temp_converted_path = Path(tmp_file.name)
                insert_path = temp_converted_path
            except Exception as exc:
                return {
                    "error": "Failed to convert SVG pictogram to PNG.",
                    "pictogram_name": pictogram_name,
                    "image_path": str(pictogram_path),
                    "details": str(exc),
                }
        elif pictogram_path.suffix.lower() not in supported_extensions:
            return {
                "error": "Unsupported pictogram image format for python-pptx.",
                "pictogram_name": pictogram_name,
                "image_path": str(pictogram_path),
                "supported_extensions": sorted(supported_extensions),
            }

        try:
            left_emu = int(left)
            top_emu = int(top)
        except (TypeError, ValueError):
            return {"error": "left and top must be numeric values."}

        try:
            width_emu = int(width) if width is not None else None
            height_emu = int(height) if height is not None else None
        except (TypeError, ValueError):
            return {"error": "width and height must be numeric values when provided."}

        try:
            inserted_picture = slide_to_update.shapes.add_picture(
                str(insert_path),
                left_emu,
                top_emu,
                width_emu,
                height_emu,
            )
        except Exception as exc:
            return {
                "error": "Failed to add pictogram to slide.",
                "pictogram_name": pictogram_name,
                "image_path": str(pictogram_path),
                "details": str(exc),
            }
        finally:
            if temp_converted_path and temp_converted_path.exists():
                try:
                    temp_converted_path.unlink()
                except OSError:
                    pass

        return {
            "message": "Pictogram added to slide successfully",
            "presentation_id": presentation_id,
            "slide_id": slide_id,
            "pictogram_name": pictogram_name,
            "image_path": str(pictogram_path),
            "shape": {
                "shape_id": inserted_picture.shape_id,
                "name": inserted_picture.name,
                "shape_type": str(inserted_picture.shape_type),
                "shape_type_value": int(inserted_picture.shape_type),
                "left": int(inserted_picture.left),
                "top": int(inserted_picture.top),
                "width": int(inserted_picture.width),
                "height": int(inserted_picture.height),
            },
        }