from pathlib import Path
from typing import Dict

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
        description="Returns a list of all available pictograms in the library."
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
        Position values are expected in EMUs."
    )
    def add_pictogram_to_slide(
        presentation_id: str,
        slide_id: int,
        pictogram_name: str,
        top: int = 0,
        left: int = 0,
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

        pictogram_path = Path(pictogram.image_path).expanduser().resolve()
        if not pictogram_path.exists() or not pictogram_path.is_file():
            return {
                "error": "Pictogram image file not found.",
                "pictogram_name": pictogram_name,
                "image_path": str(pictogram_path),
            }

        try:
            left_emu = int(left)
            top_emu = int(top)
        except (TypeError, ValueError):
            return {"error": "left and top must be numeric values."}

        inserted_picture = slide_to_update.shapes.add_picture(
            str(pictogram_path),
            left_emu,
            top_emu,
        )

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