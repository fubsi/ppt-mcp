from mcp import FastMCP

from typing import Dict

from mcp.types import ToolAnnotations

from utils.models import PresentationFile, PictogramLibrary

def register_pictogram_tools(app: FastMCP, presentations: Dict[str, PresentationFile]):
    """Register simple PowerPoint file management tools."""

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
        description="Adds a specified pictogram to a slide in the presentation. \
        Requires the presentation_id, slide_index, and pictogram_name as input. \
        Uses top and left positioning for the pictogram, which can be adjusted as needed."
    )
    def add_pictogram_to_slide(presentation_id: str, slide_index: int, pictogram_name: str, top: int = 0, left: int = 0) -> Dict:
        """Add a specified pictogram to a slide in the presentation."""
        if presentation_id not in presentations:
            return {"error": "Presentation ID not found."}
        
        presentation_file = presentations[presentation_id]
        pptx = presentation_file.get_pptx_object()
        
        if slide_index < 0 or slide_index >= len(pptx.slides):
            return {"error": "Invalid slide index."}
        
        if pictogram_name not in pictogram_library.pictograms:
            return {"error": "Pictogram not found in library."}
        
        pictogram = pictogram_library.pictograms[pictogram_name]
        slide = pptx.slides[slide_index]
        
        # Add the pictogram image to the slide
        slide.shapes.add_picture(pictogram.image_path, left, top)
        
        # Save the updated presentation
        presentation_file.save()
        
        return {
            "message": f"Pictogram '{pictogram_name}' added to slide {slide_index} of presentation '{presentation_id}'."
        }