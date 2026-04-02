from mcp.types import ToolAnnotations
from utils.models import PresentationFile

def register_slide_tools(app, presentations: dict[str, PresentationFile]):
    """Register slide tools for PowerPoint file management."""

    def _extract_text_from_shape(shape) -> str | None:
        """Safely extract visible text from a shape."""
        if hasattr(shape, "text") and shape.text:
            return shape.text
        if hasattr(shape, "text_frame") and shape.text_frame is not None:
            return shape.text_frame.text
        return None

    def _serialize_slide(slide, slide_index: int) -> dict:
        notes_text = None
        if slide.has_notes_slide and slide.notes_slide is not None:
            notes_text = slide.notes_slide.notes_text_frame.text

        placeholders = [_serialize_placeholder(p) for p in slide.placeholders]
        shapes = [_serialize_shape(s) for s in slide.shapes]

        return {
            "slide_id": slide.slide_id,
            "slide_index": slide_index,
            "name": slide.name,
            "slide_layout": slide.slide_layout.name,
            "has_notes": slide.has_notes_slide,
            "slide_notes": notes_text,
            "placeholders_count": len(placeholders),
            "shapes_count": len(shapes),
            "slide_placeholders": placeholders,
            "slide_shapes": shapes,
        }

    def _serialize_placeholder(placeholder) -> dict:
        return {
            "placeholder_shape_id": placeholder.shape_id,
            "placeholder_type": str(placeholder.placeholder_format.type),
            "placeholder_type_value": int(placeholder.placeholder_format.type),
            "placeholder_text": _extract_text_from_shape(placeholder),
        }

    def _serialize_shape(shape) -> dict:
        return {
            "shape_id": shape.shape_id,
            "name": shape.name,
            "shape_type": str(shape.shape_type),
            "shape_type_value": int(shape.shape_type),
            "left": int(shape.left),
            "top": int(shape.top),
            "width": int(shape.width),
            "height": int(shape.height),
            "has_text": hasattr(shape, "text_frame") and shape.text_frame is not None,
            "shape_text": _extract_text_from_shape(shape),
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Read possible slide layouts",
            readOnlyHint=False,
        ),
        description="Reads the possible slide layouts for a given presentation file."
    )
    def get_slide_layouts(presentation_id: str) -> dict:
        """Get the possible slide layouts for a presentation file."""
        if presentation_id not in presentations:
            return {"error": "Presentation ID not found."}
        
        presentation_file = presentations[presentation_id]
        slide_layouts = presentation_file.get_pptx_object().slide_layouts

        layouts_info = []
        for i, layout in enumerate(slide_layouts):
            layouts_info.append({"index": i, "name": layout.name})

        return {
            "message": "Slide layouts retrieved successfully",
            "presentation_id": presentation_id,
            "slide_layouts": layouts_info
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Get present slides",
            readOnlyHint=False,
        ),
        description="Reads the slides in a given presentation file."
    )
    def get_slides(presentation_id: str) -> dict:
        """Get the slides in a presentation file."""
        if presentation_id not in presentations:
            return {"error": "Presentation ID not found."}

        presentation_file = presentations[presentation_id]
        slides = presentation_file.get_slides()

        slides_info = [_serialize_slide(slide, index) for index, slide in enumerate(slides)]

        return {
            "message": "Slides retrieved successfully",
            "presentation_id": presentation_id,
            "slides": slides_info
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Add a slide",
            readOnlyHint=False,
        ),
        description="Adds a new slide to a given presentation file.\
            The presentation_id and layout_index must be provided for this operation.\
            Use the get_slide_layouts tool to find the correct layout_index for the desired slide layout."
    )
    def add_slide(presentation_id: str, layout_index: int, slide_name: str = None) -> dict:
        """Add a new slide to a presentation file."""
        if presentation_id not in presentations:
            return {"error": "Presentation ID not found."}

        presentation_file = presentations[presentation_id]
        pptx_object = presentation_file.get_pptx_object()

        if layout_index < 0 or layout_index >= len(pptx_object.slide_layouts):
            return {"error": "Invalid layout index."}

        slide_layout = pptx_object.slide_layouts[layout_index]
        new_slide = pptx_object.slides.add_slide(slide_layout)

        if slide_name is not None:
            new_slide.name = slide_name

        return {
            "message": "Slide added successfully",
            "presentation_id": presentation_id,
            "new_slide": {
                "slide_id": new_slide.slide_id,
                "name": new_slide.name,
                "slide_layout": new_slide.slide_layout.name
            }
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Remove a slide",
            readOnlyHint=False,
        ),
        description="Removes a slide from a given presentation file.\
            The presentation_id and slide_id must be provided for this operation.\
            Use the get_slides tool to find the correct slide_id for the desired slide."
    )
    def remove_slide(presentation_id: str, slide_id: int) -> dict:
        """Remove a slide from a presentation file."""
        if presentation_id not in presentations:
            return {"error": "Presentation ID not found."}

        presentation_file = presentations[presentation_id]
        pptx_object = presentation_file.get_pptx_object()

        slide_to_remove = None
        for slide in pptx_object.slides:
            if slide.slide_id == slide_id:
                slide_to_remove = slide
                break

        if not slide_to_remove:
            return {"error": "Slide ID not found."}

        pptx_object.slides.remove(slide_to_remove)

        return {
            "message": "Slide removed successfully",
            "presentation_id": presentation_id,
            "removed_slide_id": slide_id
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Rename a slide",
            readOnlyHint=False,
        ),
        description="Renames a slide in a given presentation file.\
            The presentation_id, slide_id, and new_name must be provided for this operation.\
            Use the get_slides tool to find the correct slide_id for the desired slide."
    )
    def rename_slide(presentation_id: str, slide_id: int, new_name: str) -> dict:
        """Rename a slide in a presentation file."""
        if presentation_id not in presentations:
            return {"error": "Presentation ID not found."}

        presentation_file = presentations[presentation_id]
        pptx_object = presentation_file.get_pptx_object()

        slide_to_rename = None
        for slide in pptx_object.slides:
            if slide.slide_id == slide_id:
                slide_to_rename = slide
                break

        if not slide_to_rename:
            return {"error": "Slide ID not found."}

        slide_to_rename.name = new_name

        return {
            "message": "Slide renamed successfully",
            "presentation_id": presentation_id,
            "renamed_slide_id": slide_id,
            "new_name": new_name
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Get slide content",
            readOnlyHint=True,
        ),
        description="Retrieves the content of a slide in a given presentation file.\
            The presentation_id and slide_id must be provided for this operation.\
            Use the get_slides tool to find the correct slide_id for the desired slide."
    )
    def get_slide_content(presentation_id: str, slide_id: int) -> dict:
        """Get the content of a slide in a presentation file."""
        if presentation_id not in presentations:
            return {"error": "Presentation ID not found."}

        presentation_file = presentations[presentation_id]
        pptx_object = presentation_file.get_pptx_object()

        slide_to_get = None
        slide_index = None
        for index, slide in enumerate(pptx_object.slides):
            if slide.slide_id == slide_id:
                slide_to_get = slide
                slide_index = index
                break

        if slide_to_get is None:
            return {"error": "Slide ID not found."}

        return {
            "message": "Slide content retrieved successfully",
            "presentation_id": presentation_id,
            "slide_content": _serialize_slide(slide_to_get, slide_index)
        }