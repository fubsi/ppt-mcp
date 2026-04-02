from mcp.types import ToolAnnotations
from io import BytesIO
from pathlib import Path
from pptx.enum.shapes import PP_PLACEHOLDER
from urllib.error import URLError, HTTPError
from urllib.parse import urlparse
from urllib.request import urlopen

def register_placeholder_tools(app, presentations):
    """Register placeholder tools for PowerPoint file management."""

    @app.tool(
        annotations=ToolAnnotations(
            title="Insert picture into placeholder",
            readOnlyHint=False,
        ),
        description="Inserts an image into a PICTURE placeholder on a specific slide.\
            The presentation_id, slide_id, placeholder_shape_id, and image_path must be provided for this operation.\
            image_path can be a local file path or an http/https URL.\
            Use get_slides to find slide_id and placeholder_shape_id values."
    )
    def insert_picture_into_placeholder(
        presentation_id: str,
        slide_id: int,
        placeholder_shape_id: int,
        image_path: str,
    ) -> dict:
        """Insert an image from a local path or URL into a picture placeholder."""
        if presentation_id not in presentations:
            return {"error": "Presentation ID not found."}

        parsed_image_path = urlparse(image_path)
        image_source = None
        image_source_label = image_path

        if parsed_image_path.scheme in ("http", "https"):
            try:
                with urlopen(image_path, timeout=10) as response:
                    image_data = response.read()
                if not image_data:
                    return {"error": "Image URL returned no data."}
                image_source = BytesIO(image_data)
            except (HTTPError, URLError, TimeoutError, ValueError) as exc:
                return {
                    "error": "Failed to download image from URL.",
                    "details": str(exc),
                }
        else:
            resolved_image_path = Path(image_path).expanduser().resolve()
            if not resolved_image_path.exists() or not resolved_image_path.is_file():
                return {"error": "Image file not found."}
            image_source = str(resolved_image_path)
            image_source_label = str(resolved_image_path)

        presentation_file = presentations[presentation_id]
        pptx_object = presentation_file.get_pptx_object()

        slide_to_update = None
        for slide in pptx_object.slides:
            if slide.slide_id == slide_id:
                slide_to_update = slide
                break

        if slide_to_update is None:
            return {"error": "Slide ID not found."}

        target_placeholder = None
        for placeholder in slide_to_update.placeholders:
            if placeholder.shape_id == placeholder_shape_id:
                target_placeholder = placeholder
                break

        if target_placeholder is None:
            return {"error": "Placeholder shape ID not found on this slide."}

        placeholder_type = target_placeholder.placeholder_format.type
        if int(placeholder_type) != int(PP_PLACEHOLDER.PICTURE):
            return {
                "error": "Target placeholder is not a PICTURE placeholder.",
                "placeholder_type": str(placeholder_type),
                "placeholder_type_value": int(placeholder_type),
            }

        inserted_picture = target_placeholder.insert_picture(image_source)

        return {
            "message": "Picture inserted into placeholder successfully",
            "presentation_id": presentation_id,
            "slide_id": slide_id,
            "placeholder_shape_id": placeholder_shape_id,
            "image_path": image_source_label,
            "inserted_picture": {
                "shape_id": inserted_picture.shape_id,
                "name": inserted_picture.name,
                "shape_type": str(inserted_picture.shape_type),
                "left": int(inserted_picture.left),
                "top": int(inserted_picture.top),
                "width": int(inserted_picture.width),
                "height": int(inserted_picture.height),
            },
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Insert text into placeholder",
            readOnlyHint=False,
        ),
        description="Inserts text into a text-capable placeholder on a specific slide.\
            The presentation_id, slide_id, placeholder_shape_id, and text must be provided for this operation.\
            Use get_slides to find slide_id and placeholder_shape_id values."
    )
    def insert_text_into_placeholder(
        presentation_id: str,
        slide_id: int,
        placeholder_shape_id: int,
        text: str,
    ) -> dict:
        """Insert text into an existing text-capable placeholder."""
        if presentation_id not in presentations:
            return {"error": "Presentation ID not found."}

        presentation_file = presentations[presentation_id]
        pptx_object = presentation_file.get_pptx_object()

        slide_to_update = None
        for slide in pptx_object.slides:
            if slide.slide_id == slide_id:
                slide_to_update = slide
                break

        if slide_to_update is None:
            return {"error": "Slide ID not found."}

        target_placeholder = None
        for placeholder in slide_to_update.placeholders:
            if placeholder.shape_id == placeholder_shape_id:
                target_placeholder = placeholder
                break

        if target_placeholder is None:
            return {"error": "Placeholder shape ID not found on this slide."}

        placeholder_type = target_placeholder.placeholder_format.type
        if not hasattr(target_placeholder, "text_frame") or target_placeholder.text_frame is None:
            return {
                "error": "Target placeholder does not support text insertion.",
                "placeholder_type": str(placeholder_type),
                "placeholder_type_value": int(placeholder_type),
            }

        target_placeholder.text_frame.clear()
        target_placeholder.text_frame.text = text

        return {
            "message": "Text inserted into placeholder successfully",
            "presentation_id": presentation_id,
            "slide_id": slide_id,
            "placeholder_shape_id": placeholder_shape_id,
            "text": target_placeholder.text_frame.text,
            "placeholder_type": str(placeholder_type),
            "placeholder_type_value": int(placeholder_type),
        }