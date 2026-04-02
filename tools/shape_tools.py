from mcp.types import ToolAnnotations

from utils.helper_methods import get_slide_by_id, resolve_picture_source
from utils.models import PresentationFile


def register_shape_tools(app, presentations: dict[str, PresentationFile]):
	"""Register shape tools for PowerPoint file management."""

	def _remove_shapes_by_ids(
		presentation_id: str,
		slide_id: int,
		shape_ids: list[int],
	) -> dict:
		"""Shared implementation for single and bulk shape removal."""
		if presentation_id not in presentations:
			return {"error": "Presentation ID not found."}

		presentation_file = presentations[presentation_id]
		pptx_object = presentation_file.get_pptx_object()

		slide_to_update = get_slide_by_id(pptx_object, slide_id)
		if slide_to_update is None:
			return {"error": "Slide ID not found."}

		if not isinstance(shape_ids, list) or len(shape_ids) == 0:
			return {"error": "shape_ids must be a non-empty list."}

		shape_ids_int: list[int] = []
		invalid_shape_ids: list = []
		for raw_shape_id in shape_ids:
			try:
				shape_ids_int.append(int(raw_shape_id))
			except (TypeError, ValueError):
				invalid_shape_ids.append(raw_shape_id)

		if invalid_shape_ids:
			return {
				"error": "All shape_ids must be numeric values.",
				"invalid_shape_ids": invalid_shape_ids,
			}

		requested_shape_ids = list(dict.fromkeys(shape_ids_int))
		shape_map = {shape.shape_id: shape for shape in slide_to_update.shapes}

		removed_shapes: list[dict] = []
		not_found_shape_ids: list[int] = []

		for shape_id_int in requested_shape_ids:
			shape_to_remove = shape_map.get(shape_id_int)
			if shape_to_remove is None:
				not_found_shape_ids.append(shape_id_int)
				continue

			shape_element = shape_to_remove.element
			shape_parent = shape_element.getparent()
			if shape_parent is None:
				not_found_shape_ids.append(shape_id_int)
				continue

			removed_shapes.append(
				{
					"shape_id": shape_to_remove.shape_id,
					"name": shape_to_remove.name,
					"shape_type": str(shape_to_remove.shape_type),
					"shape_type_value": int(shape_to_remove.shape_type),
				}
			)
			shape_parent.remove(shape_element)

		if not removed_shapes:
			return {
				"error": "No matching shapes were removed.",
				"presentation_id": presentation_id,
				"slide_id": slide_id,
				"requested_shape_ids": requested_shape_ids,
				"not_found_shape_ids": not_found_shape_ids,
			}

		return {
			"message": "Shapes removed from slide successfully",
			"presentation_id": presentation_id,
			"slide_id": slide_id,
			"requested_shape_ids": requested_shape_ids,
			"removed_shapes": removed_shapes,
			"not_found_shape_ids": not_found_shape_ids,
		}

	@app.tool(
		annotations=ToolAnnotations(
			title="Add image to slide",
			readOnlyHint=False,
		),
		description="Adds an image shape to a slide at an exact position and size.\
			The presentation_id, slide_id, image_source, left, top, width, and height must be provided.\
			image_source can be a local file path or an http/https URL.\
			Position and size values are expected in EMUs."
	)
	def add_image_to_slide(
		presentation_id: str,
		slide_id: int,
		image_source: str,
		left: int,
		top: int,
		width: int,
		height: int,
	) -> dict:
		"""Add an image shape to a slide at exact coordinates and dimensions."""
		if presentation_id not in presentations:
			return {"error": "Presentation ID not found."}

		presentation_file = presentations[presentation_id]
		pptx_object = presentation_file.get_pptx_object()

		slide_to_update = get_slide_by_id(pptx_object, slide_id)
		if slide_to_update is None:
			return {"error": "Slide ID not found."}

		try:
			left_emu = int(left)
			top_emu = int(top)
			width_emu = int(width)
			height_emu = int(height)
		except (TypeError, ValueError):
			return {"error": "left, top, width, and height must be numeric values."}

		picture_source, picture_source_label, source_error = resolve_picture_source(image_source)
		if source_error is not None:
			return source_error

		inserted_picture = slide_to_update.shapes.add_picture(
			picture_source,
			left_emu,
			top_emu,
			width_emu,
			height_emu,
		)

		return {
			"message": "Image added to slide successfully",
			"presentation_id": presentation_id,
			"slide_id": slide_id,
			"image_source": picture_source_label,
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

	@app.tool(
		annotations=ToolAnnotations(
			title="Add text to slide",
			readOnlyHint=False,
		),
		description="Adds a text box shape to a slide at an exact position and size.\
			The presentation_id, slide_id, text, left, top, width, and height must be provided.\
			Position and size values are expected in EMUs."
	)
	def add_text_to_slide(
		presentation_id: str,
		slide_id: int,
		text: str,
		left: int,
		top: int,
		width: int,
		height: int,
	) -> dict:
		"""Add a text box shape to a slide at exact coordinates and dimensions."""
		if presentation_id not in presentations:
			return {"error": "Presentation ID not found."}

		presentation_file = presentations[presentation_id]
		pptx_object = presentation_file.get_pptx_object()

		slide_to_update = get_slide_by_id(pptx_object, slide_id)
		if slide_to_update is None:
			return {"error": "Slide ID not found."}

		try:
			left_emu = int(left)
			top_emu = int(top)
			width_emu = int(width)
			height_emu = int(height)
		except (TypeError, ValueError):
			return {"error": "left, top, width, and height must be numeric values."}

		new_textbox = slide_to_update.shapes.add_textbox(
			left_emu,
			top_emu,
			width_emu,
			height_emu,
		)
		new_textbox.text_frame.clear()
		new_textbox.text_frame.text = text

		return {
			"message": "Text added to slide successfully",
			"presentation_id": presentation_id,
			"slide_id": slide_id,
			"shape": {
				"shape_id": new_textbox.shape_id,
				"name": new_textbox.name,
				"shape_type": str(new_textbox.shape_type),
				"shape_type_value": int(new_textbox.shape_type),
				"left": int(new_textbox.left),
				"top": int(new_textbox.top),
				"width": int(new_textbox.width),
				"height": int(new_textbox.height),
				"text": new_textbox.text_frame.text,
			},
		}

	@app.tool(
		annotations=ToolAnnotations(
			title="Remove shapes from slide",
			readOnlyHint=False,
		),
		description="Removes multiple shapes from a slide using shape_ids.\
			The presentation_id, slide_id, and shape_ids must be provided.\
			Use get_slide_content to find the correct shape_ids on a slide."
	)
	def remove_shapes_from_slide(
		presentation_id: str,
		slide_id: int,
		shape_ids: list[int],
	) -> dict:
		"""Remove multiple shapes from a slide by shape_id."""
		return _remove_shapes_by_ids(presentation_id, slide_id, shape_ids)
