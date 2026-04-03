from mcp.types import ToolAnnotations

from utils.helper_methods import (
	analyze_text_overflow_in_shape,
	get_slide_by_id,
	remove_shapes_by_ids,
	resolve_picture_source,
	serialize_shape,
)
from utils.models import PresentationFile


def register_shape_tools(app, presentations: dict[str, PresentationFile]):
	"""Register shape tools for PowerPoint file management."""

	def _shapes_collide(shape_a, shape_b) -> bool:
		"""Return True when two shape bounding boxes overlap (strict overlap)."""
		left_a = int(shape_a.left)
		top_a = int(shape_a.top)
		right_a = left_a + int(shape_a.width)
		bottom_a = top_a + int(shape_a.height)

		left_b = int(shape_b.left)
		top_b = int(shape_b.top)
		right_b = left_b + int(shape_b.width)
		bottom_b = top_b + int(shape_b.height)

		horizontal_overlap = left_a < right_b and right_a > left_b
		vertical_overlap = top_a < bottom_b and bottom_a > top_b
		return horizontal_overlap and vertical_overlap

	@app.tool(
		annotations=ToolAnnotations(
			title="Add image to slide",
			readOnlyHint=False,
		),
		description="Adds an image shape to a slide at an exact position and size.\
			The presentation_id, slide_id, image_source, left, top, width, and height must be provided.\
			image_source can be a local file path or an http/https URL.\
			Position and size values are expected in EMUs.\
			Placeholder-first policy: prefer placeholder tools before adding shapes manually."
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
			Position and size values are expected in EMUs.\
			Placeholder-first policy: prefer placeholder tools before adding shapes manually."
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
		overflow_analysis = analyze_text_overflow_in_shape(new_textbox)

		return {
			"message": "Text added to slide successfully",
			"presentation_id": presentation_id,
			"slide_id": slide_id,
			"text_overflow_detected": overflow_analysis["overflow_detected"],
			"text_overflow_analysis": overflow_analysis,
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
			Use get_slide_content to find the correct shape_ids on a slide.\
			Placeholder-first policy: prefer placeholder tools before adding shapes manually."
	)
	def remove_shapes_from_slide(
		presentation_id: str,
		slide_id: int,
		shape_ids: list[int],
	) -> dict:
		"""Remove multiple shapes from a slide by shape_id."""
		return remove_shapes_by_ids(presentations, presentation_id, slide_id, shape_ids)

	@app.tool(
		annotations=ToolAnnotations(
			title="Move and resize shape",
			readOnlyHint=False,
		),
		description="Moves a shape and updates its size on a slide in one operation.\
			The presentation_id, slide_id, shape_id, left, top, width, and height must be provided.\
			Position and size values are expected in EMUs.\
			Placeholder-first policy: prefer placeholder tools before adding shapes manually."
	)
	def move_and_resize_shape(
		presentation_id: str,
		slide_id: int,
		shape_id: int,
		left: int,
		top: int,
		width: int,
		height: int,
	) -> dict:
		"""Move and resize an existing shape in one operation."""
		if presentation_id not in presentations:
			return {"error": "Presentation ID not found."}

		presentation_file = presentations[presentation_id]
		pptx_object = presentation_file.get_pptx_object()

		slide_to_update = get_slide_by_id(pptx_object, slide_id)
		if slide_to_update is None:
			return {"error": "Slide ID not found."}

		try:
			shape_id_int = int(shape_id)
			left_emu = int(left)
			top_emu = int(top)
			width_emu = int(width)
			height_emu = int(height)
		except (TypeError, ValueError):
			return {
				"error": "shape_id, left, top, width, and height must be numeric values."
			}

		target_shape = None
		for shape in slide_to_update.shapes:
			if shape.shape_id == shape_id_int:
				target_shape = shape
				break

		if target_shape is None:
			return {
				"error": "Shape ID not found on this slide.",
				"presentation_id": presentation_id,
				"slide_id": slide_id,
				"shape_id": shape_id_int,
			}

		shape_before = serialize_shape(target_shape)

		target_shape.left = left_emu
		target_shape.top = top_emu
		target_shape.width = width_emu
		target_shape.height = height_emu

		return {
			"message": "Shape moved and resized successfully",
			"presentation_id": presentation_id,
			"slide_id": slide_id,
			"shape_id": target_shape.shape_id,
			"shape_before": shape_before,
			"shape_after": serialize_shape(target_shape),
		}

	@app.tool(
		annotations=ToolAnnotations(
			title="Check shape collisions",
			readOnlyHint=True,
		),
		description="Checks all slides in a presentation for shape collisions (overlapping bounds).\
			The presentation_id must be provided.\
			Returns per-slide collision pairs including shape metadata for each colliding shape.\
			Run this at the end of every modifying iteration to validate that no collisions are occurring."
	)
	def check_shape_collisions(presentation_id: str) -> dict:
		"""Check all slides for colliding shapes and return detailed collision pairs."""
		if presentation_id not in presentations:
			return {"error": "Presentation ID not found."}

		presentation_file = presentations[presentation_id]
		pptx_object = presentation_file.get_pptx_object()

		slides_with_collisions: list[dict] = []
		total_collision_pairs = 0

		for slide_index, slide in enumerate(pptx_object.slides):
			shape_list = list(slide.shapes)
			slide_collisions: list[dict] = []

			for i, shape_a in enumerate(shape_list):
				for shape_b in shape_list[i + 1 :]:
					if _shapes_collide(shape_a, shape_b):
						slide_collisions.append(
							{
								"shape_a": serialize_shape(shape_a),
								"shape_b": serialize_shape(shape_b),
							}
						)

			if slide_collisions:
				total_collision_pairs += len(slide_collisions)
				slides_with_collisions.append(
					{
						"slide_id": slide.slide_id,
						"slide_index": slide_index,
						"slide_name": slide.name,
						"collision_pairs": slide_collisions,
						"collision_pairs_count": len(slide_collisions),
					}
				)

		return {
			"message": "Shape collision check completed",
			"presentation_id": presentation_id,
			"slides_checked": len(pptx_object.slides),
			"slides_with_collisions_count": len(slides_with_collisions),
			"total_collision_pairs": total_collision_pairs,
			"slides_with_collisions": slides_with_collisions,
		}
