from __future__ import annotations

from io import BytesIO
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import urlopen


def get_slide_by_id(pptx_object, slide_id: int):
	"""Return the slide object matching slide_id or None."""
	for slide in pptx_object.slides:
		if slide.slide_id == slide_id:
			return slide
	return None


def get_slide_with_index_by_id(pptx_object, slide_id: int):
	"""Return (slide, index) for slide_id or (None, None) when missing."""
	for index, slide in enumerate(pptx_object.slides):
		if slide.slide_id == slide_id:
			return slide, index
	return None, None


def get_placeholder_by_shape_id(slide, placeholder_shape_id: int):
	"""Return placeholder object matching placeholder_shape_id or None."""
	for placeholder in slide.placeholders:
		if placeholder.shape_id == placeholder_shape_id:
			return placeholder
	return None


def resolve_picture_source(image_source: str):
	"""Resolve local/URL image source to (source, source_label, error_dict)."""
	parsed = urlparse(image_source)
	if parsed.scheme in ("http", "https"):
		try:
			with urlopen(image_source, timeout=10) as response:
				image_data = response.read()
			if not image_data:
				return None, None, {"error": "Image URL returned no data."}
			return BytesIO(image_data), image_source, None
		except (HTTPError, URLError, TimeoutError, ValueError) as exc:
			return None, None, {
				"error": "Failed to download image from URL.",
				"details": str(exc),
			}

	resolved_path = Path(image_source).expanduser().resolve()
	if not resolved_path.exists() or not resolved_path.is_file():
		return None, None, {"error": "Image file not found."}

	return str(resolved_path), str(resolved_path), None


def remove_shapes_by_ids(
	presentations: dict,
	presentation_id: str,
	slide_id: int,
	shape_ids: list[int],
) -> dict:
	"""Remove one or many shapes by shape IDs from a specific slide."""
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


def extract_text_from_shape(shape) -> str | None:
	"""Safely extract visible text from a shape."""
	if hasattr(shape, "text") and shape.text:
		return shape.text
	if hasattr(shape, "text_frame") and shape.text_frame is not None:
		return shape.text_frame.text
	return None


def serialize_placeholder(placeholder) -> dict:
	"""Convert a placeholder into a serializable dictionary."""
	return {
		"placeholder_shape_id": placeholder.shape_id,
		"placeholder_type": str(placeholder.placeholder_format.type),
		"placeholder_type_value": int(placeholder.placeholder_format.type),
		"placeholder_text": extract_text_from_shape(placeholder),
	}


def serialize_shape(shape) -> dict:
	"""Convert a shape into a serializable dictionary."""
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
		"shape_text": extract_text_from_shape(shape),
	}


def serialize_slide(slide, slide_index: int) -> dict:
	"""Convert a slide into a serializable dictionary."""
	notes_text = None
	if slide.has_notes_slide and slide.notes_slide is not None:
		notes_text = slide.notes_slide.notes_text_frame.text

	placeholders = [serialize_placeholder(p) for p in slide.placeholders]
	shapes = [serialize_shape(s) for s in slide.shapes]

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
