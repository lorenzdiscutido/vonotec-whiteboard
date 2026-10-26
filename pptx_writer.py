# pptx_writer.py
"""Generates a PowerPoint (.pptx) report from chronologically sorted daily photo chunks.

Loads `template.pptx`, resizes high-resolution photos in memory using PIL
(`Image.thumbnail`, max width 1024px) so original files are untouched and the
presentation stays lightweight, and injects both images and extracted metadata
(Date, Drop, Floor, Defect) into each slide's placeholders.
"""
import io
import os
from pathlib import Path

from PIL import Image, ImageOps
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches, Pt

TEMPLATE_PATH = Path(__file__).with_name("template.pptx")
SLIDE_LAYOUT_INDEX = 0
MAX_IMAGE_WIDTH = 1024
JPEG_QUALITY = 85

MATERIALS = ("Sealant", "Concrete", "Paint", "Gasket", "With Film", "Broken Glass")


def optimize_image_in_memory(image_source, max_width=MAX_IMAGE_WIDTH):
    """Resizes a raw image in memory using PIL `Image.thumbnail` (max width 1024px)
    and returns a seeked `io.BytesIO` JPEG stream ready for python-pptx."""
    if isinstance(image_source, (bytes, bytearray)):
        src_stream = io.BytesIO(image_source)
    elif hasattr(image_source, "read"):
        image_source.seek(0)
        src_stream = image_source
    else:
        src_stream = image_source  # file path

    with Image.open(src_stream) as img:
        img = ImageOps.exif_transpose(img)
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        elif img.mode == "L":
            img = img.convert("RGB")

        # Proportional in-memory downscaling capped at max_width (1024px)
        max_height = int(round(img.height * (max_width / max(img.width, 1)))) if img.width > max_width else img.height
        img.thumbnail((max_width, max(max_height, max_width)), Image.Resampling.LANCZOS)

        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=JPEG_QUALITY, optimize=True)
        buffer.seek(0)
        return buffer


def format_defect_text(parsed_data):
    """Formats the extracted defect summary from `parsed_data`.
    Checks for an explicit 'Defect' key first, then aggregates material damages & dimensions."""
    if not isinstance(parsed_data, dict):
        return "None"

    direct_defect = parsed_data.get("Defect")
    if isinstance(direct_defect, str) and direct_defect.strip():
        return direct_defect.strip()
    if isinstance(direct_defect, list) and direct_defect:
        cleaned = [str(x).strip() for x in direct_defect if str(x).strip()]
        if cleaned:
            return ", ".join(cleaned)

    entries = []
    for material in MATERIALS:
        damages = parsed_data.get(f"{material} Damage") or []
        dimensions = parsed_data.get(f"{material} Dimension") or []
        for idx, dmg in enumerate(damages):
            code = str(dmg).strip()
            if not code:
                continue
            dim = str(dimensions[idx]).strip() if idx < len(dimensions) else ""
            entries.append(f"{code} ({dim})" if dim else code)

    return ", ".join(entries) if entries else "No defect recorded"


def _extract_metadata(photo_item, chunk_date=""):
    """Returns normalized (date_str, drop_str, floor_str, defect_str) for one photo item."""
    parsed = photo_item.get("parsed_data") or {}
    date_str = (
        str(photo_item.get("date_display") or photo_item.get("date") or parsed.get("Date") or chunk_date or "N/A").strip()
    )
    drop_str = str(photo_item.get("drop") or parsed.get("Drop") or "N/A").strip() or "N/A"
    floor_str = str(photo_item.get("floor") or parsed.get("Floor") or "N/A").strip() or "N/A"
    defect_str = str(photo_item.get("defect") or format_defect_text(parsed) or "N/A").strip() or "N/A"
    return date_str, drop_str, floor_str, defect_str


def _format_caption(date_str, drop_str, floor_str, defect_str):
    """Formats the metadata caption injected into a photo's text placeholder."""
    return (
        f"Date: {date_str}\n"
        f"Drop: {drop_str} | Floor: {floor_str}\n"
        f"Defect: {defect_str}"
    )


def _sort_placeholders(placeholders):
    """Sorts slide placeholders in visual reading order (top-to-bottom, left-to-right, then idx)."""
    return sorted(
        placeholders,
        key=lambda ph: (
            int(getattr(ph, "top", 0) or 0),
            int(getattr(ph, "left", 0) or 0),
            int(ph.placeholder_format.idx),
        ),
    )


def _select_slide_layout(prs, chunk_len):
    """Selects the custom layout from the template that best matches a photo chunk."""
    layouts = list(prs.slide_layouts)
    if not layouts:
        raise RuntimeError("The PowerPoint template contains no slide layouts.")

    # Prefer a layout that has picture placeholders matching or accommodating the chunk
    for layout in layouts:
        pic_count = sum(
            1 for ph in layout.placeholders
            if ph.placeholder_format.type == PP_PLACEHOLDER.PICTURE or hasattr(ph, "insert_picture")
        )
        if pic_count >= chunk_len and pic_count > 0:
            return layout

    for layout in layouts:
        pic_count = sum(
            1 for ph in layout.placeholders
            if ph.placeholder_format.type == PP_PLACEHOLDER.PICTURE or hasattr(ph, "insert_picture")
        )
        if pic_count > 0:
            return layout

    idx = min(SLIDE_LAYOUT_INDEX, len(layouts) - 1)
    return layouts[idx]


def _fallback_grid_geometry(slot_index):
    """2x2 grid coordinates (left, top, width, height) when a template layout lacks picture placeholders."""
    col = slot_index % 2
    row = slot_index // 2
    left = Inches(0.6 + col * 6.1)
    top = Inches(1.1 + row * 3.0)
    width = Inches(3.2)
    height = Inches(2.5)
    return left, top, width, height


def _inject_chunk_into_slide(slide, chunk_items, chunk_date=""):
    """Injects up to 4 optimized images and their Date/Drop/Floor/Defect metadata into `slide`."""
    picture_placeholders = []
    text_placeholders = []

    for ph in slide.placeholders:
        ph_type = ph.placeholder_format.type
        if ph_type == PP_PLACEHOLDER.PICTURE or hasattr(ph, "insert_picture"):
            picture_placeholders.append(ph)
        elif ph.has_text_frame:
            text_placeholders.append(ph)

    picture_placeholders = _sort_placeholders(picture_placeholders)
    text_placeholders = _sort_placeholders(text_placeholders)

    # If there is a title/header placeholder in addition to per-photo text placeholders, populate it with the Date
    if text_placeholders and len(text_placeholders) == len(picture_placeholders) + 1 and chunk_date:
        header_ph = text_placeholders.pop(0)
        header_ph.text = f"Date: {chunk_date}"
    elif slide.shapes.title and slide.shapes.title in text_placeholders and len(text_placeholders) > len(chunk_items):
        title_ph = slide.shapes.title
        text_placeholders.remove(title_ph)
        if chunk_date:
            title_ph.text = f"Date: {chunk_date}"

    for idx, item in enumerate(chunk_items):
        date_str, drop_str, floor_str, defect_str = _extract_metadata(item, chunk_date)
        image_source = item.get("image_bytes") or item.get("image_path")

        # 1. Optimize image in memory (max width 1024px) and inject into picture placeholder
        if image_source is not None:
            img_stream = optimize_image_in_memory(image_source, max_width=MAX_IMAGE_WIDTH)
            if idx < len(picture_placeholders):
                picture_placeholders[idx].insert_picture(img_stream)
            else:
                left, top, width, _height = _fallback_grid_geometry(idx)
                slide.shapes.add_picture(img_stream, left, top, width=width)

        # 2. Inject extracted metadata (Date, Drop, Floor, Defect) into text placeholders
        # Case A: 4 separate text placeholders per photo slot (Date, Drop, Floor, Defect)
        if len(text_placeholders) >= len(chunk_items) * 4:
            base = idx * 4
            text_placeholders[base].text = date_str
            text_placeholders[base + 1].text = drop_str
            text_placeholders[base + 2].text = floor_str
            text_placeholders[base + 3].text = defect_str
        # Case B: 1 combined text placeholder per photo slot
        elif idx < len(text_placeholders):
            text_placeholders[idx].text = _format_caption(date_str, drop_str, floor_str, defect_str)
        # Case C: Fallback textbox when the slide layout does not provide enough text placeholders
        else:
            left, top, _width, _height = _fallback_grid_geometry(idx)
            tb = slide.shapes.add_textbox(left + Inches(3.3), top, Inches(2.6), Inches(2.4))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.text = _format_caption(date_str, drop_str, floor_str, defect_str)
            for paragraph in tf.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(11)


def write_chunks_to_pptx(daily_chunks, output_path, template_path=TEMPLATE_PATH):
    """Builds a PowerPoint presentation from `daily_chunks` and saves it to `output_path`.

    Parameters
    ----------
    daily_chunks : list[dict] or list[list[dict]]
        Chronologically ordered chunks of up to 4 photo records per day.
        Each item may be a dict `{"date": "...", "photos": [...]}` or a raw list of photo dicts.
    output_path : str or Path
        Destination `.pptx` file path.
    template_path : str or Path
        Path to `template.pptx` (loaded as the base presentation when present).
    """
    resolved_template = Path(template_path) if template_path else TEMPLATE_PATH
    if resolved_template.exists():
        prs = Presentation(str(resolved_template))
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

    for chunk_entry in daily_chunks:
        if isinstance(chunk_entry, dict):
            chunk_date = str(chunk_entry.get("date") or "")
            chunk_photos = chunk_entry.get("photos") or []
        else:
            chunk_photos = list(chunk_entry)
            chunk_date = str(chunk_photos[0].get("date_display") or chunk_photos[0].get("date") or "") if chunk_photos else ""

        if not chunk_photos:
            continue

        layout = _select_slide_layout(prs, len(chunk_photos))
        slide = prs.slides.add_slide(layout)
        _inject_chunk_into_slide(slide, chunk_photos, chunk_date=chunk_date)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(str(output_path))
    return str(output_path)


# Alias for external callers
generate_pptx_report = write_chunks_to_pptx
