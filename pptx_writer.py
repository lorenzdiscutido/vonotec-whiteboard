# pptx_writer.py
"""Generates a PowerPoint (.pptx) report from chronologically sorted daily photo chunks.

Loads `template.pptx` (when present), resizes high-resolution photos in memory using PIL
(`Image.thumbnail`, max width 1024px) so original files are untouched and the
presentation stays lightweight, and lays each chunk out as up to 4 PORTRAIT photos
side by side, evenly spaced, with the extracted metadata (Date, Drop, Floor, Defect)
captioned under each photo. The day's date goes into the slide title.
"""
import io
import os
from pathlib import Path

from PIL import Image, ImageOps
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

TEMPLATE_PATH = Path(__file__).with_name("template.pptx")
MAX_IMAGE_WIDTH = 1024
JPEG_QUALITY = 85

# Phone photos are shown in portrait. A photo that is still landscape after its EXIF
# orientation is applied is turned 90° clockwise (how most phones store sideways shots).
FORCE_PORTRAIT = True

# --- Slide grid (4 portrait slots in one row) ---
SLOTS_PER_SLIDE = 4
PORTRAIT_RATIO = 4 / 3          # frame height / width (typical phone photo 3:4)
SIDE_MARGIN = Inches(0.6)       # left / right slide margin
SLOT_GAP = Inches(0.45)         # space between neighbouring photos
TITLE_TOP = Inches(0.3)
TITLE_HEIGHT = Inches(0.7)
PHOTO_TOP = Inches(1.15)        # photos start below the title
CAPTION_GAP = Inches(0.12)      # space between a photo and its caption
CAPTION_HEIGHT = Inches(1.15)
BOTTOM_MARGIN = Inches(0.3)
CAPTION_FONT_SIZE = Pt(10.5)
TITLE_FONT_SIZE = Pt(24)

MATERIALS = ("Sealant", "Concrete", "Paint", "Gasket", "With Film", "Broken Glass")


def optimize_image_in_memory(image_source, max_width=MAX_IMAGE_WIDTH, force_portrait=FORCE_PORTRAIT):
    """Resizes a raw image in memory using PIL `Image.thumbnail` (max width 1024px), turning
    it to portrait when needed. Returns (jpeg_stream, width_px, height_px); the stream is a
    seeked `io.BytesIO` ready for python-pptx. The original bytes/file are never modified."""
    if isinstance(image_source, (bytes, bytearray)):
        src_stream = io.BytesIO(image_source)
    elif hasattr(image_source, "read"):
        image_source.seek(0)
        src_stream = image_source
    else:
        src_stream = image_source  # file path

    with Image.open(src_stream) as img:
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")

        if force_portrait and img.width > img.height:
            img = img.transpose(Image.Transpose.ROTATE_270)  # 90° clockwise

        # Width capped at max_width; height follows the aspect ratio
        img.thumbnail((max_width, max_width * 4), Image.Resampling.LANCZOS)

        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=JPEG_QUALITY, optimize=True)
        buffer.seek(0)
        return buffer, img.width, img.height


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


# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------

def _select_slide_layout(prs):
    """Picks a template layout with a title and no picture/body placeholders ("Title Only"),
    falling back to a blank layout, then to the layout with the fewest placeholders."""
    layouts = list(prs.slide_layouts)
    if not layouts:
        raise RuntimeError("The PowerPoint template contains no slide layouts.")

    for layout in layouts:
        if layout.name.strip().lower() == "title only":
            return layout
    for layout in layouts:
        if layout.name.strip().lower() == "blank":
            return layout
    return min(layouts, key=lambda layout: len(layout.placeholders))


def _slot_geometry(slide_width, slide_height, photo_count):
    """Returns one (left, top, width, height) portrait frame per photo, plus the caption top.

    Frames are sized for SLOTS_PER_SLIDE photos in one row so every slide uses the same photo
    size; when a chunk has fewer photos the row is centred horizontally."""
    usable_width = slide_width - 2 * SIDE_MARGIN - (SLOTS_PER_SLIDE - 1) * SLOT_GAP
    slot_w = usable_width / SLOTS_PER_SLIDE

    usable_height = slide_height - PHOTO_TOP - CAPTION_GAP - CAPTION_HEIGHT - BOTTOM_MARGIN
    slot_h = slot_w * PORTRAIT_RATIO
    if slot_h > usable_height:  # short slides (e.g. 4:3) -> shrink to fit
        slot_h = usable_height
        slot_w = slot_h / PORTRAIT_RATIO

    row_width = photo_count * slot_w + (photo_count - 1) * SLOT_GAP
    start_left = (slide_width - row_width) / 2

    frames = []
    for i in range(photo_count):
        left = start_left + i * (slot_w + SLOT_GAP)
        frames.append((int(left), int(PHOTO_TOP), int(slot_w), int(slot_h)))
    caption_top = int(PHOTO_TOP + slot_h + CAPTION_GAP)
    return frames, caption_top


def _fit_inside(frame, img_w, img_h):
    """Largest (left, top, width, height) with the image's aspect ratio that fits inside the
    frame, centred, so photos are never cropped or stretched."""
    left, top, frame_w, frame_h = frame
    scale = min(frame_w / img_w, frame_h / img_h)
    pic_w, pic_h = int(img_w * scale), int(img_h * scale)
    return left + (frame_w - pic_w) // 2, top + (frame_h - pic_h) // 2, pic_w, pic_h


def _clear_unused_placeholders(slide, keep=()):
    """Removes empty placeholders (e.g. 'Click to add text') that the layout put on the slide."""
    for ph in list(slide.placeholders):
        if ph in keep:
            continue
        ph._element.getparent().remove(ph._element)


def _set_title(slide, prs, chunk_date):
    """Writes the day's date into the title placeholder, or into a textbox if there is none."""
    title_text = f"Date: {chunk_date}" if chunk_date else "Photo Report"
    title = slide.shapes.title
    if title is None:
        for ph in slide.placeholders:
            if ph.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                title = ph
                break

    if title is not None:
        title.text = title_text
        # Keep the title clear of the photo row (and inside the slide on widescreen decks)
        overlaps_photos = (title.top or 0) + (title.height or 0) > PHOTO_TOP
        too_wide = (title.left or 0) + (title.width or 0) > prs.slide_width
        if overlaps_photos or too_wide:
            title.left, title.top = SIDE_MARGIN, TITLE_TOP
            title.width, title.height = prs.slide_width - 2 * SIDE_MARGIN, TITLE_HEIGHT
            for paragraph in title.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT
                for run in paragraph.runs:
                    run.font.size = TITLE_FONT_SIZE
                    run.font.bold = True
        return title

    box = slide.shapes.add_textbox(SIDE_MARGIN, TITLE_TOP, prs.slide_width - 2 * SIDE_MARGIN, TITLE_HEIGHT)
    tf = box.text_frame
    tf.text = title_text
    tf.paragraphs[0].runs[0].font.size = TITLE_FONT_SIZE
    tf.paragraphs[0].runs[0].font.bold = True
    return None


def _add_caption(slide, left, top, width, date_str, drop_str, floor_str, defect_str):
    """Adds the Date / Drop / Floor / Defect caption under one photo."""
    box = slide.shapes.add_textbox(left, top, width, CAPTION_HEIGHT)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = Emu(0)

    lines = [
        ("Date: ", date_str),
        ("Drop: ", drop_str),
        ("Floor: ", floor_str),
        ("Defect: ", defect_str),
    ]
    for i, (label, value) in enumerate(lines):
        paragraph = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        paragraph.alignment = PP_ALIGN.LEFT
        label_run = paragraph.add_run()
        label_run.text = label
        label_run.font.bold = True
        label_run.font.size = CAPTION_FONT_SIZE
        value_run = paragraph.add_run()
        value_run.text = value
        value_run.font.size = CAPTION_FONT_SIZE


def _inject_chunk_into_slide(slide, prs, chunk_items, chunk_date=""):
    """Places up to 4 portrait photos side by side with their metadata captions."""
    title = _set_title(slide, prs, chunk_date)
    _clear_unused_placeholders(slide, keep=(title,) if title is not None else ())

    items = list(chunk_items)[:SLOTS_PER_SLIDE]
    frames, caption_top = _slot_geometry(prs.slide_width, prs.slide_height, len(items))

    for item, frame in zip(items, frames):
        date_str, drop_str, floor_str, defect_str = _extract_metadata(item, chunk_date)
        image_source = item.get("image_bytes") or item.get("image_path")

        if image_source is not None:
            img_stream, img_w, img_h = optimize_image_in_memory(image_source)
            pic_left, pic_top, pic_w, pic_h = _fit_inside(frame, img_w, img_h)
            slide.shapes.add_picture(img_stream, pic_left, pic_top, width=pic_w, height=pic_h)

        frame_left, _frame_top, frame_w, _frame_h = frame
        _add_caption(slide, frame_left, caption_top, frame_w, date_str, drop_str, floor_str, defect_str)


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

    layout = _select_slide_layout(prs)

    for chunk_entry in daily_chunks:
        if isinstance(chunk_entry, dict):
            chunk_date = str(chunk_entry.get("date") or "")
            chunk_photos = chunk_entry.get("photos") or []
        else:
            chunk_photos = list(chunk_entry)
            chunk_date = str(chunk_photos[0].get("date_display") or chunk_photos[0].get("date") or "") if chunk_photos else ""

        if not chunk_photos:
            continue

        slide = prs.slides.add_slide(layout)
        _inject_chunk_into_slide(slide, prs, chunk_photos, chunk_date=chunk_date)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(str(output_path))
    return str(output_path)


# Alias for external callers
generate_pptx_report = write_chunks_to_pptx
