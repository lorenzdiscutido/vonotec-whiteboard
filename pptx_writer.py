# pptx_writer.py
"""Generates a PowerPoint (.pptx) report from chronologically sorted daily photo chunks.

Loads `template.pptx` (or creates a 16:9 presentation), resizes high-resolution photos
in memory using PIL (`Image.thumbnail`, max width 1024px) so original files are untouched
and the presentation stays lightweight.

Each slide features an architectural/engineering drawing sheet title block:
- Left-most cell: VONOTEC Logo
- Right-most cell:
    CREATED BY: MARJORIE BELOY
    CHECKED BY: ANNA CAOILE
    APPROVED BY:
- Cell 2: Date
- Upper content area: Up to 6 portrait photos arranged side by side with even spacing
  and metadata captions (Date, Drop, Floor, Defect) underneath each photo.
"""
import io
import os
from pathlib import Path

from PIL import Image, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

TEMPLATE_PATH = Path(__file__).with_name("template.pptx")
# Set to None to preserve 100% full native camera resolution without downscaling
MAX_IMAGE_WIDTH = None
JPEG_QUALITY = 95

# Phone photos are shown in portrait. A photo that is still landscape after its EXIF
# orientation is applied is turned 90° clockwise (how most phones store sideways shots).
FORCE_PORTRAIT = True

# --- Slide & Sheet Title Block Geometry ---
SLOTS_PER_SLIDE = 6
PORTRAIT_RATIO = 4 / 3          # frame height / width (typical phone photo 3:4)

BORDER_MARGIN_X = Inches(0.25)
BORDER_MARGIN_Y = Inches(0.20)
BORDER_WIDTH = Inches(12.833)
BORDER_HEIGHT = Inches(7.10)
LINE_COLOR = RGBColor(0, 0, 0)
LINE_WIDTH = Pt(1.5)

TB_TOP = Inches(6.25)
TB_HEIGHT = BORDER_MARGIN_Y + BORDER_HEIGHT - TB_TOP

# Vertical divider X-positions
DIVIDER_X_POSITIONS = (Inches(3.34), Inches(7.49), Inches(9.02), Inches(11.23))

TITLE_BLOCK_CREATED_BY = "CREATED BY: MARJORIE BELOY"
TITLE_BLOCK_CHECKED_BY = "CHECKED BY: ANNA CAOILE"
TITLE_BLOCK_APPROVED_BY = "APPROVED BY:"

# Upper Photo Grid
CONTENT_LEFT = BORDER_MARGIN_X + Inches(0.35)
CONTENT_WIDTH = BORDER_WIDTH - Inches(0.70)
PHOTO_TOP = Inches(0.55)
SLOT_GAP = Inches(0.26)
CAPTION_GAP = Inches(0.12)
CAPTION_HEIGHT = Inches(1.35)
CAPTION_FONT_SIZE = Pt(9.5)

MATERIALS = ("Sealant", "Concrete", "Paint", "Gasket", "With Film", "Broken Glass")


def optimize_image_in_memory(image_source, max_width=MAX_IMAGE_WIDTH, quality=JPEG_QUALITY, force_portrait=FORCE_PORTRAIT):
    """Prepares a raw image in memory for insertion into PowerPoint while preserving its full
    original resolution and crisp visual fidelity. Applies EXIF orientation and turns sideways
    photos to portrait when needed. Returns (jpeg_stream, width_px, height_px); the stream is a
    seeked `io.BytesIO` ready for python-pptx. The original file bytes are never modified."""
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

        # Only downscale if max_width is explicitly set (None retains 100% original camera resolution)
        if max_width is not None and max_width > 0:
            img.thumbnail((max_width, max_width * 4), Image.Resampling.LANCZOS)

        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=quality, subsampling=0)
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


def _get_logo_path():
    """Finds the best available Vonotec banner logo asset."""
    base_dir = Path(__file__).parent
    candidates = [
        base_dir / "vonotec_crisp_logo.png",
        base_dir / "vonotec_slide_logo.png",
        base_dir / "vonotec.png",
    ]
    for path in candidates:
        if path.exists():
            return str(path)
    return None


def _draw_sheet_title_block(slide, chunk_date="", title_block_info=None):
    """Draws the architectural sheet border and bottom title block on `slide`:
    - Outer border rectangle
    - Title block horizontal line
    - 4 vertical dividers creating 5 cells
    - Field 1 (Left-most): VONOTEC logo
    - Field 2: Custom "Title" (if provided; no indicator label)
    - Field 3: Custom "Service offered" (if provided; no indicator label)
    - Field 4: Custom location / client logo (if uploaded)
    - Field 5 (Right-most): Created by, Checked by, Approved by
    """
    info = title_block_info or {}
    title_text = str(info.get("title") or "").strip()
    service_text = str(info.get("service_offered") or "").strip()
    loc_logo_bytes = info.get("location_logo_bytes")

    # 1. Outer Border
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, BORDER_MARGIN_X, BORDER_MARGIN_Y, BORDER_WIDTH, BORDER_HEIGHT)
    rect.fill.background()
    rect.line.color.rgb = LINE_COLOR
    rect.line.width = LINE_WIDTH

    # 2. Horizontal divider for the bottom title block
    h_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, BORDER_MARGIN_X, TB_TOP, BORDER_WIDTH, Pt(1.5))
    h_line.fill.solid()
    h_line.fill.fore_color.rgb = LINE_COLOR
    h_line.line.fill.background()

    # 3. Vertical divider lines (creating 5 cells)
    for dx in DIVIDER_X_POSITIONS:
        v_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, dx, TB_TOP, Pt(1.5), TB_HEIGHT)
        v_line.fill.solid()
        v_line.fill.fore_color.rgb = LINE_COLOR
        v_line.line.fill.background()

    # 4. Field 1 (Left-most, x: 0.25" to 3.34"): VONOTEC Logo
    logo_path = _get_logo_path()
    if logo_path:
        slide.shapes.add_picture(logo_path, Inches(0.55), Inches(6.38), width=Inches(2.5), height=Inches(0.75))

    # 5. Field 2 (x: 3.34" to 7.49"): User-specified "Title" (empty if none; no label printed)
    if title_text:
        c2_box = slide.shapes.add_textbox(Inches(3.38), Inches(6.28), Inches(4.07), TB_HEIGHT - Inches(0.06))
        c2_tf = c2_box.text_frame
        c2_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        c2_tf.word_wrap = True
        c2_tf.margin_left = c2_tf.margin_right = Inches(0.1)
        c2_p = c2_tf.paragraphs[0]
        c2_p.alignment = PP_ALIGN.CENTER
        c2_run = c2_p.add_run()
        c2_run.text = title_text
        c2_run.font.bold = True
        c2_run.font.size = Pt(13 if len(title_text) < 40 else 11)
        c2_run.font.color.rgb = RGBColor(15, 23, 42)

    # 6. Field 3 (x: 7.49" to 9.02"): User-specified "Service offered" (empty if none; no label printed)
    if service_text:
        c3_box = slide.shapes.add_textbox(Inches(7.53), Inches(6.28), Inches(1.45), TB_HEIGHT - Inches(0.06))
        c3_tf = c3_box.text_frame
        c3_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        c3_tf.word_wrap = True
        c3_tf.margin_left = c3_tf.margin_right = Inches(0.08)
        c3_p = c3_tf.paragraphs[0]
        c3_p.alignment = PP_ALIGN.CENTER
        c3_run = c3_p.add_run()
        c3_run.text = service_text
        c3_run.font.bold = True
        c3_run.font.size = Pt(9.5 if len(service_text) < 30 else 8.5)
        c3_run.font.color.rgb = RGBColor(15, 23, 42)

    # 7. Field 4 (x: 9.02" to 11.23"): Location logo (empty if none uploaded)
    if loc_logo_bytes is not None:
        try:
            f4_left, f4_top, f4_w, f4_h = Inches(9.08), Inches(6.30), Inches(2.10), TB_HEIGHT - Inches(0.10)
            im = Image.open(io.BytesIO(loc_logo_bytes))
            im = ImageOps.exif_transpose(im)
            scale = min(f4_w / im.width, f4_h / im.height)
            pw, ph = int(im.width * scale), int(im.height * scale)
            px = f4_left + (f4_w - pw) // 2
            py = f4_top + (f4_h - ph) // 2
            slide.shapes.add_picture(io.BytesIO(loc_logo_bytes), px, py, width=pw, height=ph)
        except Exception:
            pass

    # 8. Field 5 (Right-most, x: 11.23" to 13.08"): Sign-off block
    cell5_box = slide.shapes.add_textbox(Inches(11.28), Inches(6.28), Inches(1.75), TB_HEIGHT - Inches(0.08))
    c5_tf = cell5_box.text_frame
    c5_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    c5_tf.word_wrap = True
    c5_tf.margin_left = c5_tf.margin_right = c5_tf.margin_top = c5_tf.margin_bottom = Emu(0)

    lines_c5 = [
        TITLE_BLOCK_CREATED_BY,
        TITLE_BLOCK_CHECKED_BY,
        TITLE_BLOCK_APPROVED_BY,
    ]
    for idx, line_text in enumerate(lines_c5):
        p = c5_tf.paragraphs[0] if idx == 0 else c5_tf.add_paragraph()
        r = p.add_run()
        r.text = line_text
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(30, 30, 30)


def _slot_geometry(photo_count):
    """Calculates coordinates for up to SLOTS_PER_SLIDE portrait photo slots inside the sheet content area.
    When a chunk has fewer than 6 photos, the row is centered horizontally."""
    slot_w = (CONTENT_WIDTH - (SLOTS_PER_SLIDE - 1) * SLOT_GAP) / SLOTS_PER_SLIDE
    slot_h = slot_w * PORTRAIT_RATIO

    row_width = photo_count * slot_w + (photo_count - 1) * SLOT_GAP
    start_left = CONTENT_LEFT + (CONTENT_WIDTH - row_width) / 2

    frames = []
    for i in range(photo_count):
        left = start_left + i * (slot_w + SLOT_GAP)
        frames.append((int(left), int(PHOTO_TOP), int(slot_w), int(slot_h)))
    caption_top = int(PHOTO_TOP + slot_h + CAPTION_GAP)
    return frames, caption_top


def _fit_inside(frame, img_w, img_h):
    """Centers and scales the image inside the target frame while preserving its aspect ratio."""
    left, top, frame_w, frame_h = frame
    scale = min(frame_w / img_w, frame_h / img_h)
    pic_w, pic_h = int(img_w * scale), int(img_h * scale)
    return left + (frame_w - pic_w) // 2, top + (frame_h - pic_h) // 2, pic_w, pic_h


def _clear_unused_placeholders(slide):
    """Removes leftover template text placeholders so they don't cover drawing content."""
    for ph in list(slide.placeholders):
        try:
            ph._element.getparent().remove(ph._element)
        except Exception:
            pass


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


def _inject_chunk_into_slide(slide, chunk_items, chunk_date="", title_block_info=None):
    """Draws the title block and places up to 6 portrait photos side by side with their metadata captions."""
    _clear_unused_placeholders(slide)
    _draw_sheet_title_block(slide, chunk_date=chunk_date, title_block_info=title_block_info)

    items = list(chunk_items)[:SLOTS_PER_SLIDE]
    frames, caption_top = _slot_geometry(len(items))

    for item, frame in zip(items, frames):
        date_str, drop_str, floor_str, defect_str = _extract_metadata(item, chunk_date)
        image_source = item.get("image_bytes") or item.get("image_path")

        if image_source is not None:
            img_stream, img_w, img_h = optimize_image_in_memory(image_source)
            pic_left, pic_top, pic_w, pic_h = _fit_inside(frame, img_w, img_h)
            slide.shapes.add_picture(img_stream, pic_left, pic_top, width=pic_w, height=pic_h)

        frame_left, _frame_top, frame_w, _frame_h = frame
        _add_caption(slide, frame_left, caption_top, frame_w, date_str, drop_str, floor_str, defect_str)


def _select_blank_layout(prs):
    """Selects the cleanest layout available (blank or fewest placeholders)."""
    layouts = list(prs.slide_layouts)
    if not layouts:
        raise RuntimeError("The PowerPoint template contains no slide layouts.")
    for layout in layouts:
        if layout.name.strip().lower() == "blank":
            return layout
    return min(layouts, key=lambda layout: len(layout.placeholders))


def write_chunks_to_pptx(daily_chunks, output_path, template_path=TEMPLATE_PATH, title_block_info=None):
    """Builds a PowerPoint presentation from `daily_chunks` and saves it to `output_path`.

    Parameters
    ----------
    daily_chunks : list[dict] or list[list[dict]]
        Chronologically ordered chunks of up to 6 photo records per day.
        Each item may be a dict `{"date": "...", "photos": [...]}` or a raw list of photo dicts.
    output_path : str or Path
        Destination `.pptx` file path.
    template_path : str or Path
        Path to `template.pptx` (loaded as the base presentation when present).
    title_block_info : dict, optional
        Metadata for slide title block: "title" (Field 2), "service_offered" (Field 3),
        "location_logo_bytes" (Field 4).
    """
    resolved_template = Path(template_path) if template_path else TEMPLATE_PATH
    if resolved_template.exists():
        prs = Presentation(str(resolved_template))
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

    # Clean out any initial dummy slides in template so only real batch slides are exported
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    layout = _select_blank_layout(prs)

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
        _inject_chunk_into_slide(slide, chunk_photos, chunk_date=chunk_date, title_block_info=title_block_info)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(str(output_path))
    return str(output_path)


# Alias for external callers
generate_pptx_report = write_chunks_to_pptx
