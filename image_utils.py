from io import BytesIO
from PIL import Image, ImageOps
from openpyxl.drawing.image import Image as ExcelImage

# Longest side (in pixels) of the photo copy embedded in Excel
MAX_EMBED_PIXELS = 1000


def _safe_exif_transpose(img):
    """Applies EXIF rotation if present. Some photos (often ones re-saved by
    Facebook/Messenger) carry malformed EXIF data that can make Pillow's EXIF
    reader raise an error. If that happens, we just use the image as-is
    instead of failing the whole photo."""
    try:
        return ImageOps.exif_transpose(img)
    except Exception:
        return img


def load_image(filepath):
    """Loads the image for the Gemini AI model (fully read, so the file is released)."""
    img = Image.open(filepath)
    img.load()
    # Phone photos often store rotation in EXIF; apply it so the AI sees the board upright
    return _safe_exif_transpose(img)


def prepare_excel_image(filepath, max_entries):
    """Embeds and scales the image for the Excel output.
    A downsized in-memory copy is embedded, so the master file stays small
    and does not depend on the temp photo still existing when Excel is saved."""
    with Image.open(filepath) as original:
        small = _safe_exif_transpose(original).convert("RGB")
        small.thumbnail((MAX_EMBED_PIXELS, MAX_EMBED_PIXELS))
        buffer = BytesIO()
        small.save(buffer, format="JPEG", quality=85)
    buffer.seek(0)

    excel_img = ExcelImage(buffer)
    excel_img.width = 140
    excel_img.height = max(140, max_entries * 25)
    return excel_img
