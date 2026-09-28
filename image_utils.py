from io import BytesIO
from PIL import Image
from openpyxl.drawing.image import Image as ExcelImage

# Longest side (in pixels) of the photo copy embedded in Excel
MAX_EMBED_PIXELS = 1000

def load_image(filepath):
    """Loads the image for the Gemini AI model (fully read, so the file is released)."""
    img = Image.open(filepath)
    img.load()
    return img

def prepare_excel_image(filepath, max_entries):
    """Embeds and scales the image for the Excel output.
    A downsized in-memory copy is embedded, so the master file stays small
    and does not depend on the temp photo still existing when Excel is saved."""
    with Image.open(filepath) as original:
        small = original.convert("RGB")
        small.thumbnail((MAX_EMBED_PIXELS, MAX_EMBED_PIXELS))
        buffer = BytesIO()
        small.save(buffer, format="JPEG", quality=85)
    buffer.seek(0)

    excel_img = ExcelImage(buffer)
    excel_img.width = 140
    excel_img.height = max(140, max_entries * 25)
    return excel_img
