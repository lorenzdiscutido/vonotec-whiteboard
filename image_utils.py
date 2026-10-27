import io
from PIL import Image, ImageOps


def load_image(image_source, max_dimension=1600):
    """Loads and prepares an image for Gemini AI extraction from raw bytes,
    a BytesIO stream, or a file path.
    Auto-corrects EXIF orientation so rotated photos are right-side-up,
    and scales the API payload to max 1600px for fast uploads
    (PowerPoint still retains 100% original full-resolution camera bytes)."""
    if isinstance(image_source, (bytes, bytearray)):
        stream = io.BytesIO(image_source)
    elif hasattr(image_source, "read"):
        image_source.seek(0)
        stream = image_source
    else:
        stream = image_source

    img = Image.open(stream)
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    if max_dimension and max(img.size) > max_dimension:
        img.thumbnail((max_dimension, max_dimension), Image.Resampling.LANCZOS)
    img.load()
    return img
