from PIL import Image, ImageOps


def load_image(filepath, max_dimension=1600):
    """Loads and prepares an image for Gemini AI extraction.
    Auto-corrects EXIF orientation so rotated photos are right-side-up,
    and scales the API payload to max 1600px for lightning-fast uploads
    (PowerPoint still retains 100% original full-resolution camera bytes)."""
    img = Image.open(filepath)
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    if max_dimension and max(img.size) > max_dimension:
        img.thumbnail((max_dimension, max_dimension), Image.Resampling.LANCZOS)
    img.load()
    return img
