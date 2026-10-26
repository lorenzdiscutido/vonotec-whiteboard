from PIL import Image


def load_image(filepath):
    """Loads the image for the Gemini AI model (fully read, so the file is released)."""
    img = Image.open(filepath)
    img.load()
    return img
