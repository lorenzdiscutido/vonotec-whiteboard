# extraction.py
import json

from image_utils import load_image
from gemini_client import get_raw_response, get_review_response, parse_and_clean_json, find_problems
from settings import EXTRACTION_RETRIES, ENABLE_REVIEW_PASS


def extract_data(filepath):
    """Reads one whiteboard photo with Gemini (no Streamlit calls here)."""
    img = load_image(filepath)

    # Pass 1: read the board (retry if the answer is not valid JSON)
    parsed_data = None
    for attempt in range(EXTRACTION_RETRIES + 1):
        try:
            raw_text = get_raw_response(img)
            parsed_data = parse_and_clean_json(raw_text)
            break
        except json.JSONDecodeError:
            if attempt == EXTRACTION_RETRIES:
                raise

    # Pass 2: a reviewer re-reads the board, only when a real data problem was found
    problems = find_problems(parsed_data)
    structural = [p for p in problems if p[2] == "structural"]
    if ENABLE_REVIEW_PASS and structural:
        try:
            raw_review = get_review_response(img, parsed_data, problems)
            parsed_data = parse_and_clean_json(raw_review)
        except Exception:
            pass  # keep the first-pass result if the review fails
        problems = find_problems(parsed_data)

    # "structural" and "manual" problems turn that material's rows red in Excel;
    # "info" problems are listed to the user but change no color.
    parsed_data["Review Notes"] = [msg for _material, msg, _kind in problems]
    parsed_data["Flagged Materials"] = sorted(
        {mat for mat, _msg, kind in problems if mat and kind in ("structural", "manual")}
    )
    return parsed_data
