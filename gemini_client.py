# gemini_client.py
import json
import time
import random
import google.generativeai as genai
import streamlit as st
from config import JSON_KEYS

# Errors from Gemini that are temporary (busy / rate limited) and worth retrying
try:
    from google.api_core import exceptions as google_exceptions
    RETRYABLE_ERRORS = (
        google_exceptions.ResourceExhausted,
        google_exceptions.TooManyRequests,
        google_exceptions.ServiceUnavailable,
        google_exceptions.DeadlineExceeded,
        google_exceptions.InternalServerError,
    )
except ImportError:
    RETRYABLE_ERRORS = ()

MAX_ATTEMPTS = 4

# Initialize the Gemini Client (key comes from Streamlit Secrets)
genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
model = genai.GenerativeModel('gemini-flash-lite-latest')

def get_raw_response(img):
    """Sends the image and prompt to Gemini and extracts the raw text block."""
    prompt_text = (
        f"Extract the data from this whiteboard grid into a flat JSON object using exactly these keys: {JSON_KEYS}. "
        "Follow these strict extraction rules based on this specific whiteboard layout: "
        "1. 'Submitter': Extract ONLY the person's name from the bottom right corner (e.g., 'Rufino Zuñiga Jr'). Strip away any prefixes like 'R.A.T.S. drop by'. "
        "2. 'Date': Scan the entire photo to locate and extract the date, regardless of its position. Ensure the year is correct (e.g., 'August 1, 2026'). "
        "3. 'Tower': This is the tower code usually written in blue marker (e.g., 'TOWER 1'). Do not include any name here. Explicitly ignore the words 'RAT' or 'RATS'. If a valid tower number is not clearly written, return an empty string so the system can infer it later. "
        "4. 'Drop': This is the specific number located strictly under the 'DROP' column header. "
        "5. 'Floor': This is the specific number located strictly under the 'FLR' column header. "
        "6. Materials (Sealant, Concrete, Paint, Gasket): You MUST extract multiple damage entries as JSON ARRAYS. "
        "- Each material gets two keys: '[Material] Damage' and '[Material] Dimension'. Both must be formatted as arrays of strings. "
        "- Example exact JSON output for Sealant:\n"
        "\"Sealant Damage\": [\"DS C-C\", \"DS F-C\"],\n"
        "\"Sealant Dimension\": [\"340 CM\", \"120 CM\"]\n"
        "CRITICAL FORMATTING FOR RULE 6: "
        "- FORCE UPPERCASE: Convert all text to uppercase (e.g., 'ds' to 'DS'). "
        "- TARGETED UNITS: For 'Sealant Dimension', output the unit simply as 'CM'. For 'Concrete Dimension', 'Paint Dimension', and 'Gasket Dimension', output the unit as 'CM²' (squared). "
        "If a field is empty on the board, return an empty array [] for materials, or an empty string \"\" for static fields. Return ONLY raw JSON. "
        "After the word \"DS\" there should be a space, then the next characters. If there is no space after \"DS\", add one. "
        "In the concrete row, it is not 'CT' it is 'C+'. If you see 'CT' in the concrete row, replace it with 'C+'. "
        "Also in the concrete row, it is not 'DS', it is 'US' (Uneven Surface). If you see 'DS' in the concrete row, replace it with 'US'. "
        "CRITICAL RULE FOR DEFECT CODES (DAMAGE COLUMN): "
        "If you detect multiple known defect codes written closely together without spaces (e.g., 'C+C-', 'BPFP', 'DSMS'), you MUST insert a single space between them in your final JSON output (e.g., output 'C+ C-', 'BP FP', 'DS MS'). "
        "The valid defect codes to watch for are: CC-, C-, CC+, C+, BH, US, DP, FP, BP, DG, DS, MS, BG. "
        "DO NOT treat location modifiers like 'CC', 'C-C', or 'F-C' as separate damage entries. They must remain attached to the main defect code in the same string (e.g., output [\"DS CC\"], NEVER [\"DS\", \"CC\"])."
    )

    # Retry a few times if Gemini is busy or rate-limiting (likely with several users)
    response = None
    for attempt in range(MAX_ATTEMPTS):
        try:
            response = model.generate_content([prompt_text, img])
            break
        except RETRYABLE_ERRORS:
            if attempt == MAX_ATTEMPTS - 1:
                raise
            time.sleep((2 ** attempt) + random.random())

    try:
        raw_text = response.text.strip()
    except ValueError:
        raise RuntimeError(
            "Gemini returned no text for this image (it may have been blocked). Please try again."
        )
    
    # Strip markdown block formatting if present
    if raw_text.startswith("```json"):
        raw_text = raw_text[7:-3].strip()
    elif raw_text.startswith("```"):
        raw_text = raw_text[3:-3].strip()
        
    return raw_text

def parse_and_clean_json(raw_text):
    """Converts the raw text into JSON and enforces array formatting."""
    parsed_data = json.loads(raw_text)
    
    # Enforce array formatting and capitalization for all material columns
    for material in ["Sealant", "Concrete", "Paint", "Gasket"]:
        for suffix in ["Damage", "Dimension"]:
            key = f"{material} {suffix}"
            # If Gemini returned a single string by accident, convert it to a list
            if isinstance(parsed_data.get(key), str):
                parsed_data[key] = [parsed_data[key].upper()] if parsed_data[key] else []
            elif isinstance(parsed_data.get(key), list):
                parsed_data[key] = [str(item).upper() for item in parsed_data[key]]
            else:
                parsed_data[key] = []
                
    return parsed_data
