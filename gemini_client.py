# gemini_client.py
import json
import re
import time
import random
import threading
import collections
import google.generativeai as genai
import streamlit as st
from config import JSON_KEYS, REFERENCE_DATA, MATERIAL_RULES

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

MAX_ATTEMPTS = 5

# Shared speed limit for ALL users. The free Gemini tier allows 15 requests
# per minute per model, so we stay a little under it.
REQUESTS_PER_MINUTE = 12
RATE_WINDOW_SECONDS = 60
_call_times = collections.deque()
_rate_lock = threading.Lock()

MATERIALS = ["Sealant", "Concrete", "Paint", "Gasket"]


def _wait_for_slot():
    """Blocks until another Gemini call is allowed under the per-minute limit."""
    while True:
        with _rate_lock:
            now = time.monotonic()
            while _call_times and now - _call_times[0] >= RATE_WINDOW_SECONDS:
                _call_times.popleft()
            if len(_call_times) < REQUESTS_PER_MINUTE:
                _call_times.append(now)
                return
            wait = RATE_WINDOW_SECONDS - (now - _call_times[0])
        time.sleep(max(wait, 0.5))


def _retry_wait_seconds(error, attempt):
    """Uses the delay Gemini asks for ("Please retry in 27.7s"), else backs off."""
    match = re.search(r"retry in ([\d.]+)s", str(error))
    if match:
        return float(match.group(1)) + 1
    return (2 ** attempt) + random.random()


# Initialize the Gemini Client (key comes from Streamlit Secrets)
genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

# Change this one line to try another model (e.g. a Flash model instead of Flash-Lite)
MODEL_NAME = 'gemini-flash-lite-latest'
model = genai.GenerativeModel(MODEL_NAME)


# ==========================================
# PROMPTS
# ==========================================
def _valid_codes_text():
    """Builds the 'valid codes per row' list from config.py so it never goes out of date."""
    code_names = {row[0]: row[1] for row in REFERENCE_DATA[1:]}
    lines = []
    for material, codes in MATERIAL_RULES.items():
        described = [f"{c} ({code_names[c]})" for c in codes if c in code_names]
        if described:
            lines.append(f"- {material}: " + ", ".join(described))
    return "\n".join(lines)


def _rules_text():
    """The reading rules shared by the first pass and the review pass."""
    all_codes = ", ".join(row[0] for row in REFERENCE_DATA[1:])
    return f"""
HOW TO READ THE BOARD (do this before answering)
1. Look at the whole photo first. Locate the header information (name, date, tower), the DROP and FLR columns, and the four material rows: Sealant, Concrete, Paint, Gasket. Each material row has a DAMAGE part and a DIMENSION part.
2. If the photo is tilted, rotated, glary, shadowed or slightly blurry, compensate mentally and keep reading.
3. Read row by row and cell by cell, zooming in on small, faint, smudged or crowded handwriting. Text in any marker color counts. Never stop at the first entry: one row can hold several entries written on separate lines or side by side. Do not skip anything.
4. Match every dimension to the damage entry it belongs to (same line or position on the board).
5. Return an empty array only when a row is truly blank. A faint or hard-to-read entry is NOT blank: read it as best you can and mention it in "Uncertain".

FIELD RULES (use exactly these keys: {JSON_KEYS}, plus the extra key "Uncertain")
1. 'Submitter': Extract ONLY the person's name from the bottom right corner (e.g., 'Rufino Zuñiga Jr'). Strip away any prefixes like 'R.A.T.S. drop by'.
2. 'Date': Scan the entire photo to locate and extract the date, regardless of its position. Ensure the year is correct (e.g., 'August 1, 2026').
3. 'Elevation': Copy exactly what is written for the elevation on the board. Empty string if it is not written.
4. 'Tower': This is the tower code usually written in blue marker (e.g., 'TOWER 1'). Do not include any name here. Explicitly ignore the words 'RAT' or 'RATS'. If a valid tower number is not clearly written, return an empty string so the system can infer it later.
5. 'Drop': The specific number located strictly under the 'DROP' column header.
6. 'Floor': The specific number located strictly under the 'FLR' column header.
7. Materials (Sealant, Concrete, Paint, Gasket): each material gets two keys, '[Material] Damage' and '[Material] Dimension'. Both MUST be JSON ARRAYS of strings, with one damage entry per dimension entry in the same order. If a damage entry has no dimension written, use an empty string "" in that position so the arrays stay aligned.
8. 'Uncertain': an array of short strings (at most 5) describing anything you could not read with confidence, for example "Concrete entry 2: could be C+ or CC+". Use an empty array [] when everything is clear.

VALID DEFECT CODES PER ROW (use this to sanity-check every entry you read)
{_valid_codes_text()}
All known defect codes: {all_codes}.
If a code you read is not valid for its row, look at the handwriting again: it is probably a misread. If, after careful re-reading, the writing truly does not match a valid code, output exactly what is written and add a note to "Uncertain".

FORMATTING RULES FOR MATERIALS
- FORCE UPPERCASE: convert all text to uppercase (e.g., 'ds' becomes 'DS').
- UNITS: 'Sealant Dimension' uses the unit 'CM'. 'Concrete Dimension', 'Paint Dimension' and 'Gasket Dimension' use the unit 'CM²' (squared). Format like "340 CM" (number, space, unit).
- After the word "DS" there must be a space, then the next characters. If there is no space after "DS", add one (e.g., 'DSC-C' becomes 'DS C-C').
- In the Concrete row, it is not 'CT', it is 'C+'. Replace 'CT' with 'C+' in the Concrete row.
- In the Concrete row, it is not 'DS', it is 'US' (Uneven Surface). Replace 'DS' with 'US' in the Concrete row.
- If several known defect codes are written closely together without spaces (e.g., 'C+C-', 'BPFP', 'DSMS'), insert a single space between them (e.g., 'C+ C-', 'BP FP', 'DS MS').
- Location modifiers like 'CC', 'C-C', 'F-C' or 'C-F' are NOT separate damage entries. They stay attached to the main defect code in the same string (e.g., ["DS CC"], NEVER ["DS", "CC"]).
- SEALANT 'DS' ONLY RULE: in the Sealant row, if an entry is just "DS" by itself with no location modifier after it (no 'C-C', 'F-C', 'C-F' or any other characters), do NOT output that entry. Leave out both that damage entry and its matching dimension so the arrays stay aligned. Apply this only after fixing the spacing (so 'DSC-C' becomes 'DS C-C' and is kept), and only to the Sealant row. Entries like "DS C-C" or "DS F-C" are output normally.
- Format illustration only (not real data): "Sealant Damage": ["DS C-C", "DS F-C"], "Sealant Dimension": ["340 CM", "120 CM"]
"""


def build_extraction_prompt():
    return (
        "You are an expert data-entry analyst reading a photo of a handwritten construction-defect "
        "inspection whiteboard. Transcribe it into ONE flat JSON object with exact fidelity to what is "
        "written. Accuracy matters more than completeness: never guess, and never invent values that are "
        "not on the board.\n"
        + _rules_text()
        + "\nSILENT SELF-CHECK BEFORE YOU ANSWER\n"
        "a) Re-read each material row one more time. Did you miss any entry?\n"
        "b) For each material, do the Damage and Dimension arrays have the same length and order?\n"
        "c) Is every code valid for its row? Fix misreads (CT to C+, DS to US in Concrete).\n"
        "d) Uppercase, spacing and units correct?\n"
        "e) Did you list every doubtful reading in \"Uncertain\"?\n\n"
        "OUTPUT: return ONLY the raw JSON object. No markdown, no explanations."
    )


def build_review_prompt(first_pass, problems):
    problem_text = "\n".join(f"- {p}" for p in problems) if problems else "- (none)"
    return (
        "You are a meticulous QA reviewer. You are given a photo of a handwritten construction-defect "
        "inspection whiteboard and a FIRST-PASS JSON extraction made by another analyst. The first pass may "
        "contain mistakes: misread handwriting, missed entries, invented entries, entries in the wrong row, "
        "wrong units, or damage and dimension entries that are not matched up.\n\n"
        "AUTOMATIC CHECKS FLAGGED THESE POSSIBLE PROBLEMS:\n"
        + problem_text
        + "\n\nYOUR TASK\n"
        "1. Read the photo yourself first, row by row, using the rules below.\n"
        "2. Compare your reading with the first-pass JSON. For every difference, look at the photo again and "
        "decide which reading is right.\n"
        "3. Fix only what the photo clearly supports: add missed entries, remove entries that are not on the "
        "board, correct misread codes, numbers, units, date or name, and re-align damage with dimension.\n"
        "4. Keep first-pass values that are correct exactly as they are. Do not change something just to be "
        "different. If the photo is genuinely unclear, keep the first-pass value and describe the doubt in "
        "\"Uncertain\".\n"
        "5. Return the COMPLETE corrected JSON object with the same keys.\n"
        + _rules_text()
        + "\nFIRST-PASS JSON:\n"
        + json.dumps(first_pass, ensure_ascii=False)
        + "\n\nOUTPUT: return ONLY the raw corrected JSON object. No markdown, no explanations."
    )


# ==========================================
# GEMINI CALLS
# ==========================================
def _generate(prompt_text, img):
    """Sends a prompt + image to Gemini with speed limiting and retries, returns clean JSON text."""
    response = None
    for attempt in range(MAX_ATTEMPTS):
        _wait_for_slot()
        try:
            response = model.generate_content(
                [prompt_text, img],
                generation_config={"response_mime_type": "application/json"},
            )
            break
        except RETRYABLE_ERRORS as e:
            if attempt == MAX_ATTEMPTS - 1:
                raise
            time.sleep(_retry_wait_seconds(e, attempt))

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


def get_raw_response(img):
    """First pass: read the whiteboard."""
    return _generate(build_extraction_prompt(), img)


def get_review_response(img, first_pass, problems):
    """Second pass: a reviewer re-reads the board and corrects the first pass."""
    return _generate(build_review_prompt(first_pass, problems), img)


# ==========================================
# CLEANING & CHECKING
# ==========================================
def parse_and_clean_json(raw_text):
    """Converts the raw text into JSON and enforces array formatting."""
    parsed_data = json.loads(raw_text)

    # Enforce array formatting and capitalization for all material columns
    for material in MATERIALS:
        for suffix in ["Damage", "Dimension"]:
            key = f"{material} {suffix}"
            # If Gemini returned a single string by accident, convert it to a list
            if isinstance(parsed_data.get(key), str):
                parsed_data[key] = [parsed_data[key].upper()] if parsed_data[key] else []
            elif isinstance(parsed_data.get(key), list):
                parsed_data[key] = [str(item).upper() for item in parsed_data[key]]
            else:
                parsed_data[key] = []

    # "Uncertain" must always be a list of strings
    uncertain = parsed_data.get("Uncertain", [])
    if isinstance(uncertain, list):
        parsed_data["Uncertain"] = [str(x) for x in uncertain if str(x).strip()]
    elif uncertain:
        parsed_data["Uncertain"] = [str(uncertain)]
    else:
        parsed_data["Uncertain"] = []

    return parsed_data


def find_problems(parsed_data):
    """Automatic checks that decide whether a photo needs a second look."""
    problems = []

    if not any(parsed_data.get(f"{m} Damage") for m in MATERIALS):
        problems.append("No defects were detected on the whole board.")

    for material in MATERIALS:
        damages = parsed_data.get(f"{material} Damage", [])
        dimensions = parsed_data.get(f"{material} Dimension", [])
        if len(damages) != len(dimensions):
            problems.append(
                f"{material}: {len(damages)} damage entries but {len(dimensions)} dimension entries."
            )
        allowed = MATERIAL_RULES.get(material, [])
        for entry in damages:
            tokens = str(entry).split()
            if tokens and tokens[0] not in allowed:
                problems.append(f"{material}: code '{tokens[0]}' is not valid for this row.")

    if not str(parsed_data.get("Date", "")).strip():
        problems.append("Date was not found.")
    if not str(parsed_data.get("Submitter", "")).strip():
        problems.append("Submitter was not found.")

    for note in parsed_data.get("Uncertain", []):
        problems.append(f"The AI was unsure: {note}")

    return problems
