# gemini_client.py
import json
import re
import time
import random
import threading
import collections
import google.generativeai as genai
import streamlit as st
from config import JSON_KEYS, MATERIAL_RULES, REFERENCE_DATA

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
VALID_CODES = {row[0] for row in REFERENCE_DATA[1:]}
# Words/characters that show the AI was unsure about what it read
HEDGE_MARKERS = ("?", "UNCLEAR", "ILLEGIBLE", "UNSURE", "UNREADABLE")


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
model = genai.GenerativeModel('gemini-flash-lite-latest')


def _extraction_prompt():
    """The reading rules, shared by the first pass and the review pass."""
    return (
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
        "SEALANT 'DS' ONLY RULE: In the Sealant row, if a damage entry is just \"DS\" by itself, with no location modifier after it (such as 'C-C', 'F-C', 'C-F' or any other characters), do NOT output that entry at all. "
        "Leave out both that damage entry and its matching dimension so the 'Sealant Damage' and 'Sealant Dimension' arrays stay aligned. If nothing remains, return empty arrays []. "
        "Apply this rule only after fixing the spacing (for example, 'DSC-C' becomes 'DS C-C' and is kept). This rule applies ONLY to the Sealant row. Entries like \"DS C-C\" or \"DS F-C\" must still be output normally. "
        "In the concrete row, it is not 'CT' it is 'C+'. If you see 'CT' in the concrete row, replace it with 'C+'. "
        "Also in the concrete row, it is not 'DS', it is 'US' (Uneven Surface). If you see 'DS' in the concrete row, replace it with 'US'. "
        "CRITICAL RULE FOR DEFECT CODES (DAMAGE COLUMN): "
        "If you detect multiple known defect codes written closely together without spaces (e.g., 'C+C-', 'BPFP', 'DSMS'), you MUST insert a single space between them in your final JSON output (e.g., output 'C+ C-', 'BP FP', 'DS MS'). "
        "The valid defect codes to watch for are: CC-, C-, CC+, C+, BH, US, DP, FP, BP, DG, DS, MS, BG. "
        "DO NOT treat location modifiers like 'CC', 'C-C', or 'F-C' as separate damage entries. They must remain attached to the main defect code in the same string (e.g., output [\"DS CC\"], NEVER [\"DS\", \"CC\"])."
    )


def _generate_text(parts):
    """Calls Gemini (retrying when it is busy) and returns the reply text
    with any markdown code fence removed."""
    response = None
    for attempt in range(MAX_ATTEMPTS):
        _wait_for_slot()
        try:
            response = model.generate_content(parts)
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
    """Sends the image and prompt to Gemini and extracts the raw text block."""
    return _generate_text([_extraction_prompt(), img])


def get_review_response(img, parsed_data, problems):
    """Second pass: shows Gemini the photo again, plus the first reading and
    the problems the automatic checks found, and asks for a corrected JSON."""
    first_pass = {k: parsed_data.get(k) for k in JSON_KEYS}
    problem_lines = "\n".join(f"- {msg}" for _material, msg, _kind in problems)
    review_prompt = (
        _extraction_prompt()
        + "\n\nREVIEW TASK: A first reading of this same photo produced the JSON below, "
        "but automatic checks found possible problems with it.\n"
        f"FIRST READING: {json.dumps(first_pass, ensure_ascii=False)}\n"
        f"PROBLEMS FOUND:\n{problem_lines}\n"
        "Look at the photo again carefully and return the corrected JSON, using exactly the same keys "
        "and all of the rules above. Only change values that you can actually see on the board; "
        "fix anything the first reading misread, but keep entries the board genuinely shows even if "
        "they were flagged. Do not invent entries. Return ONLY raw JSON."
    )
    return _generate_text([review_prompt, img])


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

    return parsed_data


def find_problems(parsed_data):
    """Automatic checks on one extracted whiteboard.

    Returns a list of (material, message, kind) tuples.
      kind "structural": a real data problem. Triggers the review pass and
                         turns that material's rows red in Excel.
      kind "info":       shown to the user for awareness only.
    material is None when the problem is not tied to one material row.
    """
    problems = []
    any_entries = False
    seen_dimension_lists = {}

    for mat in MATERIALS:
        dmg = parsed_data.get(f"{mat} Damage", []) or []
        dim = parsed_data.get(f"{mat} Dimension", []) or []
        allowed = MATERIAL_RULES.get(mat, [])
        if dmg or dim:
            any_entries = True

        # Mismatched entries: every damage needs exactly one dimension
        if len(dmg) != len(dim):
            problems.append((
                mat,
                f"{mat}: {len(dmg)} damage entr{'y' if len(dmg) == 1 else 'ies'} "
                f"but {len(dim)} dimension(s)",
                "structural",
            ))
        else:
            for d, size in zip(dmg, dim):
                if str(d).strip() and not str(size).strip():
                    problems.append((mat, f"{mat}: '{d}' has no dimension", "structural"))

        # Wrong code: the code must belong to this material's row
        for d in dmg:
            tokens = str(d).split()
            if not tokens:
                continue
            codes_in_entry = [t for t in tokens if t in VALID_CODES]
            if tokens[0] not in allowed or any(c not in allowed for c in codes_in_entry):
                problems.append((mat, f"{mat}: code '{d}' does not belong in the {mat} row", "structural"))

        # Likely copy error: two materials with identical lists of dimensions
        if len(dim) >= 2:
            key = tuple(dim)
            if key in seen_dimension_lists:
                problems.append((
                    mat,
                    f"{mat}: dimensions are identical to {seen_dimension_lists[key]} (possible copy error)",
                    "structural",
                ))
            else:
                seen_dimension_lists[key] = mat

        # AI hedges: the model marked something as uncertain
        for text in list(dmg) + list(dim):
            if any(marker in str(text).upper() for marker in HEDGE_MARKERS):
                problems.append((mat, f"{mat}: the AI was unsure about '{text}'", "info"))

    # Nothing detected at all
    if not any_entries:
        problems.append((None, "No damage entries were detected on this whiteboard", "structural"))

    # Missing date / submitter: shown for awareness only
    if not str(parsed_data.get("Date", "")).strip():
        problems.append((None, "No date was found on the board", "info"))
    if not str(parsed_data.get("Submitter", "")).strip():
        problems.append((None, "No submitter name was found on the board", "info"))

    return problems
