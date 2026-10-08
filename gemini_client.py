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

# ==========================================================================
# DEFECTS, LOCATION MODIFIERS AND UNITS: edit ONLY this block when they change.
# The Gemini prompt below is built from it.
# ==========================================================================
# Location modifiers that follow a SEALANT defect code (e.g. "DS CC", "MS FG")
SEALANT_LOCATION_MODIFIERS = ["CC", "CF", "GG", "FG", "FF"]

# (material, defect codes, unit of the dimension written next to those codes)
DEFECT_UNITS = [
    ("Sealant",  ["DS", "MS"],                 "CM"),
    ("Concrete", ["US", "BH"],                 "CM²"),
    ("Concrete", ["CC-", "CC+", "C-", "C+"],   "CM"),
    ("Paint",    ["DP", "FP", "BP"],           "CM²"),
    ("Gasket",   ["DG"],                       "CM"),
]
# Defect codes that exist on the reference sheet but are not read per material
OTHER_DEFECT_CODES = ["BG"]
# ==========================================================================


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


def _defect_reference_text():
    """Turns the DEFECT_UNITS / SEALANT_LOCATION_MODIFIERS tables into prompt text."""
    lines = []
    for material, codes, unit in DEFECT_UNITS:
        lines.append(f"- {material} defect codes {', '.join(codes)}: dimension unit is {unit}.")
    modifiers = ", ".join(SEALANT_LOCATION_MODIFIERS)
    all_codes = []
    for _material, codes, _unit in DEFECT_UNITS:
        all_codes += [c for c in codes if c not in all_codes]
    all_codes += [c for c in OTHER_DEFECT_CODES if c not in all_codes]
    return (
        "DEFECT CODES, LOCATION MODIFIERS AND UNITS (this list is the current, official one): "
        + " ".join(lines)
        + f" The ONLY valid sealant location modifiers are: {modifiers}. They follow the sealant defect code in the same string "
        "(e.g., 'DS CC', 'MS FG'). Note that 'CC-' and 'CC+' (with a minus or plus sign) are CONCRETE crack defect codes, "
        "while 'CC' with no sign written after a sealant code is a location modifier. "
        f"The valid defect codes to watch for are: {', '.join(all_codes)}. "
    )


def _extraction_prompt():
    """The reading rules, shared by the first pass and the review pass."""
    modifiers_quoted = ", ".join(f"'{m}'" for m in SEALANT_LOCATION_MODIFIERS)
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
        "\"Sealant Damage\": [\"DS CC\", \"DS FG\"],\n"
        "\"Sealant Dimension\": [\"340 CM\", \"120 CM\"]\n"
        + _defect_reference_text()
        + "CRITICAL FORMATTING FOR RULE 6: "
        "- FORCE UPPERCASE: Convert all text to uppercase (e.g., 'ds' to 'DS'). "
        "- UNITS DEPEND ON THE DEFECT CODE: write each dimension with the unit that belongs to the defect code in the SAME position of the damage array, "
        "exactly as listed above (for example Concrete 'US' and 'BH' use 'CM²', but Concrete 'CC-', 'CC+', 'C-', 'C+' use 'CM'; Paint uses 'CM²'; Gasket and Sealant use 'CM'). "
        "Use the unit with the superscript ² for square units (CM²). "
        "If a field is empty on the board, return an empty array [] for materials, or an empty string \"\" for static fields. Return ONLY raw JSON. "
        "After the word \"DS\" there should be a space, then the next characters. If there is no space after \"DS\", add one. "
        f"SEALANT 'DS' ONLY RULE: In the Sealant row, if a damage entry is just \"DS\" by itself, with no location modifier after it (one of {modifiers_quoted}, or any other characters), do NOT output that entry at all. "
        "Leave out both that damage entry and its matching dimension so the 'Sealant Damage' and 'Sealant Dimension' arrays stay aligned. If nothing remains, return empty arrays []. "
        "Apply this rule only after fixing the spacing (for example, 'DSCC' becomes 'DS CC' and is kept). This rule applies ONLY to the Sealant row. Entries like \"DS CC\" or \"DS FG\" must still be output normally. "
        "In the concrete row, it is not 'CT' it is 'C+'. If you see 'CT' in the concrete row, replace it with 'C+'. "
        "Also in the concrete row, it is not 'DS', it is 'US' (Uneven Surface). If you see 'DS' in the concrete row, replace it with 'US'. "
        "CRITICAL RULE FOR DEFECT CODES (DAMAGE COLUMN): "
        "If you detect multiple known defect codes written closely together without spaces (e.g., 'C+C-', 'BPFP', 'DSMS'), you MUST insert a single space between them in your final JSON output (e.g., output 'C+ C-', 'BP FP', 'DS MS'). "
        f"DO NOT treat the sealant location modifiers ({modifiers_quoted}) as separate damage entries. They must remain attached to the main defect code in the same string (e.g., output [\"DS CC\"], NEVER [\"DS\", \"CC\"]). "
        "SHARED DIMENSION RULE: Every dimension you output must be physically written on the board for that defect. "
        "NEVER copy or repeat one dimension onto other defects. If a material row lists several defect codes but only ONE dimension is written, "
        "output every defect code in the Damage array, put the dimension on the FIRST defect only, and use an empty string \"\" for the other defects, "
        "so both arrays have the same length. Example: board shows 'US BH C-' and '200 CM' -> "
        "\"Concrete Damage\": [\"US\", \"BH\", \"C-\"], \"Concrete Dimension\": [\"200 CM\", \"\", \"\"]."
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


# (material, defect code) -> the unit its dimension must be written in
_UNIT_TABLE = {
    (material, code): unit
    for material, codes, unit in DEFECT_UNITS
    for code in codes
}
# A trailing centimetre unit in any common spelling: CM, CM2, CM^2, CM², SQ CM, SQ.CM, SQUARE CM
_UNIT_SUFFIX = re.compile(r"\s*(?:(?:SQ\.?|SQUARE)\s*)?CM(?:\s*(?:\^\s*2|²|2))?\s*$")
# A dimension that is only numbers (and x / * between them), with no unit at all
_NUMBER_ONLY = re.compile(r"^[\d.,\s xX×*/+-]+$")


def _apply_unit(dimension, unit):
    """Rewrites one dimension so it ends in `unit`. Leaves it untouched when it
    is empty, has no number, or uses a unit we don't recognise (e.g. meters)."""
    text = str(dimension).strip()
    if not text:
        return dimension
    match = _UNIT_SUFFIX.search(text)
    if match:
        number_part = text[:match.start()].strip()
        return f"{number_part} {unit}" if number_part else dimension
    if _NUMBER_ONLY.match(text):
        return f"{text} {unit}"
    return dimension


def _fix_dimension_units(parsed_data):
    """Makes every dimension use the unit that belongs to its defect code
    (from DEFECT_UNITS), no matter what unit Gemini wrote."""
    for material in MATERIALS:
        damages = parsed_data[f"{material} Damage"]
        dimensions = parsed_data[f"{material} Dimension"]
        for i in range(min(len(damages), len(dimensions))):
            units = {
                _UNIT_TABLE[(material, token)]
                for token in str(damages[i]).split()
                if (material, token) in _UNIT_TABLE
            }
            # Only when the entry has exactly one kind of unit: an entry with
            # mixed codes (e.g. US and CC-) is ambiguous, so it is left alone.
            if len(units) == 1:
                dimensions[i] = _apply_unit(dimensions[i], units.pop())


_HAS_DIGIT = re.compile(r"\d")


def _group_defect_codes(entry):
    """Splits one damage string that holds several defect codes into one string
    per defect, keeping location modifiers with their own code:
        'C+ US'          -> ['C+', 'US']
        'DS CC MS FG'    -> ['DS CC', 'MS FG']
        'DS CC'          -> ['DS CC']   (only one defect code, so nothing to split)"""
    entry = str(entry)
    tokens = entry.split()
    if sum(1 for t in tokens if t in VALID_CODES) < 2:
        return [entry]
    groups, leading = [], []
    for token in tokens:
        if token in VALID_CODES:
            groups.append(leading + [token])
            leading = []
        elif groups:
            groups[-1].append(token)
        else:
            leading.append(token)
    return [" ".join(g) for g in groups]


def _normalize_dimensions(parsed_data):
    """Handles several defects that share a single dimension on the board.
    
    Rule: the dimension belongs to the FIRST defect only; every other defect is
    left with a blank dimension and flagged for a human to fill in.
    """
    notes = []
    for material in MATERIALS:
        damages = parsed_data[f"{material} Damage"]
        dimensions = parsed_data[f"{material} Dimension"]
        new_damages, new_dimensions = [], []

        for i, entry in enumerate(damages):
            size = dimensions[i] if i < len(dimensions) else None  # None = dimensions ran out
            groups = _group_defect_codes(entry)

            if len(groups) > 1:
                # Several defects written together: only the first gets the dimension
                first_index = len(new_damages)
                for g, group in enumerate(groups):
                    new_damages.append(group)
                    new_dimensions.append((size or "") if g == 0 else "")
                blanks = list(range(first_index + 1, first_index + len(groups)))
                
                # Check if it's missing a number or entirely blank
                if size is None or size.strip() == "" or (size.strip() and not _HAS_DIGIT.search(size)):
                    new_dimensions[first_index] = ""
                    blanks = [first_index] + blanks
                notes.append({
                    "material": material,
                    "indices": blanks,
                    "message": (
                        f"{material}: {' and '.join(groups)} were written with only one dimension. "
                        f"It was applied to '{groups[0]}' only; the others were left blank "
                        "(yellow cell) for manual entry."
                    ),
                })
            else:
                no_number = size is not None and size.strip() != "" and not _HAS_DIGIT.search(size)
                
                # Trigger manual note if size is None, completely empty string (""), or has no number
                if size is None or size.strip() == "" or no_number:
                    notes.append({
                        "material": material,
                        "indices": [len(new_damages)],
                        "message": (
                            f"{material}: '{entry}' has no dimension of its own (one dimension was "
                            "shared by several defects). Left blank (yellow cell) for manual entry."
                        ),
                    })
                    new_damages.append(entry)
                    new_dimensions.append("")
                else:
                    new_damages.append(entry)
                    new_dimensions.append(size)

        new_dimensions.extend(dimensions[len(damages):])

        parsed_data[f"{material} Damage"] = new_damages
        parsed_data[f"{material} Dimension"] = new_dimensions

    parsed_data["Manual Dimension"] = notes

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

    # Several defects sharing one dimension: first defect keeps it, the rest are
    # left blank and flagged for manual entry
    _normalize_dimensions(parsed_data)

    # Force each dimension's unit to match its defect code (CM vs CM²)
    _fix_dimension_units(parsed_data)

    return parsed_data


def find_problems(parsed_data):
    """Automatic checks on one extracted whiteboard.

    Returns a list of (material, message, kind) tuples.
      kind "structural": a real data problem. Triggers the review pass and
                         turns that material's rows red in Excel.
      kind "manual":     needs a human (e.g. a dimension left blank on purpose).
                         Turns the rows red but does not trigger the review pass.
      kind "info":       shown to the user for awareness only.
    material is None when the problem is not tied to one material row.
    """
    problems = []
    any_entries = False
    seen_dimension_lists = {}

    # Blank dimensions that the "one dimension for several defects" rule already
    # handled: these get their own "manual" flag instead of the generic check
    manual_notes = parsed_data.get("Manual Dimension", []) or []
    manual_indices = {}
    for note in manual_notes:
        manual_indices.setdefault(note["material"], set()).update(note["indices"])

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
            for idx, (d, size) in enumerate(zip(dmg, dim)):
                if idx in manual_indices.get(mat, ()):
                    continue
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

    # Blank dimensions left on purpose for a human to fill in. Kind "manual"
    # highlights the rows red like "structural" does, but does NOT trigger the
    # extra AI review pass (the rule already decided what to do).
    for note in manual_notes:
        problems.append((note["material"], note["message"], "manual"))

    # Nothing detected at all
    if not any_entries:
        problems.append((None, "No damage entries were detected on this whiteboard", "structural"))

    # Missing date / submitter: shown for awareness only
    if not str(parsed_data.get("Date", "")).strip():
        problems.append((None, "No date was found on the board", "info"))
    if not str(parsed_data.get("Submitter", "")).strip():
        problems.append((None, "No submitter name was found on the board", "info"))

    return problems
