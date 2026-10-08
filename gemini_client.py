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

REQUESTS_PER_MINUTE = 12
RATE_WINDOW_SECONDS = 60
_call_times = collections.deque()
_rate_lock = threading.Lock()

MATERIALS = ["Sealant", "Concrete", "Paint", "Gasket", "With Film"]
VALID_CODES = {row[0] for row in REFERENCE_DATA[1:]}

HEDGE_MARKERS = ("?", "UNCLEAR", "ILLEGIBLE", "UNSURE", "UNREADABLE")

# ==========================================================================
# DEFECTS, LOCATION MODIFIERS AND UNITS
# ==========================================================================

SEALANT_LOCATION_MODIFIERS = ["CC", "CF", "GG", "FG", "FF"]

DEFECT_UNITS = [
    ("Sealant",   ["DS", "MS"],                 "CM"),
    ("Concrete",  ["US", "BH"],                 "CM²"),
    ("Concrete",  ["CC-", "CC+", "C-", "C+"],   "CM"),
    ("Paint",     ["DP", "FP", "BP"],           "CM²"),
    ("Paint",     ["NP"],                       ""),
    ("Gasket",    ["DG"],                       "CM"),
    ("With Film", ["GLASS-FRAME", "FRAME"],     "CM"),
]

OTHER_DEFECT_CODES = ["BG"]

# ==========================================================================

def _wait_for_slot():
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
    match = re.search(r"retry in ([\d.]+)s", str(error))
    if match:
        return float(match.group(1)) + 1
    return (2 ** attempt) + random.random()

genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
model = genai.GenerativeModel('gemini-flash-lite-latest')

def _defect_reference_text():
    lines = []
    for material, codes, unit in DEFECT_UNITS:
        if unit:
            lines.append(f"- {material} defect codes {', '.join(codes)}: dimension unit is {unit}.")
    modifiers = ", ".join(SEALANT_LOCATION_MODIFIERS)
    all_codes = []
    for _material, codes, _unit in DEFECT_UNITS:
        all_codes += [c for c in codes if c not in all_codes]
    all_codes += [c for c in OTHER_DEFECT_CODES if c not in all_codes]
    
    return (
        "DEFECT CODES, LOCATION MODIFIERS AND UNITS (this list is the current, official one): "
        + " ".join(lines) + 
        f" The ONLY valid sealant location modifiers are: {modifiers}. They follow the sealant defect code in the same string "
        "(e.g., 'DS CC', 'MS FG'). Note that 'CC-' and 'CC+' (with a minus or plus sign) are CONCRETE crack defect codes, "
        "while 'CC' with no sign written after a sealant code is a location modifier. "
        f"The valid defect codes to watch for are: {', '.join(all_codes)}. "
    )

def _extraction_prompt():
    modifiers_quoted = ", ".join(f"'{m}'" for m in SEALANT_LOCATION_MODIFIERS)
    return (
        f"Extract the data from this whiteboard grid into a flat JSON object using exactly these keys: {JSON_KEYS}. "
        "Follow these strict extraction rules based on this specific whiteboard layout: "
        "1. 'Submitter': Extract ONLY the person's name from the bottom right corner (e.g., 'Rufino Zuñiga Jr'). Strip away any prefixes like 'R.A.T.S. drop by'. "
        "2. 'Date': Scan the entire photo to locate and extract the date, regardless of its position. Ensure the year is correct (e.g., 'August 1, 2026'). "
        "3. 'Tower': This is the tower code usually written in blue marker (e.g., 'TOWER 1'). Do not include any name here. Explicitly ignore the words 'RAT' or 'RATS'. If a valid tower number is not clearly written, return an empty string so the system can infer it later. "
        "4. 'Drop': This is the specific number located strictly under the 'DROP' column header. "
        "5. 'Floor': This is the specific number located strictly under the 'FLR' column header. "
        "6. Materials (Sealant, Concrete, Paint, Gasket, With Film): You MUST extract multiple damage entries as JSON ARRAYS. "
        "- Each material gets two keys: '[Material] Damage' and '[Material] Dimension'. Both must be formatted as arrays of strings. "
        "- Example exact JSON output for Sealant:\n"
        "\"Sealant Damage\": [\"DS CC\", \"MS FG\"],\n"
        "\"Sealant Dimension\": [\"340 CM\", \"120 CM\"]\n"
        + _defect_reference_text()
        + "CRITICAL FORMATTING FOR RULE 6: "
        "- FORCE UPPERCASE: Convert all text to uppercase (e.g., 'ds' to 'DS'). "
        "- UNITS DEPEND ON THE DEFECT CODE: write each dimension with the unit that belongs to the defect code in the SAME position of the damage array, "
        "exactly as listed above (for example Concrete 'US' and 'BH' use 'CM²', but Concrete 'CC-', 'CC+', 'C-', 'C+' use 'CM'; Paint uses 'CM²'; Gasket and Sealant use 'CM'). "
        "Use the unit with the superscript ² for square units (CM²). "
        "If a field is empty on the board, return an empty array [] for materials, or an empty string \"\" for static fields. Return ONLY raw JSON. "
        "After the word \"DS\" or \"MS\" there should be a space, then the next characters. If there is no space, add one. "
        "HYPHENATED MODIFIER RULE: If a sealant location modifier contains a hyphen (e.g., 'C-C', 'c-f', 'F-G'), you MUST remove the hyphen and output the standard two-letter code (e.g., 'CC', 'CF', 'FG'). "
        "WITH FILM RULE: If the board has a row for 'WITH FILM' and the worker wrote 'Glass Frame', you MUST extract it as 'GLASS-FRAME' (with a hyphen) so it functions as a single code. If they wrote 'Frame', extract it as 'FRAME'. "
        "In the concrete row, it is not 'CT' it is 'C+'. If you see 'CT' in the concrete row, replace it with 'C+'. "
        "Also in the concrete row, it is not 'DS', it is 'US' (Uneven Surface). If you see 'DS' in the concrete row, replace it with 'US'. "
        "NO PAINT (NP) RULE: The defect code 'NP' under Paint has no dimension. Whenever you extract 'NP', you MUST output an empty string \"\" for its matching dimension to keep the 'Paint Damage' and 'Paint Dimension' arrays perfectly aligned. "
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
        raise RuntimeError("Gemini returned no text for this image (it may have been blocked). Please try again.")

    if raw_text.startswith("```json"):
        raw_text = raw_text[7:-3].strip()
    elif raw_text.startswith("```"):
        raw_text = raw_text[3:-3].strip()

    return raw_text

def get_raw_response(img):
    return _generate_text([_extraction_prompt(), img])

def get_review_response(img, parsed_data, problems):
    first_pass = {k: parsed_data.get(k) for k in JSON_KEYS}
    problem_lines = "\n".join(f"- {msg}" for _material, msg, _kind in problems)
    review_prompt = (
        _extraction_prompt() +
        "\n\nREVIEW TASK: A first reading of this same photo produced the JSON below, "
        "but automatic checks found possible problems with it.\n"
        f"FIRST READING: {json.dumps(first_pass, ensure_ascii=False)}\n"
        f"PROBLEMS FOUND:\n{problem_lines}\n"
        "Look at the photo again carefully and return the corrected JSON, using exactly the same keys "
        "and all of the rules above. Only change values that you can actually see on the board; "
        "fix anything the first reading misread, but keep entries the board genuinely shows even if "
        "they were flagged. Do not invent entries. Return ONLY raw JSON."
    )
    return _generate_text([review_prompt, img])

_UNIT_TABLE = {
    (material, code): unit
    for material, codes, unit in DEFECT_UNITS
    for code in codes
}

_UNIT_SUFFIX = re.compile(r"\s*(?:(?:SQ.?|SQUARE)\s*)?CM(?:\s*(?:^\s *2|²|2))?\s*$")
_NUMBER_ONLY = re.compile(r"^[\d.,\s xX×*/+-]+$")

def _apply_unit(dimension, unit):
    text = str(dimension).strip()
    if not text or not unit:
        return dimension

    match = _UNIT_SUFFIX.search(text)
    if match:
        number_part = text[:match.start()].strip()
        return f"{number_part} {unit}" if number_part else dimension

    if _NUMBER_ONLY.match(text):
        return f"{text} {unit}"

    return dimension

def _fix_dimension_units(parsed_data):
    for material in MATERIALS:
        damages = parsed_data[f"{material} Damage"]
        dimensions = parsed_data[f"{material} Dimension"]
        
        for i in range(min(len(damages), len(dimensions))):
            units = {
                _UNIT_TABLE[(material, token)]
                for token in str(damages[i]).split()
                if (material, token) in _UNIT_TABLE
            }
            if len(units) == 1:
                dimensions[i] = _apply_unit(dimensions[i], units.pop())

_HAS_DIGIT = re.compile(r"\d")

def _group_defect_codes(entry):
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
    notes = []
    for material in MATERIALS:
        damages = parsed_data[f"{material} Damage"]
        dimensions = parsed_data[f"{material} Dimension"]
        new_damages, new_dimensions = [], []

        for i, entry in enumerate(damages):
            size = dimensions[i] if i < len(dimensions) else None
            groups = _group_defect_codes(entry)

            if len(groups) > 1:
                first_index = len(new_damages)
                for g, group in enumerate(groups):
                    new_damages.append(group)
                    new_dimensions.append((size or "") if g == 0 else "")
                
                blanks = [
                    first_index + idx 
                    for idx, group in enumerate(groups) 
                    if idx > 0 and group != "NP"
                ]
                
                if size is None or size.strip() == "" or (size.strip() and not _HAS_DIGIT.search(size)):
                    if groups[0] != "NP":
                        new_dimensions[first_index] = ""
                        blanks = [first_index] + blanks

                if blanks:
                    notes.append({
                        "material": material,
                        "indices": blanks,
                        "message": (
                            f"{material}: {' and '.join(groups)} were written with only one dimension. "
                            f"It was applied to '{groups[0]}' only; the others were left blank "
                            "(manual entry required)."
                        ),
                    })
            else:
                no_number = size is not None and size.strip() != "" and not _HAS_DIGIT.search(size)
                
                if entry == "NP":
                    new_damages.append(entry)
                    new_dimensions.append("")
                elif size is None or size.strip() == "" or no_number:
                    notes.append({
                        "material": material,
                        "indices": [len(new_damages)],
                        "message": (
                            f"{material}: '{entry}' has no dimension of its own (one dimension was "
                            "shared by several defects). Left blank (manual entry required)."
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
    parsed_data = json.loads(raw_text)

    for material in MATERIALS:
        for suffix in ["Damage", "Dimension"]:
            key = f"{material} {suffix}"
            if isinstance(parsed_data.get(key), str):
                parsed_data[key] = [parsed_data[key].upper()] if parsed_data[key] else []
            elif isinstance(parsed_data.get(key), list):
                parsed_data[key] = [str(item).upper() for item in parsed_data[key]]
            else:
                parsed_data[key] = []

    _normalize_dimensions(parsed_data)
    _fix_dimension_units(parsed_data)

    return parsed_data

def find_problems(parsed_data):
    problems = []
    any_entries = False
    seen_dimension_lists = {}

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
                if str(d).strip() == "NP":
                    continue
                if str(d).strip() and not str(size).strip():
                    problems.append((mat, f"{mat}: '{d}' has no dimension", "structural"))

        for d in dmg:
            tokens = str(d).split()
            if not tokens:
                continue
            codes_in_entry = [t for t in tokens if t in VALID_CODES]
            if tokens[0] not in allowed or any(c not in allowed for c in codes_in_entry):
                problems.append((mat, f"{mat}: code '{d}' does not belong in the {mat} row", "structural"))

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

        for text in list(dmg) + list(dim):
            if any(marker in str(text).upper() for marker in HEDGE_MARKERS):
                problems.append((mat, f"{mat}: the AI was unsure about '{text}'", "info"))

    for note in manual_notes:
        problems.append((note["material"], note["message"], "manual"))

    if not any_entries:
        problems.append((None, "No damage entries were detected on this whiteboard", "structural"))

    if not str(parsed_data.get("Date", "")).strip():
        problems.append((None, "No date was found on the board", "info"))
    if not str(parsed_data.get("Submitter", "")).strip():
        problems.append((None, "No submitter name was found on the board", "info"))

    return problems
