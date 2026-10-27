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
from defect_rules import DEFECT_UNITS

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

MATERIALS = ["Sealant", "Concrete", "Paint", "Gasket", "With Film", "Broken Glass"]
SEALANT_LOCATION_MODIFIERS = ["CC", "CF", "GG", "FG", "FF", "FC", "GF"]

# Defect codes that are written on the board WITHOUT a dimension (NP now uses CM² -> SQ.M. like other Paint defects)
NO_DIMENSION_CODES = {"BG", "GLASS-FRAME", "FRAME"}

# Exclude modifiers and multi-word aliases from VALID_CODES token splitting
VALID_CODES = {row[0] for row in REFERENCE_DATA[1:]} - set(SEALANT_LOCATION_MODIFIERS) - {"NO PAINT"}

HEDGE_MARKERS = ("?", "UNCLEAR", "ILLEGIBLE", "UNSURE", "UNREADABLE")

OTHER_DEFECT_CODES = ["BG"]


def _retry_wait_seconds(error, attempt):
    match = re.search(r"retry in ([\d.]+)s", str(error))
    if match:
        return float(match.group(1)) + 1
    return (2 ** attempt) + random.random()


class KeySlot:
    """Manages an individual Gemini API key, its dedicated GenerativeModel client,
    and its independent rate-limiting queue."""

    def __init__(self, key: str, index: int):
        self.key = key
        self.index = index
        self.masked = f"...{key[-6:]}" if len(key) >= 6 else f"Key #{index + 1}"
        self.model = None
        self.is_exhausted = False
        self.exhausted_reason = ""
        self.call_times = collections.deque()
        self.lock = threading.Lock()
        self.total_calls = 0
        self.last_call_time = 0.0
        self.cooldown_until = 0.0

    def get_model(self):
        if self.model is None:
            from google.generativeai import client as genai_client
            genai.configure(api_key=self.key)
            m = genai.GenerativeModel("gemini-flash-lite-latest")
            m._client = genai_client.get_default_generative_client()
            self.model = m
        return self.model

    def wait_for_slot(self, max_rpm=REQUESTS_PER_MINUTE, window_sec=RATE_WINDOW_SECONDS):
        """Enforces both sliding-window RPM limit and smooth minimum-interval pacing per key."""
        min_interval = window_sec / max(max_rpm, 1)  # e.g. 60 / 12 = 5.0 seconds per key
        while True:
            with self.lock:
                now = time.monotonic()
                if now < self.cooldown_until:
                    wait = self.cooldown_until - now
                else:
                    while self.call_times and now - self.call_times[0] >= window_sec:
                        self.call_times.popleft()
                    if len(self.call_times) < max_rpm:
                        elapsed = now - self.last_call_time
                        if elapsed < min_interval:
                            wait = min_interval - elapsed
                        else:
                            self.call_times.append(now)
                            self.last_call_time = now
                            return
                    else:
                        wait = window_sec - (now - self.call_times[0])
            time.sleep(max(wait, 0.1))

    def set_cooldown(self, seconds: float):
        """Puts this key slot on temporary cooldown after hitting a rate-limit error."""
        with self.lock:
            self.cooldown_until = max(self.cooldown_until, time.monotonic() + seconds)

    def record_success(self):
        with self.lock:
            self.total_calls += 1

    def mark_exhausted(self, reason):
        with self.lock:
            self.is_exhausted = True
            self.exhausted_reason = str(reason)


class KeyPool:
    """Thread-safe pool of Gemini API keys supporting round-robin rotation,
    independent per-key rate limiting, automatic failover, and quota exhaustion handling."""

    def __init__(self):
        self._slots = []
        self._current_idx = -1
        self._lock = threading.Lock()
        self._initialized = False

    def _discover_keys(self):
        import os
        sources = []
        # 1. Streamlit secrets
        try:
            if hasattr(st, "secrets"):
                if "GEMINI_API_KEYS" in st.secrets:
                    sources.append(st.secrets["GEMINI_API_KEYS"])
                if "GEMINI_API_KEY" in st.secrets:
                    sources.append(st.secrets["GEMINI_API_KEY"])
                for k, v in st.secrets.items():
                    if k.startswith("GEMINI_API_KEY_") and isinstance(v, str):
                        sources.append(v)
        except Exception:
            pass

        # 2. Environment variables
        for var in ("GEMINI_API_KEYS", "GEMINI_API_KEY"):
            val = os.environ.get(var)
            if val:
                sources.append(val)
        for k, val in os.environ.items():
            if k.startswith("GEMINI_API_KEY_") and val:
                sources.append(val)

        # 3. config.py fallback
        try:
            import config
            cfg_keys = getattr(config, "GEMINI_API_KEYS", None)
            if cfg_keys:
                sources.append(cfg_keys)
            cfg_key = getattr(config, "GEMINI_API_KEY", None)
            if cfg_key:
                sources.append(cfg_key)
        except Exception:
            pass

        clean_keys = []
        seen = set()
        for item in sources:
            if isinstance(item, (list, tuple)):
                raw_items = item
            elif isinstance(item, str):
                raw_items = re.split(r"[\n,;]+", item)
            else:
                continue
            for r in raw_items:
                if isinstance(r, str):
                    clean = r.strip().strip("'\"")
                    if clean and clean not in seen:
                        seen.add(clean)
                        clean_keys.append(clean)
        return clean_keys

    def initialize(self):
        with self._lock:
            if self._initialized and self._slots:
                return
            keys = self._discover_keys()
            self._slots = [KeySlot(k, i) for i, k in enumerate(keys)]
            self._current_idx = -1
            self._initialized = True

    def reset_status(self):
        """Clears exhausted flags (e.g. for a new day)."""
        with self._lock:
            for s in self._slots:
                s.is_exhausted = False
                s.exhausted_reason = ""

    def active_slots(self):
        self.initialize()
        return [s for s in self._slots if not s.is_exhausted]

    def all_slots(self):
        self.initialize()
        return list(self._slots)

    def pool_stats(self):
        """Returns dict with active_count, total_count, approx_daily_quota, details."""
        self.initialize()
        active = [s for s in self._slots if not s.is_exhausted]
        total = self._slots
        return {
            "active_count": len(active),
            "total_count": len(total),
            "daily_quota": len(active) * 500,
            "slots": [
                {
                    "index": s.index + 1,
                    "masked": s.masked,
                    "active": not s.is_exhausted,
                    "calls": s.total_calls,
                    "reason": s.exhausted_reason,
                }
                for s in total
            ],
        }

    def next_slot(self):
        self.initialize()
        with self._lock:
            active = [s for s in self._slots if not s.is_exhausted]
            if not active:
                if not self._slots:
                    raise RuntimeError(
                        "No Gemini API key found. Please set GEMINI_API_KEY or GEMINI_API_KEYS "
                        "in Streamlit secrets or environment."
                    )
                raise RuntimeError(
                    f"All configured Gemini API keys ({len(self._slots)} key(s)) have exhausted "
                    f"their daily quota. Please check Google AI Studio or add additional keys."
                )
            now = time.monotonic()
            ready_slots = [s for s in active if now >= s.cooldown_until]
            candidates = ready_slots if ready_slots else active
            self._current_idx = (self._current_idx + 1) % len(candidates)
            return candidates[self._current_idx]


key_pool = KeyPool()


def get_key_pool_stats():
    """Returns status information about the active Gemini API key pool."""
    return key_pool.pool_stats()


def _get_model():
    """Retrieves an active model instance from the key pool."""
    return key_pool.next_slot().get_model()


class _LazyModelProxy:
    def __getattr__(self, name):
        m = _get_model()
        return getattr(m, name)


model = _LazyModelProxy()


def _defect_reference_text():
    lines = []
    for material, codes, unit in DEFECT_UNITS:
        visible_codes = [c for c in codes if c != "NO PAINT"]
        if unit and visible_codes:
            lines.append(f"- {material} defect codes {', '.join(visible_codes)}: dimension unit is {unit}.")
    modifiers = ", ".join(SEALANT_LOCATION_MODIFIERS)
    all_codes = []
    for _material, codes, _unit in DEFECT_UNITS:
        all_codes += [c for c in codes if c not in all_codes and c != "NO PAINT"]
    all_codes += [c for c in OTHER_DEFECT_CODES if c not in all_codes]

    return (
        "DEFECT CODES, LOCATION MODIFIERS AND UNITS (this list is the current, official one): "
        + " ".join(lines)
        + f" The ONLY valid sealant location modifiers are: {modifiers}. They follow the sealant defect code in the same string "
        "(e.g., 'DS CC', 'MS FG'). Note that 'CC-' and 'CC+' (with a minus or plus sign) are CONCRETE crack defect codes, "
        "while 'CC' with no sign written after a sealant code is a location modifier. "
        f"The valid defect codes to watch for are: {', '.join(all_codes)}. "
    )


SORTER_JSON_KEYS = ["Date", "Drop", "Floor", "Tower", "Submitter"]


def _extraction_prompt():
    return (
        "Extract inspection photo metadata from this image into a flat JSON object with exactly these keys: "
        f"{SORTER_JSON_KEYS}. "
        "The photo may be a construction whiteboard OR an on-site inspection photo with a camera watermark/timestamp overlay (such as Timemark). "
        "Strict extraction rules: "
        "1. 'Date': Scan the photo to locate the date (from whiteboard or camera timestamp). "
        "CRITICAL DATE FORMATTING RULES: "
        "- DAY/MONTH/YEAR ORDER: When dates are written numerically with slashes, dots, or dashes (e.g., '01/07/2026' or '04/07/2004'), "
        "the format is strictly DAY/MONTH/YEAR (DD/MM/YYYY). For example, '01/07/2026' is 1st of July (July 1, 2026), NOT January 7. '04/07/2026' is 4th of July (July 4, 2026), NOT April 7. "
        "Convert numeric dates into standard 'Month Day, Year' format (e.g., 'July 1, 2026' or 'July 4, 2026') or output as 'DD/MM/YYYY'. "
        "- TEXTUAL MONTHS: If the board or watermark explicitly writes out the month name (e.g., 'June 5, 2026' or '5 June 2026'), format as 'Month Day, Year'. "
        "- YEAR: Ensure the year is standardized to 2026 (e.g., convert '04/07/2004' or '04/07/24' to 'July 4, 2026'). "
        "- IGNORE WEEKDAYS AND TIME: Do not include weekdays (e.g. 'Wed', 'Wednesday') or times (e.g. '14:28') in the Date field. "
        "2. 'Drop': Extract the drop number/code (e.g. '6' from 'DROP' or 'Drop 6'). "
        "3. 'Floor': Extract the floor number/code (e.g. '36D' or '36' from 'FLR' or notes like '36D'). "
        "4. 'Tower': Extract the tower code if clearly present (e.g., 'TOWER 1'), else \"\". Ignore words like 'RAT' or 'RATS'. "
        "5. 'Submitter': Extract the person name if written (e.g., from bottom right corner of whiteboard or 'Name:' in timestamp), else \"\". "
        "If a field is not found in the photo, return an empty string \"\" for that key. Return ONLY valid raw JSON."
    )


def _generate_text(parts):
    key_pool.initialize()
    active_keys = key_pool.active_slots()
    if not active_keys:
        key_pool.next_slot()  # raises descriptive error if no keys configured

    max_attempts = max(MAX_ATTEMPTS, len(active_keys) * 2)
    response = None
    last_error = None

    for attempt in range(max_attempts):
        active = key_pool.active_slots()
        if not active:
            raise RuntimeError(
                f"All configured Gemini API keys ({len(key_pool.all_slots())} key(s)) have exhausted "
                f"their daily quota. Last error: {last_error}"
            )

        slot = key_pool.next_slot()
        slot.wait_for_slot(REQUESTS_PER_MINUTE, RATE_WINDOW_SECONDS)

        try:
            model_instance = slot.get_model()
            response = model_instance.generate_content(parts)
            slot.record_success()
            break
        except RETRYABLE_ERRORS as e:
            last_error = e
            err_msg = str(e).lower()

            # Check if this error is daily quota exhaustion vs temporary per-minute throttle
            is_daily_exhausted = (
                "quota exceeded" in err_msg
                or "limit per day" in err_msg
                or "quota metric" in err_msg
            )

            if is_daily_exhausted:
                slot.mark_exhausted(e)
                print(f"[Gemini Key Pool] Key #{slot.index + 1} ({slot.masked}) daily quota reached. Switching to next key...")
                continue

            # Temporary per-minute throttle / 429
            retry_after = _retry_wait_seconds(e, attempt)
            slot.set_cooldown(retry_after)

            # If other active keys are ready without cooldown, rotate immediately
            now = time.monotonic()
            other_ready = [s for s in key_pool.active_slots() if s.index != slot.index and now >= s.cooldown_until]
            if other_ready:
                continue

            # Single key or all keys throttling: do standard backoff
            if attempt == max_attempts - 1:
                raise
            time.sleep(retry_after)
        except Exception as e:
            err_msg = str(e).lower()
            if "429" in err_msg or "resource_exhausted" in err_msg:
                if "quota exceeded" in err_msg or "limit per day" in err_msg:
                    slot.mark_exhausted(e)
                    print(f"[Gemini Key Pool] Key #{slot.index + 1} ({slot.masked}) quota exhausted. Switching to next key...")
                    continue
            raise

    if response is None:
        raise RuntimeError(f"Failed to generate text after {max_attempts} attempts. Last error: {last_error}")

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


_UNIT_TABLE = {
    (material, code): unit
    for material, codes, unit in DEFECT_UNITS
    for code in codes
}

_UNIT_SUFFIX = re.compile(r"\s*(?:(?:SQ\.?|SQUARE)\s*)?CM(?:\s*(?:\^\s*2|\u00b2|2))?\s*$")
_NUMBER_ONLY = re.compile(r"^[\d.,\s xX\u00d7*/+-]+$")


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
    entry = re.sub(r"\bNO[\s-]+PAINT\b", "NP", str(entry).upper())
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
            entry = re.sub(r"\bNO[\s-]+PAINT\b", "NP", str(entry).upper()).strip()
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
                    if idx > 0 and group not in NO_DIMENSION_CODES
                ]

                if size is None or size.strip() == "" or (size.strip() and not _HAS_DIGIT.search(size)):
                    if groups[0] not in NO_DIMENSION_CODES:
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

                if entry in NO_DIMENSION_CODES:
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
                if str(d).strip() in NO_DIMENSION_CODES:
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
            non_empty_dims = [x for x in dim if str(x).strip()]
            if len(non_empty_dims) >= 2:
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

    if not str(parsed_data.get("Date", "")).strip():
        problems.append((None, "No date was found on the board", "info"))
    if not str(parsed_data.get("Submitter", "")).strip():
        problems.append((None, "No submitter name was found on the board", "info"))

    return problems
