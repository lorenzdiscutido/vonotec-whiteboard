# ui_auto_watch.py
"""Dedicated horizontal status strip for SharePoint Auto-Extract (auto_watcher.py)
positioned above the main workspace, matching the enterprise Figma layout."""
import datetime
import html
import importlib
import time
import streamlit as st

import auto_watcher
import settings as _settings_mod

if not hasattr(_settings_mod, "ENABLE_AUTO_WATCH"):
    importlib.reload(_settings_mod)

if not hasattr(auto_watcher, "trigger_check_now"):
    importlib.reload(auto_watcher)
    auto_watcher.start_if_due()

ENABLE_SHAREPOINT_SYNC = getattr(_settings_mod, "ENABLE_SHAREPOINT_SYNC", True)
ENABLE_AUTO_WATCH = getattr(_settings_mod, "ENABLE_AUTO_WATCH", True)
AUTO_WATCH_PHOTO_ROOT = getattr(
    _settings_mod,
    "AUTO_WATCH_PHOTO_ROOT",
    ["07 - Employee Folders (to be reviewed)", "Lorenz", "Whiteboard Photos"],
)
AUTO_WATCH_INTERVAL_SECONDS = getattr(_settings_mod, "AUTO_WATCH_INTERVAL_SECONDS", 120)

ENABLED_KEY = "auto_watch_enabled_toggle"
ERROR_KEY = "auto_watch_action_error"
_LOCAL_TZ = datetime.timezone(datetime.timedelta(hours=8))


def _short_monitored_path(parts):
    """Formats the watched folder path into a clean display label."""
    if not parts:
        return "SharePoint"
    cleaned = []
    for p in parts:
        label = p.split(" - ", 1)[-1] if " - " in p and p[:2].isdigit() else p
        if "(" in label:
            label = label.split("(")[0].strip()
        cleaned.append(label)
    if len(cleaned) >= 2:
        return f"{cleaned[0]} / {cleaned[-1]}"
    return " / ".join(cleaned)


def _relative_last_check(status):
    """Computes a friendly 'Last checked X minutes ago' string."""
    next_at = float(status.get("next_check_at") or 0.0)
    if status.get("state") in ("scanning", "working"):
        return "Checking now\u2026"
    if next_at > 0:
        last_ts = next_at - AUTO_WATCH_INTERVAL_SECONDS
        elapsed = max(0, int(time.time() - last_ts))
        if elapsed < 45:
            return "Last checked just now"
        mins = max(1, elapsed // 60)
        return f"Last checked {mins} minute{'s' if mins != 1 else ''} ago"
    if status.get("last_check"):
        return f"Last checked {status['last_check']}"
    return "Starting monitor\u2026"


def _on_enabled_change():
    try:
        enabled = bool(st.session_state.get(ENABLED_KEY, True))
        auto_watcher.set_paused(not enabled)
        st.session_state.pop(ERROR_KEY, None)
    except Exception as error:
        st.session_state[ERROR_KEY] = f"Could not update Auto-Extract state: {error}"


@st.fragment(run_every=4)
def render_auto_watch_strip():
    """Renders the full-width SharePoint Auto-Extract status strip above the workspace."""
    if not (ENABLE_SHAREPOINT_SYNC and ENABLE_AUTO_WATCH):
        return

    status = auto_watcher.get_status()
    is_enabled = not bool(status.get("paused", False))
    st.session_state[ENABLED_KEY] = is_enabled

    curr_idx = int(status.get("current_index") or 0)
    curr_total = int(status.get("current_total") or 0)
    curr_photo = status.get("current_photo") or ""
    curr_file = status.get("current_file") or ""
    waiting = int(status.get("waiting") or 0)
    state = status.get("state", "idle")

    monitored_label = _short_monitored_path(AUTO_WATCH_PHOTO_ROOT)
    if not is_enabled:
        dot_class = "paused"
        sub_text = f"Paused \u00b7 Monitoring {monitored_label}"
    elif state == "error":
        dot_class = "error"
        sub_text = f"Connection issue \u00b7 Retrying {monitored_label}"
    elif state == "working" and curr_total > 0:
        dot_class = "active"
        sub_text = f"Extracting {curr_idx} of {curr_total} photo(s) \u00b7 {monitored_label}"
    elif state == "scanning":
        dot_class = "active"
        sub_text = f"Scanning folders \u00b7 {monitored_label}"
    elif waiting > 0:
        dot_class = "active"
        sub_text = f"Active \u00b7 {waiting} new photo(s) queued in {monitored_label}"
    else:
        dot_class = "active"
        sub_text = f"Active \u00b7 Monitoring {monitored_label}"

    last_checked_str = _relative_last_check(status)

    with st.container(border=True, key="auto_watch_strip"):
        col_left, col_meta, col_toggle, col_btn = st.columns([4.8, 2.2, 1.9, 1.1], vertical_alignment="center")

        with col_left:
            st.markdown(
                f'<div class="aw-strip-brand">'
                f'  <div class="aw-strip-icon-wrap">'
                f'    <span class="aw-strip-dot {dot_class}"></span>'
                f'    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#16a34a" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
                f'      <polygon points="12 2 2 7 12 12 22 7 12 2"/>'
                f'      <polyline points="2 12 12 17 22 12"/>'
                f'      <polyline points="2 17 12 22 22 17"/>'
                f'    </svg>'
                f'  </div>'
                f'  <div class="aw-strip-text">'
                f'    <div class="aw-strip-title">SharePoint Auto-Extract</div>'
                f'    <div class="aw-strip-sub">{html.escape(sub_text)}</div>'
                f'  </div>'
                f'</div>',
                unsafe_allow_html=True,
            )

        with col_meta:
            st.markdown(
                f'<div class="aw-strip-checked">{html.escape(last_checked_str)}</div>',
                unsafe_allow_html=True,
            )

        with col_toggle:
            st.toggle(
                "Auto-Extract enabled",
                key=ENABLED_KEY,
                on_change=_on_enabled_change,
                help="Automatically extracts new whiteboard photos uploaded to the monitored SharePoint folder.",
            )

        with col_btn:
            busy = state in ("scanning", "working")
            if st.button(
                "Check now",
                key="auto_watch_check_now",
                disabled=not is_enabled or busy,
                use_container_width=True,
            ):
                if hasattr(auto_watcher, "trigger_check_now"):
                    auto_watcher.trigger_check_now(skip_settle=True)
                st.rerun()

        if state == "working" and curr_total > 0:
            fraction = min(max(curr_idx / curr_total, 0.08), 1.0)
            target_note = f" \u2192 {curr_file}" if curr_file else ""
            photo_note = f": {curr_photo}" if curr_photo else ""
            st.progress(fraction, text=f"Processing {curr_idx} of {curr_total}{photo_note}{target_note}")

        if status.get("error"):
            st.error(f"Last check failed: {status['error']}. Retrying automatically.")
        if st.session_state.get(ERROR_KEY):
            st.error(st.session_state[ERROR_KEY])

        if status.get("failed"):
            f_col1, f_col2 = st.columns([4, 1], vertical_alignment="center")
            with f_col1:
                st.warning(f"{status['failed']} photo(s) could not be processed after 3 attempts.")
            with f_col2:
                if st.button("Retry failed", key="auto_watch_retry", use_container_width=True):
                    auto_watcher.request_retry_failed()
                    st.rerun()


render_auto_watch_card = render_auto_watch_strip