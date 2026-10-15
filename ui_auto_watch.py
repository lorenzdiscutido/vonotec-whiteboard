# ui_auto_watch.py
"""The small status card for the auto-extract watcher (auto_watcher.py): is it running,
when it last looked, what it did, a Pause switch and a "Retry failed photos" button.

It is a "fragment", so it refreshes itself every few seconds without reloading the page."""
import streamlit as st

import auto_watcher
import ui_components
from settings import (
    ENABLE_SHAREPOINT_SYNC, ENABLE_AUTO_WATCH, AUTO_WATCH_PHOTO_ROOT, AUTO_WATCH_INTERVAL_SECONDS,
)

PAUSE_KEY = "auto_watch_pause_toggle"
ERROR_KEY = "auto_watch_action_error"

_STATE_TEXT = {
    "starting": "Starting...",
    "idle": "Running. Waiting for new photos",
    "working": "Running. Reading photos now",
    "paused": "Paused",
    "error": "Could not check the folder",
}
_LEVEL_MARK = {"success": "Added", "info": "Note", "warning": "Check", "error": "Problem"}


def _on_pause_change():
    try:
        auto_watcher.set_paused(bool(st.session_state.get(PAUSE_KEY)))
        st.session_state.pop(ERROR_KEY, None)
    except Exception as error:
        st.session_state[ERROR_KEY] = f"Could not change the pause setting: {error}"


@st.fragment(run_every=15)
def render_auto_watch_card():
    if not (ENABLE_SHAREPOINT_SYNC and ENABLE_AUTO_WATCH):
        return

    status = auto_watcher.get_status()
    # Follow the real state, so the switch is right even if someone else paused it
    st.session_state[PAUSE_KEY] = bool(status["paused"])

    with st.container(border=True, key="auto_watch_card"):
        ui_components.card_heading("Auto-extract", "SharePoint photo watcher")
        st.caption(
            f'Watching: {" / ".join(AUTO_WATCH_PHOTO_ROOT)}. New photos are read and added to the right '
            f"Excel file by themselves, about every {AUTO_WATCH_INTERVAL_SECONDS // 60} minutes."
        )

        lines = [f'Status: {_STATE_TEXT.get(status["state"], status["state"])}']
        if status["last_check"]:
            lines.append(f'Last check: {status["last_check"]} (Manila time)')
        lines.append(f'Added since the app started: {status["processed_session"]}')
        if status["waiting"]:
            lines.append(f'Waiting for upload to finish: {status["waiting"]} photo(s)')
        st.caption("  \n".join(lines))

        if status["error"]:
            st.error(f'The last check failed: {status["error"]}. It tries again automatically.')
        if st.session_state.get(ERROR_KEY):
            st.error(st.session_state[ERROR_KEY])

        st.toggle("Pause auto-extract", key=PAUSE_KEY, on_change=_on_pause_change,
                  help="Stops new photos from being read. Photos stay in the folder and are "
                       "picked up when you switch it back on.")

        if status["failed"]:
            st.warning(f'{status["failed"]} photo(s) could not be processed and were set aside.')
            if st.button("Retry failed photos", type="primary", key="auto_watch_retry"):
                try:
                    auto_watcher.request_retry_failed()
                    st.info("They will be tried again on the next check.")
                except Exception as error:
                    st.error(f"Could not request a retry: {error}")

        if status["events"]:
            st.caption("Latest activity")
            for event in status["events"][:6]:
                st.caption(f'{event["time"]}  |  {_LEVEL_MARK.get(event["level"], "Note")}: {event["text"]}')
