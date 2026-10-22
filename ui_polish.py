# ui_polish.py
"""Small helpers that make the Streamlit app look finished:

  configure_page()    - page title, favicon, collapsed sidebar, no default menu items
  load_polish_css()   - injects theme_polish.css (call it AFTER your existing CSS)
  empty_state(...)    - a proper empty/zero-results card instead of grey caption text
  live_chip(...)      - a small "Live" status chip for the Dashboard header
"""
import html
from pathlib import Path

import streamlit as st

_CSS_PATH = Path(__file__).with_name("theme_polish.css")

_ICON_INBOX = (
    '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M3 13l3-8h12l3 8"/><path d="M3 13v6h18v-6h-5l-1.5 2h-5L8 13H3z"/></svg>'
)


def configure_page(page_icon="\U0001F3D7\uFE0F"):
    """Must be the FIRST Streamlit call in app.py, and replaces any existing st.set_page_config.
    Pass a path to your Vonotec favicon (e.g. "assets/favicon.png") to use it instead of the emoji."""
    st.set_page_config(
        page_title="Vonotec | Automated Defect Logger",
        page_icon=page_icon,
        layout="wide",
        initial_sidebar_state="collapsed",
        menu_items={"Get help": None, "Report a bug": None, "About": None},
    )


def load_polish_css():
    """Injects theme_polish.css. Call after your existing stylesheet so these rules win."""
    css = _CSS_PATH.read_text(encoding="utf-8")
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def empty_state(title, body=""):
    """A friendly empty/zero-results card with an icon, a title and a one-line explanation."""
    body_html = f'<div class="empty-body">{html.escape(body)}</div>' if body else ""
    st.markdown(
        '<div class="empty-state">'
        f'<div class="empty-icon">{_ICON_INBOX}</div>'
        f'<div class="empty-title">{html.escape(title)}</div>'
        f"{body_html}"
        "</div>",
        unsafe_allow_html=True,
    )


def live_chip(text):
    """A small green status chip, e.g. live_chip("Live \u00b7 refreshes every 15s")."""
    st.markdown(
        f'<div class="live-chip"><span class="live-dot"></span>{html.escape(text)}</div>',
        unsafe_allow_html=True,
    )
