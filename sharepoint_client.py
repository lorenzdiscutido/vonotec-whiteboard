# sharepoint_client.py
"""
Uploads the master Excel file to a SharePoint document library automatically,
using the Microsoft Graph API with app-only authentication (client-credentials
flow). This runs with no one signed in, which is what makes it possible to
sync in the background every time a photo is processed.

Required Streamlit secrets (set in Streamlit Cloud > Settings > Secrets):
    SHAREPOINT_TENANT_ID
    SHAREPOINT_CLIENT_ID
    SHAREPOINT_CLIENT_SECRET

Required config.py values (not secret, so they live in code):
    SHAREPOINT_HOSTNAME
    SHAREPOINT_SITE_NAME
    SHAREPOINT_FOLDER_PATH
"""
import os
import time
import random
from urllib.parse import quote

import requests
import streamlit as st

from config import SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_NAME, SHAREPOINT_FOLDER_PATH

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Microsoft requires a resumable "upload session" above 4 MB; a single PUT
# works below that. The master file starts small but grows as photos are
# embedded, so both paths are needed.
SIMPLE_UPLOAD_LIMIT_BYTES = 4 * 1024 * 1024
# Each chunk in a resumable upload must be a multiple of 320 KiB.
# 10 MB is Microsoft's own recommended chunk size.
CHUNK_SIZE_BYTES = 10 * 1024 * 1024

XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Writing to the same file again very soon after the previous write can get a
# transient "423 Locked" while SharePoint finishes processing that previous
# write, or a "429"/"503" if it's briefly busy. These clear on their own
# within a few seconds, so they are retried rather than treated as failures.
MAX_UPLOAD_ATTEMPTS = 5
RETRYABLE_STATUS_CODES = {423, 429, 503, 504}


def _retry_wait_seconds(response, attempt):
    """Uses the Retry-After header when SharePoint provides one, else backs off."""
    retry_after = response.headers.get("Retry-After") if response is not None else None
    if retry_after:
        try:
            return float(retry_after) + 0.5
        except ValueError:
            pass
    return min((2 ** attempt) + random.random(), 20)

# The access token and the resolved site id are reused across calls instead
# of being fetched every time, since both are slow network round-trips and
# neither changes between photos in the same run.
_token_cache = {"access_token": None, "expires_at": 0}
_site_id_cache = {"site_id": None}


def _get_access_token():
    """Gets an app-only Graph API token, reusing it until shortly before it expires."""
    now = time.time()
    if _token_cache["access_token"] and now < _token_cache["expires_at"] - 60:
        return _token_cache["access_token"]

    tenant_id = st.secrets["SHAREPOINT_TENANT_ID"]
    client_id = st.secrets["SHAREPOINT_CLIENT_ID"]
    client_secret = st.secrets["SHAREPOINT_CLIENT_SECRET"]

    response = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    response.raise_for_status()
    token_data = response.json()

    _token_cache["access_token"] = token_data["access_token"]
    _token_cache["expires_at"] = now + token_data.get("expires_in", 3600)
    return _token_cache["access_token"]


def _get_site_id():
    """Resolves the SharePoint site's Graph API id once, then reuses it."""
    if _site_id_cache["site_id"]:
        return _site_id_cache["site_id"]

    token = _get_access_token()
    url = f"{GRAPH_BASE}/sites/{SHAREPOINT_HOSTNAME}:/sites/{SHAREPOINT_SITE_NAME}"
    response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    response.raise_for_status()
    site_id = response.json()["id"]

    _site_id_cache["site_id"] = site_id
    return site_id


def _encode_item_path(folder_path, filename):
    """URL-encodes each segment of the folder path + filename (spaces,
    parentheses, etc.), while keeping the slashes between segments intact."""
    segments = [s for s in folder_path.split("/") if s] + [filename]
    return "/".join(quote(segment, safe="") for segment in segments)


def download_file(remote_filename, local_path):
    """Downloads the file from SharePoint to local_path, if it exists there.

    This matters because this app's local disk is temporary: it can be wiped
    by a reboot or redeploy, while the SharePoint copy persists. Without this,
    a wiped-and-restarted app would start a brand new, empty master file and
    overwrite the fuller one already in SharePoint the next time it synced.

    Returns True if an existing file was downloaded, False if nothing exists
    there yet (a normal first-ever run, not an error)."""
    # This is often the very first SharePoint network call of a fresh session,
    # before the connection has "warmed up," so a one-off transient failure
    # here is retried rather than surfaced as a scary warning right away.
    last_error = None
    for attempt in range(3):
        try:
            token = _get_access_token()
            site_id = _get_site_id()
            item_path = _encode_item_path(SHAREPOINT_FOLDER_PATH, remote_filename)
            url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{item_path}:/content"

            response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120, stream=True)
            if response.status_code == 404:
                return False
            response.raise_for_status()

            with open(local_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    f.write(chunk)
            return True
        except requests.exceptions.HTTPError as e:
            # A real HTTP error (not just "not found") is worth retrying once or
            # twice, but not endlessly, unlike a plain connection/timeout issue.
            last_error = e
            if attempt < 2:
                time.sleep(2 + attempt * 2)
        except requests.exceptions.RequestException as e:
            # Connection/timeout-type issues: the likely "cold start" case.
            last_error = e
            if attempt < 2:
                time.sleep(2 + attempt * 2)

    raise last_error


def upload_file(local_path, remote_filename):
    """Uploads local_path to the configured SharePoint folder, replacing any
    existing file with the same name.

    Raises an exception on failure. The caller should treat this as
    best-effort: the local Excel file is always the source of truth and is
    saved successfully regardless of whether this upload succeeds."""
    token = _get_access_token()
    site_id = _get_site_id()
    item_path = _encode_item_path(SHAREPOINT_FOLDER_PATH, remote_filename)
    file_size = os.path.getsize(local_path)

    for attempt in range(MAX_UPLOAD_ATTEMPTS):
        try:
            if file_size <= SIMPLE_UPLOAD_LIMIT_BYTES:
                _upload_small_file(local_path, site_id, item_path, token)
            else:
                _upload_large_file(local_path, site_id, item_path, token, file_size)
            return
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            if status in RETRYABLE_STATUS_CODES and attempt < MAX_UPLOAD_ATTEMPTS - 1:
                time.sleep(_retry_wait_seconds(e.response, attempt))
                continue
            raise


def _upload_small_file(local_path, site_id, item_path, token):
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{item_path}:/content"
    with open(local_path, "rb") as f:
        response = requests.put(
            url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": XLSX_MIME_TYPE},
            data=f,
            timeout=120,
        )
    response.raise_for_status()


def _upload_large_file(local_path, site_id, item_path, token, file_size):
    session_url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{item_path}:/createUploadSession"
    response = requests.post(
        session_url,
        headers={"Authorization": f"Bearer {token}"},
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
        timeout=30,
    )
    response.raise_for_status()
    upload_url = response.json()["uploadUrl"]

    with open(local_path, "rb") as f:
        start = 0
        while start < file_size:
            chunk = f.read(CHUNK_SIZE_BYTES)
            end = start + len(chunk) - 1
            # No Authorization header here: the uploadUrl itself is pre-authenticated.
            chunk_response = requests.put(
                upload_url,
                headers={
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {start}-{end}/{file_size}",
                },
                data=chunk,
                timeout=120,
            )
            chunk_response.raise_for_status()
            start += len(chunk)
