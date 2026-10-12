# sharepoint_client.py
"""
Uploads the master Excel file to a SharePoint document library automatically,
using the Microsoft Graph API with app-only authentication (client-credentials
flow). This runs with no one signed in, which is what makes it possible to
sync in the background every time a photo is processed.

Files can be saved straight into the landing folder, or into a subfolder of
it. Subfolders can be listed and created from here, so people can group
their batches (for example one folder per project or per month).

Whiteboard photos that coworkers put in those folders can also be listed
(list_image_files) and read straight into memory (download_file_to_bytes).

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
import re
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

# Photo types the app can read (matched ignoring capitals)
IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png")

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


def _encode_folder_path(segments):
    return "/".join(quote(s, safe="") for s in segments if s)


def _landing_segments():
    return [s for s in SHAREPOINT_FOLDER_PATH.split("/") if s]


def _remote_folder(subfolder=""):
    """The landing folder, plus the chosen subfolder inside it (if any)."""
    parts = _landing_segments()
    if subfolder:
        parts += [s for s in subfolder.split("/") if s]
    return "/".join(parts)


# ---------------------------------------------------------------------------
# Subfolders: listing and creating
# ---------------------------------------------------------------------------

def sanitize_folder_name(name):
    """Turns what the person typed into a safe SharePoint folder name, or
    None if nothing usable was entered."""
    name = (name or "").strip()
    name = re.sub(r'[\\/:*?"<>|]', "_", name)  # characters not allowed in SharePoint names
    name = name.rstrip(". ")  # trailing dots/spaces aren't allowed either
    name = name[:100]
    return name or None


def list_subfolders():
    """Returns the names of the folders directly inside the landing folder,
    sorted A-Z. Returns an empty list if the landing folder doesn't exist yet."""
    site_id = _get_site_id()
    encoded = _encode_folder_path(_landing_segments())
    if encoded:
        url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded}:/children?$select=name,folder&$top=200"
    else:
        url = f"{GRAPH_BASE}/sites/{site_id}/drive/root/children?$select=name,folder&$top=200"

    names = []
    while url:
        token = _get_access_token()
        response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
        if response.status_code == 404:
            return []
        response.raise_for_status()
        data = response.json()
        # Children that have a "folder" facet are folders; the rest are files
        names += [item["name"] for item in data.get("value", []) if "folder" in item]
        url = data.get("@odata.nextLink")  # more pages, if there are many folders
    return sorted(names, key=str.lower)


def _create_folder_if_missing(parent_segments, name):
    site_id = _get_site_id()
    token = _get_access_token()
    if parent_segments:
        url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{_encode_folder_path(parent_segments)}:/children"
    else:
        url = f"{GRAPH_BASE}/sites/{site_id}/drive/root/children"

    response = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}"},
        json={"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
        timeout=30,
    )
    if response.status_code == 409:
        return  # already exists, which is fine
    response.raise_for_status()


def ensure_subfolder(name):
    """Creates the subfolder inside the landing folder. Does nothing if it
    already exists. Also creates the landing folder itself if it's missing."""
    parent = []
    for segment in _landing_segments() + [name]:
        _create_folder_if_missing(parent, segment)
        parent.append(segment)


# ---------------------------------------------------------------------------
# Files: listing
# ---------------------------------------------------------------------------

def list_xlsx_files(subfolder=""):
    """Returns the Excel files directly inside the landing folder (subfolder="")
    or inside one subfolder of it, sorted A-Z by name.

    Each item: {"name", "modified", "size", "subfolder"}. "modified" is
    SharePoint's last-modified time; the dashboard uses it to know when a
    cached copy of the file is out of date. Excel's hidden "~$..." lock files
    (created while someone has a workbook open) are left out. Returns an empty
    list if the folder doesn't exist."""
    site_id = _get_site_id()
    encoded = _encode_folder_path(_remote_folder(subfolder).split("/"))
    select = "$select=name,size,lastModifiedDateTime,file&$top=200"
    if encoded:
        url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded}:/children?{select}"
    else:
        url = f"{GRAPH_BASE}/sites/{site_id}/drive/root/children?{select}"

    files = []
    while url:
        token = _get_access_token()
        response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
        if response.status_code == 404:
            return []
        response.raise_for_status()
        data = response.json()
        for item in data.get("value", []):
            name = item.get("name", "")
            # Children with a "file" facet are files; folders have a "folder" facet instead
            if "file" not in item:
                continue
            if name.startswith("~$") or not name.lower().endswith(".xlsx"):
                continue
            files.append({
                "name": name,
                "modified": item.get("lastModifiedDateTime", ""),
                "size": item.get("size", 0),
                "subfolder": subfolder,
            })
        url = data.get("@odata.nextLink")  # more pages, if there are many files
    return sorted(files, key=lambda f: f["name"].lower())


def list_image_files(subfolder=""):
    """Returns the photos (.jpg, .jpeg, .png) directly inside the landing folder
    (subfolder="") or inside one subfolder of it, sorted A-Z by name.

    Each item: {"name", "modified", "size", "subfolder"}. Sub-folders are not
    searched, and only items with a "file" facet count (folders have a "folder"
    facet instead). Returns an empty list if the folder doesn't exist."""
    site_id = _get_site_id()
    encoded = _encode_folder_path(_remote_folder(subfolder).split("/"))
    select = "$select=name,size,lastModifiedDateTime,file&$top=200"
    if encoded:
        url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded}:/children?{select}"
    else:
        url = f"{GRAPH_BASE}/sites/{site_id}/drive/root/children?{select}"

    files = []
    while url:
        token = _get_access_token()
        response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
        if response.status_code == 404:
            return []
        response.raise_for_status()
        data = response.json()
        for item in data.get("value", []):
            name = item.get("name", "")
            if "file" not in item:
                continue  # a folder
            if not name.lower().endswith(IMAGE_EXTENSIONS):
                continue
            files.append({
                "name": name,
                "modified": item.get("lastModifiedDateTime", ""),
                "size": item.get("size", 0),
                "subfolder": subfolder,
            })
        url = data.get("@odata.nextLink")  # more pages, if there are many files
    return sorted(files, key=lambda f: f["name"].lower())


# ---------------------------------------------------------------------------
# Browsing ANY folder of the document library (for photos that live outside the landing folder)
# ---------------------------------------------------------------------------

# Safety stop for "include subfolders": never walk more folders than this
MAX_RECURSIVE_FOLDERS = 300


def _children_url(segments, select):
    """Graph URL for the children of the folder at `segments` (a list of folder
    names counted from the top of the document library; [] = the top itself)."""
    site_id = _get_site_id()
    encoded = _encode_folder_path(segments)
    if encoded:
        return f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded}:/children?{select}"
    return f"{GRAPH_BASE}/sites/{site_id}/drive/root/children?{select}"


def _iter_children(segments, select):
    """Yields every child item of a folder, following all result pages.
    Yields nothing if the folder doesn't exist."""
    url = _children_url(segments, select)
    while url:
        token = _get_access_token()
        response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
        if response.status_code == 404:
            return
        response.raise_for_status()
        data = response.json()
        for item in data.get("value", []):
            yield item
        url = data.get("@odata.nextLink")


def list_folders_at(segments):
    """Names of the folders directly inside the folder at `segments`
    (counted from the top of the document library), sorted A-Z."""
    names = [
        item["name"]
        for item in _iter_children(list(segments), "$select=name,folder&$top=200")
        if "folder" in item
    ]
    return sorted(names, key=str.lower)


def list_images_at(segments, recursive=False):
    """Photos (.jpg, .jpeg, .png) inside the folder at `segments`, optionally
    including every folder below it.

    Each item: {"id", "name", "display_name", "folder", "size", "modified"}.
    "id" is SharePoint's id for the file (use download_item_to_bytes with it);
    "display_name" is the file name with its subfolder path relative to the starting
    folder, so two photos called IMG_001.jpg in different folders stay distinguishable.
    At most MAX_RECURSIVE_FOLDERS folders are walked."""
    base = list(segments)
    select = "$select=id,name,size,lastModifiedDateTime,file,folder&$top=200"
    queue = [base]
    walked = 0
    images = []
    while queue:
        current = queue.pop(0)
        walked += 1
        for item in _iter_children(current, select):
            name = item.get("name", "")
            if "folder" in item:
                if recursive and walked + len(queue) < MAX_RECURSIVE_FOLDERS:
                    queue.append(current + [name])
                continue
            if "file" not in item or not name.lower().endswith(IMAGE_EXTENSIONS):
                continue
            relative = current[len(base):]
            images.append({
                "id": item["id"],
                "name": name,
                "display_name": "/".join(relative + [name]),
                "folder": "/".join(current),
                "size": item.get("size", 0),
                "modified": item.get("lastModifiedDateTime", ""),
            })
    return sorted(images, key=lambda i: i["display_name"].lower())


# ---------------------------------------------------------------------------
# Files: download and upload
# ---------------------------------------------------------------------------

def download_file(remote_filename, local_path, subfolder=""):
    """Downloads the file from SharePoint to local_path, if it exists there.
    subfolder is the folder inside the landing folder ("" = the landing folder itself).

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
            item_path = _encode_item_path(_remote_folder(subfolder), remote_filename)
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


def download_file_to_bytes(remote_filename, subfolder=""):
    """Downloads one file from SharePoint straight into memory and returns its
    raw bytes, so nothing is written to the local disk. subfolder is the folder
    inside the landing folder ("" = the landing folder itself).

    Returns None if the file no longer exists there. Raises an exception if the
    download keeps failing after a few tries."""
    last_error = None
    for attempt in range(3):
        try:
            token = _get_access_token()
            site_id = _get_site_id()
            item_path = _encode_item_path(_remote_folder(subfolder), remote_filename)
            url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{item_path}:/content"

            response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120)
            if response.status_code == 404:
                return None
            response.raise_for_status()
            return response.content
        except requests.exceptions.RequestException as e:
            # HTTP errors and connection/timeout problems alike: try again a couple of times
            last_error = e
            if attempt < 2:
                time.sleep(2 + attempt * 2)

    raise last_error


def download_item_to_bytes(item_id):
    """Downloads one file straight into memory using SharePoint's id for it (the
    "id" from list_images_at), so it works for a file in any folder.

    Returns None if the file no longer exists. Raises an exception if the
    download keeps failing after a few tries."""
    last_error = None
    for attempt in range(3):
        try:
            token = _get_access_token()
            site_id = _get_site_id()
            url = f"{GRAPH_BASE}/sites/{site_id}/drive/items/{quote(item_id, safe='')}/content"

            response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120)
            if response.status_code == 404:
                return None
            response.raise_for_status()
            return response.content
        except requests.exceptions.RequestException as e:
            last_error = e
            if attempt < 2:
                time.sleep(2 + attempt * 2)

    raise last_error


def upload_file(local_path, remote_filename, subfolder=""):
    """Uploads local_path to the landing folder (or a subfolder of it),
    replacing any existing file with the same name.

    Raises an exception on failure. The caller should treat this as
    best-effort: the local Excel file is always the source of truth and is
    saved successfully regardless of whether this upload succeeds."""
    token = _get_access_token()
    site_id = _get_site_id()
    item_path = _encode_item_path(_remote_folder(subfolder), remote_filename)
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
