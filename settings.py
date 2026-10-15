# settings.py
"""App-level switches and constants (kept apart from config.py, which holds
the defect/reference data)."""

LOG_FILE = "processed_log.json"

# If Gemini returns broken JSON, try again this many extra times
EXTRACTION_RETRIES = 2
# Second AI "review" pass on photos where the automatic checks find problems
ENABLE_REVIEW_PASS = True
# Push the master Excel file to SharePoint after every photo
ENABLE_SHAREPOINT_SYNC = True

# Local working copies: one subfolder per SharePoint folder
LOCAL_BATCH_DIR = "local_batches"

# Remembers which local batch files have been uploaded to SharePoint
SYNC_STATUS_FILE = "sync_status.json"

# How often (in seconds) the Analytics Dashboard checks SharePoint for new or
# changed Excel files. Lower = more current, but more SharePoint requests.
DASHBOARD_REFRESH_SECONDS = 15

# Choices shown in the "Save into which folder?" dropdown
MAIN_FOLDER_LABEL = "(Main folder, no subfolder)"
NEW_FOLDER_LABEL = "+ Create a new folder..."

# Where "Process from SharePoint" starts browsing for photos, as a path counted from the
# top of the site's document library, e.g. "07 - Employee Folders (to be reviewed)".
# "" starts at the very top of the library. People can open subfolders from there
# but cannot go above this folder.
PHOTO_BROWSE_ROOT = ""

# Only top-level folders of the document library whose names start with one of these
# are shown when browsing for photos (e.g. "07 - Employee Folders ..."). Folders deeper
# down are never filtered. An empty tuple () shows every top-level folder.
ALLOWED_ROOT_PREFIXES = ("00", "01", "02", "04", "07")

# Remembers the last photo folder and save folder used, so the app opens where you left off
PREFS_FILE = "user_prefs.json"

# ---------------------------------------------------------------------------
# AUTO-EXTRACT (auto_watcher.py): photos uploaded to the folder below are read
# and added to the right Excel file by themselves.
# ---------------------------------------------------------------------------

# Master switch. False = the watcher never starts.
ENABLE_AUTO_WATCH = True

# The folder people upload to, counted from the top of the document library:
#   Whiteboard Photos / Project Name / [Tower #] / Elevation / Drop
# It is created automatically if it does not exist yet.
AUTO_WATCH_PHOTO_ROOT = ["07 - Employee Folders (to be reviewed)", "Lorenz", "Whiteboard Photos"]

# How long (seconds) the watcher waits between checks of the folder
AUTO_WATCH_INTERVAL_SECONDS = 120

# A photo modified less than this many seconds ago is left for the next check,
# so a photo that is still uploading is never read half-finished
AUTO_WATCH_SETTLE_SECONDS = 120

# Safety stop: never walk more folders than this in one check
AUTO_WATCH_MAX_FOLDERS = 2000

# A photo whose reading or Excel step fails is tried again on later checks, up to
# this many times in total, then it is set aside (see "Retry failed photos" in the app)
AUTO_WATCH_MAX_ATTEMPTS = 3

# Small files the watcher keeps in the Excel landing folder in SharePoint, so nothing
# is lost when the app restarts (the app's own disk is wiped on every restart):
#   state   = which photos were already processed (written only by the watcher)
#   control = the Pause switch and the retry request (written only by the app's buttons)
AUTO_WATCH_STATE_FILE = "auto_watch_state.json"
AUTO_WATCH_CONTROL_FILE = "auto_watch_control.json"