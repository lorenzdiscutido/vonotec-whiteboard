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
