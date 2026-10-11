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

# Choices shown in the "Save into which folder?" dropdown
MAIN_FOLDER_LABEL = "(Main folder, no subfolder)"
NEW_FOLDER_LABEL = "+ Create a new folder..."
