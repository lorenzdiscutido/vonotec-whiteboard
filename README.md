# VONOTEC Whiteboard AI — Data Extraction & Master Log Automator

An enterprise Streamlit web application that automates extracting facade defect inspection data from site whiteboard photos using **Google Gemini**, generates structured engineering Master Log workbooks (`.xlsx`) with embedded photos and live formulas, syncs bidirectionally with **SharePoint Online** via the Microsoft Graph API, and provides a real-time **Analytics Dashboard**.

---

## Key Features

### 1. Three Ways to Process Whiteboard Photos
* **Upload from Device:** Drag and drop `.jpg`, `.jpeg`, or `.png` whiteboard photos directly from a computer or mobile device with live thumbnail previews.
* **Import from SharePoint:** Browse and search any SharePoint Online document library folder using a multi-level cascading folder selector (pre-filled to `01 - Projects > 04 - Execution`), with optional recursive subfolder scanning.
* **Zero-Click Auto-Extract Watcher (`auto_watcher.py`):** A background daemon thread continuously monitors the designated SharePoint `Whiteboard Photos` folder every 2 minutes. When coworkers drop photos into project folders (`Project Name / [Tower #] / Elevation / Drop`), the watcher automatically extracts them and appends the rows to `<Project Name> / <Project - Tower - Elevation - Drop>.xlsx` in SharePoint—with no human clicks required.

### 2. Two-Pass Gemini AI Extraction & Validation
* **Pass 1 (`gemini_client.py`):** Uses `gemini-flash-lite-latest` with strict prompt rules to extract `Submitter`, `Date`, `Elevation`, `Drop`, `Floor`, `Tower`, and paired `Damage` / `Dimension` arrays across all 6 facade material categories (`Concrete`, `Paint`, `Gasket`, `Sealant`, `Broken Glass`, `With Film`).
* **Pass 2 (`extraction.py`):** Automatically triggers a second AI verification pass whenever missing dimensions, shared dimensions, or cross-category defect codes are detected.
* **Smart Deduplication:** Tracks SHA-256 content fingerprints (`processed_log.json` and SharePoint `auto_watch_state.json`) so renamed or re-uploaded duplicate photos are never double-counted.

### 3. Dynamic Excel Master Log Engine (`excel_writer.py`)
* **Embedded Whiteboard Thumbnails:** Downsizes (`image_utils.py`) and embeds each whiteboard photo directly into Column J beside its defect rows.
* **Live Formula Linking:**
  * **Column H (`Dimension`):** Plain numeric cell for easy manual editing.
  * **Column I (`Unit`) & Column P (`Repair Unit`):** Live `VLOOKUP` formulas referencing `Reference Data` (`CM`, `CM²`, `L.M.`, `SQ.M.`).
  * **Columns K–N (`Likely Cause`, `Justification`, `Recommended Repair`, `Findings`):** Live `VLOOKUP` formulas linked to `$G{row}` (`Damage`).
  * **Column O (`Repair Dimension`):** Dynamic Excel formula (`ROUNDUP` / `MROUND`) watching both `$G{row}` and `$H{row}`, automatically applying engineering multipliers and repair limits (`DG`, `CC`, `CF`, `GG`, `FG`, `FF`).
  * **One-Keystroke Deletion:** Deleting a defect code in Column G automatically blanks out the entire row's units, descriptions, and repair calculations.
* **Visual Flagging:** Rows with missing or shared dimensions are highlighted in red for rapid engineering review.

### 4. Bidirectional SharePoint Sync & Concurrency Safety
* **App-Only Graph API Authentication (`sharepoint_client.py`):** Uses OAuth2 client-credentials flow—no user sign-in required.
* **Resumable Uploads & Backoff:** Uses simple upload (`<= 4 MB`) or chunked resumable upload sessions (`> 4 MB`) with exponential backoff for `423 Locked`, `429 Too Many Requests`, `503`, and `504` responses.
* **Write Locking (`write_lock.py`):** Prevents manual batch runs and the background auto-watcher from writing to the same `.xlsx` file simultaneously.
* **Stranded File Recovery (`sync_tracker.py`):** Tracks local batch files and provides a one-click **Push to SharePoint** retry queue if a network interruption occurs.

### 5. Real-Time Engineering Analytics Dashboard (`ui_dashboard.py`)
* **Background Warm-Up (`dashboard_loader.py`):** Pre-fetches and caches parsed SharePoint `.xlsx` workbooks across all nested subfolders (`max_depth=8`) in a background thread so the dashboard opens instantly.
* **Engineering Overview (4 KPI Cards):**
  1. **Total Defects**
  2. **Linear Repair Total (`LM`)** *(with Board Total in `LM`)*
  3. **Area Repair Total (`SQ LM`)** *(with Board Total in `SQ LM`)*
  4. **Missing Dimensions** *(with a collapsible warm-amber review queue pinpointing the exact Folder, File, Tower, Floor, Drop, and Defect Code)*
* **Repair Quantity Summary Table:**
  * Ordered with all **Linear Meter (`LM`)** defects first (`C-`, `CC+`, `C+`, `DG`, `DS CC..FF`, `MS CC..FF`), followed by **Square Meter (`SQ LM`)** defects (`BH`, `US`, `DP`, `FP`, `BP`), and non-dimensional codes (`NP`, `BG`, `GLASS-FRAME`).
  * Includes a **Download summary XLSX** export button with formatted headers and portfolio totals.
* **Location Analysis & Audit Log:** Interactive Altair charts for **Defects by Material**, **Defects by Tower**, and **Defects by Floor**, plus an expandable **Raw Audit Log** drawer.

---

## Defect Code & Formula Reference

| Material | Defect Codes | Board Unit | Repair Unit | Repair Formula / Rule |
| :--- | :--- | :---: | :---: | :--- |
| **Concrete (Cracks)** | `C-`, `CC-` *(merged as `C-`)*, `CC+`, `C+` | `CM` | `L.M.` | `ROUNDUP(MROUND(H*1.5, 10)/100, 2)` |
| **Concrete (Surface)** | `BH`, `US` | `CM²` | `SQ.M.` | `ROUNDUP(H*1.15/1000, 2)` |
| **Paint** | `DP`, `FP`, `BP` | `CM²` | `SQ.M.` | `IF(ROUNDUP(H*1.3/1000, 2)<1, 1, ROUNDUP(H*1.3/1000, 2))` |
| **Paint (No Paint)** | `NP` | — | — | Non-dimensional |
| **Gasket** | `DG` | `CM` | `L.M.` | `ROUNDUP(MIN(MROUND(H*8, 10), 3.2)/100, 2)` *(Limit: 3.2)* |
| **Sealant** | `DS`, `MS` *(+ `CC`, `CF`/`FC`, `GG`, `FG`/`GF`, `FF`)* | `CM` | `L.M.` | `MROUND(MIN(H*8/100, Limit), 0.05)` *(Limits: `CC`=4.8, `CF`/`FF`=3.2, `GG`/`FG`=1.8)* |
| **Broken Glass** | `BG` | — | — | Non-dimensional |
| **With Film** | `GLASS-FRAME`, `FRAME` *(merged)* | — | — | Non-dimensional |

---

## Repository Structure

```text
vonotec-whiteboard/
├── .streamlit/
│   └── config.toml          # Streamlit light theme & typography configuration
├── app.py                   # Main entry point, header view switcher,stepper & workspace
├── auto_naming.py           # Maps watched SharePoint photo folder paths to target .xlsx names
├── auto_watcher.py          # Background SharePoint photo folder watcher & auto-extractor
├── batch_processor.py       # Per-photo extraction, Excel writing, and SharePoint sync loop
├── config.py                # Reference data table, material rules, and Excel formula templates
├── dashboard_data.py        # Parses .xlsx workbooks and aggregates engineering/BOQ tables
├── dashboard_loader.py      # Background multi-threaded SharePoint .xlsx cache & warm-up
├── defect_rules.py          # Python repair dimension calculator (Decimal exact parity)
├── excel_writer.py          # OpenPyXL workbook generator (Whiteboard Data & Reference Data)
├── extraction.py            # Two-pass Gemini extraction and validation orchestrator
├── file_utils.py            # Processed photo hash log, filename sanitizer, user preferences
├── gemini_client.py         # Gemini API client, rate limiter, prompt rules & JSON normalizer
├── image_utils.py           # EXIF-aware photo resizing for Excel embedding
├── requirements.txt         # Python package dependencies
├── settings.py              # Application switches, folder roots, and auto-watch constants
├── sharepoint_client.py     # Microsoft Graph API client (auth, browse, upload, download)
├── style.css                # Enterprise Figma stylesheet (DM Sans, pills, KPI cards, tables)
├── sync_tracker.py          # Tracks and retries pending SharePoint .xlsx uploads
├── ui_auto_watch.py         # Live status card, progress bar & controls for Auto-Extract
├── ui_components.py         # Top header bar, workflow stepper, headings & contextual tips
├── ui_dashboard.py          # Analytics Dashboard view, KPI cards, BOQ export & Altair charts
├── ui_sections.py           # Cascading output folder picker, batch results & file management
├── ui_source_browser.py     # Cascading SharePoint photo source browser with type-to-search
├── vonotec.png              # VONOTEC brand logo
└── write_lock.py            # Per-file thread locks between manual batches and auto-watcher
```

---

## Local Setup & Deployment

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure Streamlit Secrets (`.streamlit/secrets.toml`)
Create `.streamlit/secrets.toml` locally (or configure under **Settings → Secrets** on Streamlit Community Cloud):

```toml
GEMINI_API_KEY = "your-google-gemini-api-key"
SHAREPOINT_TENANT_ID = "your-azure-ad-tenant-id"
SHAREPOINT_CLIENT_ID = "your-entra-app-client-id"
SHAREPOINT_CLIENT_SECRET = "your-entra-app-client-secret"
```

### 3. Configure SharePoint Paths (`config.py` & `settings.py`)
* **`config.py`:**
  * `SHAREPOINT_HOSTNAME`: e.g., `"vonotecph.sharepoint.com"`
  * `SHAREPOINT_SITE_NAME`: e.g., `"Vonotec"`
  * `SHAREPOINT_FOLDER_PATH`: Master Excel output root (e.g., `"07 - Employee Folders (to be reviewed)/Lorenz/Whiteboard Excel Files"`)
* **`settings.py`:**
  * `AUTO_WATCH_PHOTO_ROOT`: Watched SharePoint folder for automatic background extraction (e.g., `["07 - Employee Folders (to be reviewed)", "Lorenz", "Whiteboard Photos"]`).

### 4. Run the Application
```bash
streamlit run app.py
```
