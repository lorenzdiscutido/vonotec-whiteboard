import os
import re
import json
import uuid
import hashlib
import datetime
import base64
import traceback

import google.generativeai as genai
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

from config import REFERENCE_DATA, MATERIAL_RULES
from image_utils import load_image, prepare_excel_image
from gemini_client import get_raw_response, get_review_response, parse_and_clean_json, find_problems
import sharepoint_client

LOG_FILE = "processed_log.json"

# Single-user version: photos are processed one at a time, in order.
# (The old multi-user branch added a shared write lock and a Gemini call
# limiter across sessions; neither is needed here.)
# If Gemini returns broken JSON, try again this many extra times
EXTRACTION_RETRIES = 2

# Second AI "review" pass on photos where the automatic checks find problems
ENABLE_REVIEW_PASS = True

# Automatically push the master Excel file to SharePoint after every photo,
# so data is never only sitting on this app's temporary disk waiting to be
# downloaded manually. A sync failure never blocks or undoes the local save.
ENABLE_SHAREPOINT_SYNC = True

# Local working copies live here, in one subfolder per SharePoint folder, so
# two batches with the same name in different folders never mix.
LOCAL_BATCH_DIR = "local_batches"

# Choices shown in the "Save into which folder?" dropdown
MAIN_FOLDER_LABEL = "(Main folder, no subfolder)"
NEW_FOLDER_LABEL = "+ Create a new folder..."

def load_processed_log():
    """Loads the list of already processed photo fingerprints."""
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return []
    return []

def save_processed_log(log_list):
    """Saves the log safely (write to a temp file, then swap it in)."""
    tmp_path = LOG_FILE + ".tmp"
    with open(tmp_path, "w") as f:
        json.dump(log_list, f)
    os.replace(tmp_path, LOG_FILE)

# ==========================================
# 1. CONFIGURATION & RULES
# ==========================================
# Securely pull the API key from Streamlit Secrets
genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

def _populate_reference_sheet(wb):
    ws_ref = wb.create_sheet(title="Reference Data")
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    for row_idx, row_values in enumerate(REFERENCE_DATA, start=1):
        ws_ref.append(row_values)
        for col_idx in range(1, len(REFERENCE_DATA[0]) + 1):
            cell = ws_ref.cell(row=row_idx, column=col_idx)
            if row_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    ws_ref.column_dimensions['A'].width = 15
    ws_ref.column_dimensions['B'].width = 26
    ws_ref.column_dimensions['C'].width = 24
    ws_ref.column_dimensions['D'].width = 45
    ws_ref.column_dimensions['E'].width = 40
    ws_ref.column_dimensions['F'].width = 18
    ws_ref.row_dimensions[1].height = 28

def _refresh_reference_sheet(wb):
    """Makes sure the Reference Data sheet matches config.py. An older master file (fewer defects, or no unit column) gets the sheet rebuilt, so newly added defect codes are found by the lookup formulas."""
    if "Reference Data" not in wb.sheetnames:
        _populate_reference_sheet(wb)
        return
    ws_ref = wb["Reference Data"]
    if ws_ref.max_row != len(REFERENCE_DATA) or ws_ref.max_column != len(REFERENCE_DATA[0]):
        wb.remove(ws_ref)
        _populate_reference_sheet(wb)

def extract_data(filepath):
    """Reads one whiteboard photo with Gemini (no Streamlit calls here)."""
    img = load_image(filepath)

    # Pass 1: read the board (retry if the answer is not valid JSON)
    parsed_data = None
    for attempt in range(EXTRACTION_RETRIES + 1):
        try:
            raw_text = get_raw_response(img)
            parsed_data = parse_and_clean_json(raw_text)
            break
        except json.JSONDecodeError:
            if attempt == EXTRACTION_RETRIES:
                raise

    # Pass 2: a reviewer re-reads the board, but only when a real data problem
    # was found (wrong code, mismatched entries, a likely copy error, or nothing
    # detected at all). A minor AI hedge or a missing date/submitter is shown to
    # the user for awareness but does not by itself trigger this extra AI call.
    problems = find_problems(parsed_data)
    structural = [p for p in problems if p[2] == "structural"]
    if ENABLE_REVIEW_PASS and structural:
        try:
            raw_review = get_review_response(img, parsed_data, problems)
            parsed_data = parse_and_clean_json(raw_review)
        except Exception:
            pass  # keep the first-pass result if the review fails
        problems = find_problems(parsed_data)

    # Whatever is still doubtful is shown to the user after the batch. Only
    # "structural" problems highlight their material row red in Excel; "info"
    # problems (AI hedges, missing date/submitter) are listed but change no color.
    parsed_data["Review Notes"] = [msg for _material, msg, _kind in problems]
    parsed_data["Flagged Materials"] = sorted({mat for mat, _msg, kind in problems if mat and kind in ("structural", "manual")})
    return parsed_data

def split_dimension(text):
    """Splits a dimension like '60 CM' or '120 CM²' into (60, 'CM') / (120, 'CM²'), so the Dimension cell holds a real number and the unit sits in its own cell. Anything that doesn't look like 'number + unit' (for example '30 X 20 CM') is kept whole as text, with no unit."""
    text = str(text or "").strip()
    if not text:
        return "", ""
    match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*([A-Z]+(?:²|^2|2)?)?", text, flags=re.IGNORECASE)
    if not match:
        return text, ""
    number = float(match.group(1))
    if number.is_integer():
        number = int(number)
    return number, (match.group(2) or "").upper()

def write_parsed_data_to_excel(parsed_data, filepath, output_xlsx_path):
    """Adds one whiteboard's data to the master Excel file. Must be called while holding the write lock.
    
    Layout (new files): A-F details | G Damage | H Dimension (number only) |
    I Unit | J Whiteboard Photo | K-N lookup columns | O Repair Dimension. Older files that were
    created before the Unit column existed keep their old layout (unit written
    inside the Dimension cell), so appending to them never shifts any column."""
    # Only the specific material row(s) a problem points to are highlighted red,
    # not the whole whiteboard entry.
    flagged_materials = set(parsed_data.get("Flagged Materials", []))
    
    if os.path.exists(output_xlsx_path):
        wb = openpyxl.load_workbook(output_xlsx_path)
        ws = wb["Whiteboard Data"] if "Whiteboard Data" in wb.sheetnames else wb.active
        if "Reference Data" not in wb.sheetnames:
            _populate_reference_sheet(wb)
        # New layout has the photo column in J; the old layout had it in I
        has_unit_col = (ws["J1"].value == "Whiteboard Photo")
    else:
        has_unit_col = True
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Whiteboard Data"

        ws.append([
            "Submitter", "Date", "Elevation", "Drop", "Floor", "Tower",
            "Defects", "", "", "Whiteboard Photo",
            "POSSIBLE CAUSE", "Possible Cause Justification", "Recommended Repair", "FINDINGS", "Repair Dimension"
        ])
        ws.append([
            "", "", "", "", "", "",
            "Damage", "Dimension", "", "",
            "", "", "", "", ""
        ])

        ws.merge_cells('G1:I1')   # "Defects" spans Damage, Dimension and Unit
        ws.merge_cells('H2:I2')   # "Dimension" spans the number and its unit
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'J', 'K', 'L', 'M', 'N', 'O']:
            ws.merge_cells(f'{col}1:{col}2')

        gray_fill = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
        for row in ws['A1':'O2']:
            for cell in row:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        ws['N1'].fill = gray_fill
        ws['N2'].fill = gray_fill
        ws['O1'].fill = gray_fill
        ws['O2'].fill = gray_fill

        ws.column_dimensions['I'].width = 9
        ws.column_dimensions['K'].width = 24
        ws.column_dimensions['L'].width = 30
        ws.column_dimensions['M'].width = 28
        ws.column_dimensions['N'].width = 22
        ws.column_dimensions['O'].width = 18

        _populate_reference_sheet(wb)

    # Column positions depend on the layout of this file
    unit_col = 9 if has_unit_col else None
    photo_col = 10 if has_unit_col else 9
    last_col = photo_col + 5
    photo_col_letter = get_column_letter(photo_col)
    title_last_col = 9 if has_unit_col else 8

    start_row = ws.max_row + 1
    VALID_CODES = {row[0] for row in REFERENCE_DATA[1:]}

    rows_data = []
    for mat in ["Sealant", "Concrete", "Paint", "Gasket"]:
        rows_data.append((mat, "", True, False, mat, False)) 
        
        dmg_list = parsed_data.get(f"{mat} Damage", [])
        dim_list = parsed_data.get(f"{mat} Dimension", [])
        # Entries whose dimension was left blank on purpose, for a human to fill in
        manual_rows = {
            idx
            for note in (parsed_data.get("Manual Dimension") or [])
            if note["material"] == mat
            for idx in note["indices"]
        }
        
        expanded_dmg = []
        expanded_dim = []
        
        for i in range(max(len(dmg_list), len(dim_list))):
            dmg_str = dmg_list[i] if i < len(dmg_list) else ""
            dim_str = dim_list[i] if i < len(dim_list) else ""
            
            tokens = str(dmg_str).split()
            found_codes = [t for t in tokens if t in VALID_CODES]
            
            if len(found_codes) > 1:
                # Several defects sharing one dimension: only the first gets it
                for n, code in enumerate(found_codes):
                    expanded_dmg.append(code)
                    expanded_dim.append(dim_str if n == 0 else "")
            else:
                expanded_dmg.append(dmg_str)
                expanded_dim.append(dim_str)
        
        max_len = max(len(expanded_dmg), 1) 
        for i in range(max_len):
            dmg_val = expanded_dmg[i] if i < len(expanded_dmg) else ""
            dim_val = expanded_dim[i] if i < len(expanded_dim) else ""
            rows_data.append((dmg_val, dim_val, False, True, mat, i in manual_rows)) 

    total_rows = len(rows_data)

    static_keys = ["Submitter", "Date", "Elevation", "Drop", "Floor", "Tower"]
    for i, key in enumerate(static_keys):
        ws.cell(row=start_row, column=i+1).value = parsed_data.get(key, "")
        if total_rows > 1:
            ws.merge_cells(start_row=start_row, start_column=i+1, end_row=start_row + total_rows - 1, end_column=i+1)

    for i, (g_val, h_val, is_title, is_data, current_mat, needs_manual) in enumerate(rows_data):
        r = start_row + i
        cell_g = ws.cell(row=r, column=7)
        cell_h = ws.cell(row=r, column=8)
        
        cell_g.value = g_val

        # Setup dimension values
        number_val = None
        if has_unit_col and is_data:
            # Number only in the Dimension cell, unit in its own cell next to it
            number_val, unit = split_dimension(h_val)
            cell_h.value = number_val
            ws.cell(row=r, column=unit_col).value = unit
        else:
            cell_h.value = h_val

        is_invalid = is_data and current_mat in flagged_materials
        if is_data and g_val:
            base_code = str(g_val).split()[0]
            allowed_codes = MATERIAL_RULES.get(current_mat, [])
            if base_code not in allowed_codes:
                is_invalid = True

        if is_title:
            ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=title_last_col)
            cell_g.font = Font(bold=True)
        elif is_data:
            lookup_val = f'LEFT($G{r}, FIND(" ", $G{r}&" ") - 1)'
            ws.cell(row=r, column=photo_col + 1).value = f'=IFERROR(VLOOKUP({lookup_val}, \'Reference Data\'!$A$2:$E$15, 3, FALSE), "")'
            ws.cell(row=r, column=photo_col + 2).value = f'=IFERROR(VLOOKUP({lookup_val}, \'Reference Data\'!$A$2:$E$15, 4, FALSE), "")'
            ws.cell(row=r, column=photo_col + 3).value = f'=IFERROR(VLOOKUP({lookup_val}, \'Reference Data\'!$A$2:$E$15, 5, FALSE), "")'
            ws.cell(row=r, column=photo_col + 4).value = f'=IFERROR(VLOOKUP({lookup_val}, \'Reference Data\'!$A$2:$E$15, 2, FALSE), "")'

            # -------- NEW FEATURE: HIDDEN REPAIR DIMENSION COMPUTATION --------
            repair_val = ""
            parts = str(g_val).split()
            if parts and parts[0] in ["MS", "DS"] and has_unit_col and isinstance(number_val, (int, float)):
                limits = {"CC": 4.8, "CF": 3.2, "GG": 1.8, "FG": 1.8, "FF": 3.2}
                modifier = parts[1] if len(parts) > 1 else None
                
                computed = number_val * 0.08
                if modifier in limits:
                    limit = limits[modifier]
                    if computed > limit:
                        computed = limit
                        
                # MROUND to 0.05 logic directly in Python:
                repair_val = round(0.05 * round(computed / 0.05), 2)
                
            ws.cell(row=r, column=photo_col + 5).value = repair_val

            if is_invalid:
                for col_idx in range(7, last_col + 1):
                    ws.cell(row=r, column=col_idx).font = Font(color="FF0000")

    excel_img = prepare_excel_image(filepath, total_rows)
    ws.add_image(excel_img, f"{photo_col_letter}{start_row}")

    if total_rows > 1:
        ws.merge_cells(f"{photo_col_letter}{start_row}:{photo_col_letter}{start_row + total_rows - 1}")

    base_height = max(110 / total_rows, 20)
    for i, (_g, _h, _is_title, is_data, _mat, _manual) in enumerate(rows_data):
        r = start_row + i
        ws.row_dimensions[r].height = base_height
        for c in range(1, last_col + 1):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        if has_unit_col and is_data:
            # The unit sits right next to the number, like "60 | CM"
            ws.cell(row=r, column=unit_col).alignment = Alignment(horizontal='left', vertical='center')
            
    ws.column_dimensions[photo_col_letter].width = 25

    # Save to a temp file first, then swap it in, so a crash mid-save
    # can never leave a half-written master file.
    tmp_path = os.path.splitext(output_xlsx_path)[0] + "_tmp.xlsx"
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, output_xlsx_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise

# --- UI CONFIGURATION ---
st.set_page_config(page_title="Vonotec Whiteboard Extractor", layout="centered")

# Inject Custom CSS for readability and closer vertical spacing in header
st.markdown(
    """
    <style>
    .stApp {
        background-color: #f4f7f9;
    }
    .header-container {
        background-color: #ffffff;
        padding: 20px 24px;
        border-radius: 14px;
        box-shadow: 0 4px 14px rgba(15,23,42,0.08);
        margin-bottom: 24px;
    }
    /* Logo and title block share the same vertical center */
    .header-flex {
        display: flex;
        align-items: center;
        gap: 20px;
    }
    .header-flex img {
        display: block;
        height: auto;
        margin: 0;
    }
    .header-flex > div {
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .main-title {
        color: #1E3A8A !important;
        font-weight: 800 !important;
        margin: 0 !important;
        padding: 0 !important;
        line-height: 1.2 !important;
    }
    .sub-title {
        color: #475569 !important;
        margin: 4px 0 0 0 !important;
        padding: 0 !important;
        line-height: 1.3 !important;
        font-size: 1rem;
    }
    /* Solid Blue Button with White Text */
    div.stButton > button:first-child, .stDownloadButton > button:first-child {
        background-color: #2F5597 !important;
        color: #ffffff !important;
        border: none !important;
        font-weight: 700 !important;
        padding: 0.75rem 2rem !important;
        border-radius: 8px !important;
        width: 100%;
        box-shadow: 0 2px 6px rgba(47,85,151,0.25);
        transition: background-color 0.15s ease, transform 0.15s ease, box-shadow 0.15s ease;
    }
    div.stButton > button:first-child p, .stDownloadButton > button:first-child p {
        color: #ffffff !important;
        font-weight: 700 !important;
    }
    div.stButton > button:first-child:hover, .stDownloadButton > button:first-child:hover {
        background-color: #1E3A8A !important;
        transform: translateY(-1px);
        box-shadow: 0 4px 10px rgba(30,58,138,0.3);
    }
    .instruction-text {
        color: #1e293b !important;
        background-color: #ffffff;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #2F5597;
        margin-bottom: 24px;
    }
    /* Readable labels above inputs, e.g. "Name this batch..." above the text box */
    div[data-testid="stWidgetLabel"] p,
    div[data-testid="stWidgetLabel"] label,
    label[data-testid="stWidgetLabel"],
    .stTextInput label p,
    .stTextInput label {
        color: #1e293b !important;
        font-weight: 600 !important;
        opacity: 1 !important;
    }
    /* Readable status text: "Processing...", "Batch Complete." */
    div[data-testid="stText"],
    div[data-testid="stText"] p {
        color: #1e293b !important;
        font-weight: 600 !important;
    }
    /* Readable "N image(s) selected." line */
    .status-msg {
        color: #1e293b !important;
        font-weight: 600;
        margin: 1.25rem 0 1.25rem 0;
    }
    /* Space below the file uploader dropzone */
    div[data-testid="stFileUploader"] {
        margin-bottom: 1.5rem;
    }
    /* Space above the action buttons */
    .st-key-action_buttons {
        margin-top: 0.75rem;
        margin-bottom: 1rem;
    }
    /* Readable Skipped / Success / Error messages */
    div[data-testid="stAlert"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-left: 5px solid #2F5597 !important;
    }
    div[data-testid="stAlert"] p,
    div[data-testid="stAlert"] div[data-testid="stMarkdownContainer"] {
        color: #1e293b !important;
    }
    /* Bring Extract and Clear Photos buttons closer together */
    .st-key-action_buttons div[data-testid="stHorizontalBlock"] {
        gap: 0.5rem !important;
        justify-content: flex-start !important;
        flex-wrap: wrap !important;
    }
    .st-key-action_buttons div[data-testid="stColumn"],
    .st-key-action_buttons div[data-testid="column"] {
        flex: 0 0 auto !important;
        width: auto !important;
        min-width: 0 !important;
    }
    .st-key-action_buttons div.stButton {
        width: auto !important;
    }
    /* Orange "Clear All Photos" button (mild action) */
    .st-key-clear_photos div.stButton > button:first-child {
        background-color: #F97316 !important;
        color: #FFFFFF !important;
    }
    .st-key-clear_photos div.stButton > button:first-child p {
        color: #FFFFFF !important;
    }
    .st-key-clear_photos div.stButton > button:first-child:hover {
        background-color: #EA580C !important;
    }
    /* Red "Start Fresh" and "Yes, delete everything" buttons (destructive action) */
    .st-key-start_fresh div.stButton > button:first-child,
    .st-key-confirm_delete div.stButton > button:first-child {
        background-color: #DC2626 !important;
        color: #FFFFFF !important;
    }
    .st-key-start_fresh div.stButton > button:first-child p,
    .st-key-confirm_delete div.stButton > button:first-child p {
        color: #FFFFFF !important;
    }
    .st-key-start_fresh div.stButton > button:first-child:hover,
    .st-key-confirm_delete div.stButton > button:first-child:hover {
        background-color: #B91C1C !important;
    }
    h3 {
        color: #1E3A8A !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

try:
    with open("vonotec.png", "rb") as f:
        logo_base64 = base64.b64encode(f.read()).decode()
    st.markdown(
        f"""
        <div class="header-container">
            <div class="header-flex">
                <img src="data:image/png;base64,{logo_base64}" width="180">
                <div>
                    <h1 class="main-title">Whiteboard AI</h1>
                    <p class="sub-title">Data Extraction and Master Log Automator</p>
                </div>
            </div>
        </div>
        """, 
        unsafe_allow_html=True
    )
except FileNotFoundError:
    st.markdown(
        """
        <div class="header-container">
            <h1 class="main-title">VONOTEC Whiteboard AI</h1>
            <p class="sub-title">Data Extraction and Master Log Automator</p>
        </div>
        """, 
        unsafe_allow_html=True
    )

st.markdown(
    """
    <div class="instruction-text">
        Upload one or multiple whiteboard photos below. The AI will automatically extract the data 
        and append it to the Master Excel file. You can also drag and drop images directly into the uploader.
    </div>
    """, 
    unsafe_allow_html=True
)

st.subheader("Upload Whiteboard Photos")

# Changing the uploader's key resets it, which clears all selected photos at once
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "batch_results" not in st.session_state:
    st.session_state.batch_results = []
if "confirm_reset" not in st.session_state:
    st.session_state.confirm_reset = False

# The file this batch will be saved as, remembered after a batch finishes so
# "Master File Management" below can always point at the right file.
if "last_output_path" not in st.session_state:
    st.session_state.last_output_path = None
if "last_output_label" not in st.session_state:
    st.session_state.last_output_label = None

# The SharePoint folder that file went into ("" = the main landing folder)
if "last_output_folder" not in st.session_state:
    st.session_state.last_output_folder = ""

# The folder picked last time, so the dropdown starts there next round
if "last_folder_choice" not in st.session_state:
    st.session_state.last_folder_choice = MAIN_FOLDER_LABEL

# Folder names found in SharePoint (None = not loaded yet)
if "sp_folders" not in st.session_state:
    st.session_state.sp_folders = None
if "sp_folders_error" not in st.session_state:
    st.session_state.sp_folders_error = False

def sanitize_batch_filename(name):
    """Turns what the person typed into a safe .xlsx filename, or None if nothing usable was entered."""
    name = (name or "").strip()
    if not name:
        return None
    name = re.sub(r'[\/:*?"<>|]', "_", name)  # characters not allowed in Windows/SharePoint filenames
    name = name.rstrip(". ")  # trailing dots/spaces aren't allowed either
    name = name[:150]
    if not name:
        return None
    return name if name.lower().endswith(".xlsx") else f"{name}.xlsx"

def refresh_folder_list():
    """Reads the current subfolders from SharePoint. If that fails, the app carries on
    with an empty list (main folder + creating a new folder still work)."""
    try:
        st.session_state.sp_folders = sharepoint_client.list_subfolders()
        st.session_state.sp_folders_error = False
    except Exception:
        traceback.print_exc()
        st.session_state.sp_folders = []
        st.session_state.sp_folders_error = True

uploaded_files = st.file_uploader(
    "Choose whiteboard images...", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True,
    label_visibility="collapsed",
    key=f"uploader_{st.session_state.uploader_key}"
)

# As soon as new photos are selected, drop the previous batch's messages
# so "Batch Complete." can't be mistaken for the new photos being processed
if uploaded_files:
    st.session_state.batch_results = []

# Show the results of the last batch (kept even after the uploader is cleared)
for kind, text in st.session_state.batch_results:
    if kind == "warning":
        st.warning(text)
    elif kind == "error":
        st.error(text)
    elif kind == "success":
        st.success(text)
    else:
        st.markdown(f'<p class="status-msg">{text}</p>', unsafe_allow_html=True)

if uploaded_files:
    st.markdown(
        f'<p class="status-msg">{len(uploaded_files)} image(s) selected.</p>',
        unsafe_allow_html=True
    )

# Suggest a fresh default name each time a new round of photos is selected
# (after Clear Photos or after a batch finishes), without overwriting
# whatever the person is actively typing in the meantime.
if st.session_state.get("batch_name_round") != st.session_state.uploader_key:
    st.session_state.batch_name_round = st.session_state.uploader_key
    st.session_state.batch_name_input = f"Batch_{datetime.datetime.now().strftime('%Y-%m-%d_%H%M')}"
    # Re-read the folder list each round, so folders other people created show up
    st.session_state.sp_folders = None

# ---- Which SharePoint folder should this batch go into? ----
folder_choice = MAIN_FOLDER_LABEL
new_folder_raw = ""
if ENABLE_SHAREPOINT_SYNC:
    if st.session_state.sp_folders is None:
        with st.spinner("Loading SharePoint folders..."):
            refresh_folder_list()
    if st.session_state.sp_folders_error:
        st.warning(
            "Could not load the folder list from SharePoint right now. You can still save to "
            "the main folder, or create a new folder."
        )

    folder_options = [MAIN_FOLDER_LABEL] + list(st.session_state.sp_folders) + [NEW_FOLDER_LABEL]
    default_choice = st.session_state.last_folder_choice
    default_index = folder_options.index(default_choice) if default_choice in folder_options else 0
    folder_choice = st.selectbox(
        "Save into which SharePoint folder?",
        folder_options,
        index=default_index,
        key="folder_choice_input"
    )
    if folder_choice == NEW_FOLDER_LABEL:
        new_folder_raw = st.text_input(
            "Name for the new folder",
            key="new_folder_name_input"
        )

batch_name_raw = st.text_input(
    "Name this batch (this becomes the Excel file name, locally and in SharePoint)",
    key="batch_name_input"
)

with st.container(key="action_buttons"):
    btn_col1, btn_col2 = st.columns(2, gap="small")
    with btn_col1:
        extract_clicked = st.button("Extract Data & Update Excel", type="primary")
    with btn_col2:
        clear_clicked = st.button("Clear All Photos", type="primary", key="clear_photos")

if clear_clicked:
    st.session_state.uploader_key += 1
    st.session_state.batch_results = []
    st.rerun()

if extract_clicked and not sanitize_batch_filename(batch_name_raw):
    st.error("Please enter a name for this batch before extracting.")
    extract_clicked = False

# Work out which folder was chosen (and tidy up a newly typed name)
subfolder = ""
creating_new_folder = False
if extract_clicked and ENABLE_SHAREPOINT_SYNC:
    if folder_choice == NEW_FOLDER_LABEL:
        typed = sharepoint_client.sanitize_folder_name(new_folder_raw)
        if not typed:
            st.error("Please enter a name for the new folder before extracting.")
            extract_clicked = False
        else:
            # A name that matches an existing folder (ignoring capitals) just uses that folder
            match = next((f for f in st.session_state.sp_folders if f.lower() == typed.lower()), None)
            subfolder = match or typed
            creating_new_folder = match is None
    elif folder_choice != MAIN_FOLDER_LABEL:
        subfolder = folder_choice

if extract_clicked:
    # Local working copy: one local folder per SharePoint folder
    local_dir = os.path.join(LOCAL_BATCH_DIR, subfolder) if subfolder else LOCAL_BATCH_DIR
    os.makedirs(local_dir, exist_ok=True)
    output_xlsx_path = os.path.join(local_dir, sanitize_batch_filename(batch_name_raw))
    remote_filename = os.path.basename(output_xlsx_path)

    st.session_state.last_output_path = output_xlsx_path
    st.session_state.last_output_label = batch_name_raw.strip()
    st.session_state.last_output_folder = subfolder
    st.session_state.last_folder_choice = subfolder if subfolder else MAIN_FOLDER_LABEL

    total_files = len(uploaded_files)
    progress_bar = st.progress(0)
    status_text = st.empty()
    results = []

    processed_count = 0
    skipped_count = 0

    # Create the new folder in SharePoint first, so it exists before the
    # first upload. If this fails the batch still runs: the local save is
    # unaffected and the upload after each photo is retried.
    if ENABLE_SHAREPOINT_SYNC and creating_new_folder:
        try:
            sharepoint_client.ensure_subfolder(subfolder)
            if subfolder.lower() not in [f.lower() for f in st.session_state.sp_folders]:
                st.session_state.sp_folders = sorted(
                    st.session_state.sp_folders + [subfolder], key=str.lower
                )
            results.append(("status", f'Created the folder "{subfolder}" in SharePoint.'))
        except Exception:
            traceback.print_exc()
            results.append((
                "warning",
                f'Could not create the folder "{subfolder}" in SharePoint just now. '
                "The photos are still being processed and saved in the app; syncing will be tried again after each photo."
            ))

    # This app's local disk is temporary (a reboot or redeploy wipes it),
    # while the SharePoint copy persists. If there's no local file yet,
    # pull down whatever's already in SharePoint first, so new entries get
    # appended to it instead of silently starting over and overwriting it.
    if ENABLE_SHAREPOINT_SYNC and not os.path.exists(output_xlsx_path):
        try:
            found = sharepoint_client.download_file(remote_filename, output_xlsx_path, subfolder)
            if found:
                results.append((
                    "status",
                    "Found an existing master file in SharePoint and loaded it before continuing."
                ))
        except Exception:
            traceback.print_exc()
            results.append((
                "warning",
                "Could not check SharePoint for an existing master file before starting. "
                "If one already exists there, please verify afterward that no data was lost."
            ))

    # Photos are read by content, not filename, so the same photo is
    # recognized as a duplicate even if it was renamed or re-uploaded.
    processed_log = load_processed_log()
    seen_hashes = set()

    for i, uploaded_file in enumerate(uploaded_files):
        status_text.markdown(
            f'<p class="status-msg">Processing image {i + 1} of {total_files}: {uploaded_file.name}</p>',
            unsafe_allow_html=True
        )

        file_bytes = uploaded_file.getvalue()
        file_hash = hashlib.sha256(file_bytes).hexdigest()

        if file_hash in processed_log or file_hash in seen_hashes:
            results.append(("warning", f"Skipping '{uploaded_file.name}' - already processed."))
            skipped_count += 1
            progress_bar.progress((i + 1) / total_files)
            continue
        seen_hashes.add(file_hash)

        # A unique temp name avoids any clash with a leftover file from a
        # previous run under the same original filename.
        extension = os.path.splitext(uploaded_file.name)[1].lower() or ".jpg"
        temp_path = f"temp_{uuid.uuid4().hex}{extension}"
        with open(temp_path, "wb") as f:
            f.write(file_bytes)

        try:
            parsed_data = extract_data(temp_path)
            write_parsed_data_to_excel(parsed_data, temp_path, output_xlsx_path)

            processed_log.append(file_hash)
            save_processed_log(processed_log)
            processed_count += 1

            # Push the updated master file to SharePoint right away. This is
            # best-effort: the photo's data is already safely saved locally
            # above, so a SharePoint hiccup here never loses or blocks that.
            if ENABLE_SHAREPOINT_SYNC:
                try:
                    sharepoint_client.upload_file(output_xlsx_path, remote_filename, subfolder)
                except Exception:
                    traceback.print_exc()
                    results.append((
                        "warning",
                        f"'{uploaded_file.name}' was saved locally, but syncing to SharePoint failed "
                        "just now. It will be retried after the next photo; if it keeps failing, "
                        "download the file manually as a backup."
                    ))

            review_notes = parsed_data.get("Review Notes", [])
            flagged_materials = parsed_data.get("Flagged Materials", [])
            if review_notes:
                shown = "; ".join(review_notes[:3])
                extra = f" (+{len(review_notes) - 3} more)" if len(review_notes) > 3 else ""
                if flagged_materials:
                    results.append((
                        "warning",
                        f"'{uploaded_file.name}' was added to the Excel file, but its data row is "
                        f"highlighted in RED for manual validation: {shown}{extra}"
                    ))
                else:
                    results.append((
                        "warning",
                        f"'{uploaded_file.name}' was added to the Excel file. FYI only, nothing is "
                        f"highlighted red: {shown}{extra}"
                    ))

        except Exception as e:
            # Print the full traceback to the app's logs (Manage app > logs on
            # Streamlit Cloud), so the exact failing line can be found later.
            # The person only sees the short message below.
            traceback.print_exc()
            results.append(("error", f"An error occurred while processing {uploaded_file.name}: {e}"))
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

        progress_bar.progress((i + 1) / total_files)

    results.append(("status", "Batch Complete."))
    results.append(("success", f"Successfully added {processed_count} new file(s). Skipped {skipped_count} duplicate(s)."))

    # Save the messages, then clear the uploader automatically
    st.session_state.batch_results = results
    st.session_state.uploader_key += 1
    st.rerun()

current_batch_path = st.session_state.last_output_path
if current_batch_path and os.path.exists(current_batch_path):
    st.markdown("<br>", unsafe_allow_html=True)
    folder_note = f' in folder "{st.session_state.last_output_folder}"' if st.session_state.last_output_folder else ""
    st.subheader(f'File Management — "{st.session_state.last_output_label}"{folder_note}')

    col1, col2 = st.columns(2)

    with col1:
        xlsx_bytes = None
        try:
            with open(current_batch_path, "rb") as file:
                xlsx_bytes = file.read()
        except FileNotFoundError:
            pass

        if xlsx_bytes is not None:
            st.download_button(
                label=f"Download {os.path.basename(current_batch_path)}",
                data=xlsx_bytes,
                file_name=os.path.basename(current_batch_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

    with col2:
        if st.button("Delete This Batch File", type="primary", key="start_fresh"):
            st.session_state.confirm_reset = True
            st.rerun()

    # Ask before deleting, since this removes that batch's local data
    if st.session_state.confirm_reset:
        st.warning(
            f'This will permanently delete the local copy of "{os.path.basename(current_batch_path)}". It does not '
            "touch the processed-photos duplicate log, and it does not delete any copy already "
            "synced to SharePoint, you can remove that there directly if needed. Are you sure?"
        )
        confirm_col1, confirm_col2 = st.columns(2)
        with confirm_col1:
            if st.button("Yes, delete this file", type="primary", key="confirm_delete"):
                if os.path.exists(current_batch_path):
                    os.remove(current_batch_path)
                st.session_state.last_output_path = None
                st.session_state.last_output_label = None
                st.session_state.last_output_folder = ""
                st.session_state.confirm_reset = False
                st.rerun()
        with confirm_col2:
            if st.button("Cancel", type="primary", key="cancel_delete"):
                st.session_state.confirm_reset = False
                st.rerun()
