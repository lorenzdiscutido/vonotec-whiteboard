import os
import json
import uuid
import hashlib
import threading
import datetime
import base64
from concurrent.futures import ThreadPoolExecutor

import google.generativeai as genai
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

from config import REFERENCE_DATA, MATERIAL_RULES
from image_utils import load_image, prepare_excel_image
from gemini_client import get_raw_response, parse_and_clean_json

LOG_FILE = "processed_log.json"

# How many photos of ONE batch are sent to Gemini at the same time
MAX_PARALLEL_EXTRACTIONS = 3
# Max Gemini calls running at the same time across ALL users (avoids rate limits)
MAX_TOTAL_GEMINI_CALLS = 4


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


@st.cache_resource
def get_write_lock():
    """One lock shared by every user session: only one person writes to the
    master Excel file / log at a time."""
    return threading.Lock()


@st.cache_resource
def get_gemini_gate():
    """Caps how many Gemini calls run at once across all users."""
    return threading.Semaphore(MAX_TOTAL_GEMINI_CALLS)


def _populate_reference_sheet(wb):
    ws_ref = wb.create_sheet(title="Reference Data")
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for row_idx, row_values in enumerate(REFERENCE_DATA, start=1):
        ws_ref.append(row_values)
        for col_idx in range(1, 6):
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
    ws_ref.row_dimensions[1].height = 28


def extract_data(filepath, gemini_gate):
    """Runs in a background thread. Only talks to Gemini (no Streamlit calls here)."""
    img = load_image(filepath)
    with gemini_gate:
        raw_text = get_raw_response(img)
    return parse_and_clean_json(raw_text)


def write_parsed_data_to_excel(parsed_data, filepath, output_xlsx_path):
    """Adds one whiteboard's data to the master Excel file.
    Must be called while holding the write lock."""
    if os.path.exists(output_xlsx_path):
        wb = openpyxl.load_workbook(output_xlsx_path)
        ws = wb["Whiteboard Data"] if "Whiteboard Data" in wb.sheetnames else wb.active
        if "Reference Data" not in wb.sheetnames:
            _populate_reference_sheet(wb)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Whiteboard Data"
        
        ws.append([
            "Submitter", "Date", "Elevation", "Drop", "Floor", "Tower",
            "Defects", "", "Whiteboard Photo",
            "POSSIBLE CAUSE", "Possible Cause Justification", "Recommended Repair", "FINDINGS"
        ])
        ws.append([
            "", "", "", "", "", "",
            "Damage", "Dimension", "",
            "", "", "", ""
        ])

        ws.merge_cells('G1:H1') 
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'I', 'J', 'K', 'L', 'M']:
            ws.merge_cells(f'{col}1:{col}2')
        
        gray_fill = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
        for row in ws['A1':'M2']:
            for cell in row:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        ws['M1'].fill = gray_fill
        ws['M2'].fill = gray_fill

        ws.column_dimensions['J'].width = 24
        ws.column_dimensions['K'].width = 30
        ws.column_dimensions['L'].width = 28
        ws.column_dimensions['M'].width = 22

        _populate_reference_sheet(wb)

    start_row = ws.max_row + 1
    VALID_CODES = {row[0] for row in REFERENCE_DATA[1:]}

    rows_data = []
    for mat in ["Sealant", "Concrete", "Paint", "Gasket"]:
        rows_data.append((mat, "", True, False, mat)) 
        
        dmg_list = parsed_data.get(f"{mat} Damage", [])
        dim_list = parsed_data.get(f"{mat} Dimension", [])
        
        expanded_dmg = []
        expanded_dim = []
        
        for i in range(max(len(dmg_list), len(dim_list))):
            dmg_str = dmg_list[i] if i < len(dmg_list) else ""
            dim_str = dim_list[i] if i < len(dim_list) else ""
            
            tokens = str(dmg_str).split()
            found_codes = [t for t in tokens if t in VALID_CODES]
            
            if len(found_codes) > 1:
                for code in found_codes:
                    expanded_dmg.append(code)
                    expanded_dim.append(dim_str) 
            else:
                expanded_dmg.append(dmg_str)
                expanded_dim.append(dim_str)
        
        max_len = max(len(expanded_dmg), 1) 
        for i in range(max_len):
            dmg_val = expanded_dmg[i] if i < len(expanded_dmg) else ""
            dim_val = expanded_dim[i] if i < len(expanded_dim) else ""
            rows_data.append((dmg_val, dim_val, False, True, mat)) 

    total_rows = len(rows_data)

    static_keys = ["Submitter", "Date", "Elevation", "Drop", "Floor", "Tower"]
    for i, key in enumerate(static_keys):
        ws.cell(row=start_row, column=i+1).value = parsed_data.get(key, "")
        if total_rows > 1:
            ws.merge_cells(start_row=start_row, start_column=i+1, end_row=start_row + total_rows - 1, end_column=i+1)

    for i, (g_val, h_val, is_title, is_data, current_mat) in enumerate(rows_data):
        r = start_row + i
        cell_g = ws.cell(row=r, column=7)
        cell_h = ws.cell(row=r, column=8)
        
        cell_g.value = g_val
        cell_h.value = h_val

        is_invalid = False
        if is_data and g_val:
            base_code = str(g_val).split()[0]
            allowed_codes = MATERIAL_RULES.get(current_mat, [])
            if base_code not in allowed_codes:
                is_invalid = True

        if is_title:
            ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
            cell_g.font = Font(bold=True)
        elif is_data:
            lookup_val = f'LEFT($G{r}, FIND(" ", $G{r}&" ") - 1)'
            ws.cell(row=r, column=10).value = f'=IFERROR(VLOOKUP({lookup_val}, \'Reference Data\'!$A$2:$E$15, 3, FALSE), "")'
            ws.cell(row=r, column=11).value = f'=IFERROR(VLOOKUP({lookup_val}, \'Reference Data\'!$A$2:$E$15, 4, FALSE), "")'
            ws.cell(row=r, column=12).value = f'=IFERROR(VLOOKUP({lookup_val}, \'Reference Data\'!$A$2:$E$15, 5, FALSE), "")'
            ws.cell(row=r, column=13).value = f'=IFERROR(VLOOKUP({lookup_val}, \'Reference Data\'!$A$2:$E$15, 2, FALSE), "")'

            if is_invalid:
                for col_idx in range(7, 14):
                    ws.cell(row=r, column=col_idx).font = Font(color="FF0000")

    photo_col_letter = 'I' 
    excel_img = prepare_excel_image(filepath, total_rows)
    ws.add_image(excel_img, f"{photo_col_letter}{start_row}")
    
    if total_rows > 1:
        ws.merge_cells(f"{photo_col_letter}{start_row}:{photo_col_letter}{start_row + total_rows - 1}")
    
    base_height = max(110 / total_rows, 20)
    for i in range(total_rows):
        r = start_row + i
        ws.row_dimensions[r].height = base_height
        for c in range(1, 14):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
    ws.column_dimensions[photo_col_letter].width = 25

    # Save to a temp file first, then swap it in, so a crash mid-save
    # can never leave a half-written master file.
    tmp_path = output_xlsx_path.replace(".xlsx", "_tmp.xlsx")
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
    div.stButton > button:first-child p,
    .stDownloadButton > button:first-child p {
        color: #ffffff !important;
        font-weight: 700 !important;
    }
    div.stButton > button:first-child:hover,
    .stDownloadButton > button:first-child:hover {
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
        Upload one or multiple whiteboard photos below. The AI will automatically extract the data and append it to the Master Excel file. You can also drag and drop images directly into the uploader.
    </div>
    """,
    unsafe_allow_html=True
)

st.subheader("Upload Whiteboard Photos")

current_month_year = datetime.datetime.now().strftime("%B_%Y")
output_xlsx_path = f"master_output_{current_month_year}.xlsx"

write_lock = get_write_lock()
gemini_gate = get_gemini_gate()

# Changing the uploader's key resets it, which clears all selected photos at once
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "batch_results" not in st.session_state:
    st.session_state.batch_results = []
if "confirm_reset" not in st.session_state:
    st.session_state.confirm_reset = False

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

    if extract_clicked:
        total_files = len(uploaded_files)
        progress_bar = st.progress(0)
        status_text = st.empty()
        results = []

        processed_count = 0
        skipped_count = 0
        finished_count = 0

        # ---- Step 1: fingerprint every photo and skip duplicates ----
        processed_log = load_processed_log()
        seen_hashes = set()
        pending = []
        for uploaded_file in uploaded_files:
            file_bytes = uploaded_file.getvalue()
            file_hash = hashlib.sha256(file_bytes).hexdigest()

            if file_hash in processed_log or file_hash in seen_hashes:
                results.append(("warning", f"Skipping '{uploaded_file.name}' - already processed."))
                skipped_count += 1
                finished_count += 1
                continue
            seen_hashes.add(file_hash)

            # Unique temp name so two users never overwrite each other's photo
            extension = os.path.splitext(uploaded_file.name)[1].lower() or ".jpg"
            temp_path = f"temp_{uuid.uuid4().hex}{extension}"
            with open(temp_path, "wb") as f:
                f.write(file_bytes)
            pending.append({"name": uploaded_file.name, "hash": file_hash, "temp_path": temp_path})

        progress_bar.progress(finished_count / total_files)

        # ---- Step 2: extract in parallel, write to Excel one at a time ----
        with ThreadPoolExecutor(max_workers=MAX_PARALLEL_EXTRACTIONS) as executor:
            jobs = [
                (item, executor.submit(extract_data, item["temp_path"], gemini_gate))
                for item in pending
            ]

            for item, future in jobs:
                status_text.markdown(
                    f'<p class="status-msg">Processing image {finished_count + 1} of {total_files}: {item["name"]}</p>',
                    unsafe_allow_html=True
                )
                try:
                    parsed_data = future.result()

                    with write_lock:
                        # Re-check inside the lock in case another user just added the same photo
                        current_log = load_processed_log()
                        if item["hash"] in current_log:
                            results.append(("warning", f"Skipping '{item['name']}' - already processed."))
                            skipped_count += 1
                        else:
                            write_parsed_data_to_excel(parsed_data, item["temp_path"], output_xlsx_path)
                            current_log.append(item["hash"])
                            save_processed_log(current_log)
                            processed_count += 1

                except Exception as e:
                    results.append(("error", f"An error occurred while processing {item['name']}: {e}"))
                finally:
                    if os.path.exists(item["temp_path"]):
                        os.remove(item["temp_path"])

                finished_count += 1
                progress_bar.progress(finished_count / total_files)

        results.append(("status", "Batch Complete."))
        results.append(("success", f"Successfully added {processed_count} new file(s). Skipped {skipped_count} duplicate(s)."))

        # Save the messages, then clear the uploader automatically
        st.session_state.batch_results = results
        st.session_state.uploader_key += 1
        st.rerun()

if os.path.exists(output_xlsx_path):
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("Master File Management")

    col1, col2 = st.columns(2)

    with col1:
        xlsx_bytes = None
        try:
            with open(output_xlsx_path, "rb") as file:
                xlsx_bytes = file.read()
        except FileNotFoundError:
            pass

        if xlsx_bytes is not None:
            st.download_button(
                label=f"Download {output_xlsx_path}",
                data=xlsx_bytes,
                file_name=output_xlsx_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
    with col2:
        if st.button("Start Fresh (Clear Data)", type="primary", key="start_fresh"):
            st.session_state.confirm_reset = True
            st.rerun()

    # Ask before deleting, because this wipes the file for everyone
    if st.session_state.confirm_reset:
        st.warning(
            "This will permanently delete the master Excel file and the processed-photos log "
            "for everyone using the app. Are you sure?"
        )
        confirm_col1, confirm_col2 = st.columns(2)
        with confirm_col1:
            if st.button("Yes, delete everything", type="primary", key="confirm_delete"):
                with write_lock:
                    if os.path.exists(output_xlsx_path):
                        os.remove(output_xlsx_path)
                    if os.path.exists(LOG_FILE):
                        os.remove(LOG_FILE)
                st.session_state.batch_results = []
                st.session_state.confirm_reset = False
                st.rerun()
        with confirm_col2:
            if st.button("Cancel", type="primary", key="cancel_delete"):
                st.session_state.confirm_reset = False
                st.rerun()
