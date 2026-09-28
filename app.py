import os
import json
import datetime
import base64
import google.generativeai as genai
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

from config import INCOMING_FOLDER, REFERENCE_DATA, MATERIAL_RULES
from image_utils import load_image, prepare_excel_image
from gemini_client import get_raw_response, parse_and_clean_json

# --- UI CONFIGURATION & STYLING ---
st.set_page_config(page_title="Vonotec Whiteboard Extractor", layout="wide")

# Custom CSS to fix readability and button layout
st.markdown("""
    <style>
    /* 1. Set a clean, professional background for the whole app */
    .stApp {
        background-color: #f4f7f9;
    }

    /* 2. Fix the main header area */
    .header-container {
        background-color: #ffffff;
        padding: 20px 40px;
        border-radius: 0 0 15px 15px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        margin-bottom: 30px;
    }
    .main-title {
        color: #1E3A8A !important;
        margin-bottom: 0px !important;
        font-weight: 800;
    }
    .sub-title {
        color: #475569 !important;
        font-size: 1.1rem;
        margin-top: 5px;
    }

    /* 3. Solid Blue Button with Orange Text */
    div.stButton > button:first-child, .stDownloadButton > button:first-child {
        background-color: #2F5597 !important;
        color: #FFA500 !important;
        border: none !important;
        font-weight: bold !important;
        padding: 0.75rem 2rem !important;
        border-radius: 5px !important;
        width: 100%;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    /* 4. Instructions Box (High Contrast) */
    .instruction-box {
        background-color: #ffffff;
        border-left: 5px solid #2F5597;
        padding: 20px;
        border-radius: 5px;
        margin-bottom: 25px;
        color: #1e293b;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }

    /* 5. Section Headers */
    h3 {
        color: #1E3A8A !important;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 10px;
        margin-top: 30px !important;
    }
    
    /* Remove padding at the top of the block container */
    .block-container {
        padding-top: 0rem !important;
    }
    </style>
    """, unsafe_allow_html=True)

LOG_FILE = "processed_log.json"

def load_processed_log():
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, "r") as f:
            return json.load(f)
    return []

def save_processed_log(log_list):
    with open(LOG_FILE, "w") as f:
        json.dump(log_list, f)

genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

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

def process_image_to_excel(filepath, output_xlsx_path):
    img = load_image(filepath)
    raw_text = get_raw_response(img)
    parsed_data = parse_and_clean_json(raw_text)

    if os.path.exists(output_xlsx_path):
        wb = openpyxl.load_workbook(output_xlsx_path)
        ws = wb["Whiteboard Data"] if "Whiteboard Data" in wb.sheetnames else wb.active
        if "Reference Data" not in wb.sheetnames:
            _populate_reference_sheet(wb)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Whiteboard Data"
        ws.append(["Submitter", "Date", "Elevation", "Drop", "Floor", "Tower", "Defects", "", "Whiteboard Photo", "POSSIBLE CAUSE", "Possible Cause Justification", "Recommended Repair", "FINDINGS"])
        ws.append(["", "", "", "", "", "", "Damage", "Dimension", "", "", "", "", ""])
        ws.merge_cells('G1:H1') 
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'I', 'J', 'K', 'L', 'M']:
            ws.merge_cells(f'{col}1:{col}2')
        gray_fill = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
        for row in ws['A1':'M2']:
            for cell in row:
                cell.font = Font(bold=True); cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws['M1'].fill = gray_fill; ws['M2'].fill = gray_fill
        _populate_reference_sheet(wb)

    start_row = ws.max_row + 1
    VALID_CODES = {row[0] for row in REFERENCE_DATA[1:]}
    rows_data = []
    for mat in ["Sealant", "Concrete", "Paint", "Gasket"]:
        rows_data.append((mat, "", True, False, mat)) 
        dmg_list, dim_list = parsed_data.get(f"{mat} Damage", []), parsed_data.get(f"{mat} Dimension", [])
        expanded_dmg, expanded_dim = [], []
        for i in range(max(len(dmg_list), len(dim_list))):
            dmg_str, dim_str = dmg_list[i] if i < len(dmg_list) else "", dim_list[i] if i < len(dim_list) else ""
            found_codes = [t for t in str(dmg_str).split() if t in VALID_CODES]
            if len(found_codes) > 1:
                for code in found_codes: expanded_dmg.append(code); expanded_dim.append(dim_str)
            else: expanded_dmg.append(dmg_str); expanded_dim.append(dim_str)
        for i in range(max(len(expanded_dmg), 1)):
            rows_data.append((expanded_dmg[i] if i < len(expanded_dmg) else "", expanded_dim[i] if i < len(expanded_dim) else "", False, True, mat)) 

    total_rows = len(rows_data)
    static_keys = ["Submitter", "Date", "Elevation", "Drop", "Floor", "Tower"]
    for i, key in enumerate(static_keys):
        ws.cell(row=start_row, column=i+1).value = parsed_data.get(key, "")
        if total_rows > 1: ws.merge_cells(start_row=start_row, start_column=i+1, end_row=start_row + total_rows - 1, end_column=i+1)

    for i, (g_val, h_val, is_title, is_data, current_mat) in enumerate(rows_data):
        r = start_row + i
        ws.cell(row=r, column=7).value = g_val
        ws.cell(row=r, column=8).value = h_val
        if is_title:
            ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
            ws.cell(row=r, column=7).font = Font(bold=True)
        elif is_data:
            lookup_val = f'LEFT($G{r}, FIND(" ", $G{r}&" ") - 1)'
            for col, idx in [ (10, 3), (11, 4), (12, 5), (13, 2) ]:
                ws.cell(row=r, column=col).value = f'=IFERROR(VLOOKUP({lookup_val}, \'Reference Data\'!$A$2:$E$15, {idx}, FALSE), "")'

    excel_img = prepare_excel_image(filepath, total_rows)
    ws.add_image(excel_img, f"I{start_row}")
    if total_rows > 1: ws.merge_cells(f"I{start_row}:I{start_row + total_rows - 1}")
    wb.save(output_xlsx_path)

# --- UI CONTENT ---
# Custom Header with white background for readability
try:
    with open("vonotec.png", "rb") as f:
        logo_base64 = base64.b64encode(f.read()).decode()
    st.markdown(f"""
        <div class="header-container">
            <div style="display: flex; align-items: center; gap: 30px;">
                <img src="data:image/png;base64,{logo_base64}" width="180">
                <div>
                    <h1 class="main-title">Whiteboard AI</h1>
                    <p class="sub-title">Data Extraction and Master Log Automator</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
except FileNotFoundError:
    st.title("Vonotec Whiteboard AI")

# Instructions Container
st.markdown("""
    <div class="instruction-box">
        <strong>Instructions:</strong> Upload whiteboard photos below. 
        The system will process each image, extract the data, and update the Master Excel file automatically.
    </div>
    """, unsafe_allow_html=True)

st.subheader("Upload Whiteboard Photos")
uploaded_files = st.file_uploader("Select JPG or PNG files", type=["jpg", "jpeg", "png"], accept_multiple_files=True, label_visibility="collapsed")

current_month_year = datetime.datetime.now().strftime("%B_%Y") 
output_xlsx_path = f"master_output_{current_month_year}.xlsx"

if uploaded_files:
    if st.button("Process and Update Excel"):
        processed_log = load_processed_log()
        progress_bar = st.progress(0)
        for i, uploaded_file in enumerate(uploaded_files):
            if uploaded_file.name in processed_log:
                st.info(f"Skipped {uploaded_file.name} (Already Processed)")
                continue
            
            temp_path = f"temp_{uploaded_file.name}"
            with open(temp_path, "wb") as f: f.write(uploaded_file.getbuffer())
            
            try:
                process_image_to_excel(temp_path, output_xlsx_path)
                processed_log.append(uploaded_file.name)
                save_processed_log(processed_log)
            except Exception as e:
                st.error(f"Error processing {uploaded_file.name}: {e}")
            finally:
                if os.path.exists(temp_path): os.remove(temp_path)
            progress_bar.progress((i + 1) / len(uploaded_files))
        st.success("Batch Processing Complete")

if os.path.exists(output_xlsx_path):
    st.subheader("Master File Management")
    col1, col2 = st.columns(2)
    
    with col1:
        with open(output_xlsx_path, "rb") as file:
            st.download_button(
                label=f"Download {output_xlsx_path}",
                data=file,
                file_name=output_xlsx_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    with col2:
        if st.button("Start Fresh"):
            if os.path.exists(output_xlsx_path): os.remove(output_xlsx_path)
            if os.path.exists(LOG_FILE): os.remove(LOG_FILE)
            st.rerun()
