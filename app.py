import io
import json
import os
import uuid
from datetime import datetime
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# Configuration & Theme Setup
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Stock Check — Added vs Return vs Hold",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DATA_FILE = "stock_check.json"

# -----------------------------------------------------------------------------
# Data Persistence
# -----------------------------------------------------------------------------
def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return {
                        "unassigned": data.get("unassigned", []),
                        "added": data.get("added", []),
                        "returns": data.get("returns", []),
                        "hold": data.get("hold", []),
                    }
        except Exception:
            pass
    return {"unassigned": [], "added": [], "returns": [], "hold": []}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

# Initialize Session State
if "stock_data" not in st.session_state:
    st.session_state.stock_data = load_data()

if "manual_form_version" not in st.session_state:
    st.session_state.manual_form_version = 0

if "pending_return_item" not in st.session_state:
    st.session_state.pending_return_item = None

# -----------------------------------------------------------------------------
# Column Auto-Detection Logic for Excel Files
# -----------------------------------------------------------------------------
PART_NO_CANDIDATES = [
    "partno", "partnumber", "partnum", "itemcode", "partcode",
    "part", "code", "sku", "itemno", "itemnumber", "material", "articleno"
]
PART_NAME_CANDIDATES = [
    "partname", "partdescription", "description", "itemname",
    "item", "product", "productname", "name", "desc", "title", "particulars"
]
QTY_CANDIDATES = [
    "qty", "quantity", "count", "units", "pcs", "pieces", "stockqty", "totalqty"
]

def normalize_col(col):
    return "".join(c for c in str(col).lower() if c.isalnum())

def detect_columns(columns):
    norm_map = {normalize_col(c): c for c in columns}

    detected_pno = None
    for cand in PART_NO_CANDIDATES:
        for norm, orig in norm_map.items():
            if cand == norm or cand in norm:
                detected_pno = orig
                break
        if detected_pno:
            break

    detected_pname = None
    for cand in PART_NAME_CANDIDATES:
        for norm, orig in norm_map.items():
            if orig == detected_pno:
                continue
            if cand == norm or cand in norm:
                detected_pname = orig
                break
        if detected_pname:
            break

    detected_qty = None
    for cand in QTY_CANDIDATES:
        for norm, orig in norm_map.items():
            if orig in (detected_pno, detected_pname):
                continue
            if cand == norm or cand in norm:
                detected_qty = orig
                break
        if detected_qty:
            break

    return detected_pno, detected_pname, detected_qty

def generate_template_excel():
    df = pd.DataFrame([
        {"Part No": "15400-PLM-A02", "Part Name": "Engine Oil Filter", "Qty": 10},
        {"Part No": "45022-S9A-E50", "Part Name": "Front Brake Pad Set", "Qty": 4},
        {"Part No": "12290-R40-A01", "Part Name": "Spark Plug Laser Iridium", "Qty": 8},
        {"Part No": "17220-R6A-J00", "Part Name": "Engine Air Filter", "Qty": 5},
    ])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="StockSheet")
    output.seek(0)
    return output.getvalue()

# -----------------------------------------------------------------------------
# Premium Custom Styling & Dynamic Scripting
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap');

    * {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }

    /* Overall background: Deep Navy */
    .stApp {
        background: radial-gradient(circle at 50% 0%, #111E48 0%, #080E24 60%, #050816 100%);
        color: #F8FAFC;
    }

    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 2.5rem !important;
        max-width: 1280px !important;
    }

    /* Top Bar Styling */
    .top-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 8px 0 16px 0;
        border-bottom: 2px solid rgba(212, 175, 55, 0.28);
        margin-bottom: 20px;
    }

    .brand-title {
        color: #FFFFFF;
        font-size: 28px;
        font-weight: 900;
        letter-spacing: -0.5px;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 12px;
        flex-wrap: wrap;
    }

    .gold-pill {
        background: linear-gradient(135deg, #F3C64F 0%, #C59B27 100%);
        color: #070D1E;
        font-size: 13px;
        font-weight: 800;
        letter-spacing: 0.8px;
        padding: 4px 12px;
        border-radius: 999px;
        text-transform: uppercase;
        box-shadow: 0 2px 8px rgba(243, 198, 79, 0.35);
    }

    .brand-sub {
        color: #94A3B8;
        font-size: 14px;
        font-weight: 500;
        margin-top: 4px;
    }

    /* BIG Summary Cards (Shop counter visibility from distance) */
    .summary-card {
        background: #FFFFFF;
        border-radius: 18px;
        padding: 22px 24px;
        box-shadow: 0 12px 30px rgba(0, 0, 0, 0.4), 0 2px 4px rgba(0, 0, 0, 0.2);
        color: #0F172A;
        position: relative;
        min-height: 145px;
    }

    .card-added {
        border-top: 9px solid #059669;
    }

    .card-return {
        border-top: 9px solid #DC2626;
    }

    .card-hold {
        border-top: 9px solid #64748B;
    }

    .card-badge {
        font-size: 14px;
        font-weight: 800;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .badge-added { color: #047857; }
    .badge-return { color: #B91C1C; }
    .badge-hold { color: #475569; }

    .stat-flex {
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
    }

    .stat-big-num {
        font-size: 50px;
        font-weight: 900;
        line-height: 1;
        letter-spacing: -1px;
    }

    .num-green { color: #047857; }
    .num-red { color: #B91C1C; }
    .num-grey { color: #334155; }

    .stat-label-muted {
        font-size: 12px;
        font-weight: 700;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-top: 6px;
    }

    .stat-box-side {
        text-align: right;
        background: #F8FAFC;
        padding: 8px 14px;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        min-width: 95px;
    }

    .sec-parts-num {
        font-size: 24px;
        font-weight: 900;
        color: #0F172A;
        line-height: 1.1;
    }

    .sec-parts-txt {
        font-size: 11px;
        font-weight: 700;
        color: #64748B;
        text-transform: uppercase;
    }

    /* Section White Card Boxes */
    .white-card-panel {
        background: #FFFFFF;
        border-radius: 18px;
        padding: 22px 24px;
        box-shadow: 0 10px 25px rgba(0, 0, 0, 0.35);
        color: #0F172A;
        margin-bottom: 20px;
    }

    .white-card-gold-border {
        border-left: 8px solid #F3C64F;
    }

    .white-card-blue-border {
        border-left: 8px solid #3B82F6;
    }

    .panel-title {
        font-size: 17px;
        font-weight: 800;
        color: #070D1E;
        letter-spacing: 0.3px;
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 14px;
    }

    /* Unassigned Sorting Table Rows */
    .unassigned-row {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 14px 16px;
        margin-bottom: 12px;
    }

    .part-title-text {
        font-size: 16px;
        font-weight: 800;
        color: #070D1E;
    }

    .part-desc-text {
        font-size: 14px;
        color: #475569;
        font-weight: 500;
        margin-top: 2px;
    }

    .part-qty-pill {
        background: #0B132B;
        color: #FFFFFF;
        font-weight: 800;
        font-size: 13px;
        padding: 4px 10px;
        border-radius: 6px;
        display: inline-block;
        margin-top: 4px;
    }

    /* Columns for Three Lists */
    .list-card-box {
        background: #FFFFFF;
        border-radius: 18px;
        padding: 20px 20px;
        box-shadow: 0 10px 25px rgba(0, 0, 0, 0.35);
        color: #0F172A;
        min-height: 280px;
    }

    .list-card-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding-bottom: 12px;
        border-bottom: 2px solid #F1F5F9;
        margin-bottom: 14px;
    }

    .list-title {
        font-size: 16px;
        font-weight: 800;
        color: #070D1E;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .list-count-pill {
        font-size: 11px;
        font-weight: 800;
        padding: 3px 8px;
        border-radius: 6px;
    }

    .pill-added-count { background: #ECFDF5; color: #047857; border: 1px solid #A7F3D0; }
    .pill-return-count { background: #FEF2F2; color: #B91C1C; border: 1px solid #FECACA; }
    .pill-hold-count { background: #F1F5F9; color: #475569; border: 1px solid #CBD5E1; }

    .part-entry-row {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 10px 12px;
        margin-bottom: 8px;
    }

    .reason-wrong-part { background: #FEE2E2; color: #991B1B; }
    .reason-damaged { background: #FEF3C7; color: #92400E; }
    .reason-extra { background: #E0E7FF; color: #3730A3; }

    .badge-reason-tag {
        display: inline-block;
        font-size: 11px;
        font-weight: 800;
        padding: 2px 7px;
        border-radius: 6px;
        margin-top: 5px;
        margin-left: 6px;
        text-transform: uppercase;
    }

    .empty-prompt {
        text-align: center;
        padding: 34px 12px;
        color: #64748B;
        font-size: 13px;
        font-weight: 500;
        border: 2px dashed #CBD5E1;
        border-radius: 10px;
        background: #F8FAFC;
        margin-top: 8px;
    }

    /* Big buttons and inputs */
    div.stButton > button {
        border-radius: 10px !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        padding: 10px 16px !important;
        min-height: 44px !important;
        transition: transform 0.1s ease, box-shadow 0.1s ease !important;
    }

    div.stButton > button:active {
        transform: scale(0.98) !important;
    }

    /* Delete button */
    .delete-btn-wrap button {
        background-color: #FEE2E2 !important;
        color: #DC2626 !important;
        border: 1px solid #FCA5A5 !important;
        border-radius: 8px !important;
        padding: 4px 8px !important;
        min-height: 38px !important;
        width: 100% !important;
        font-size: 16px !important;
    }

    .delete-btn-wrap button:hover {
        background-color: #DC2626 !important;
        color: #FFFFFF !important;
    }

    /* Target Stock / Return / Hold buttons by custom class */
    .btn-stock-target button {
        background-color: #059669 !important;
        color: #FFFFFF !important;
        border: 1px solid #047857 !important;
    }
    .btn-stock-target button:hover {
        background-color: #047857 !important;
    }

    .btn-return-target button {
        background-color: #DC2626 !important;
        color: #FFFFFF !important;
        border: 1px solid #B91C1C !important;
    }
    .btn-return-target button:hover {
        background-color: #B91C1C !important;
    }

    .btn-hold-target button {
        background-color: #64748B !important;
        color: #FFFFFF !important;
        border: 1px solid #475569 !important;
    }
    .btn-hold-target button:hover {
        background-color: #475569 !important;
    }

    /* Hide standard Streamlit header and footer */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Confirmation Dialogs
# -----------------------------------------------------------------------------
@st.dialog("⚠️ Start New Stock Check?")
def confirm_new_check_dialog():
    st.markdown(
        """
        <div style="font-size: 15px; color: #1E293B; line-height: 1.5; margin-bottom: 20px;">
            Are you sure you want to clear everything and start a fresh check?
            <br><br>
            <span style="color: #DC2626; font-weight: 700;">
            This will permanently erase all unassigned items, Stock Added, To Return, and Hold records.
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    col_confirm, col_cancel = st.columns(2)
    with col_confirm:
        if st.button("🗑️ Yes, Clear Everything", type="primary", use_container_width=True):
            st.session_state.stock_data = {
                "unassigned": [],
                "added": [],
                "returns": [],
                "hold": [],
            }
            save_data(st.session_state.stock_data)
            st.rerun()
    with col_cancel:
        if st.button("Cancel", use_container_width=True):
            st.rerun()

@st.dialog("↩️ Return Reason")
def return_reason_dialog(item):
    st.markdown(
        f"""
        <div style="font-size: 15px; color: #1E293B; line-height: 1.5; margin-bottom: 16px;">
            Select return reason for:<br>
            <strong># {item.get('part_no', '—')}</strong> — {item.get('part_name', '—')}
            <span style="background: #0B132B; color: #fff; padding: 2px 8px; border-radius: 4px; font-size: 12px; margin-left: 6px;">Qty: {item.get('qty', 1)}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    r1, r2, r3 = st.columns(3)
    with r1:
        if st.button("❌ Wrong Part", use_container_width=True):
            item["reason"] = "wrong part"
            st.session_state.stock_data["unassigned"] = [
                x for x in st.session_state.stock_data["unassigned"] if x["id"] != item["id"]
            ]
            st.session_state.stock_data["returns"].insert(0, item)
            save_data(st.session_state.stock_data)
            st.rerun()
    with r2:
        if st.button("💥 Damaged", use_container_width=True):
            item["reason"] = "damaged"
            st.session_state.stock_data["unassigned"] = [
                x for x in st.session_state.stock_data["unassigned"] if x["id"] != item["id"]
            ]
            st.session_state.stock_data["returns"].insert(0, item)
            save_data(st.session_state.stock_data)
            st.rerun()
    with r3:
        if st.button("📦 Extra", use_container_width=True):
            item["reason"] = "extra"
            st.session_state.stock_data["unassigned"] = [
                x for x in st.session_state.stock_data["unassigned"] if x["id"] != item["id"]
            ]
            st.session_state.stock_data["returns"].insert(0, item)
            save_data(st.session_state.stock_data)
            st.rerun()

# -----------------------------------------------------------------------------
# Top Bar: Title + New Check Button
# -----------------------------------------------------------------------------
header_col1, header_col2 = st.columns([4, 1.2])

with header_col1:
    st.markdown(
        """
        <div class="top-header">
            <div>
                <div class="brand-title">
                    📦 STOCK CHECK
                    <span class="gold-pill">Stock • Return • Hold</span>
                </div>
                <div class="brand-sub">Import any Excel file or add manually • Tap Stock, Return, or Hold per item</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with header_col2:
    st.write("")
    if st.button("🔄 New Check", use_container_width=True, help="Clear everything and start fresh check"):
        confirm_new_check_dialog()

# -----------------------------------------------------------------------------
# Calculate Totals
# -----------------------------------------------------------------------------
added_list = st.session_state.stock_data.get("added", [])
returns_list = st.session_state.stock_data.get("returns", [])
hold_list = st.session_state.stock_data.get("hold", [])
unassigned_list = st.session_state.stock_data.get("unassigned", [])

total_added_parts = len(added_list)
total_added_qty = sum(item.get("qty", 0) for item in added_list)

total_return_parts = len(returns_list)
total_return_qty = sum(item.get("qty", 0) for item in returns_list)

total_hold_parts = len(hold_list)
total_hold_qty = sum(item.get("qty", 0) for item in hold_list)

# -----------------------------------------------------------------------------
# THREE BIG SUMMARY CARDS ON TOP
# -----------------------------------------------------------------------------
card_col1, card_col2, card_col3 = st.columns(3)

with card_col1:
    st.markdown(
        f"""
        <div class="summary-card card-added">
            <div class="card-badge badge-added">
                <span>🟢</span> STOCK ADDED
            </div>
            <div class="stat-flex">
                <div>
                    <div class="stat-big-num num-green">{total_added_qty}</div>
                    <div class="stat-label-muted">Total Units Added</div>
                </div>
                <div class="stat-box-side">
                    <div class="sec-parts-num">{total_added_parts}</div>
                    <div class="sec-parts-txt">Items Added</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with card_col2:
    st.markdown(
        f"""
        <div class="summary-card card-return">
            <div class="card-badge badge-return">
                <span>🔴</span> TO RETURN
            </div>
            <div class="stat-flex">
                <div>
                    <div class="stat-big-num num-red">{total_return_qty}</div>
                    <div class="stat-label-muted">Total Units To Return</div>
                </div>
                <div class="stat-box-side">
                    <div class="sec-parts-num">{total_return_parts}</div>
                    <div class="sec-parts-txt">Items To Return</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with card_col3:
    st.markdown(
        f"""
        <div class="summary-card card-hold">
            <div class="card-badge badge-hold">
                <span>⚪</span> HOLD
            </div>
            <div class="stat-flex">
                <div>
                    <div class="stat-big-num num-grey">{total_hold_qty}</div>
                    <div class="stat-label-muted">Total Units On Hold</div>
                </div>
                <div class="stat-box-side">
                    <div class="sec-parts-num">{total_hold_parts}</div>
                    <div class="sec-parts-txt">Items On Hold</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.write("")

# -----------------------------------------------------------------------------
# Section: Excel Import & Add One Item Form
# -----------------------------------------------------------------------------
top_action_c1, top_action_c2 = st.columns([1.6, 1.4])

# --- Excel Upload Box ---
with top_action_c1:
    st.markdown(
        """
        <div class="white-card-panel white-card-blue-border">
            <div class="panel-title">
                <span>📁</span> IMPORT EXCEL FILE (.xlsx)
            </div>
        """,
        unsafe_allow_html=True,
    )

    up_col1, up_col2 = st.columns([2.5, 1.2])
    with up_col1:
        uploaded_file = st.file_uploader(
            "Upload Excel (.xlsx)",
            type=["xlsx", "xls"],
            help="Upload ANY Excel file. Columns will be auto-detected.",
            label_visibility="collapsed",
        )
    with up_col2:
        template_bytes = generate_template_excel()
        st.download_button(
            label="📥 Template",
            data=template_bytes,
            file_name="stock_check_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Download standard sample Excel template",
        )

    if uploaded_file is not None:
        try:
            file_bytes = uploaded_file.read()
            excel_df = pd.read_excel(io.BytesIO(file_bytes))
            excel_df.dropna(how="all", inplace=True)
            cols = list(excel_df.columns)

            auto_pno, auto_pname, auto_qty = detect_columns(cols)

            # Check if columns matched automatically
            is_auto_matched = all([auto_pno, auto_pname, auto_qty])

            if is_auto_matched:
                st.success(
                    f"✅ **Auto-detected columns:** Part No: `{auto_pno}` | Part Name: `{auto_pname}` | Qty: `{auto_qty}`"
                )
                chosen_pno = auto_pno
                chosen_pname = auto_pname
                chosen_qty = auto_qty
            else:
                st.warning("⚠️ Could not auto-detect all columns. Please pick from the dropdowns below:")
                col_sel1, col_sel2, col_sel3 = st.columns(3)
                with col_sel1:
                    chosen_pno = st.selectbox(
                        "Part No Column",
                        options=cols,
                        index=cols.index(auto_pno) if auto_pno in cols else 0,
                    )
                with col_sel2:
                    pname_idx = cols.index(auto_pname) if auto_pname in cols else (1 if len(cols) > 1 else 0)
                    chosen_pname = st.selectbox(
                        "Part Name Column",
                        options=cols,
                        index=pname_idx,
                    )
                with col_sel3:
                    qty_idx = cols.index(auto_qty) if auto_qty in cols else (2 if len(cols) > 2 else 0)
                    chosen_qty = st.selectbox(
                        "Qty Column",
                        options=cols,
                        index=qty_idx,
                    )

            if st.button("🚀 Import Items to Queue", type="primary", use_container_width=True):
                imported_rows = []
                for _, r in excel_df.iterrows():
                    val_pno = str(r[chosen_pno]).strip() if pd.notna(r[chosen_pno]) else "—"
                    val_pname = str(r[chosen_pname]).strip() if pd.notna(r[chosen_pname]) else "—"
                    try:
                        val_qty = int(float(r[chosen_qty])) if pd.notna(r[chosen_qty]) else 1
                        if val_qty < 1:
                            val_qty = 1
                    except Exception:
                        val_qty = 1

                    if val_pno != "—" or val_pname != "—":
                        imported_rows.append({
                            "id": str(uuid.uuid4())[:8],
                            "part_no": val_pno,
                            "part_name": val_pname,
                            "qty": val_qty,
                            "time": datetime.now().strftime("%H:%M"),
                        })

                if imported_rows:
                    st.session_state.stock_data["unassigned"].extend(imported_rows)
                    save_data(st.session_state.stock_data)
                    st.toast(f"✅ Imported {len(imported_rows)} items to sort!", icon="📦")
                    st.rerun()
                else:
                    st.error("No valid data rows found in the uploaded file.")
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")

    st.markdown("</div>", unsafe_allow_html=True)

# --- Add Single Item Form ---
with top_action_c2:
    st.markdown(
        """
        <div class="white-card-panel white-card-gold-border">
            <div class="panel-title">
                <span>⚡</span> ADD ONE ITEM (MANUAL)
            </div>
        """,
        unsafe_allow_html=True,
    )

    m_ver = st.session_state.manual_form_version
    mc1, mc2, mc3 = st.columns([1.3, 1.8, 0.9])
    with mc1:
        man_part_no = st.text_input("Part No.", key=f"man_pno_{m_ver}", placeholder="e.g. 15400-PLM")
    with mc2:
        man_part_name = st.text_input("Part Name", key=f"man_pname_{m_ver}", placeholder="e.g. Oil Filter")
    with mc3:
        man_qty = st.number_input("Qty", min_value=1, step=1, value=1, key=f"man_qty_{m_ver}")

    act_col, opt_col = st.columns([1.6, 1.4])
    with act_col:
        man_action = st.radio(
            "Destination",
            options=["Stock", "Return", "Hold"],
            horizontal=True,
            key=f"man_action_{m_ver}",
            label_visibility="collapsed",
        )

    with opt_col:
        if man_action == "Return":
            man_reason = st.selectbox(
                "Return Reason",
                options=["wrong part", "damaged", "extra"],
                key=f"man_reason_{m_ver}",
                label_visibility="collapsed",
            )
        else:
            man_reason = None
            st.markdown(
                f"<div style='color: #64748B; font-size: 13px; font-weight: 600; padding-top: 6px;'>Action: {man_action}</div>",
                unsafe_allow_html=True,
            )

    if st.button("➕ Add Item", use_container_width=True, type="primary" if man_action == "Stock" else "secondary"):
        clean_no = man_part_no.strip()
        clean_name = man_part_name.strip()

        if not clean_no and not clean_name:
            st.toast("⚠️ Enter at least Part No. or Part Name", icon="⚠️")
        else:
            new_item = {
                "id": str(uuid.uuid4())[:8],
                "part_no": clean_no if clean_no else "—",
                "part_name": clean_name if clean_name else "—",
                "qty": int(man_qty),
                "time": datetime.now().strftime("%H:%M"),
            }
            if man_action == "Stock":
                st.session_state.stock_data["added"].insert(0, new_item)
                st.toast(f"✅ Added {new_item['qty']}x {new_item['part_no']} to Stock", icon="📦")
            elif man_action == "Return":
                new_item["reason"] = man_reason if man_reason else "wrong part"
                st.session_state.stock_data["returns"].insert(0, new_item)
                st.toast(f"🚨 Recorded {new_item['qty']}x {new_item['part_no']} to Return", icon="↩️")
            else:  # Hold
                st.session_state.stock_data["hold"].insert(0, new_item)
                st.toast(f"⚪ Held {new_item['qty']}x {new_item['part_no']} for review", icon="⏸️")

            save_data(st.session_state.stock_data)
            st.session_state.manual_form_version += 1
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Section: Incoming Items Table (ALL Items Appear Here with THREE Buttons)
# -----------------------------------------------------------------------------
if unassigned_list:
    st.markdown(
        f"""
        <div class="white-card-panel" style="border-left: 8px solid #059669; margin-bottom: 24px;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div class="panel-title" style="margin-bottom: 0;">
                    <span>📋</span> ITEMS TO SORT
                    <span style="font-size: 13px; font-weight: 800; background: #FEF3C7; color: #92400E; padding: 4px 10px; border-radius: 8px; margin-left: 8px;">
                        {len(unassigned_list)} items waiting for decision
                    </span>
                </div>
                <div style="color: #64748B; font-size: 13px; font-weight: 600;">
                    Tap <strong>Stock</strong>, <strong>Return</strong>, or <strong>Hold</strong> per item
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    for item in unassigned_list:
        row_c1, row_c2, row_c3, row_c4, row_c5 = st.columns([3.5, 1.2, 1.2, 1.2, 0.6])

        with row_c1:
            st.markdown(
                f"""
                <div class="unassigned-row" style="margin-bottom: 0px;">
                    <div class="part-title-text"># {item.get('part_no', '—')}</div>
                    <div class="part-desc-text">{item.get('part_name', '—')}</div>
                    <span class="part-qty-pill">Qty: {item.get('qty', 1)}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with row_c2:
            st.markdown('<div class="btn-stock-target">', unsafe_allow_html=True)
            if st.button("🟢 Stock", key=f"btn_stock_{item['id']}", use_container_width=True, help="Add to stock"):
                st.session_state.stock_data["unassigned"] = [
                    x for x in st.session_state.stock_data["unassigned"] if x["id"] != item["id"]
                ]
                st.session_state.stock_data["added"].insert(0, item)
                save_data(st.session_state.stock_data)
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        with row_c3:
            st.markdown('<div class="btn-return-target">', unsafe_allow_html=True)
            if st.button("🔴 Return", key=f"btn_return_{item['id']}", use_container_width=True, help="Return item"):
                return_reason_dialog(item)
            st.markdown('</div>', unsafe_allow_html=True)

        with row_c4:
            st.markdown('<div class="btn-hold-target">', unsafe_allow_html=True)
            if st.button("⚪ Hold", key=f"btn_hold_{item['id']}", use_container_width=True, help="Decide later"):
                st.session_state.stock_data["unassigned"] = [
                    x for x in st.session_state.stock_data["unassigned"] if x["id"] != item["id"]
                ]
                st.session_state.stock_data["hold"].insert(0, item)
                save_data(st.session_state.stock_data)
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        with row_c5:
            st.markdown('<div class="delete-btn-wrap">', unsafe_allow_html=True)
            if st.button("🗑️", key=f"btn_del_unasgn_{item['id']}", help="Remove from queue"):
                st.session_state.stock_data["unassigned"] = [
                    x for x in st.session_state.stock_data["unassigned"] if x["id"] != item["id"]
                ]
                save_data(st.session_state.stock_data)
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        st.write("")  # Row spacing

# -----------------------------------------------------------------------------
# THREE LISTS BELOW CARDS (Added, Return, Hold)
# -----------------------------------------------------------------------------
col_list_added, col_list_return, col_list_hold = st.columns(3)

# --- 1. Stock Added List ---
with col_list_added:
    st.markdown(
        f"""
        <div class="list-card-box">
            <div class="list-card-header">
                <div class="list-title">
                    <span style="color: #059669;">●</span> Stock Added List
                </div>
                <div class="list-count-pill pill-added-count">
                    {total_added_parts} parts • {total_added_qty} pcs
                </div>
            </div>
        """,
        unsafe_allow_html=True,
    )

    if not added_list:
        st.markdown(
            """
            <div class="empty-prompt">
                No parts added to stock yet.<br>Tap "Stock" on imported items or add manually.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        for item in added_list:
            r1, r2 = st.columns([4.8, 1.2])
            with r1:
                st.markdown(
                    f"""
                    <div class="part-entry-row">
                        <div class="part-title-text" style="font-size: 15px;"># {item.get('part_no', '—')}</div>
                        <div class="part-desc-text" style="font-size: 13px;">{item.get('part_name', '—')}</div>
                        <span class="part-qty-pill" style="font-size: 12px; padding: 2px 8px;">Qty: {item.get('qty', 0)}</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with r2:
                st.markdown('<div class="delete-btn-wrap">', unsafe_allow_html=True)
                if st.button("🗑️", key=f"del_add_{item['id']}", help="Delete item"):
                    st.session_state.stock_data["added"] = [
                        x for x in st.session_state.stock_data["added"] if x["id"] != item["id"]
                    ]
                    save_data(st.session_state.stock_data)
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

# --- 2. To Return List ---
with col_list_return:
    st.markdown(
        f"""
        <div class="list-card-box">
            <div class="list-card-header">
                <div class="list-title">
                    <span style="color: #DC2626;">●</span> To Return List
                </div>
                <div class="list-count-pill pill-return-count">
                    {total_return_parts} parts • {total_return_qty} pcs
                </div>
            </div>
        """,
        unsafe_allow_html=True,
    )

    if not returns_list:
        st.markdown(
            """
            <div class="empty-prompt">
                No items marked for return.<br>Tap "Return" on items to send back.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        for item in returns_list:
            reason_str = item.get("reason", "wrong part")
            r_class = "reason-wrong-part"
            if "damaged" in reason_str.lower():
                r_class = "reason-damaged"
            elif "extra" in reason_str.lower():
                r_class = "reason-extra"

            r1, r2 = st.columns([4.8, 1.2])
            with r1:
                st.markdown(
                    f"""
                    <div class="part-entry-row">
                        <div class="part-title-text" style="font-size: 15px;"># {item.get('part_no', '—')}</div>
                        <div class="part-desc-text" style="font-size: 13px;">{item.get('part_name', '—')}</div>
                        <div>
                            <span class="part-qty-pill" style="background: #991B1B; font-size: 12px; padding: 2px 8px;">Qty: {item.get('qty', 0)}</span>
                            <span class="badge-reason-tag {r_class}">{reason_str}</span>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with r2:
                st.markdown('<div class="delete-btn-wrap">', unsafe_allow_html=True)
                if st.button("🗑️", key=f"del_ret_{item['id']}", help="Delete item"):
                    st.session_state.stock_data["returns"] = [
                        x for x in st.session_state.stock_data["returns"] if x["id"] != item["id"]
                    ]
                    save_data(st.session_state.stock_data)
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

# --- 3. Hold List ---
with col_list_hold:
    st.markdown(
        f"""
        <div class="list-card-box">
            <div class="list-card-header">
                <div class="list-title">
                    <span style="color: #64748B;">●</span> Hold List
                </div>
                <div class="list-count-pill pill-hold-count">
                    {total_hold_parts} parts • {total_hold_qty} pcs
                </div>
            </div>
        """,
        unsafe_allow_html=True,
    )

    if not hold_list:
        st.markdown(
            """
            <div class="empty-prompt">
                No items on hold.<br>Tap "Hold" to decide later.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        for item in hold_list:
            r1, r2 = st.columns([4.8, 1.2])
            with r1:
                st.markdown(
                    f"""
                    <div class="part-entry-row">
                        <div class="part-title-text" style="font-size: 15px;"># {item.get('part_no', '—')}</div>
                        <div class="part-desc-text" style="font-size: 13px;">{item.get('part_name', '—')}</div>
                        <span class="part-qty-pill" style="background: #475569; font-size: 12px; padding: 2px 8px;">Qty: {item.get('qty', 0)}</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with r2:
                st.markdown('<div class="delete-btn-wrap">', unsafe_allow_html=True)
                if st.button("🗑️", key=f"del_hold_{item['id']}", help="Delete item"):
                    st.session_state.stock_data["hold"] = [
                        x for x in st.session_state.stock_data["hold"] if x["id"] != item["id"]
                    ]
                    save_data(st.session_state.stock_data)
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)
