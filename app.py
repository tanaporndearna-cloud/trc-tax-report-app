import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import google.generativeai as genai
import json
import re
from datetime import date
import calendar
import io

# ===== CONFIG =====
SOURCE_SHEET_ID = "1lD6YrCoSbA5RI79PvtG2WtWTVd0q83xvPmESLDMcu-0"
DEST_SHEET_ID   = "1YJNipc9ndkrn9ZYu9pG3coykWh5oMj7uwrnlzViZPgQ"
CACHE_TAB_NAME  = "receipt_cache"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

BRANCH_MAP = {
    "BNA": 3,  "BNB": 4,  "BNC": 7,  "BND": 5,  "CJH": 10,
    "PAN": 9,  "PRS": 6,  "PTA": 1,  "PTB": 2,  "PTC": 8,
    "PTY": 11, "TBY": 16, "TCB": 13, "TEM": 12, "TLX": 17,
    "TNW": 14, "TNY": 15, "TSP": 18,
}

MONTH_NAMES_TH = {
    1: "à¸¡à¸à¸£à¸²à¸à¸¡", 2: "à¸à¸¸à¸¡à¸ à¸²à¸à¸±à¸à¸à¹", 3: "à¸¡à¸µà¸à¸²à¸à¸¡", 4: "à¹à¸¡à¸©à¸²à¸¢à¸",
    5: "à¸à¸¤à¸©à¸ à¸²à¸à¸¡", 6: "à¸¡à¸´à¸à¸¸à¸à¸²à¸¢à¸", 7: "à¸à¸£à¸à¸à¸²à¸à¸¡", 8: "à¸ªà¸´à¸à¸«à¸²à¸à¸¡",
    9: "à¸à¸±à¸à¸¢à¸²à¸¢à¸", 10: "à¸à¸¸à¸¥à¸²à¸à¸¡", 11: "à¸à¸¤à¸¨à¸à¸´à¸à¸²à¸¢à¸", 12: "à¸à¸±à¸à¸§à¸²à¸à¸¡",
}

MONTH_NAMES_EN = {
    1:"january",2:"february",3:"march",4:"april",5:"may",6:"june",
    7:"july",8:"august",9:"september",10:"october",11:"november",12:"december",
}

# ===== AUTH =====
@st.cache_resource
def get_clients():
    creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    gc = gspread.authorize(creds)
    drive = build("drive", "v3", credentials=creds)
    return gc, drive

# ===== CACHE =====
def load_cache(gc):
    """Load receipt cache from destination sheet tab."""
    try:
        wb = gc.open_by_key(DEST_SHEET_ID)
        try:
            ws = wb.worksheet(CACHE_TAB_NAME)
            rows = ws.get_all_values()
            return {r[0]: r[1] for r in rows if len(r) >= 2 and r[0]}
        except gspread.exceptions.WorksheetNotFound:
            wb.add_worksheet(CACHE_TAB_NAME, rows=1000, cols=2)
            return {}
    except Exception as e:
        st.warning(f"à¹à¸«à¸¥à¸ cache à¹à¸¡à¹à¹à¸à¹: {e}")
        return {}

def save_cache(gc, cache: dict):
    """Save receipt cache back to destination sheet tab."""
    wb = gc.open_by_key(DEST_SHEET_ID)
    ws = wb.worksheet(CACHE_TAB_NAME)
    ws.clear()
    if cache:
        ws.update([[k, v] for k, v in cache.items()])

# ===== BRANCH =====
def get_branch(doc_no: str, employee_name: str = "") -> int:
    if "HCT" in doc_no or "HPW" in doc_no:
        m = re.search(r'T(\d+)', employee_name or "")
        return int(m.group(1)) if m else 0
    for code, num in BRANCH_MAP.items():
        if code in doc_no:
            return num
    return 0

# ===== INVOICE NUMBER =====
def get_working_days_in_month(year: int, month: int):
    num_days = calendar.monthrange(year, month)[1]
    return [date(year, month, d) for d in range(1, num_days + 1)
            if date(year, month, d).weekday() != 6]

def compute_inv_no(year: int, month: int, wd_index: int, slot: int) -> int:
    be_year = year + 543
    prefix = be_year * 100000 + month * 1000
    return prefix + wd_index * 25 + slot

# ===== GEMINI RECEIPT PARSER =====
def parse_receipt_with_gemini(file_bytes: bytes, mime_type: str) -> dict:
    """Use Gemini vision to extract receipt data."""
    prompt = """à¸­à¹à¸²à¸à¹à¸à¹à¸ªà¸£à¹à¸à¸à¸µà¹à¹à¸¥à¹à¸§à¸à¸­à¸à¹à¸à¹à¸ JSON à¹à¸à¹à¸²à¸à¸±à¹à¸ à¹à¸¡à¹à¸à¹à¸­à¸à¸­à¸à¸´à¸à¸²à¸¢à¹à¸à¸´à¹à¸¡:
{
  "doc_no": "à¹à¸¥à¸à¸à¸µà¹à¹à¸­à¸à¸ªà¸²à¸£ à¹à¸à¹à¸ ABBPTC26070003",
  "customer": "à¸à¸·à¹à¸­à¸¥à¸¹à¸à¸à¹à¸²",
  "employee": "à¸à¸·à¹à¸­à¸à¸à¸±à¸à¸à¸²à¸ (à¸à¹à¸²à¸¡à¸µ)",
  "items": [
    {"name": "à¸à¸·à¹à¸­à¸ªà¸´à¸à¸à¹à¸²", "qty": 1, "price": 1000}
  ]
}
à¸à¹à¸²à¸¡à¸£à¸²à¸¢à¸à¸²à¸£ VAT, deposit, à¸£à¸²à¸à¸² 0 à¸à¸²à¸"""

    import base64
    b64 = base64.b64encode(file_bytes).decode()
    import base64
    b64 = base64.b64encode(file_bytes).decode()
    model = genai.GenerativeModel("gemini-2.0-flash")
    response = model.generate_content([
        {"mime_type": mime_type, "data": b64},
        prompt
    ])
    text = response.text.strip()
    # extract JSON
    m = re.search(r'\{[\s\S]+\}', text)
    if m:
        return json.loads(m.group())
    return {}

def download_drive_file(drive, file_id: str):
    """Download file from Google Drive, return (bytes, mime_type)."""
    meta = drive.files().get(fileId=file_id, fields="mimeType,name").execute()
    mime = meta.get("mimeType", "")

    if "google-apps" in mime:
        # Export Google Docs/Slides as PDF
        export_mime = "application/pdf"
        data = drive.files().export(fileId=file_id, mimeType=export_mime).execute()
    else:
        data = drive.files().get_media(fileId=file_id).execute()
        export_mime = mime

    return data, export_mime

def extract_file_id(url: str) -> str:
    patterns = [
        r'/file/d/([a-zA-Z0-9_-]+)',
        r'[?&]id=([a-zA-Z0-9_-]+)',
        r'/d/([a-zA-Z0-9_-]+)',
    ]
    for p in patterns:
        m = re.search(p, url or "")
        if m:
            return m.group(1)
    return ""

# ===== WRITE TO SHEET =====
def get_or_create_report_sheet(gc, year: int, month: int):
    sheet_name = f"{MONTH_NAMES_TH[month]} {year + 543}"
    wb = gc.open_by_key(DEST_SHEET_ID)
    try:
        ws = wb.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = wb.add_worksheet(sheet_name, rows=2000, cols=15)
        # Write header
        headers = ["à¸§.à¸.à¸.", "à¹à¸¥à¸à¸à¸µà¹", "à¸ªà¸²à¸à¸²", "à¹à¸¥à¸à¸à¸µà¹à¹à¸­à¸à¸ªà¸²à¸£", "à¸à¸·à¹à¸­à¸¥à¸¹à¸à¸à¹à¸²",
                   "à¸à¸µà¹à¸­à¸¢à¸¹à¹", "TAX ID", "à¸£à¸²à¸¢à¸à¸²à¸£à¸ªà¸´à¸à¸à¹à¸²", "à¸à¸³à¸à¸§à¸", "à¸£à¸²à¸à¸²/à¸«à¸à¹à¸§à¸¢",
                   "à¸£à¸§à¸¡", "VAT 7%", "à¸£à¸§à¸¡à¸à¸±à¹à¸à¸ªà¸´à¹à¸", "à¸à¹à¸­à¸à¸à¸²à¸", "à¸­à¸µà¹à¸¡à¸¥/à¸à¸±à¸à¸ªà¹à¸"]
        ws.append_row(headers)
    return ws

def rows_already_in_sheet(ws) -> set:
    """Return set of doc_no already written."""
    vals = ws.col_values(4)  # Col D = doc_no
    return set(v for v in vals[1:] if v)

def append_bills_to_sheet(ws, bills: list, year: int, month: int, existing_doc_nos: set):
    """bills = list of dicts with all bill data. Appends only new ones."""
    working_days = get_working_days_in_month(year, month)
    # Group bills by purchase_date
    from collections import defaultdict
    by_date = defaultdict(list)
    for b in bills:
        by_date[b["purchase_date"]].append(b)

    rows_to_add = []
    slot_counters = {}  # date -> slot index

    for bill in sorted(bills, key=lambda x: (x["purchase_date"], x["doc_no"])):
        doc_no = bill["doc_no"]
        if doc_no in existing_doc_nos:
            continue

        d = bill["purchase_date"]
        be_date_str = f"{d.day:02d}/{d.month:02d}/{d.year + 543}"

        # Find working day index
        if d in working_days:
            wd_idx = working_days.index(d)
        else:
            wd_idx = 0  # fallback

        slot = slot_counters.get(d, 0)
        slot_counters[d] = slot + 1
        inv_no = compute_inv_no(year, month, wd_idx, slot + 1)

        items = bill.get("items", [])
        for i, item in enumerate(items):
            qty   = item.get("qty", 1)
            price = item.get("price", 0)
            total = qty * price
            vat   = round(total * 7 / 100, 2)
            grand = total + vat

            if i == 0:
                row = [
                    be_date_str,
                    inv_no,
                    bill.get("branch", ""),
                    doc_no,
                    bill.get("customer", ""),
                    bill.get("address", ""),
                    bill.get("tax_id", ""),
                    item.get("name", ""),
                    qty, price, total, vat, grand,
                    bill.get("channel", ""),
                    bill.get("email", ""),
                ]
            else:
                row = ["", "", "", "", "", "", "",
                       item.get("name", ""), qty, price, total, vat, grand,
                       "", ""]
            rows_to_add.append(row)

    if rows_to_add:
        ws.append_rows(rows_to_add, value_input_option="USER_ENTERED")
    return len([b for b in bills if b["doc_no"] not in existing_doc_nos])

# ===== MAIN APP =====
st.set_page_config(page_title="à¸£à¸²à¸¢à¸à¸²à¸à¸ à¸²à¸©à¸µà¸à¸²à¸¢ TRC", page_icon="ð", layout="wide")
st.title("ð à¸£à¸²à¸¢à¸à¸²à¸à¸ à¸²à¸©à¸µà¸à¸²à¸¢à¸£à¸²à¸¢à¹à¸à¸·à¸­à¸ â TRC Motorsport")

# Sidebar
with st.sidebar:
    st.header("âï¸ à¸à¸±à¹à¸à¸à¹à¸²")
    year  = st.number_input("à¸à¸µ (à¸.à¸¨.)", value=2026, min_value=2024, max_value=2030)
    month = st.selectbox("à¹à¸à¸·à¸­à¸", list(MONTH_NAMES_TH.keys()),
                         format_func=lambda x: MONTH_NAMES_TH[x], index=6)

    st.markdown("---")
    st.caption("Gemini API Key à¸à¸±à¹à¸à¹à¸ Streamlit Secrets")
    st.caption(f"Sheet à¸à¸¥à¸²à¸¢à¸à¸²à¸: [à¹à¸à¸´à¸](https://docs.google.com/spreadsheets/d/{DEST_SHEET_ID})")

# Main
col1, col2 = st.columns([2, 1])
with col1:
    run_btn = st.button("ð à¸ªà¸£à¹à¸²à¸/à¸­à¸±à¸à¹à¸à¸à¸£à¸²à¸¢à¸à¸²à¸", type="primary", use_container_width=True)

if run_btn:
    try:
        gc, drive = get_clients()
        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    except Exception as e:
        st.error(f"à¹à¸à¸·à¹à¸­à¸¡à¸à¹à¸­à¹à¸¡à¹à¹à¸à¹: {e}")
        st.stop()

    progress = st.progress(0, text="à¹à¸«à¸¥à¸ cache...")

    # 1. Load cache
    cache = load_cache(gc)
    st.write(f"Cache: {len(cache)} à¸£à¸²à¸¢à¸à¸²à¸£")

    # 2. Read source sheet
    progress.progress(10, text="à¸­à¹à¸²à¸ Google Sheet...")
    try:
        src_wb = gc.open_by_key(SOURCE_SHEET_ID)
        src_ws = src_wb.get_worksheet(0)
        rows = src_ws.get_all_values()
    except Exception as e:
        st.error(f"à¸­à¹à¸²à¸ Sheet à¸à¹à¸à¸à¸²à¸à¹à¸¡à¹à¹à¸à¹: {e}")
        st.stop()

    # Parse sheet rows (skip header row 0)
    form_rows = []
    seen_doc = {}
    for row in rows[1:]:
        if len(row) < 7:
            continue
        _, purchase_date_str, address, tax_id, channel, email, drive_url = row[:7]
        file_id = extract_file_id(drive_url)
        if not file_id:
            continue
        form_rows.append({
            "purchase_date_str": purchase_date_str,
            "address": address,
            "tax_id": tax_id,
            "channel": channel,
            "email": email,
            "file_id": file_id,
            "drive_url": drive_url,
        })

    # 3. Find new receipts
    new_rows = [r for r in form_rows if r["file_id"] not in cache]
    st.write(f"à¹à¸à¹à¸«à¸¡à¹à¸à¸µà¹à¸à¹à¸­à¸à¸­à¹à¸²à¸: {len(new_rows)} à¹à¸")
    progress.progress(20, text=f"à¸­à¹à¸²à¸à¹à¸à¹à¸ªà¸£à¹à¸à¹à¸«à¸¡à¹ {len(new_rows)} à¹à¸...")

    # 4. Parse new receipts with Gemini
    bills = []
    for idx, fr in enumerate(form_rows):
        file_id = fr["file_id"]

        if file_id in cache:
            # Cache hit â skip (we'll need cached data separately for full rebuild)
            continue

        pct = 20 + int(60 * idx / max(len(form_rows), 1))
        progress.progress(pct, text=f"à¸­à¹à¸²à¸à¹à¸à¹à¸ªà¸£à¹à¸ {idx+1}/{len(new_rows)}...")

        try:
            file_bytes, mime_type = download_drive_file(drive, file_id)
            parsed = parse_receipt_with_gemini(file_bytes, mime_type)
        except Exception as e:
            st.warning(f"à¸­à¹à¸²à¸à¹à¸à¹à¸¡à¹à¹à¸à¹ (file_id={file_id}): {e}")
            continue

        doc_no = parsed.get("doc_no", "")
        if not doc_no:
            st.warning(f"à¹à¸¡à¹à¸à¸ doc_no à¹à¸à¹à¸ file_id={file_id}")
            continue

        # Handle duplicate doc_no
        if doc_no in seen_doc:
            doc_no = doc_no + "B"
        seen_doc[doc_no] = True

        # Parse purchase date
        try:
            from datetime import datetime
            d = datetime.strptime(fr["purchase_date_str"].strip(), "%d/%m/%Y").date()
        except Exception:
            d = date.today()

        branch = get_branch(doc_no, parsed.get("employee", ""))

        bills.append({
            "doc_no": doc_no,
            "purchase_date": d,
            "address": fr["address"],
            "tax_id": fr["tax_id"],
            "channel": fr["channel"],
            "email": fr["email"],
            "customer": parsed.get("customer", ""),
            "branch": branch,
            "items": parsed.get("items", []),
        })

        # Update cache
        cache[file_id] = date.today().strftime("%d/%m/%Y")

    progress.progress(85, text="à¹à¸à¸µà¸¢à¸à¸¥à¸ Google Sheet...")

    # 5. Write to destination sheet
    ws = get_or_create_report_sheet(gc, year, month)
    existing = rows_already_in_sheet(ws)
    added = append_bills_to_sheet(ws, bills, year, month, existing)

    # 6. Save cache
    progress.progress(95, text="à¸à¸±à¸à¸à¸¶à¸ cache...")
    save_cache(gc, cache)

    progress.progress(100, text="à¹à¸ªà¸£à¹à¸à¹à¸¥à¹à¸§!")
    st.success(f"â à¹à¸à¸´à¹à¸¡à¹à¸à¹à¸«à¸¡à¹ {added} à¹à¸ à¸¥à¸à¹à¸ Sheet à¹à¸£à¸µà¸¢à¸à¸£à¹à¸­à¸¢à¸à¹à¸°")
    st.markdown(f"[à¹à¸à¸´à¸ Google Sheet à¸à¸¥à¸²à¸¢à¸à¸²à¸](https://docs.google.com/spreadsheets/d/{DEST_SHEET_ID})")
