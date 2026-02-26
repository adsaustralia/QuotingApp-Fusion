
import streamlit as st
import pandas as pd
import numpy as np
import re, json, math, io
from pathlib import Path
from datetime import datetime
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

APP_VERSION = "row-based-v5-eval-sum-formulas-and-export-safe"

APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
MAPPING_DIR = DATA_DIR / "mappings"
MAPPING_DIR.mkdir(parents=True, exist_ok=True)
DEFAULT_STANDARD_STOCKS_CSV = DATA_DIR / "standard_stocks.csv"

# ---------- helpers ----------
def clean_text(s) -> str:
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("Ø", "dia ").replace("⌀", "dia ")
    return s

def load_mapping(customer: str) -> dict:
    fp = MAPPING_DIR / f"{clean_text(customer).replace(' ', '_')}_stock_map.json"
    if fp.exists():
        return json.loads(fp.read_text(encoding="utf-8"))
    return {"customer": customer, "mappings": {}}

def save_mapping(customer: str, mapping: dict) -> None:
    fp = MAPPING_DIR / f"{clean_text(customer).replace(' ', '_')}_stock_map.json"
    fp.write_text(json.dumps(mapping, indent=2), encoding="utf-8")

def parse_size_text(text: str):
    """
    Supports:
      - "1200 x 900"
      - "H1700 x W800mm" (letter adjacent)
      - "1125W X 508 Hmm (2pcs)" (suffix)
      - "585mm x W662.5mm"
      - "DIA 600", "diameter 600", "450 round"
    """
    t = clean_text(text)
    if not t:
        return {"shape":"unknown","width_mm":np.nan,"height_mm":np.nan,"diameter_mm":np.nan}

    # remove bracket notes like (2pcs)
    t = re.sub(r"\([^)]*\)", " ", t)
    t = t.replace("mm", " ")
    t = re.sub(r"\s+", " ", t).strip()

    # circle
    m = re.search(r"(dia(?:meter)?)\s*[:=]?\s*(\d+(?:\.\d+)?)", t)
    if m:
        d = float(m.group(2))
        return {"shape":"circle","width_mm":np.nan,"height_mm":np.nan,"diameter_mm":d}
    m = re.search(r"(\d+(?:\.\d+)?)\s*round", t)
    if m:
        d = float(m.group(1))
        return {"shape":"circle","width_mm":np.nan,"height_mm":np.nan,"diameter_mm":d}

    # labelled width/height (any order), including adjacent patterns: w800, 800w, h1700, 1700h
    w = None
    h = None

    # width patterns
    mw = re.search(r"(?:\bwidth\b\s*[:=]?\s*(\d+(?:\.\d+)?))", t)
    if mw:
        w = float(mw.group(1))
    if w is None:
        mw = re.search(r"w\s*(\d+(?:\.\d+)?)", t)  # w800 or w 800
        if mw:
            w = float(mw.group(1))
    if w is None:
        mw = re.search(r"(\d+(?:\.\d+)?)\s*w", t)  # 800w
        if mw:
            w = float(mw.group(1))

    # height patterns
    mh = re.search(r"(?:\bheight\b\s*[:=]?\s*(\d+(?:\.\d+)?))", t)
    if mh:
        h = float(mh.group(1))
    if h is None:
        mh = re.search(r"h\s*(\d+(?:\.\d+)?)", t)  # h1700 or h 1700
        if mh:
            h = float(mh.group(1))
    if h is None:
        mh = re.search(r"(\d+(?:\.\d+)?)\s*h", t)  # 1700h
        if mh:
            h = float(mh.group(1))

    if w is not None and h is not None:
        return {"shape":"rectangle","width_mm":w,"height_mm":h,"diameter_mm":np.nan}

    # generic rectangle: "1200 x 900"
    m = re.search(r"(\d+(?:\.\d+)?)\s*(x|\*|by)\s*(\d+(?:\.\d+)?)", t)
    if m:
        a = float(m.group(1))
        b = float(m.group(3))
        # if we found one of w/h, use it to decide which is missing
        if w is not None and h is None:
            return {"shape":"rectangle","width_mm":w,"height_mm":b,"diameter_mm":np.nan}
        if h is not None and w is None:
            return {"shape":"rectangle","width_mm":a,"height_mm":h,"diameter_mm":np.nan}
        return {"shape":"rectangle","width_mm":a,"height_mm":b,"diameter_mm":np.nan}

    # fallback: first two numbers
    nums = re.findall(r"\d+(?:\.\d+)?", t)
    if len(nums) >= 2:
        return {"shape":"rectangle","width_mm":float(nums[0]),"height_mm":float(nums[1]),"diameter_mm":np.nan}

    return {"shape":"unknown","width_mm":np.nan,"height_mm":np.nan,"diameter_mm":np.nan}

def sqm_calc(shape: str, width_mm=None, height_mm=None, diameter_mm=None):
    if shape == "rectangle" and pd.notna(width_mm) and pd.notna(height_mm):
        return (float(width_mm)/1000.0) * (float(height_mm)/1000.0)
    if shape == "circle" and pd.notna(diameter_mm):
        r = (float(diameter_mm)/1000.0)/2.0
        return math.pi * (r**2)
    return np.nan

def sides_normalize(val: str, default="SS"):
    t = clean_text(val)
    if t in ("ds","2s","double","2pp","two","double sided","double-sided"):
        return "DS"
    if t in ("ss","1s","single","1pp","one","single sided","single-sided"):
        return "SS"
    return default

def eval_qty_value(qty_cell_val, ws):
    """
    If qty is a number -> return it.
    If qty is a simple Excel SUM formula like "=SUM(AB13:AB56)" -> compute it from sheet cells.
    Otherwise -> try numeric conversion.
    """
    if qty_cell_val is None:
        return np.nan

    # direct number
    if isinstance(qty_cell_val, (int, float)) and not (isinstance(qty_cell_val, float) and np.isnan(qty_cell_val)):
        return float(qty_cell_val)

    s = str(qty_cell_val).strip()
    # try numeric string
    n = pd.to_numeric(s, errors="coerce")
    if pd.notna(n):
        return float(n)

    # SUM(range)
    m = re.match(r"^=\s*sum\s*\(\s*([A-Z]+\d+)\s*:\s*([A-Z]+\d+)\s*\)\s*$", s, flags=re.IGNORECASE)
    if m:
        a1, b1 = m.group(1).upper(), m.group(2).upper()
        try:
            cells = ws[a1:b1]
            total = 0.0
            for row in cells:
                for c in row:
                    v = c.value
                    vv = pd.to_numeric(v, errors="coerce")
                    if pd.notna(vv):
                        total += float(vv)
            return total
        except Exception:
            return np.nan

    return np.nan

def export_quote_pdf(df_summary: pd.DataFrame, df_lines: pd.DataFrame, title="Quote") -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm

    bio = io.BytesIO()
    c = canvas.Canvas(bio, pagesize=A4)
    w, h = A4

    y = h - 18*mm
    c.setFont("Helvetica-Bold", 14)
    c.drawString(18*mm, y, title)
    y -= 7*mm
    c.setFont("Helvetica", 9)
    c.drawString(18*mm, y, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  {APP_VERSION}")
    y -= 8*mm

    c.setFont("Helvetica-Bold", 11)
    c.drawString(18*mm, y, "Summary")
    y -= 6*mm
    c.setFont("Helvetica", 10)
    for _, row in df_summary.iterrows():
        c.drawString(20*mm, y, f"{row['Label']}: {row['Value']}")
        y -= 5*mm

    y -= 4*mm
    c.setFont("Helvetica-Bold", 11)
    c.drawString(18*mm, y, "Line Items (Material)")
    y -= 6*mm

    headers = ["Col", "Qty", "Sides", "Size", "Stock", "Total SQM", "Rate", "DS%", "Line Total"]
    col_x = [18, 30, 42, 60, 98, 150, 170, 182, 198]  # mm
    c.setFont("Helvetica-Bold", 7)
    for hx, head in zip(col_x, headers):
        c.drawString(hx*mm, y, head)
    y -= 4*mm
    c.setFont("Helvetica", 7)

    def money(x):
        try: return f"{float(x):,.2f}"
        except: return str(x)

    for _, r in df_lines.head(100).iterrows():
        if y < 18*mm:
            c.showPage()
            y = h - 18*mm
            c.setFont("Helvetica-Bold", 7)
            for hx, head in zip(col_x, headers):
                c.drawString(hx*mm, y, head)
            y -= 4*mm
            c.setFont("Helvetica", 7)

        vals = [
            str(r.get("col_letter","")),
            str(int(r.get("qty",0))),
            str(r.get("sides","")),
            str(r.get("size_text",""))[:16],
            str(r.get("stock_std",""))[:22],
            money(r.get("total_sqm",0)),
            money(r.get("sqm_rate",0)),
            str(int(round((float(r.get("ds_factor",1.0))-1.0)*100))) if pd.notna(r.get("ds_factor")) else "0",
            money(r.get("line_total",0)),
        ]
        for hx, v in zip(col_x, vals):
            c.drawString(hx*mm, y, v)
        y -= 4*mm

    c.showPage()
    c.save()
    return bio.getvalue()

# ---------- UI ----------
st.set_page_config(page_title="Quoting App (Row-based)", layout="wide")
st.title("Quoting App — Row-based Input (Pick ROWS + Column Range)")
st.caption("Example: Size row 5, Material row 6, Sides row 10, Qty row 57. Each COLUMN = one line item.")

with st.sidebar:
    st.header("Upload")
    customer = st.text_input("Customer", value="AU Holiday")
    uploaded = st.file_uploader("Customer Excel", type=["xlsx"])

    st.divider()
    st.header("Standard stock rates")
    st.caption("CSV columns required: stock_name_std, sqm_rate")
    std_upload = st.file_uploader("Optional: upload standard_stocks.csv", type=["csv"])

if uploaded is None:
    st.info("Upload an Excel file to start.")
    st.stop()

# Standard stocks
if std_upload is not None:
    df_std = pd.read_csv(std_upload)
else:
    if DEFAULT_STANDARD_STOCKS_CSV.exists():
        df_std = pd.read_csv(DEFAULT_STANDARD_STOCKS_CSV)
    else:
        df_std = pd.DataFrame(columns=["stock_name_std","sqm_rate"])

df_std["stock_name_std"] = df_std["stock_name_std"].astype(str)
df_std["sqm_rate"] = pd.to_numeric(df_std["sqm_rate"], errors="coerce")
std_options = df_std["stock_name_std"].dropna().tolist()
std_rate_map = dict(zip(df_std["stock_name_std"], df_std["sqm_rate"]))

# Sheet names
uploaded.seek(0)
wb_ro = openpyxl.load_workbook(uploaded, read_only=True, data_only=True)
sheet_names = wb_ro.sheetnames
wb_ro.close()

top1, top2 = st.columns([2,1])
sheet_name = top1.selectbox("Sheet", sheet_names, index=0)
units = top2.selectbox("Units", ["mm","cm","m"], index=0)
u = {"mm":1.0,"cm":10.0,"m":1000.0}[units]

# Preview (best-effort)
with st.expander("Preview (all rows) — best effort", expanded=False):
    try:
        uploaded.seek(0)
        df_preview = pd.read_excel(uploaded, sheet_name=sheet_name, header=None)
        st.dataframe(df_preview, use_container_width=True)
    except Exception as e:
        st.warning("Preview could not be rendered (sheet may be highly formatted). Row-based extraction will still work.")

st.subheader("Pick ROW numbers (1-indexed)")

c1, c2, c3, c4, c5 = st.columns([1.1,1.1,1.1,1.1,1.6])
size_row = c1.number_input("Size row", 1, 5000, 5, 1)
mat_row  = c2.number_input("Material row", 1, 5000, 6, 1)
sides_row= c3.number_input("Sides row", 1, 5000, 10, 1)
qty_row  = c4.number_input("Qty row", 1, 5000, 57, 1)
default_sides = c5.radio("Default sides (if blank)", ["SS","DS"], index=0, horizontal=True)

ds_loading_pct = st.slider("DS loading (%)", 0.0, 100.0, 20.0, 1.0) / 100.0

st.subheader("Pick COLUMN range")
r1, r2, r3 = st.columns([1.2,1.2,1.6])
start_col_letter = r1.text_input("Start column letter", value="A")
end_col_letter   = r2.text_input("End column letter", value="Z")
skip_zero_qty    = r3.checkbox("Skip columns with Qty <= 0", value=True)

st.subheader("Export output location")
st.caption("By default, price is written **next to the Qty row** (Qty row + 1).")
price_row = st.number_input("Write price into row (1-indexed)", 1, 5000, int(qty_row)+1, 1)

# ---------- Extract row-based items ----------
uploaded.seek(0)
wb = openpyxl.load_workbook(uploaded, read_only=False, data_only=False)
ws = wb[sheet_name]

def col_idx(letter: str) -> int:
    letter = (letter or "").strip().upper()
    return column_index_from_string(letter)

c_start = col_idx(start_col_letter)
c_end = col_idx(end_col_letter)
if c_start > c_end:
    c_start, c_end = c_end, c_start

items = []
for c in range(c_start, c_end+1):
    size_val = ws.cell(row=int(size_row), column=c).value
    mat_val  = ws.cell(row=int(mat_row), column=c).value
    sides_val= ws.cell(row=int(sides_row), column=c).value
    qty_val  = ws.cell(row=int(qty_row), column=c).value

    qty = pd.to_numeric(eval_qty_value(qty_val, ws), errors="coerce")
    if skip_zero_qty and (pd.isna(qty) or float(qty) <= 0):
        continue

    size_text = "" if size_val is None else str(size_val)
    geo = parse_size_text(size_text)
    shape = geo["shape"] if geo["shape"] != "unknown" else "rectangle"

    width_mm = geo["width_mm"] * u if pd.notna(geo["width_mm"]) else np.nan
    height_mm = geo["height_mm"] * u if pd.notna(geo["height_mm"]) else np.nan
    diameter_mm = geo["diameter_mm"] * u if pd.notna(geo["diameter_mm"]) else np.nan

    sides = sides_normalize(sides_val, default=default_sides)

    items.append({
        "origin_col": c,
        "col_letter": get_column_letter(c),
        "qty": float(qty) if pd.notna(qty) else 0.0,
        "sides": sides,
        "size_text": size_text,
        "shape": shape,
        "width_mm": width_mm,
        "height_mm": height_mm,
        "diameter_mm": diameter_mm,
        "stock_customer": "" if mat_val is None else str(mat_val),
    })

lines = pd.DataFrame(items)
if len(lines) == 0:
    st.warning("No line items found in the selected column range (check Qty row and column range).")
    st.stop()

lines["sqm_each"] = lines.apply(lambda r: sqm_calc(r["shape"], r["width_mm"], r["height_mm"], r["diameter_mm"]), axis=1)
lines["total_sqm"] = pd.to_numeric(lines["sqm_each"], errors="coerce") * pd.to_numeric(lines["qty"], errors="coerce")

if pd.to_numeric(lines["sqm_each"], errors="coerce").notna().sum() == 0:
    st.error("Size parsing failed: SQM is blank for all columns. Check the Size row values and Units. Supported: W x H, H/W labels (H1700 x W800), DIA 600, 450 round.")

# ---------- Mapping ----------
st.subheader("Stock Mapping (Customer name → Standard stock with sqm rate)")
mapping = load_mapping(customer)
unique_cs = sorted({s for s in lines["stock_customer"].dropna().astype(str).tolist() if clean_text(s)})

map_df = pd.DataFrame([{
    "customer_stock": cs,
    "standard_stock": mapping["mappings"].get(clean_text(cs), "")
} for cs in unique_cs])

edited = st.data_editor(
    map_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "standard_stock": st.column_config.SelectboxColumn("Standard stock", options=[""] + std_options)
    }
)

if st.button("Save mappings", type="primary"):
    new_map = load_mapping(customer)
    for _, r in edited.iterrows():
        cs_key = clean_text(r["customer_stock"])
        std = str(r["standard_stock"] or "").strip()
        if cs_key and std:
            new_map["mappings"][cs_key] = std
    save_mapping(customer, new_map)
    st.success("Mappings saved.")

mapping = load_mapping(customer)
lines["stock_std"] = lines["stock_customer"].apply(lambda x: mapping["mappings"].get(clean_text(x), ""))
lines["sqm_rate"] = lines["stock_std"].map(std_rate_map)

# DS loading factor
lines["ds_factor"] = np.where(lines["sides"].astype(str) == "DS", 1.0 + ds_loading_pct, 1.0)

lines["line_total"] = (
    pd.to_numeric(lines["total_sqm"], errors="coerce")
    * pd.to_numeric(lines["sqm_rate"], errors="coerce")
    * pd.to_numeric(lines["ds_factor"], errors="coerce")
)

# ---------- Review ----------
st.subheader("Quote Review")
review = lines[["col_letter","qty","sides","ds_factor","shape","size_text","stock_customer","stock_std","sqm_each","total_sqm","sqm_rate","line_total"]].copy()
st.dataframe(review, use_container_width=True)

total_sqm = float(pd.to_numeric(review["total_sqm"], errors="coerce").fillna(0).sum())
subtotal = float(pd.to_numeric(review["line_total"], errors="coerce").fillna(0).sum())

summary = pd.DataFrame([
    {"Label":"Customer", "Value": customer},
    {"Label":"Sheet", "Value": sheet_name},
    {"Label":"Total SQM", "Value": f"{total_sqm:,.3f}"},
    {"Label":"Subtotal (Material)", "Value": f"{subtotal:,.2f}"},
    {"Label":"DS loading %", "Value": f"{ds_loading_pct*100:.0f}%"},
])

st.markdown("**Totals**")
st.table(summary)

def export_preserving_excel() -> bytes:
    # write prices across columns into the chosen row (next to qty by default)
    for _, r in lines.iterrows():
        c = int(r["origin_col"])
        val = r.get("line_total")
        if pd.isna(val):
            continue
        cell = ws.cell(row=int(price_row), column=c)
        cell.value = float(val)
        cell.number_format = "0.00"

    # add summary/detail sheets
    for name in ["Quote Summary", "Line Items"]:
        if name in wb.sheetnames:
            del wb[name]
    ws_sum = wb.create_sheet("Quote Summary")
    for r_idx, row in enumerate(summary.itertuples(index=False), start=1):
        ws_sum.cell(row=r_idx, column=1, value=row.Label)
        ws_sum.cell(row=r_idx, column=2, value=row.Value)

    ws_li = wb.create_sheet("Line Items")
    for c_idx, col in enumerate(review.columns.tolist(), start=1):
        ws_li.cell(row=1, column=c_idx, value=col)
    for r_idx, row in enumerate(review.itertuples(index=False), start=2):
        for c_idx, v in enumerate(row, start=1):
            ws_li.cell(row=r_idx, column=c_idx, value=v)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

b1, b2 = st.columns(2)
with b1:
    xbytes = export_preserving_excel()
    st.download_button(
        "Download Excel (preserve format)",
        data=xbytes,
        file_name=f"Quote_{customer}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
with b2:
    pbytes = export_quote_pdf(summary, review, title=f"Quote - {customer}")
    st.download_button(
        "Download Quote PDF",
        data=pbytes,
        file_name=f"Quote_{customer}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf"
    )

with st.expander("Version / Debug", expanded=False):
    st.write("APP VERSION:", APP_VERSION)
    st.write("Columns processed:", f"{start_col_letter}:{end_col_letter}")
    st.write("Rows:", dict(size_row=int(size_row), material_row=int(mat_row), sides_row=int(sides_row), qty_row=int(qty_row), price_row=int(price_row)))
