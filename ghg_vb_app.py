
# ghg_vb_app.py
# Streamlit prototype for a GHG Verification Body (single-file demo)
# Run locally with:  streamlit run ghg_vb_app.py
# Author: ChatGPT (GPT-5 Thinking)
#
# Features:
# 1) Upload & parse Excel, Word, PDF, and URL sources
# 2) Extract emissions-relevant data (fuel, electricity, travel, PGS, commuting, 3PL)
# 3) Standardize + validate to a unified schema
# 4) Quantify CO2e with EPA/GHG Protocol-like emission factors and unit conversions
# 5) Simple ISO 14064-3:2019 verification logic (engagement type, assurance level, materiality, sampling)
# 6) Draft outputs: GHG Inventory Report, Verification Statement, Summary Verification Report
# 7) Auditor tools: annotate findings, upload notes, approve/reject data sources
# 8) Streamlit UI with drag-and-drop upload
# 9) All in one file for demo; uses mock/sample data and factors
# 10) Downloadable PDF reports (via ReportLab if available; otherwise fallback to .txt)

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import json
import math
import datetime as dt

# Optional imports with graceful fallback
try:
    import docx  # python-docx
except Exception:
    docx = None

try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    import requests
    from bs4 import BeautifulSoup
except Exception:
    requests = None
    BeautifulSoup = None

# PDF generation
REPORTLAB_AVAILABLE = True
try:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
except Exception:
    REPORTLAB_AVAILABLE = False

st.set_page_config(page_title="GHG VB Prototype", layout="wide")

# -----------------------------
# Mock emission factors & units
# -----------------------------
EMISSION_FACTORS = {
    # factors ~illustrative only (kg CO2e per unit)
    "diesel_liter": 2.68,       # kg CO2e / liter
    "gasoline_liter": 2.31,     # kg CO2e / liter
    "natural_gas_therm": 5.30,  # kg CO2e / therm (illustrative)
    "electricity_kwh": 0.40,    # kg CO2e / kWh (grid avg placeholder)
    "air_travel_km": 0.15,      # kg CO2e / km
    "car_travel_km": 0.18,      # kg CO2e / km (generic)
    "hotel_night": 15.0,        # kg CO2e / night (illustrative)
    "pgs_usd": 0.5,             # kg CO2e / $ (very rough illustrative)
    "commute_km": 0.14,         # kg CO2e / km
    "3pl_ton_km": 0.06          # kg CO2e / ton-km
}

UNIT_CONVERSIONS = {
    # input_unit -> (normalized_unit, factor_to_normalized)
    "gallon_diesel": ("liter", 3.78541),
    "gallon_gasoline": ("liter", 3.78541),
    "m3_natgas": ("therm", 0.0366),  # ~1 m3 ~ 0.0366 therm (illustrative)
    "kwh": ("kwh", 1.0),
    "km": ("km", 1.0),
    "mile": ("km", 1.60934),
    "usd": ("usd", 1.0),
    "night": ("night", 1.0),
    "ton_km": ("ton_km", 1.0),
    "liter": ("liter", 1.0),
    "therm": ("therm", 1.0)
}

CATEGORIES = [
    "Fuel: Diesel",
    "Fuel: Gasoline",
    "Fuel: Natural Gas",
    "Electricity",
    "Business Travel",
    "Purchased Goods & Services",
    "Employee Commuting",
    "Third-Party Logistics"
]

SCOPE_MAP = {
    "Fuel: Diesel": "Scope 1",
    "Fuel: Gasoline": "Scope 1",
    "Fuel: Natural Gas": "Scope 1",
    "Electricity": "Scope 2",
    "Business Travel": "Scope 3",
    "Purchased Goods & Services": "Scope 3",
    "Employee Commuting": "Scope 3",
    "Third-Party Logistics": "Scope 3"
}

# -----------------------------
# Utility helpers
# -----------------------------
def readable_size(num_bytes: int) -> str:
    for unit in ["B", "KB", "MB", "GB"]:
        if num_bytes < 1024.0:
            return f"{num_bytes:3.1f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.1f} TB"

def to_float_safe(val):
    try:
        if isinstance(val, str):
            val = val.replace(",", "").strip()
        return float(val)
    except Exception:
        return None

def normalize_unit(value, unit_hint):
    """Convert value to normalized unit from UNIT_CONVERSIONS mapping."""
    if unit_hint is None:
        return None, None, None
    unit_hint = unit_hint.lower()
    if unit_hint in UNIT_CONVERSIONS:
        normalized, factor = UNIT_CONVERSIONS[unit_hint]
        return value * factor, normalized, factor
    # try partial matches
    for k, (norm, factor) in UNIT_CONVERSIONS.items():
        if k in unit_hint:
            return value * factor, norm, factor
    return value, unit_hint, 1.0

# -----------------------------
# Parsing functions
# -----------------------------
def parse_excel(file_bytes):
    dfs = {}
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
        for sheet in xls.sheet_names:
            try:
                df = xls.parse(sheet)
                dfs[sheet] = df
            except Exception:
                pass
    except Exception as e:
        st.warning(f"Excel parse error: {e}")
    return dfs

def parse_word(file_bytes):
    tables = []
    text_blocks = []
    if docx is None:
        st.warning("python-docx not installed; cannot parse .docx.")
        return {"tables": tables, "text": text_blocks}
    try:
        document = docx.Document(io.BytesIO(file_bytes))
        # tables
        for t in document.tables:
            rows = []
            for row in t.rows:
                rows.append([cell.text.strip() for cell in row.cells])
            if rows:
                try:
                    df = pd.DataFrame(rows[1:], columns=rows[0])
                except Exception:
                    df = pd.DataFrame(rows)
                tables.append(df)
        # paragraphs
        paragraphs = [p.text.strip() for p in document.paragraphs if p.text.strip()]
        text_blocks = paragraphs
    except Exception as e:
        st.warning(f"Word parse error: {e}")
    return {"tables": tables, "text": text_blocks}

def parse_pdf(file_bytes):
    tables = []
    lines = []
    if pdfplumber is None:
        st.warning("pdfplumber not installed; cannot parse PDFs. Falling back to text extraction not available.")
        return {"tables": tables, "lines": lines}
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                try:
                    tbls = page.extract_tables() or []
                    for t in tbls:
                        try:
                            df = pd.DataFrame(t[1:], columns=t[0])
                        except Exception:
                            df = pd.DataFrame(t)
                        tables.append(df)
                except Exception:
                    pass
                try:
                    text = page.extract_text() or ""
                    for ln in text.splitlines():
                        if ln.strip():
                            lines.append(ln.strip())
                except Exception:
                    pass
    except Exception as e:
        st.warning(f"PDF parse error: {e}")
    return {"tables": tables, "lines": lines}

def parse_url(url):
    if requests is None or BeautifulSoup is None:
        st.warning("requests/BeautifulSoup not installed; cannot fetch URLs.")
        return {"tables": [], "text": ""}
    try:
        resp = requests.get(url, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        # Find tables
        tables = []
        for html_table in soup.find_all("table"):
            rows = []
            for tr in html_table.find_all("tr"):
                cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
                if cells:
                    rows.append(cells)
            if rows:
                try:
                    df = pd.DataFrame(rows[1:], columns=rows[0])
                except Exception:
                    df = pd.DataFrame(rows)
                tables.append(df)
        # Get text content
        text = soup.get_text(" ", strip=True)
        return {"tables": tables, "text": text}
    except Exception as e:
        st.warning(f"URL parse error: {e}")
        return {"tables": [], "text": ""}

# -----------------------------
# Extraction heuristics
# -----------------------------
def extract_emissions_records(from_df: pd.DataFrame, source_name: str):
    """
    Try to detect rows/columns related to categories. Look for keywords in headers and rows.
    Return standardized records list of dicts: {source, category, activity_value, unit_hint, period, notes}
    """
    records = []
    if from_df is None or from_df.empty:
        return records

    df = from_df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Try to infer column roles
    text_cols = [c for c in df.columns if any(k in c for k in ["desc", "category", "type", "item", "activity", "fuel", "name"])]
    qty_cols = [c for c in df.columns if any(k in c for k in ["qty", "quantity", "amount", "value", "usage", "use", "consumption"])]
    unit_cols = [c for c in df.columns if "unit" in c or "uom" in c]
    period_cols = [c for c in df.columns if any(k in c for k in ["period", "month", "date", "year"])]
    note_cols = [c for c in df.columns if "note" in c or "comment" in c]

    # Fallback guesses
    if not text_cols and "category" in df.columns:
        text_cols = ["category"]
    if not qty_cols:
        for c in df.columns:
            if df[c].dtype.kind in "fi":
                qty_cols.append(c); break

    # Iterate rows
    for _, row in df.iterrows():
        text_blob = " ".join([str(row[c]) for c in text_cols if c in df.columns]).lower()
        qty = None
        unit_hint = None
        period = None
        notes = " ".join([str(row[c]) for c in note_cols if c in df.columns if not pd.isna(row[c])]).strip()

        for qc in qty_cols:
            val = row.get(qc, None)
            if pd.isna(val): 
                continue
            f = to_float_safe(val)
            if f is not None:
                qty = f
                break
        for uc in unit_cols:
            u = row.get(uc, None)
            if not pd.isna(u):
                unit_hint = str(u)

        for pc in period_cols:
            p = row.get(pc, None)
            if not pd.isna(p):
                period = str(p)

        # classify category
        category = None
        if any(k in text_blob for k in ["diesel"]):
            category = "Fuel: Diesel"
            unit_hint = unit_hint or "liter"
        elif any(k in text_blob for k in ["gasoline", "petrol"]):
            category = "Fuel: Gasoline"
            unit_hint = unit_hint or "liter"
        elif any(k in text_blob for k in ["natural gas", "nat gas", "ng"]):
            category = "Fuel: Natural Gas"
            unit_hint = unit_hint or "therm"
        elif any(k in text_blob for k in ["electric", "kwh", "grid"]):
            category = "Electricity"
            unit_hint = unit_hint or "kwh"
        elif any(k in text_blob for k in ["travel", "flight", "air", "mileage"]):
            category = "Business Travel"
            unit_hint = unit_hint or "km"
        elif any(k in text_blob for k in ["purchased goods", "pgs", "spend", "procure", "supplier"]):
            category = "Purchased Goods & Services"
            unit_hint = unit_hint or "usd"
        elif any(k in text_blob for k in ["commute", "employee commute"]):
            category = "Employee Commuting"
            unit_hint = unit_hint or "km"
        elif any(k in text_blob for k in ["logistics", "3pl", "freight", "ton-km", "ton_km"]):
            category = "Third-Party Logistics"
            unit_hint = unit_hint or "ton_km"

        if category and qty is not None:
            records.append({
                "source": source_name,
                "category": category,
                "activity_value": qty,
                "unit_hint": unit_hint,
                "period": period,
                "notes": notes
            })
    return records

def extract_from_text_lines(lines, source_name):
    records = []
    patt = re.compile(r"(diesel|gasoline|natural gas|electricity|travel|commute|purchased goods|pgs|logistics|3pl|ton-?km)\s*[:\-]\s*([0-9\.,]+)\s*([A-Za-z_/-]+)?", re.I)
    for ln in lines:
        m = patt.search(ln)
        if m:
            kw = m.group(1).lower()
            qty = to_float_safe(m.group(2))
            unit = m.group(3) or ""
            unit = unit.replace("/", "_").replace("-", "_")
            if "diesel" in kw:
                cat, default = "Fuel: Diesel", "liter"
            elif "gasoline" in kw:
                cat, default = "Fuel: Gasoline", "liter"
            elif "natural gas" in kw:
                cat, default = "Fuel: Natural Gas", "therm"
            elif "electricity" in kw:
                cat, default = "Electricity", "kwh"
            elif "travel" in kw:
                cat, default = "Business Travel", "km"
            elif "commute" in kw:
                cat, default = "Employee Commuting", "km"
            elif "purchased goods" in kw or "pgs" in kw:
                cat, default = "Purchased Goods & Services", "usd"
            else:
                cat, default = "Third-Party Logistics", "ton_km"
            if qty is not None:
                records.append({
                    "source": source_name,
                    "category": cat,
                    "activity_value": qty,
                    "unit_hint": unit or default,
                    "period": None,
                    "notes": ln
                })
    return records

# -----------------------------
# Standardize & quantify
# -----------------------------
def standardize_and_quantify(records):
    """
    Convert activity to normalized units and apply EFs to get kgCO2e.
    Return DataFrame with columns:
    source, category, scope, activity_value, unit_hint, normalized_value, normalized_unit, ef_used, kg_co2e
    """
    rows = []
    for r in records:
        cat = r["category"]
        scope = SCOPE_MAP.get(cat, "Scope 3")
        val = r["activity_value"]
        unit_hint = (r["unit_hint"] or "").lower()

        # Normalize unit
        norm_val, norm_unit, factor = normalize_unit(val, unit_hint)
        # Map to EF key
        ef_key = None
        if cat == "Fuel: Diesel":
            ef_key = "diesel_liter"
        elif cat == "Fuel: Gasoline":
            ef_key = "gasoline_liter"
        elif cat == "Fuel: Natural Gas":
            ef_key = "natural_gas_therm"
        elif cat == "Electricity":
            ef_key = "electricity_kwh"
        elif cat == "Business Travel":
            # assume km based by default
            if norm_unit != "km":
                # try convert miles to km was covered in normalize_unit
                pass
            ef_key = "air_travel_km"  # placeholder
        elif cat == "Purchased Goods & Services":
            ef_key = "pgs_usd"
        elif cat == "Employee Commuting":
            ef_key = "commute_km"
        elif cat == "Third-Party Logistics":
            ef_key = "3pl_ton_km"

        ef = EMISSION_FACTORS.get(ef_key, 0.0)
        kg_co2e = (norm_val or 0.0) * ef if norm_val is not None else None

        rows.append({
            "source": r["source"],
            "category": cat,
            "scope": scope,
            "activity_value": val,
            "unit_hint": unit_hint,
            "normalized_value": norm_val,
            "normalized_unit": norm_unit,
            "ef_used": ef_key,
            "kg_co2e": kg_co2e,
            "period": r.get("period"),
            "notes": r.get("notes", "")
        })
    return pd.DataFrame(rows)

# -----------------------------
# Verification logic (ISO 14064-3 lite)
# -----------------------------
def verification_plan(df, assurance_level: str, materiality_pct: float):
    """
    Simple risk-based selection:
    - Rank categories by kgCO2e, select categories contributing until cumulative >= materiality % of total.
    - Tag these as 'High' risk, else 'Moderate/Low' based on cut.
    """
    if df.empty:
        return pd.DataFrame(columns=["category","scope","kg_co2e","contribution_pct","risk","sample_size"])

    summary = df.groupby(["category","scope"], dropna=False)["kg_co2e"].sum().reset_index()
    total = summary["kg_co2e"].sum()
    summary["contribution_pct"] = np.where(total>0, 100.0 * summary["kg_co2e"] / total, 0.0)
    summary = summary.sort_values("contribution_pct", ascending=False)

    cum = 0.0
    risks = []
    for _, row in summary.iterrows():
        cum += row["contribution_pct"]
        if cum <= materiality_pct:
            risks.append("High")
        else:
            # set Moderate for next 20%, Low after
            if cum <= min(100.0, materiality_pct + 20.0):
                risks.append("Moderate")
            else:
                risks.append("Low")
    summary["risk"] = risks

    # Sample size heuristic based on assurance
    base_n = 5 if assurance_level.lower().startswith("limited") else 15
    summary["sample_size"] = summary["risk"].map({"High": base_n+10, "Moderate": base_n, "Low": max(2, base_n//2)})
    return summary

def make_inventory_tables(df):
    scopes = df.groupby("scope")["kg_co2e"].sum().reset_index().sort_values("scope")
    cats = df.groupby(["scope","category"])["kg_co2e"].sum().reset_index()
    total = df["kg_co2e"].sum()
    return scopes, cats, total

# -----------------------------
# Reporting (PDF/text fallback)
# -----------------------------
def build_pdf(filename, title, sections):
    if not REPORTLAB_AVAILABLE:
        # Fallback: write a .txt file
        with open(filename.replace(".pdf", ".txt"), "w", encoding="utf-8") as f:
            f.write(title + "\n" + "="*len(title) + "\n\n")
            for heading, body in sections:
                f.write(heading + "\n" + "-"*len(heading) + "\n")
                if isinstance(body, str):
                    f.write(body + "\n\n")
                else:
                    # assume DataFrame
                    f.write(body.to_string(index=False))
                    f.write("\n\n")
        return filename.replace(".pdf", ".txt")

    doc = SimpleDocTemplate(filename, pagesize=LETTER,
                            rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    elems = [Paragraph(title, styles["Title"]), Spacer(1, 12)]
    for heading, body in sections:
        elems.append(Paragraph(heading, styles["Heading2"]))
        elems.append(Spacer(1, 6))
        if isinstance(body, str):
            # wrap long body
            for para in body.split("\n\n"):
                elems.append(Paragraph(para, styles["BodyText"]))
                elems.append(Spacer(1, 6))
        elif isinstance(body, pd.DataFrame) and not body.empty:
            data = [list(body.columns)] + body.fillna("").values.tolist()
            t = Table(data, repeatRows=1)
            t.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
                ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
                ("ALIGN", (0,0), (-1,-1), "LEFT"),
                ("FONTSIZE", (0,0), (-1,-1), 8),
            ]))
            elems.append(t)
            elems.append(Spacer(1, 12))
        else:
            elems.append(Paragraph("No data.", styles["BodyText"]))
            elems.append(Spacer(1, 12))
    doc.build(elems)
    return filename

# -----------------------------
# Session state
# -----------------------------
if "records" not in st.session_state:
    st.session_state.records = []
if "sources" not in st.session_state:
    st.session_state.sources = {}  # source_name -> {"status": "Pending/Approved/Rejected", "notes": ""}
if "auditor_notes" not in st.session_state:
    st.session_state.auditor_notes = []  # list of dicts with message, ts

# -----------------------------
# UI - Sidebar
# -----------------------------
st.sidebar.title("GHG VB Prototype")
st.sidebar.markdown("**ISO 14064-3:2019** demo — not for production.")

with st.sidebar.expander("Engagement & Assurance", expanded=True):
    engagement_type = st.selectbox("Type of Engagement", ["Verification", "Validation"], index=0)
    assurance_level = st.selectbox("Level of Assurance", ["Limited", "Reasonable"], index=0)
    materiality_pct = st.slider("Materiality threshold (% of total footprint)", 1, 20, 5, help="Used for risk-based sampling in this demo.")

with st.sidebar.expander("Options"):
    grid_ef = st.number_input("Grid EF (kg CO2e/kWh, for Electricity)", value=EMISSION_FACTORS["electricity_kwh"], step=0.01)
    EMISSION_FACTORS["electricity_kwh"] = grid_ef
    st.caption("Adjust only for demonstration.")

st.sidebar.markdown("---")
st.sidebar.info("Upload multiple files: Excel, Word, PDF. You can also paste a URL below.")

# -----------------------------
# Main layout
# -----------------------------
st.title("GHG Verification Body – Prototype (Streamlit)")

col_up1, col_up2 = st.columns(2)
with col_up1:
    uploaded = st.file_uploader("Upload data files (Excel .xlsx, Word .docx, PDF .pdf). You can upload multiple files.", type=["xlsx", "xls", "csv", "docx", "pdf"], accept_multiple_files=True)
with col_up2:
    url_input = st.text_input("Optional: Enter a URL to parse (webpage with emissions data):", placeholder="https://example.com/sustainability")

parse_btn = st.button("Parse Inputs")

if parse_btn:
    for f in uploaded or []:
        source_name = f.name
        st.session_state.sources.setdefault(source_name, {"status": "Pending", "notes": ""})
        bytes_data = f.getvalue()

        if f.type in ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel"] or f.name.lower().endswith((".xlsx",".xls",".csv")):
            dfs = {}
            if f.name.lower().endswith(".csv"):
                try:
                    df = pd.read_csv(io.BytesIO(bytes_data))
                    dfs["CSV"] = df
                except Exception as e:
                    st.warning(f"CSV parse error for {f.name}: {e}")
            else:
                dfs = parse_excel(bytes_data)
            with st.expander(f"Parsed Excel: {f.name}"):
                for sname, df in dfs.items():
                    st.write(f"**Sheet:** {sname}")
                    st.dataframe(df, use_container_width=True)
                    recs = extract_emissions_records(df, f"{source_name}:{sname}")
                    st.session_state.records.extend(recs)

        elif f.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document" or f.name.lower().endswith(".docx"):
            parsed = parse_word(bytes_data)
            with st.expander(f"Parsed Word: {f.name}"):
                for i, tdf in enumerate(parsed["tables"]):
                    st.write(f"**Table {i+1}**")
                    st.dataframe(tdf, use_container_width=True)
                    recs = extract_emissions_records(tdf, f"{source_name}:table{i+1}")
                    st.session_state.records.extend(recs)
                if parsed["text"]:
                    st.write("**Paragraphs (first 10):**")
                    st.write(parsed["text"][:10])

        elif f.type == "application/pdf" or f.name.lower().endswith(".pdf"):
            parsed = parse_pdf(bytes_data)
            with st.expander(f"Parsed PDF: {f.name}"):
                for i, tdf in enumerate(parsed["tables"]):
                    st.write(f"**Table {i+1}**")
                    st.dataframe(tdf, use_container_width=True)
                    recs = extract_emissions_records(tdf, f"{source_name}:table{i+1}")
                    st.session_state.records.extend(recs)
                if parsed["lines"]:
                    st.write("**Detected Lines (first 15):**")
                    st.write(parsed["lines"][:15])
                    recs2 = extract_from_text_lines(parsed["lines"], source_name)
                    st.session_state.records.extend(recs2)
        else:
            st.warning(f"Unsupported file type for {f.name}")

    if url_input.strip():
        parsed = parse_url(url_input.strip())
        with st.expander("Parsed URL"):
            for i, tdf in enumerate(parsed["tables"]):
                st.write(f"**Table {i+1}**")
                st.dataframe(tdf, use_container_width=True)
                recs = extract_emissions_records(tdf, f"URL:{url_input}:table{i+1}")
                st.session_state.records.extend(recs)
            if parsed["text"]:
                st.write("**Extracted text sample (first 500 chars):**")
                st.write(parsed["text"][:500])

    st.success("Parsing complete. Scroll down to review extracted records.")

# -----------------------------
# Records review & approvals
# -----------------------------
st.header("1) Extracted Activity Records")
if st.session_state.records:
    df_records = pd.DataFrame(st.session_state.records)
    st.dataframe(df_records, use_container_width=True, height=300)
    st.caption("These are raw extracted records prior to standardization and quantification.")
else:
    st.info("No records yet. Upload files and click **Parse Inputs**.")

st.subheader("Source Approvals")
for source_name, meta in st.session_state.sources.items():
    with st.expander(source_name):
        c1, c2 = st.columns([1,2])
        with c1:
            status = st.selectbox("Status", ["Pending","Approved","Rejected"], index=["Pending","Approved","Rejected"].index(meta["status"]), key=f"status_{source_name}")
        with c2:
            notes = st.text_area("Reviewer notes", value=meta.get("notes",""), key=f"notes_{source_name}")
        st.session_state.sources[source_name]["status"] = status
        st.session_state.sources[source_name]["notes"] = notes

# -----------------------------
# Quantification
# -----------------------------
st.header("2) Standardize & Quantify")
quant_btn = st.button("Run Quantification")
if quant_btn:
    df_records = pd.DataFrame(st.session_state.records)
    if df_records.empty:
        st.warning("No records to quantify.")
    else:
        quantified = standardize_and_quantify(st.session_state.records)
        st.session_state["quantified"] = quantified
        st.success("Quantification complete.")

if "quantified" in st.session_state:
    quantified = st.session_state["quantified"]
    st.dataframe(quantified, use_container_width=True, height=320)

    scopes, cats, total = make_inventory_tables(quantified)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**By Scope (kg CO2e)**")
        st.dataframe(scopes, use_container_width=True)
    with c2:
        st.markdown("**By Category (kg CO2e)**")
        st.dataframe(cats, use_container_width=True)
    with c3:
        st.metric("Total Footprint (t CO2e)", f"{total/1000.0:,.2f}")

    # -----------------------------
    # Verification planning
    # -----------------------------
    st.header("3) Verification Plan")
    vplan = verification_plan(quantified, assurance_level, materiality_pct)
    st.dataframe(vplan, use_container_width=True)
    st.caption("Risk bands and sample sizes are demonstrative only.")

    # -----------------------------
    # Auditor annotations
    # -----------------------------
    st.header("4) Auditor Annotations & Evidence")
    note = st.text_area("Add a finding / observation / assumption:")
    if st.button("Add Note"):
        if note.strip():
            st.session_state.auditor_notes.append({"ts": dt.datetime.now().isoformat(timespec="seconds"), "note": note.strip()})
            st.success("Note added.")
    if st.session_state.auditor_notes:
        notes_df = pd.DataFrame(st.session_state.auditor_notes)
        st.dataframe(notes_df, use_container_width=True, height=200)
    evidence = st.file_uploader("Optional: Upload supporting evidence files (any type)", accept_multiple_files=True)
    if evidence:
        st.info(f"Evidence received: {', '.join([f.name for f in evidence])} (stored in memory for demo)")

    # -----------------------------
    # Report generation
    # -----------------------------
    st.header("5) Generate Reports")
    org_name = st.text_input("Reporting Entity Name", value="Demo Company LLC")
    reporting_year = st.text_input("Reporting Period", value="FY 2024 (01-Jul-2024 to 30-Jun-2025)")
    verifier_name = st.text_input("Verification Body", value="Demo Verification Body, Inc.")
    materiality_str = f"{materiality_pct}% of total inventory"

    if st.button("Build PDF Reports"):
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        inv_sections = [
            ("Organizational Boundaries",
             f"Entity: {org_name}\nOwnership/Control Approach: Operational Control\nReporting Period: {reporting_year}"),
            ("Operational Boundaries",
             "Scopes included: Scope 1 (stationary & mobile fuel), Scope 2 (purchased electricity), Scope 3 (Business Travel, Purchased Goods & Services, Employee Commuting, Third-Party Logistics)."),
            ("GHG Quantification",
             f"Emission factors are demonstrative defaults.\nMateriality Threshold: {materiality_str}\nAssurance Level: {assurance_level}\nEngagement Type: {engagement_type}"),
            ("By Scope (kg CO2e)", scopes),
            ("By Category (kg CO2e)", cats),
            ("Total", pd.DataFrame([{'Total kg CO2e': total, 'Total t CO2e': total/1000.0}])),
            ("Source Approvals", pd.DataFrame([{'source': s, **meta} for s, meta in st.session_state.sources.items()]))
        ]
        inv_path = f"GHG_Inventory_{ts}.pdf"
        inv_file = build_pdf(inv_path, f"{org_name} – GHG Inventory Report", inv_sections)

        vstmt_sections = [
            ("Verification Statement",
             f"This Verification Statement is provided by {verifier_name} to {org_name} for the period {reporting_year}. "
             f"The engagement type was {engagement_type} with {assurance_level} assurance. "
             f"The Verification Body applied risk-based procedures and materiality of {materiality_str}. "
             "Based on the procedures performed and evidence obtained, nothing has come to our attention that causes us to believe "
             "that the GHG assertion is materially misstated (demonstrative language for limited assurance)."),
            ("Scope & Boundary",
             "Scope 1, Scope 2, and selected Scope 3 categories were included as defined above."),
            ("Summary of Results",
             pd.DataFrame([{'Total t CO2e': total/1000.0, 'Assurance': assurance_level, 'Materiality %': materiality_pct}]))
        ]
        vstmt_path = f"Verification_Statement_{ts}.pdf"
        vstmt_file = build_pdf(vstmt_path, f"{org_name} – Verification Statement", vstmt_sections)

        svrep_sections = [
            ("Objective, Criteria, and Scope",
             f"Objective: Verify {org_name}'s GHG assertion.\nCriteria: ISO 14064-1/2/3 principles and GHG Protocol.\nScope: See Operational Boundaries."),
            ("Methodology",
             "Activities included planning, risk assessment, sampling based on category contribution, analytical procedures, "
             "inspection of underlying data, and review of calculations against emission factors."),
            ("Findings",
             "No material inconsistencies noted in this demonstrative example. Some assumptions applied to normalize units."),
            ("Verification Plan Snapshot", vplan),
            ("Auditor Notes (excerpt)",
             pd.DataFrame(st.session_state.auditor_notes) if st.session_state.auditor_notes else pd.DataFrame([{'note':'(none recorded in demo)'}]))
        ]
        svrep_path = f"Summary_Verification_Report_{ts}.pdf"
        svrep_file = build_pdf(svrep_path, f"{org_name} – Summary Verification Report", svrep_sections)

        st.success("Reports generated.")
        if inv_file.endswith(".pdf"):
            with open(inv_file, "rb") as f:
                st.download_button("Download GHG Inventory Report (PDF)", f, file_name=inv_file, mime="application/pdf")
        else:
            with open(inv_file, "r", encoding="utf-8") as f:
                st.download_button("Download GHG Inventory Report (TXT)", f, file_name=inv_file, mime="text/plain")

        if vstmt_file.endswith(".pdf"):
            with open(vstmt_file, "rb") as f:
                st.download_button("Download Verification Statement (PDF)", f, file_name=vstmt_file, mime="application/pdf")
        else:
            with open(vstmt_file, "r", encoding="utf-8") as f:
                st.download_button("Download Verification Statement (TXT)", f, file_name=vstmt_file, mime="text/plain")

        if svrep_file.endswith(".pdf"):
            with open(svrep_file, "rb") as f:
                st.download_button("Download Summary Verification Report (PDF)", f, file_name=svrep_file, mime="application/pdf")
        else:
            with open(svrep_file, "r", encoding="utf-8") as f:
                st.download_button("Download Summary Verification Report (TXT)", f, file_name=svrep_file, mime="text/plain")


st.markdown("---")
st.caption("This prototype uses illustrative emission factors and simplified verification logic for demonstration only.")
