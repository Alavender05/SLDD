import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np
import io
import requests
from bs4 import BeautifulSoup
import re

# ==========================================
# 1. SHARED EXCEL STYLES & HELPERS
# ==========================================

def get_header_style():
    return {
        "font": Font(bold=True, color="FFFFFF", size=11),
        "fill": PatternFill("solid", fgColor="4472C4"),
        "border": Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin')),
        "alignment": Alignment(horizontal="center", vertical="center", wrap_text=True)
    }

def style_sheet_columns(ws):
    """Auto-adjust column widths."""
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[column].width = (max_length + 2)

# ==========================================
# 2. TSP ANALYSIS LOGIC (FROM UPLOAD)
# ==========================================

def get_value(wb, sheet_name, cell_ref):
    if sheet_name in wb.sheetnames:
        return wb[sheet_name][cell_ref].value
    return None

def clean_val(val):
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float, np.integer, np.floating)):
        return float(val)
    try:
        clean_str = str(val).replace("$", "").replace(",", "").replace("%", "").strip()
        return float(clean_str)
    except (ValueError, TypeError):
        return 0.0

def calc_growth(current, previous):
    prev_float = clean_val(previous)
    curr_float = clean_val(current)
    if prev_float == 0:
        return "N/A"
    growth = (curr_float - prev_float) / prev_float
    return f"{growth:.2%}"

def get_midpoint(label):
    if not isinstance(label, str): return 0
    clean = label.replace('$', '').replace(',', '').strip()
    if "Negative" in clean or "Nil" in clean: return 0
    elif "or more" in clean:
        try: return float(clean.replace(' or more', '')) * 1.1
        except: return 0
    elif "-" in clean:
        try:
            low, high = map(float, clean.split('-'))
            return (low + high) / 2
        except: return 0
    return 0

def add_conditional_formatting(ws, start_row, end_row, start_col_idx, end_col_idx):
    start_col = get_column_letter(start_col_idx)
    end_col = get_column_letter(end_col_idx)
    cell_range = f"{start_col}{start_row}:{end_col}{end_row}"
    ws.conditional_formatting.add(cell_range, ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='FF0000'))

def write_tsp_analysis_to_sheet(wb, uploaded_file):
    """Reads uploaded file and adds a 'TSP Analysis' sheet to the wb."""
    wb_source = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws_out = wb.create_sheet("TSP Analysis")

    # --- T02 ---
    if "T02" in wb_source.sheetnames:
        sheet = wb_source["T02"]
        ws_out.append(["--- T02 MEDIAN / AVERAGE METRICS ---"])
        ws_out.append(["Metric", "2011", "2016", "2021"])
        
        # Apply Header Style
        for cell in ws_out[2]: cell.font = Font(bold=True)

        rows = [15, 17, 19, 21, 23]
        for r in rows:
            if sheet[f"A{r}"].value: ws_out.append([sheet[f"A{r}"].value, sheet[f"B{r}"].value, sheet[f"C{r}"].value, sheet[f"D{r}"].value])
            if sheet[f"F{r}"].value: ws_out.append([sheet[f"F{r}"].value, sheet[f"G{r}"].value, sheet[f"H{r}"].value, sheet[f"I{r}"].value])
        ws_out.append([]); ws_out.append([])

    # --- Summary Table ---
    ws_out.append(["--- SUMMARY (2011 / 2016 / 2021) ---"])
    header_row = ["Metric", "2011", "2016", "2021", "Growth '11-'16", "Growth '16-'21", "Total Growth '11-'21"]
    ws_out.append(header_row)
    for cell in ws_out[ws_out.max_row]: cell.font = Font(bold=True)

    summary_items = [
        ("Total Persons Divorced", ("T04", "L28"), ("T04", "L48"), ("T04", "L68")),
        ("Separate House", ("T14a", "J13"), ("T14b", "J13"), ("T14c", "J13")),
        ("Flat or Apartment", ("T14a", "J26"), ("T14b", "J26"), ("T14c", "J26")),
        ("Owned Outright", ("T18", "G15"), ("T18", "G34"), ("T18", "G53")),
        ("Owned with a Mortgage", ("T18", "G16"), ("T18", "G35"), ("T18", "G54")),
        ("Rented", ("T18", "G25"), ("T18", "G44"), ("T18", "G63")),
        ("Employed Worked Full Time", ("T29", "D15"), ("T29", "H15"), ("T29", "L15")),
        ("Unemployment %", ("T29", "D23"), ("T29", "H23"), ("T29", "L23")),
        ("Labour Force Participation", ("T29", "D24"), ("T29", "H24"), ("T29", "L24")),
    ]
    for metric, s11, s16, s21 in summary_items:
        v11, v16, v21 = get_value(wb_source, *s11), get_value(wb_source, *s16), get_value(wb_source, *s21)
        ws_out.append([metric, v11, v16, v21, calc_growth(v16, v11), calc_growth(v21, v16), calc_growth(v21, v11)])
    ws_out.append([]); ws_out.append([])

    # --- T24 Matrix ---
    if "T24" in wb_source.sheetnames:
        sheet = wb_source["T24"]
        ws_out.append(["--- DATA FROM T24 (Income x Rent Matrix) ---"])
        rent_labels = ["$1-$74", "$75-$99", "$100-$149", "$150-$199", "$200-$224", "$225-$274", "$275-$349", "$350-$449", "$450-$549", "$550-$649", "$650 or more"]
        ws_out.append(["Income Range"] + rent_labels)
        for cell in ws_out[ws_out.max_row]: cell.font = Font(bold=True)
        
        start_row = ws_out.max_row + 1
        raw_rows = []
        for row in sheet.iter_rows(min_row=55, max_row=71, min_col=1, max_col=14, values_only=True):
            lbl = "" if row[0] is None else str(row[0]).strip()
            if lbl and "CENSUS" not in lbl.upper():
                clean_row = [lbl] + [0 if v in (None, "") else v for v in row[1:12]]
                raw_rows.append(clean_row)
                ws_out.append(clean_row)
        
        if raw_rows:
            numeric_matrix = [[clean_val(v) for v in r[1:]] for r in raw_rows]
            totals = ["TOTAL"] + list(np.sum(numeric_matrix, axis=0))
            ws_out.append(totals)
            # Add Heatmap
            add_conditional_formatting(ws_out, start_row, start_row + len(raw_rows) - 1, 2, 12)

    style_sheet_columns(ws_out)

# ==========================================
# 3. SCRAPING LOGIC (ONLINE DATA)
# ==========================================

BASE_URLS = {
    2011: "https://www.abs.gov.au/census/find-census-data/quickstats/2011/{}",
    2016: "https://www.abs.gov.au/census/find-census-data/quickstats/2016/{}",
    2021: "https://www.abs.gov.au/census/find-census-data/quickstats/2021/{}",
}

METRICS = [
    {"name": "Average number of people per household", "unit": "", "variants": ["Average number of people per household", "Average people per household"]},
    {"name": "Median weekly household income", "unit": "$", "variants": ["Median weekly household income"]},
    {"name": "Less than $650 total household weekly income (a)", "unit": "%", "variants": ["Less than $650 total household weekly income (a)"]},
    {"name": "More than $3,000 total household weekly income (a)", "unit": "%", "variants": ["More than $3,000 total household weekly income (a)"]},
    {"name": "Median monthly mortgage repayments", "unit": "$", "variants": ["Median monthly mortgage repayments"]},
    {"name": "Owned outright", "unit": "%", "variants": ["Owned outright", "Owned Outright"]},
    {"name": "Owned with a mortgage", "unit": "%", "variants": ["Owned with a mortgage", "Owned with a Mortgage"]},
    {"name": "Rented", "unit": "%", "variants": ["Rented"]},
    {"name": "Median weekly rent", "unit": "$", "variants": ["Median weekly rent", "Median weekly rent (a)", "Median weekly rent (b)"]},
    {"name": "Rent payments less than 30% of household income", "unit": "%", "variants": ["Renter households where rent payments are less than or equal to 30% of household income (b)", "Households where rent payments are less than 30% of household income"]},
    {"name": "Rent payments 30% or more of household income", "unit": "%", "variants": ["Renter households with rent payments greater than 30% of household income (b)", "Households where rent payments are 30%, or greater, of household income", "Households with rent payments greater than or equal to 30% of household income"]},
    {"name": "Mortgage payments less than 30% of household income", "unit": "%", "variants": ["Owner with mortgage households where mortgage repayments are less than or equal to 30% of household income (a)", "Households where mortgage payments are less than 30% of household income", "Households where mortgage repayments are less than 30% of household income"]},
    {"name": "Mortgage payments 30% or more of household income", "unit": "%", "variants": ["Owner with mortgage households with mortgage repayments greater than 30% of household income (a)", "Households where mortgage payments are 30%, or greater, of household income"]},
    {"name": "Worked full-time", "unit": "%", "variants": ["Worked full-time"]},
    {"name": "Separate house", "unit": "%", "variants": ["Separate house"]},
    {"name": "Flat, unit or apartment", "unit": "%", "variants": ["Flat or apartment", "Flat, unit or apartment"]},
]

def get_quickstats_tables(area_code, year):
    url = BASE_URLS[year].format(area_code)
    try:
        resp = requests.get(url, timeout=15)
        resp.raise_for_status()
    except Exception as e:
        return None, None, None
    soup = BeautifulSoup(resp.text, "html.parser")
    tables = soup.find_all("table")
    h1 = soup.find("h1")
    area_name = h1.get_text(strip=True) if h1 else f"Area {area_code}"
    return tables, area_name, url

def extract_metric_value(tables, variants):
    if not tables: return None
    lower_variants = [v.lower() for v in variants]
    for table in tables:
        for tr in table.find_all("tr"):
            cells = [c.get_text(strip=True) for c in tr.find_all(["th", "td"])]
            if cells:
                label = cells[0].strip().lower()
                if any(v in label for v in lower_variants):
                    return cells[2].strip() if len(cells) > 2 else (cells[1].strip() if len(cells) > 1 else None)
    return None

def extract_all_metrics(area_code, year):
    tables, area_name, url = get_quickstats_tables(area_code, year)
    if tables is None: return None
    result = {"area_code": area_code, "area_name": area_name, "year": year, "url": url}
    for m in METRICS:
        result[m["name"]] = extract_metric_value(tables, m["variants"])
    return result

def write_scraped_data_to_sheet(wb, data_dict):
    """Adds a 'Online QuickStats' sheet to the wb."""
    ws = wb.create_sheet("Online QuickStats")
    
    # Styles
    styles = get_header_style()
    
    # Header
    headers = ["Metric", "Unit", "", "2011", "2016", "2021"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = styles["fill"]
        cell.font = styles["font"]
        cell.alignment = styles["alignment"]
        cell.border = styles["border"]

    # Info Block
    latest_data = data_dict.get(2021, data_dict.get(2016, data_dict.get(2011)))
    if latest_data:
        ws.append(["Area Code", "", "", latest_data["area_code"], "", ""])
        ws.append(["Area Name", "", "", latest_data["area_name"], "", ""])
        ws.append(["Source URL", "", "", latest_data["url"], "", ""])
        ws.append([]) 

    # Metrics
    for m in METRICS:
        row = [m["name"], m["unit"], ""]
        for year in [2011, 2016, 2021]:
            val = data_dict.get(year, {}).get(m["name"], "—")
            row.append(val if val else "—")
        ws.append(row)

    ws.column_dimensions["A"].width = 50
    for col in ["D", "E", "F"]: ws.column_dimensions[col].width = 15

# ==========================================
# 4. STREAMLIT APP LOGIC
# ==========================================

def main():
    st.set_page_config(page_title="Census Data Tool", layout="wide", page_icon="📊")
    st.title("📊 Australian Census Data Combiner")
    st.write("This tool performs two actions and merges them into a single Excel file:")
    st.markdown("1. **Analyzes an uploaded TSP Excel file** (calculates growth, heatmaps, etc.)")
    st.markdown("2. **Scrapes Online ABS Data** (fetches summary stats for a given Area Code)")

    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. File Upload")
        uploaded_file = st.file_uploader("Upload TSP_*.xlsx file", type=["xlsx"])
    
    with col2:
        st.subheader("2. Online Data")
        area_code_input = st.text_input("Enter ABS Area Code (e.g. 3GBRI):", help="Leave empty if you only want to process the file.").strip().upper()

    if st.button("Generate Combined Report", type="primary"):
        if not uploaded_file and not area_code_input:
            st.error("Please provide at least a File OR an Area Code.")
            return

        # Initialize Output Workbook
        wb_out = openpyxl.Workbook()
        # Remove default sheet created by openpyxl
        if "Sheet" in wb_out.sheetnames:
            del wb_out["Sheet"]

        # 1. PROCESS FILE
        if uploaded_file:
            with st.spinner("Processing uploaded file..."):
                try:
                    write_tsp_analysis_to_sheet(wb_out, uploaded_file)
                    st.success("✅ File Analysis Complete")
                except Exception as e:
                    st.error(f"Error processing file: {e}")

        # 2. PROCESS SCRAPER
        if area_code_input:
            with st.spinner(f"Scraping online data for {area_code_input}..."):
                data_by_year = {}
                years = [2011, 2016, 2021]
                progress_bar = st.progress(0)
                
                valid_data = False
                for i, year in enumerate(years):
                    data = extract_all_metrics(area_code_input, year)
                    if data: 
                        data_by_year[year] = data
                        valid_data = True
                    progress_bar.progress((i + 1) / len(years))
                
                if valid_data:
                    write_scraped_data_to_sheet(wb_out, data_by_year)
                    st.success(f"✅ Online Data Scraped for {area_code_input}")
                else:
                    st.warning(f"Could not find online data for {area_code_input}")

        # 3. SAVE AND DOWNLOAD
        if len(wb_out.sheetnames) > 0:
            buffer = io.BytesIO()
            wb_out.save(buffer)
            buffer.seek(0)
            
            file_label = "Combined_Report.xlsx"
            if uploaded_file and not area_code_input: file_label = "TSP_Analysis_Report.xlsx"
            if area_code_input and not uploaded_file: file_label = f"{area_code_input}_Census_Data.xlsx"

            st.download_button(
                label="📥 Download Final Excel Report",
                data=buffer,
                file_name=file_label,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No data was generated. Please check your inputs.")

if __name__ == "__main__":
    main()