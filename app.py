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
# PART 1: FILE ANALYSIS LOGIC (From Previous Code)
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

def calculate_rent_stress_stats(raw_data_rows, income_labels, rent_labels):
    try:
        df_counts = pd.DataFrame(raw_data_rows, index=income_labels, columns=rent_labels)
        df_counts = df_counts.apply(pd.to_numeric, errors='coerce').fillna(0)
        income_mids = [get_midpoint(l) for l in income_labels]
        rent_mids = [get_midpoint(l) for l in rent_labels]

        proportion_matrix = np.zeros(df_counts.shape)
        for r in range(len(income_labels)):
            for c in range(len(rent_labels)):
                inc = income_mids[r]
                rent = rent_mids[c]
                if inc > 0: proportion_matrix[r, c] = (rent / inc)
                else: proportion_matrix[r, c] = np.nan

        df_proportions = pd.DataFrame(proportion_matrix, index=income_labels, columns=rent_labels)
        percentiles_data = []
        for c, col_name in enumerate(rent_labels):
            counts = df_counts.iloc[:, c].values
            ratios = df_proportions.iloc[:, c].values
            mask = (counts > 0) & (~np.isnan(ratios))
            valid_counts = counts[mask]
            valid_ratios = ratios[mask]
            
            if len(valid_counts) == 0:
                percentiles_data.append([col_name, np.nan, np.nan, np.nan])
                continue
            
            expanded_ratios = np.repeat(valid_ratios, valid_counts.astype(int))
            percentiles_data.append([
                col_name, 
                np.percentile(expanded_ratios, 25), 
                np.percentile(expanded_ratios, 50), 
                np.percentile(expanded_ratios, 75)
            ])

        df_percentiles = pd.DataFrame(percentiles_data, columns=['Rent Range', '25th Percentile', 'Median', '75th Percentile'])
        return df_proportions, df_percentiles
    except: return None, None

def style_output_sheet(ws):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and str(cell.value).startswith("---"):
                cell.font = Font(bold=True)
            if cell.value in ["Metric", "Income Range", "Description", "Year"]:
                cell.font = Font(bold=True)
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except: pass
        ws.column_dimensions[column].width = (max_length + 2)

def add_conditional_formatting(ws, start_row, end_row, start_col_idx, end_col_idx):
    start_col = get_column_letter(start_col_idx)
    end_col = get_column_letter(end_col_idx)
    cell_range = f"{start_col}{start_row}:{end_col}{end_row}"
    ws.conditional_formatting.add(cell_range, ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='FF0000'))

def process_uploaded_file(uploaded_file):
    wb_source = openpyxl.load_workbook(uploaded_file, data_only=True)
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Summary Report"

    # T02
    if "T02" in wb_source.sheetnames:
        sheet = wb_source["T02"]
        ws_out.append(["--- T02 MEDIAN / AVERAGE METRICS ---"])
        ws_out.append(["Metric", "2011", "2016", "2021"])
        rows = [15, 17, 19, 21, 23]
        for r in rows:
            if sheet[f"A{r}"].value: ws_out.append([sheet[f"A{r}"].value, sheet[f"B{r}"].value, sheet[f"C{r}"].value, sheet[f"D{r}"].value])
            if sheet[f"F{r}"].value: ws_out.append([sheet[f"F{r}"].value, sheet[f"G{r}"].value, sheet[f"H{r}"].value, sheet[f"I{r}"].value])
        ws_out.append([]); ws_out.append([])

    # Summary Table
    ws_out.append(["--- SUMMARY (2011 / 2016 / 2021) ---"])
    ws_out.append(["Metric", "2011", "2016", "2021", "Growth '11-'16", "Growth '16-'21", "Total Growth '11-'21"])
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

    # T24 Matrix
    if "T24" in wb_source.sheetnames:
        sheet = wb_source["T24"]
        ws_out.append(["--- DATA FROM T24 (Income x Rent Matrix) ---"])
        rent_labels = ["$1-$74", "$75-$99", "$100-$149", "$150-$199", "$200-$224", "$225-$274", "$275-$349", "$350-$449", "$450-$549", "$550-$649", "$650 or more"]
        ws_out.append(["Income Range"] + rent_labels)
        
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
            add_conditional_formatting(ws_out, start_row, start_row + len(raw_rows) - 1, 2, 12)

    style_output_sheet(ws_out)
    buffer = io.BytesIO()
    wb_out.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# PART 2: SCRAPING LOGIC (New Request)
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

def create_scraped_excel(data_dict, area_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Census Data"
    
    # Styles
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="4472C4")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Header
    headers = ["Metric", "Unit", "", "2011", "2016", "2021"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    # Info Block
    latest_data = data_dict.get(2021, data_dict.get(2016, data_dict.get(2011)))
    if latest_data:
        ws.append(["Area Code", "", "", latest_data["area_code"], "", ""])
        ws.append(["Area Name", "", "", latest_data["area_name"], "", ""])
        ws.append(["Source URL", "", "", latest_data["url"], "", ""])
        ws.append([]) # Spacer

    # Metrics
    for m in METRICS:
        row = [m["name"], m["unit"], ""]
        for year in [2011, 2016, 2021]:
            val = data_dict.get(year, {}).get(m["name"], "—")
            row.append(val if val else "—")
        ws.append(row)

    # Formatting
    ws.column_dimensions["A"].width = 50
    for col in ["D", "E", "F"]: ws.column_dimensions[col].width = 15
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, f"{area_name}_Census_Data.xlsx"

# ==========================================
# 3. STREAMLIT APP LAYOUT
# ==========================================

def main():
    st.set_page_config(page_title="Census Tools", layout="wide", page_icon="📊")
    st.title("📊 Australian Census Data Tools")

    tab1, tab2 = st.tabs(["📂 **Analyze Uploaded File**", "🌐 **Scrape ABS QuickStats**"])

    # --- TAB 1: FILE UPLOAD ---
    with tab1:
        st.markdown("### Upload TSP Data (.xlsx)")
        st.write("Upload a raw Time Series Profile (TSP) Excel file to generate an analyzed report with Heatmaps and Growth Rates.")
        
        uploaded_file = st.file_uploader("Drag and drop Excel file here", type=["xlsx"], key="file_upload")
        
        if uploaded_file:
            if st.button("Process File", key="btn_process"):
                with st.spinner("Analyzing data..."):
                    try:
                        result_excel = process_uploaded_file(uploaded_file)
                        st.success("Analysis Complete!")
                        st.download_button(
                            label="📥 Download Analysis Report",
                            data=result_excel,
                            file_name="TSP_Analysis_Report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"Error: {e}")

    # --- TAB 2: SCRAPER ---
    with tab2:
        st.markdown("### Scrape Online Data")
        st.write("Enter an ABS Area Code (e.g., `3GBRI` or `POA2000`) to fetch summary statistics for 2011, 2016, and 2021.")
        
        area_code_input = st.text_input("Enter ABS Area Code:", placeholder="e.g. 3GBRI").strip().upper()
        
        if st.button("Fetch Data", key="btn_scrape"):
            if not area_code_input:
                st.warning("Please enter an Area Code.")
            else:
                data_by_year = {}
                progress_bar = st.progress(0)
                status_text = st.empty()

                years = [2011, 2016, 2021]
                for i, year in enumerate(years):
                    status_text.text(f"Fetching data for {year}...")
                    data = extract_all_metrics(area_code_input, year)
                    if data:
                        data_by_year[year] = data
                    progress_bar.progress((i + 1) / len(years))
                
                status_text.text("Processing complete.")
                
                if data_by_year:
                    # Determine Area Name safely
                    any_year = sorted(data_by_year.keys())[-1]
                    area_name = data_by_year[any_year]["area_name"]
                    
                    # Create Excel
                    excel_buffer, filename = create_scraped_excel(data_by_year, area_name)
                    
                    st.success(f"Successfully scraped data for: **{area_name}**")
                    st.download_button(
                        label=f"📥 Download {filename}",
                        data=excel_buffer,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error(f"Could not find any data for Area Code: {area_code_input}. Please check the code and try again.")

if __name__ == "__main__":
    main()
