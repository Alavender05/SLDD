import openpyxl
import os
import csv
import pandas as pd
import numpy as np

# ==========================================
# 1. HELPER FUNCTIONS
# ==========================================

def get_value(wb, sheet_name, cell_ref):
    """Helper to get a value safely."""
    if sheet_name in wb.sheetnames:
        return wb[sheet_name][cell_ref].value
    return None

def clean_val(val):
    """Converts Excel value to float, handling None and strings."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        # Remove common currency/percentage symbols
        clean_str = str(val).replace('$', '').replace(',', '').replace('%', '').strip()
        return float(clean_str)
    except ValueError:
        return 0.0

def calc_growth(current, previous):
    """Calculates percentage growth between two values."""
    prev_float = clean_val(previous)
    curr_float = clean_val(current)
    
    if prev_float == 0:
        return "N/A" # Avoid division by zero
    
    growth = (curr_float - prev_float) / prev_float
    return f"{growth:.2%}"

def get_midpoint(label):
    """Calculates the numerical midpoint of a range string."""
    if not isinstance(label, str):
        return 0
    clean = label.replace('$', '').replace(',', '').strip()
    
    if "Negative" in clean or "Nil" in clean:
        return 0
    elif "or more" in clean:
        try:
            lower = float(clean.replace(' or more', ''))
            return lower * 1.1 
        except ValueError:
            return 0
    elif "-" in clean:
        try:
            low, high = map(float, clean.split('-'))
            return (low + high) / 2
        except ValueError:
            return 0
    return 0

def calculate_rent_stress_stats(raw_data_rows, income_labels, rent_labels):
    """Performs the matrix math: Proportions -> Weighted Percentiles."""
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
                if inc > 0:
                    proportion_matrix[r, c] = (rent / inc)
                else:
                    proportion_matrix[r, c] = np.nan

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
            
            p25 = np.percentile(expanded_ratios, 25)
            p50 = np.percentile(expanded_ratios, 50)
            p75 = np.percentile(expanded_ratios, 75)
            
            percentiles_data.append([col_name, p25, p50, p75])

        df_percentiles = pd.DataFrame(
            percentiles_data, 
            columns=['Rent Range', '25th Percentile', 'Median', '75th Percentile']
        )
        return df_proportions, df_percentiles

    except Exception as e:
        print(f"Error in calculation: {e}")
        return None, None

# ==========================================
# 2. MAIN EXPORT FUNCTION
# ==========================================

def export_data_to_csv():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_filename = 'TSP_305041135.xlsx'
    csv_filename = 'extracted_data.csv'
    
    excel_path = os.path.join(script_dir, excel_filename)
    csv_path = os.path.join(script_dir, csv_filename)

    print(f"Reading from: {excel_filename}")
    print(f"Writing to:   {csv_filename}\n")

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        
        with open(csv_path, mode='w', newline='') as f:
            writer = csv.writer(f)

            # ---------------------------------------------------------
            # SECTION 1: T02 MATRIX
            # ---------------------------------------------------------
            if 'T02' in wb.sheetnames:
                writer.writerow(["--- DATA FROM T02 (Matrix A12:I21) ---"])
                sheet = wb['T02']
                for row in sheet.iter_rows(min_row=12, max_row=21, min_col=1, max_col=9, values_only=True):
                    writer.writerow(["" if cell is None else cell for cell in row])
                writer.writerow([]) 

            # ---------------------------------------------------------
            # SECTION 2: TIME SERIES PANELS (Divorced, Tenure, Labour)
            # ---------------------------------------------------------
            writer.writerow(["--- Time Series Analysis (Growth Rates) ---"])
            
            # UPDATED HEADER: Added "Total Growth '11-'21"
            writer.writerow(["Metric", "2011 Value", "2016 Value", "2021 Value", "Growth '11-'16", "Growth '16-'21", "Total Growth '11-'21"])

            time_series_items = [
                # Divorced
                ("Total Persons Divorced",       ('T04', 'L28'),  ('T04', 'L48'),  ('T04', 'L68')),
                
                # Tenure Types
                ("Separate House",               ('T14a','J13'),  ('T14b','J13'),  ('T14c','J13')),
                ("Flat or Apartment",            ('T14a','J26'),  ('T14b','J26'),  ('T14c','J26')),
                ("Owned Outright",               ('T18', 'G15'),  ('T18', 'G34'),  ('T18', 'G53')),
                ("Owned with a Mortgage",        ('T18', 'G16'),  ('T18', 'G35'),  ('T18', 'G54')),
                ("Rented",                       ('T18', 'G25'),  ('T18', 'G44'),  ('T18', 'G63')),
                
                # Labour Force
                ("Employed Worked Full Time",    ('T29', 'D15'),  ('T29', 'H15'),  ('T29', 'L15')),
                ("Unemployment %",               ('T29', 'D23'),  ('T29', 'H23'),  ('T29', 'L23')),
                ("Labour Force Participation",   ('T29', 'D24'),  ('T29', 'H24'),  ('T29', 'L24')),
            ]

            for metric, (s11, c11), (s16, c16), (s21, c21) in time_series_items:
                val11 = get_value(wb, s11, c11)
                val16 = get_value(wb, s16, c16)
                val21 = get_value(wb, s21, c21)

                growth_11_16 = calc_growth(val16, val11)
                growth_16_21 = calc_growth(val21, val16)
                
                # NEW CALCULATION: Total Growth (2021 vs 2011)
                growth_11_21 = calc_growth(val21, val11)

                writer.writerow([metric, val11, val16, val21, growth_11_16, growth_16_21, growth_11_21])
            
            writer.writerow([]) # Spacer

            # ---------------------------------------------------------
            # SECTION 3: T24 MATRIX & ANALYSIS (A55:N71)
            # ---------------------------------------------------------
            if 'T24' in wb.sheetnames:
                writer.writerow(["--- DATA FROM T24 (Matrix A55:N71) ---"])
                sheet = wb['T24']
                
                raw_rows_extracted = []
                for row in sheet.iter_rows(min_row=55, max_row=71, min_col=1, max_col=14, values_only=True):
                    clean_row = ["" if cell is None else cell for cell in row]
                    raw_rows_extracted.append(clean_row)
                    writer.writerow(clean_row)
                
                print("Extracted T24 data.")
                writer.writerow([]) 

                # ANALYSIS
                if len(raw_rows_extracted) >= 15:
                    rent_labels = [
                        "$1-$74", "$75-$99", "$100-$149", "$150-$199", "$200-$224", 
                        "$225-$274", "$275-$349", "$350-$449", "$450-$549", 
                        "$550-$649", "$650 or more"
                    ]
                    
                    analysis_matrix = []
                    income_labels = []

                    for i in range(15):
                        row = raw_rows_extracted[i]
                        income_labels.append(row[0]) 
                        row_data = row[1:12] # Cols B-L
                        row_data = [0 if (x == "" or x is None) else x for x in row_data]
                        analysis_matrix.append(row_data)

                    df_props, df_stats = calculate_rent_stress_stats(analysis_matrix, income_labels, rent_labels)

                    if df_props is not None:
                        writer.writerow(["--- T24 Proportion Heatmap (Rent / Income) ---"])
                        writer.writerow(["Income Range"] + list(df_props.columns))
                        for index, row in df_props.iterrows():
                            formatted_row = [f"{x:.2%}" if not np.isnan(x) else "N/A" for x in row]
                            writer.writerow([index] + formatted_row)
                        writer.writerow([])

                    if df_stats is not None:
                        writer.writerow(["--- T24 Weighted Percentiles (Rent Stress) ---"])
                        writer.writerow(list(df_stats.columns))
                        for index, row in df_stats.iterrows():
                            formatted_row = []
                            for x in row:
                                if isinstance(x, (int, float)):
                                    formatted_row.append(f"{x:.2%}")
                                else:
                                    formatted_row.append(x)
                            writer.writerow(formatted_row)
                else:
                    print("Warning: Not enough rows in T24 range to perform analysis.")
            else:
                print("Warning: Sheet 'T24' not found.")

        print(f"\nSuccess! Open '{csv_filename}' to see the updated report.")

    except FileNotFoundError:
        print(f"Error: The file '{excel_filename}' was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    export_data_to_csv()
