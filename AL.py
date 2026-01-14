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
    return "SHEET NOT FOUND"

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
            return lower * 1.1  # Estimate upper bound
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
        # Create DataFrame from the raw counts
        df_counts = pd.DataFrame(raw_data_rows, index=income_labels, columns=rent_labels)
        df_counts = df_counts.apply(pd.to_numeric, errors='coerce').fillna(0)

        # Midpoint Estimates
        income_mids = [get_midpoint(l) for l in income_labels]
        rent_mids = [get_midpoint(l) for l in rent_labels]

        # Proportion Matrix (Rent as % of Income)
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

        # Weighted Percentiles Table
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
    # Setup file paths
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
            # SECTION 1: T02 MATRIX (A12:I21)
            # ---------------------------------------------------------
            if 'T02' in wb.sheetnames:
                writer.writerow(["--- DATA FROM T02 (Matrix A12:I21) ---"])
                sheet = wb['T02']
                for row in sheet.iter_rows(min_row=12, max_row=21, min_col=1, max_col=9, values_only=True):
                    clean_row = ["" if cell is None else cell for cell in row]
                    writer.writerow(clean_row)
                writer.writerow([]) 
            
            # ---------------------------------------------------------
            # SECTION 2: TOTAL PERSONS DIVORCED
            # ---------------------------------------------------------
            writer.writerow(["--- Total Persons Divorced ---"])
            writer.writerow(["Year", "Value"])

            divorced_data = [
                ("2011", "T04", "L28"),
                ("2016", "T04", "L48"),
                ("2021", "T04", "L68"),
            ]

            for label, sheet, cell in divorced_data:
                val = get_value(wb, sheet, cell)
                writer.writerow([label, val])
            writer.writerow([]) 

            # ---------------------------------------------------------
            # SECTION 3: TENURE TYPES
            # ---------------------------------------------------------
            writer.writerow(["--- Tenure Types ---"])
            writer.writerow(["Description", "Value (Persons)"])

            tenure_data = [
                # 2011
                ("2011 Separate House",        "T14a", "J13"),
                ("2011 Flat or Apartment",     "T14a", "J26"),
                ("2011 Owned Outright",        "T18",  "G15"),
                ("2011 Owned with a mortgage", "T18",  "G16"),
                ("2011 Rented",                "T18",  "G25"),
                
                # 2016
                ("2016 Separate House",        "T14b", "J13"),
                ("2016 Flat or Apartment",     "T14b", "J26"),
                ("2016 Owned Outright",        "T18",  "G34"),
                ("2016 Owned with a mortgage", "T18",  "G35"),
                ("2016 Rented",                "T18",  "G44"),

                # 2021
                ("2021 Separate House",        "T14c", "J13"),
                ("2021 Flat or Apartment",     "T14c", "J26"),
                ("2021 Owned Outright",        "T18",  "G53"),
                ("2021 Owned with a mortgage", "T18",  "G54"),
                ("2021 Rented",                "T18",  "G63"),
            ]

            for label, sheet, cell in tenure_data:
                val = get_value(wb, sheet, cell)
                writer.writerow([label, val])
            writer.writerow([]) 

            # ---------------------------------------------------------
            # SECTION 4: T24 MATRIX & ANALYSIS (A55:N71)
            # ---------------------------------------------------------
            if 'T24' in wb.sheetnames:
                writer.writerow(["--- DATA FROM T24 (Matrix A55:N71) ---"])
                sheet = wb['T24']
                
                # 1. Fetch the raw rows specified by the user (A55:N71)
                raw_rows_extracted = []
                for row in sheet.iter_rows(min_row=55, max_row=71, min_col=1, max_col=14, values_only=True):
                    clean_row = ["" if cell is None else cell for cell in row]
                    raw_rows_extracted.append(clean_row)
                    writer.writerow(clean_row)
                
                print("Extracted T24 data (Rows 55-71).")
                writer.writerow([]) 

                # 2. PERFORM ANALYSIS (Heatmap & Percentiles)
                # We need the first 15 rows of the extracted block for the income buckets
                # (Negative/Nil up to $4,000 or more)
                if len(raw_rows_extracted) >= 15:
                    
                    # Define Labels
                    rent_labels = [
                        "$1-$74", "$75-$99", "$100-$149", "$150-$199", "$200-$224", 
                        "$225-$274", "$275-$349", "$350-$449", "$450-$549", 
                        "$550-$649", "$650 or more"
                    ]
                    
                    # Slice the matrix: First 15 rows, Columns B-L (Indices 1-11)
                    analysis_matrix = []
                    income_labels = []

                    for i in range(15):
                        row = raw_rows_extracted[i]
                        income_labels.append(row[0]) # Column A is the label
                        
                        # Get columns 1 to 11 (B to L) for the Rent Counts
                        # Handle potential empty strings/None
                        row_data = row[1:12]
                        row_data = [0 if (x == "" or x is None) else x for x in row_data]
                        analysis_matrix.append(row_data)

                    # Run Calculations
                    df_props, df_stats = calculate_rent_stress_stats(analysis_matrix, income_labels, rent_labels)

                    # Write Heatmap
                    if df_props is not None:
                        writer.writerow(["--- T24 Proportion Heatmap (Rent / Income) ---"])
                        writer.writerow(["Income Range"] + list(df_props.columns))
                        for index, row in df_props.iterrows():
                            formatted_row = [f"{x:.2%}" if not np.isnan(x) else "N/A" for x in row]
                            writer.writerow([index] + formatted_row)
                        writer.writerow([])

                    # Write Percentiles
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
