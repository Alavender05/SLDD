import openpyxl

def get_single_cell_text(file_path, sheet_name, cell_address):
    """
    Retrieves data from a single cell and returns it as a text string.
    """
    try:
        # Load the workbook (data_only=True ensures we get values, not formulas)
        wb = openpyxl.load_workbook(file_path, data_only=True)
        
        # Access the specific sheet
        if sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            
            # Retrieve the value
            cell_value = sheet[cell_address].value
            
            # Convert to string (handle None if cell is empty)
            result_text = str(cell_value) if cell_value is not None else ""
            return result_text
        else:
            return f"Error: Sheet '{sheet_name}' not found."

    except FileNotFoundError:
        return f"Error: File '{file_path}' not found."
    except Exception as e:
        return f"An error occurred: {e}"

# --- Configuration ---
file_name = 'TSP_305041135.xlsx'
target_sheet = 'T01'  #sheet
target_cell = 'B17'   #Cell

# --- Execution ---
cell_data = get_single_cell_text(file_name, target_sheet, target_cell)

# --- Output ---
print(f"Data from {target_sheet}!{target_cell}: {cell_data}")