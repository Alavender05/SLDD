import openpyxl
import os

def load_data():
    # 1. Get the directory where THIS script (AL.py) is located
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 2. Construct the full path to the Excel file
    # This ensures it works even if you run the script from a different folder
    filename = 'TSP_305041135.xlsx'
    file_path = os.path.join(script_dir, filename)

    print(f"Attempting to load: {file_path}")

    try:
        # 3. Load the workbook using openpyxl
        wb = openpyxl.load_workbook(file_path)
        
        # Select the active sheet
        sheet = wb.active
        
        print("Success: Workbook loaded.")
        
        # Example: Print the value of the first cell (A1) to verify
        print(f"Value in A1: {sheet['A1'].value}")
        
        return wb

    except FileNotFoundError:
        print(f"Error: The file '{filename}' was not found in {script_dir}")
        return None

if __name__ == "__main__":
    load_data()
