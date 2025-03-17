import clr
import openpyxl

# Inputs from Dynamo
excel_file = IN[0]  
sheet_name = IN[1]  

try:
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(excel_file, data_only=True)  # data_only=True reads cell values (not formulas)
    
    # Check if sheet exists
    if sheet_name not in workbook.sheetnames:
        OUT = f" Error: Sheet '{sheet_name}' not found in the file."
    else:
        sheet = workbook[sheet_name]  # Get the specified sheet

        # Extract all rows and columns into a list of lists
        excel_data = []
        for row in sheet.iter_rows(values_only=True):
            excel_data.append([cell if cell is not None else None for cell in row])  # Replace empty cells with None

        # Output the extracted data
        OUT = excel_data

except Exception as e:
    OUT = f"Error: {e}"
