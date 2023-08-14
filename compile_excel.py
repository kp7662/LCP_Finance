import os
import pandas as pd
import warnings

# Suppress UserWarnings from openpyxl
warnings.simplefilter("ignore")

# Set the directory containing your Excel files
folder_path = r"C:\Users\User\Desktop\Merge\excel_inputs"

# Get a list of all Excel files (both .xlsx and .xls) in the folder
excel_files = [file for file in os.listdir(folder_path) if file.endswith((".xlsx", ".xls"))]

# Create a new Excel writer object
output_file = "combined_excel_file.xlsx"
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

# Loop through each Excel file and copy its sheet to the combined file
for file in excel_files:
    file_path = os.path.join(folder_path, file)
    excel_data = pd.ExcelFile(file_path)
    
    print(f"Processing file: {file}")
    
    for sheet_name in excel_data.sheet_names:
        print(f"Processing sheet: {sheet_name}")
        
        combined_sheet_name = f"{file[:-5]}_{sheet_name}"
        
        # Truncate the sheet name if it's too long
        if len(combined_sheet_name) > 31:
            combined_sheet_name = combined_sheet_name[:31]
            
        df = excel_data.parse(sheet_name)
        df.to_excel(writer, sheet_name=combined_sheet_name, index=False)

# Save the combined Excel file using _save() method
writer._save()

print("Excel files combined successfully.")
