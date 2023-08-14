import os
import pandas as pd
import streamlit as st
from io import BytesIO
import base64
import warnings

# Suppress UserWarnings from openpyxl
warnings.simplefilter("ignore")

# Define the Streamlit app title
st.title("Excel Files Combiner")

# Upload multiple Excel files
uploaded_files = st.file_uploader("Upload Excel Files", type=["xlsx", "xls"], accept_multiple_files=True)

# Check if files are uploaded
if uploaded_files:
    # Create a new Excel writer object
    output_file = "combined_excel_file.xlsx"
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

    # Loop through each uploaded Excel file
    for file in uploaded_files:
        excel_data = pd.ExcelFile(file)
        
        for sheet_name in excel_data.sheet_names:
            combined_sheet_name = f"{os.path.splitext(file.name)[0]}_{sheet_name}"[:31]
            
            df = excel_data.parse(sheet_name)
            df.to_excel(writer, sheet_name=combined_sheet_name, index=False)

    # Close the Excel writer
    writer.close()

    # Provide the download link for the combined Excel file
    with open(output_file, "rb") as f:
        bytes_data = BytesIO(f.read())
        st.download_button("Download Combined Excel File", bytes_data, file_name=output_file)

    # Print success message
    st.success("Excel files combined successfully.")
