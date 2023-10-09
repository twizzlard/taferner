# streamlit_app.py
import streamlit as st
import pandas as pd
import re
import base64
import subprocess

# Function to install 'openpyxl'
def install_openpyxl():
    result = subprocess.run(['pip', 'install', 'openpyxl'], capture_output=True, text=True)
    return result

# ... (rest of your code)

# Streamlit UI
def main():
    st.title("Data Processing App")
    st.write("Upload an Excel file to process and download the resulting Excel file.")

    # File uploader for Excel file
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])

    # Install 'openpyxl' if not already installed
    try:
        import openpyxl
    except ImportError:
        st.warning("The 'openpyxl' library is not installed. Click below to install it.")
        if st.button("Install openpyxl"):
            installation_result = install_openpyxl()
            if installation_result.returncode == 0:
                st.success("Successfully installed 'openpyxl'! You may now upload and process Excel files.")
            else:
                st.error("Failed to install 'openpyxl'. Check the error message below:")
                st.code(installation_result.stderr)
        return

    if uploaded_file is not None:
        st.write(f"Uploaded file: {uploaded_file.name}")

        processed_file_name = process_data(uploaded_file)
        st.success(f"Data processed and saved to 'Final_Excel.xlsx'")
        st.write(f"Download Processed Data: [Download File](data:application/octet-stream;base64,{base64.b64encode(open(processed_file_name, 'rb').read()).decode()})")

if __name__ == "__main__":
    main()
