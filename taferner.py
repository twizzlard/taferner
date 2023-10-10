# streamlit_app.py
import streamlit as st
import pandas as pd
import re
import base64
import subprocess
import xlsxwriter
import openpyxl

# Create a function to find the matching category
def find_matching_category(category, category_list):
    category_lower = category.lower().replace(' ','')
    for cat in category_list:
        if category_lower == cat.lower():
            return cat
        elif category_lower.startswith(cat.lower() + ">"):
            return cat
        else:
            # Check for partial string matching
            parts = cat.lower().split(" > ")
            for i in range(len(parts) - 1, 0, -1):
                partial_cat = " > ".join(parts[:i])
                if category_lower.startswith(partial_cat):
                    return partial_cat
    return "No match"

# Function to install 'openpyxl'
def install_openpyxl():
    result = subprocess.run(['pip', 'install', 'openpyxl'], capture_output=True, text=True)
    return result

# Function to process and save data
def process_data(uploaded_file):
    products = pd.read_excel(uploaded_file, sheet_name='products',nrows=5)
    shipping = pd.read_excel(uploaded_file, sheet_name='shipping')
    shipping['Category'] = shipping['Category'].str.lower().str.replace(' ','')
    shipping = shipping.drop_duplicates()

    categoryList = shipping['Category'].tolist()
    categoryList = [x.replace(' ','').lower() for x in categoryList]
    categoryList = sorted(list(set(categoryList)))

    products['Images'] = "'" + products['Images']
    
    products['Categories'] = products['Categories'].ffill()

    # Create products['Matching_Category'] for matching purposes and find the match category for each row
    products['Matching_Category'] = products['Categories'].apply(lambda x: find_matching_category(x, categoryList))

    # Merge the products table with the shipping data
    products = products.merge(shipping, how='left', left_on='Matching_Category', right_on='Category')

    # Change 'Published' to '-1' where 'Weight' is 'draft'
    products.loc[products['Weight'] == 'draft', 'Published'] = "'-1"

    # Drop the old dimension products columns and rename the new ones to the old names
    products = products.drop(columns=['Category', 'Weight (kg)', 'Length (cm)', 'Width (cm)', 'Height (cm)'])
    products = products.rename(columns={'Weight':'Weight (kg)', 'Length':'Length (cm)', 'Width':'Width (cm)', 'Height':'Height (cm)'})

    # Save URLs as plaintext so they don't get cut off by Excel limitations
    products['Images'] = products['Images'].astype(str)

    # Save the DataFrame to an Excel file using Pandas with 'XlsxWriter' engine
    processed_file_name = 'Final_Excel.xlsx'
    products.to_excel(processed_file_name, index=False, engine='xlsxwriter')


    return processed_file_name

# Streamlit UI
def main():
    st.title("Data Processing App")
    st.write("Upload an Excel file to process and download the resulting Excel file.")

    # File uploader for Excel file
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])

    if uploaded_file is not None:
        st.write(f"Uploaded file: {uploaded_file.name}")

        # Process the data and get the processed DataFrame
        processed_df = process_data(uploaded_file)

        # Provide a download link for the processed DataFrame in Excel format
        excel_data = processed_df.to_excel(index=False)
        b64 = base64.b64encode(excel_data).decode()
        st.markdown(f"Download Processed Data: [Download File](data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64})", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
