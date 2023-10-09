# streamlit_app.py
import streamlit as st
import pandas as pd
import re
import base64

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

# Function to process and save data
def process_data(uploaded_file):
    products = pd.read_excel(uploaded_file, sheet_name='products')
    shipping = pd.read_excel(uploaded_file, sheet_name='shipping')
    shipping['Category'] = shipping['Category'].str.lower().str.replace(' ','')
    shipping = shipping.drop_duplicates()

    categoryList = shipping['Category'].tolist()
    categoryList = [x.replace(' ','').lower() for x in categoryList]
    categoryList = sorted(list(set(categoryList)))

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

    # Create an ExcelWriter object with specific options
    excel_writer = pd.ExcelWriter('Final_Excel.xlsx', engine='xlsxwriter')
    excel_writer.book.strings_to_urls = False  # Disable automatic conversion of strings to URLs

    # Write the DataFrame to Excel
    products.to_excel(excel_writer, index=False, sheet_name='Sheet1')

    # Save the Excel file
    excel_writer.save()

    return 'Final_Excel.xlsx'

# Streamlit UI
def main():
    st.title("Data Processing App")
    st.write("Drag and drop an Excel file to process and download the resulting Excel file.")

    # File uploader for Excel file
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])

    if uploaded_file is not None:
        st.write(f"Uploaded file: {uploaded_file.name}")
        
        if st.button("Process Data"):
            processed_file_name = process_data(uploaded_file)
            st.success(f"Data processed and saved to 'Final_Excel.xlsx'")
            st.write(f"Download Processed Data: [Final_Excel.xlsx](data:application/octet-stream;base64,{base64.b64encode(open(processed_file_name, 'rb').read()).decode()})")

if __name__ == "__main__":
    main()
