import streamlit as st
import pandas as pd
import base64
from io import BytesIO

non_matches=[]

def find_matching_category(category, category_list):

    global non_matches
    
    category_lower = category.lower().replace(' ','')
    
    while category_lower not in category_list and category_lower:
        category_lower = category_lower[:-1]
        
    if category_lower in category_list:
#         print(category)
#         print(f"Found a match: {category_lower}")
        return category_lower
    else:
        for cat in category_list:
            category_lower = category.lower().replace(' ','')
            if category_lower in cat:
                return cat
            else:
                pass
        
        print(category)
        non_matches.append(category)
        return "No Match"

# Function to process and save data
def process_data(uploaded_file):

    global non_matches
    
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
    products.loc[products['Weight'] == 'draft', 'Weight'] = ""
    
    # Filter the rows where Matching_Category is 'No Match' and now, the 'Matching_Category' column in rows with 'No Match' will be updated with the value in 'Category' preceded by '^'.
    products.loc[products['Matching_Category'] == 'No Match', 'Matching_Category'] = "^" + products['Category']

    # Drop the old dimension products columns and rename the new ones to the old names
    products = products.drop(columns=['Category', 'Weight (kg)', 'Length (cm)', 'Width (cm)', 'Height (cm)'])
    products = products.rename(columns={'Weight':'Weight (kg)', 'Length':'Length (cm)', 'Width':'Width (cm)', 'Height':'Height (cm)'})

    # Save URLs as plaintext so they don't get cut off by Excel limitations
    products['Images'] = "'" + products['Images'].astype(str)

    return products

# Streamlit UI
def main():

    global non_matches
    
    st.title("WooCommerce XLSX Processing App")
    st.write("Upload a WooCommerce XLSX to process.")

    # File uploader for Excel file
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])

    if uploaded_file is not None:
        st.write(f"Uploaded file: {uploaded_file.name}. This next part may take a few minutes.")

        # Process the data and get the processed DataFrame
        processed_df = process_data(uploaded_file)

        # Provide a download button for the processed Excel file
        excel_file_name = 'Processed_Data.xlsx'
        processed_df.to_excel(excel_file_name, index=False)
        with open(excel_file_name, "rb") as file:
            contents = file.read()
            st.download_button(
                label="Download Processed Data",
                data=contents,
                key='download_button',
                file_name=excel_file_name,
            )
        
        # # List categories without matches
        # non_matches = sorted(list(set(non_matches)))
        # st.write("Categories without matches:")
        # for category in non_matches:
        #     st.write(category)
        
if __name__ == "__main__":
    main()
