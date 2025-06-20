import streamlit as st
import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook

# Configuration
EXCEL_FILE = "MCC_QUAL_FORECAST_shared.xlsm"
TARGET_SHEET = "Reference Table 1"  # Name of your specific sheet
COLUMN_NAME = "Product"
COLUMNS = [
    COLUMN_NAME,
    "Total Positions",
    "Lee Hua's remark",
    "Note: This table is for product reference (based on HTOL Forecast)"
]
# Initialize Excel file with 12 sheets if not exists
if not os.path.isfile(EXCEL_FILE):
    # Create empty DataFrames for all sheets
    st.write("File not found")
else: 
    def add_to_sheet(product_name):
        """Add product to specific sheet in Excel file"""
        try:
            # Load workbook
            book = load_workbook(EXCEL_FILE)
            
            existing_df = pd.read_excel(EXCEL_FILE, sheet_name=TARGET_SHEET)

            new_row = {
            COLUMN_NAME: product_name,
            "Total Positions": 0,
            "Lee Hua's remark": "",
            "Note: This table is for product reference (based on HTOL Forecast)": ""
        }
        
            updated_df = pd.concat([existing_df, pd.DataFrame([new_row])], ignore_index=True)
            

            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                writer.book = book
                writer.sheets = {ws.title: ws for ws in book.worksheets}
                updated_df.to_excel(writer, sheet_name=TARGET_SHEET, index=False)
                
            return True
                
        
        except Exception as e:
            st.error(f"Error: {str(e)}")
            return False


    st.title("Product Inventory Manager")
    st.subheader(f"Add products to '{TARGET_SHEET}' sheet")

    with st.form("product_form"):
        product_name = st.text_input("Enter product name:", max_chars=100, key="product_input")
        submitted = st.form_submit_button("Add Product")
        
        if submitted:
            if product_name.strip():
                if add_to_sheet(product_name.strip()):
                    st.success(f"✅ '{product_name}' added to {TARGET_SHEET}!")
                    st.session_state.product_input = ""
            else:
                st.warning("⚠️ Please enter a valid product name")

