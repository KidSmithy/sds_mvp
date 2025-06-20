import streamlit as st
import pandas as pd
import os
import xlwings as xw

# Configuration
EXCEL_FILE = "MCC_QUAL_FORECAST_shared.xlsm"
TARGET_SHEET = "Reference Table1"
COLUMN_NAME = "Product"
COLUMNS = [
    COLUMN_NAME,
    "Total Positions",
    "Lee Hua's remark",
    "Note: This table is for product reference (based on HTOL Forecast)"
]

def add_to_sheet(product_name):
    """Add product to specific sheet using xlwings to preserve macros"""
    try:
        # Check if Excel is installed
            
        # Open Excel in the background
        app = xw.App(visible=False)
        wb = None
        
        try:
            # Open the workbook
            wb = app.books.open(EXCEL_FILE)
            
            # Check if sheet exists
            if TARGET_SHEET not in [sheet.name for sheet in wb.sheets]:
                # Create new sheet
                new_sheet = wb.sheets.add(TARGET_SHEET)
                # Add headers
                new_sheet.range('A1').value = COLUMNS
            else:
                sheet = wb.sheets[TARGET_SHEET]
                
            # Find next empty row
            sheet = wb.sheets[TARGET_SHEET]
            last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            next_row = last_row + 1 if last_row > 1 else 2
            
            # Add new data
            sheet.range(f'A{next_row}').value = [
                product_name,  # Product
                0,             # Total Positions
                "",            # Lee Hua's remark
                ""             # Note
            ]
            
            # Save and close
            wb.save()
            return True
        except Exception as e:
            st.error(f"Excel error: {str(e)}")
            return False
        finally:
            if wb:
                wb.close()
            app.quit()
    except Exception as e:
        st.error(f"Application error: {str(e)}")
        return False

# Streamlit UI
if os.path.isfile(EXCEL_FILE):
    st.title("Product Inventory Manager")
    st.subheader(f"Add products to '{TARGET_SHEET}' sheet")

    with st.form("product_form"):
        product_name = st.text_input("Enter product name:", max_chars=100, key="product_input")
        submitted = st.form_submit_button("Add Product")
        
        if submitted:
            if product_name.strip():
                if add_to_sheet(product_name.strip()):
                    st.success(f"✅ '{product_name}' added to {TARGET_SHEET}!")
                    # st.session_state.product_input = ""
            else:
                st.warning("⚠️ Please enter a valid product name")
else:
    st.error(f"Error: File {EXCEL_FILE} not found! Please check the file path.")