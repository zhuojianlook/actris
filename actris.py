import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# Function to load data from the uploaded file
def load_excel(uploaded_file):
    if uploaded_file is not None:
        return pd.ExcelFile(uploaded_file)
    else:
        return None

# Initialize session state for holding changes
if 'changes' not in st.session_state:
    st.session_state['changes'] = {}

# Streamlit application
def main():
    st.title('GMP Ancillary Materials for ACTRIS')

    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

    if uploaded_file is not None:
        xls = load_excel(uploaded_file)

        if xls is not None:
            # Get sheet names from the Excel file
            sheet_names = xls.sheet_names
            selected_sheet = st.selectbox("Select a sheet", sheet_names)

            # Read the selected sheet
            data = xls.parse(selected_sheet)

            # Dropdown to select an item
            items = data.iloc[:, 0].tolist()  # Assumes the first column contains the item names
            selected_item = st.selectbox("Select an item", items)

            if selected_item and selected_item in items:
                # Find the index of the row in the dataframe
                row_index = data[data.iloc[:, 0] == selected_item].index[0]

                # Display attributes for editing
                for col in data.columns:
                    # Use the selected item and column as a unique key, with a safer delimiter
                    key = f"{selected_sheet}__{selected_item}__{col}"
                    current_value = st.session_state['changes'].get(key, data.at[row_index, col])
                    new_value = st.text_input(col, current_value, key=key)
                    st.session_state['changes'][key] = new_value

            # Checkbox for appending today's date in filename
            append_date = st.checkbox("Append today's date to filename")

            # Button to save the edited data
            if st.button('Save All Changes'):
                # Apply changes to the dataframe
                for key, value in st.session_state['changes'].items():
                    # Ensure the key can be split correctly
                    if '__' in key:
                        sheet, item, col = key.split('__')
                        if sheet == selected_sheet and item in items:
                            row_index = data[data.iloc[:, 0] == item].index[0]
                            data.at[row_index, col] = value

                # Save the edited dataframe back to Excel
                with pd.ExcelWriter(uploaded_file.name) as writer:
                    for sheet in sheet_names:
                        if sheet == selected_sheet:
                            data.to_excel(writer, sheet_name=sheet, index=False)
                        else:
                            xls.parse(sheet).to_excel(writer, sheet_name=sheet, index=False)

                # Construct filename
                filename = 'GMP_ancillary_materials'
                if append_date:
                    today = datetime.now().strftime("%Y%m%d")
                    filename += f'_{today}'
                filename += '.xlsx'

                # Provide download link
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for sheet in sheet_names:
                        if sheet == selected_sheet:
                            data.to_excel(writer, sheet_name=sheet, index=False)
                        else:
                            xls.parse(sheet).to_excel(writer, sheet_name=sheet, index=False)
                    writer.save()
                    st.download_button(
                        label="Download Edited File",
                        data=output.getvalue(),
                        file_name=filename,
                        mime="application/vnd.ms-excel"
                    )

if __name__ == '__main__':
    main()
