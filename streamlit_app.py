import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

# Function to process and find differences
def compare_data(old_file, new_file):
    try:
        # Load Excel files
        old_data = pd.read_excel(old_file)
        new_data = pd.read_excel(new_file)

        # Filter by 'CLM Contract Type'
        old_filtered = old_data[old_data['CLM Contract Type'].isin(['National', 'Pharmacy'])]
        new_filtered = new_data[new_data['CLM Contract Type'].isin(['National', 'Pharmacy'])]

        # Drop duplicates based on 'Vendor #'
        old_df = old_filtered.drop_duplicates(subset='Vendor #', keep='first')
        new_df = new_filtered.drop_duplicates(subset='Vendor #', keep='first')

        # Identify newly added contracts
        new_added = new_df[~new_df['Contract #'].isin(old_df['Contract #'])]

        # Identify removed contracts
        removed_in_old = old_df[~old_df['Contract #'].isin(new_df['Contract #'])]

        return new_added, removed_in_old

    except Exception as e:
        st.error(f"An error occurred: {e}")
        return None, None

# Streamlit UI
def main():
    st.title("Excel Comparison Tool")
    st.write("Upload previous and current month's Excel files to find added and removed contracts.")

    # File upload
    old_file = st.file_uploader("Upload Previous Month's File (Excel format)", type=["xlsx"])
    new_file = st.file_uploader("Upload Current Month's File (Excel format)", type=["xlsx"])

    if old_file and new_file:
        st.success("Files uploaded successfully!")

        # Compare the files
        new_added, removed_in_old = compare_data(old_file, new_file)

        if new_added is not None and removed_in_old is not None:
            st.write("### Newly Added Contracts")
            st.dataframe(new_added)

            st.write("### Removed Contracts")
            st.dataframe(removed_in_old)

            # Save the newly added as an Excel file
            today = datetime.datetime.today().strftime('%Y-%m-%d')
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                new_added.to_excel(writer, index=False, sheet_name='Newly Added')
            output.seek(0)

            # Download button for Excel file
            st.download_button(
                label="Download Newly Added Contracts",
                data=output,
                file_name=f'new_vendors_{today}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    else:
        st.info("Please upload both files to begin comparison.")

if __name__ == "__main__":
    main()
