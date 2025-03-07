import streamlit as st
import pandas as pd
import os
import re

def process_excel(df):
    # Create a copy of the dataframe to avoid modifying the original
    processed_df = df.copy()
    
    # Remove spaces and special characters from NAME OF INSURED column
    def clean_name(name):
        if pd.notna(name):
            # Remove special characters and spaces using regex
            # Keep only alphanumeric characters
            return re.sub(r'[^a-zA-Z0-9]', '', str(name))
        return name
    
    processed_df['NAME OF INSURED'] = processed_df['NAME OF INSURED'].apply(clean_name)
    
    # Process MOBILE NO. column - clean and format phone numbers
    def process_mobile(number):
        if pd.notna(number):
            # Convert to string and remove any spaces, apostrophes, and other special characters
            number = str(number).strip().replace("'", "").replace(" ", "")
            # Remove any decimal points and convert to string
            number = str(int(float(number))) if '.' in number else number
            # Remove 91 prefix if number is 12 digits
            if len(number) == 12 and number.startswith('91'):
                return number[2:]
        return number
    
    processed_df['MOBILE NO.'] = processed_df['MOBILE NO.'].apply(process_mobile)
    
    # Create email IDs using processed NAME OF INSURED
    processed_df['EMAIL ID'] = processed_df['NAME OF INSURED'] + '@yahoo.com'
    
    return processed_df

def main():
    st.title("Excel File Processor")
    
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])
    
    if uploaded_file is not None:
        # Read the Excel file
        try:
            df = pd.read_excel(uploaded_file, sheet_name='TINY RENEWAL')
            
            st.subheader("First 5 rows of original data:")
            st.dataframe(df.head())
            
            if st.button("Process Excel"):
                # Process the dataframe
                processed_df = process_excel(df)
                
                st.subheader("First 5 rows of processed data:")
                st.dataframe(processed_df.head())
                
                # Generate the output filename
                original_filename = uploaded_file.name
                filename_without_ext = os.path.splitext(original_filename)[0]
                output_filename = f"{filename_without_ext}_updated.xlsx"
                
                # Save to Excel
                processed_df.to_excel(output_filename, index=False)
                
                # Create download button
                with open(output_filename, 'rb') as f:
                    file_data = f.read()  # Read file data before closing
                
                st.download_button(
                    label="Download Processed Excel",
                    data=file_data,
                    file_name=output_filename,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    on_click=lambda: cleanup_files(output_filename)
                )
                
        except Exception as e:
            st.error(f"Error: {str(e)}")
            st.error("Please make sure the file contains a sheet named 'TINY RENEWAL' and the required columns.")

def cleanup_files(output_filename):
    """Clean up temporary files"""
    try:
        if os.path.exists(output_filename):
            os.remove(output_filename)
    except Exception as e:
        st.error(f"Error cleaning up files: {str(e)}")

if __name__ == "__main__":
    main() 
