import tabula
import pandas as pd
import pdfplumber
import traceback
import io
import PyPDF2
#import camelot
import os
import base64
from PyPDF2 import PdfFileMerger
from PyPDF2 import PdfFileReader, PdfMerger
import streamlit as st
from openpyxl import load_workbook
import xlsxwriter
from io import BytesIO
from tempfile import TemporaryDirectory
from PIL import Image
import os
import fitz
import itertools
import numpy as np
import string  

st.set_page_config(layout="wide")


# def sidebar_bg(side_bg):
#     #add gif to the pages
#    side_bg_ext = 'png'
#    st.markdown(
#       f"""
#       <style>
#       [data-testid="stSidebar"] > div:first-child {{
#           background: url(data:image/{side_bg_ext};base64,{base64.b64encode(open(side_bg, "rb").read()).decode()});
#       }}
#       </style>
#       """,
#       unsafe_allow_html=True,
#       )
# side_bg = r"C:\Users\ntombesi\OneDrive - MATTEL INC\Documents\\POS Training\SA\Toyzone\barbie_hot_wheels.png"
# # "C:\Users\ntombesi\OneDrive - MATTEL INC\Documents\POS Training\SA\Toyzone\barbie.png"
# sidebar_bg(side_bg)
page = st.sidebar.radio('Navigation',['Delete PDF Page','Merge PDFs','Merge Excel Sheets','SA_Sales_Test'])
# st.sidebar.radio()

def page_1():
    st.title('Delete PDF Pages')
    st.subheader(':red[This page is for deleting specific PDF pages]')
    
    text_paragraphs = """
        \tClick Browse  and select the the file you want to change or drag and drop the file,\n
        \tOnce selected, you can preview the pages of the PDF,\n
        \tOnce you have previewed, select all the pages you want to delete in the dropdown and click Delete Selected Pages.\n
        \tOnce Done, an option to download the new PDF will appear.\n
        \tIf there are any issues, please contact sinethemba.ntombela@mattel.com
         """
    with st.expander("Hello 🙂, Click here to see instructions:"):
         st.write(text_paragraphs)
    selected_files = st.file_uploader("Browse or drag and drop PDF files", type=["pdf"])
    if selected_files:
    # Load the PDF using PyMuPDF
        doc = fitz.open(stream=selected_files.read(), filetype="pdf")
        num_pages = len(doc)
        
        st.write(f"Your PDF has **{num_pages} pages**.")
        
        # Step 1: Create a slider to preview pages
        
        page_number = st.slider("Select a page to preview", 1, num_pages, 1)
        
        # Extract the selected page as an image
        page = doc[page_number - 1]
        pix = page.get_pixmap()
        image = pix.tobytes("png")
        
        # Display the image
        st.image(image, caption=f"Page {page_number} of {num_pages}")
        
        # Step 2: Multi-select widget for deleting pages
        st.write("### Select Pages to Delete")
        pages_to_delete = st.multiselect(
            "Select the pages you want to delete:", 
            options=list(range(1, num_pages + 1)),  # Pages 1-based for user convenience
            format_func=lambda x: f"Page {x}",
        )
        
    # Display a delete button next to the multi-select
        if st.button("Click Here To Delete Selected Pages"):
            if pages_to_delete:
                progress = st.progress(0)  # Initialize the progress bar
                with st.spinner("Deleting selected pages..."):
                    # Align selected pages with 0-based indexing for PyMuPDF
                    selected_indices = [p - 1 for p in pages_to_delete]
                    
                    # Sort and delete pages in reverse order to avoid index mismatch
                    for idx, page_number_to_delete in enumerate(sorted(selected_indices, reverse=True)):
                        doc.delete_page(page_number_to_delete)
                        progress.progress((idx + 1) / len(selected_indices))  # Update progress
                
                st.success("Selected pages have been deleted!")
                
                # Step 3: Provide a button to download the new PDF
                st.write("### Download the Modified PDF")
                # Save the modified PDF to an in-memory BytesIO object
                new_pdf = BytesIO()
                doc.save(new_pdf)
                new_pdf.seek(0)
                
                st.download_button(
                    label="Download New PDF",
                    data=new_pdf,
                    file_name="modified_document.pdf",
                    mime="application/pdf",
                )
            else:
                st.warning("No pages selected for deletion. Please select pages to delete.")

def page_2():
    def merge_pdfs(selected_files, output_file):
        try:
            st.write("Merging PDFs...")
            # Create a PdfMerger object
            pdf_merger = PdfMerger()

            # Iterate through all selected files
            for file in selected_files:
                # Open each selected PDF file and append it to the PdfMerger object
                st.write(f"Merging {file.name}")
                pdf_merger.append(file)

            # Write the merged PDF to the output file
            with open(output_file, 'wb') as merged_file:
                pdf_merger.write(merged_file)

            st.success("PDF files merged successfully!")
        except Exception as e:
            # Display custom error message
            st.error("An error occurred during merging. Please check your files and try again.")

    def main():

        st.title("PDF Merger")
        st.subheader(':red[This page merges different PDF files into one document]')
        # File uploader for selecting PDF files folder
        text_paragraphs = """
        \tClick Browse files and select or drag and drop :red[ALL] the files you want to merge,\n
        \tOnce selected, you will get an Output file name box, please type the name you want for your output file and select Merge PDFs,\n
        \tOnce selected, click the Download Merged PDF at the bottom of the screen and save to your selected destination.\n
        \tIf there are any issues, please contact sinethemba.ntombela@mattel.com
         """
        with st.expander("Hello 🙂, Click here to see instructions:"):
        # Content inside the expander
            st.write(text_paragraphs)
        selected_files = st.file_uploader("Browse or drag and drop PDF files", type=["pdf"], accept_multiple_files=True)

        if selected_files:
            output_file_name = st.text_input("Enter the name for the merged file:")
            if not output_file_name.endswith('.pdf'):
                output_file_name += '.pdf'

            output_file_path = os.path.join("/tmp", output_file_name)

            # Merge button
            if st.button("Merge PDFs"):
                merge_pdfs(selected_files, output_file_path)

            # Download button
            if os.path.exists(output_file_path):
                with open(output_file_path, "rb") as file:
                    data = file.read()
                st.download_button(label="Download Merged PDF", data=data, file_name=output_file_name)

    if __name__ == "__main__":
        main()

def page_3():
    st.title('Merge multiple excel sheets')
    st.subheader(':red[Note:] if the excel sheets have different columns, the final data may be skewed.')
    st.subheader('Best if used for excel sheets with the same columns.')
    uploaded_file = st.file_uploader("Choose a file", type = 'xlsx')
    if uploaded_file is not None:
        sheets_dict = pd.read_excel(uploaded_file, sheet_name=None)
        
        # List to hold DataFrames with added sheet name column
        dfs_with_sheet_names = []
        
        # Loop through the dictionary, adding the sheet name as a column
        for sheet_name, df in sheets_dict.items():
            df['Sheet Name'] = sheet_name
            dfs_with_sheet_names.append(df)
        
        # Concatenate all DataFrames into one
        df_added = pd.concat(dfs_with_sheet_names, ignore_index=True)
        
        # Display the combined DataFrame
        st.subheader('See the final product below:')
        st.dataframe(df_added)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_added.to_excel(writer, index=False, sheet_name='Merged Data')
        
        # Ensure the data is written to the BytesIO object
        output.seek(0)
        
        # Provide a download button
        st.download_button(
            label="Download merged Excel file",
            data=output,
            file_name="merged_sheets.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.write('Please upload excel file')

# £pages = ['Merge PDFs','Merge Sheets']
# selected_page = st.sidebar.selectbox('Select Page', pages)

def page_4():
    st.header('Test Page')
    #col1,col2 = st.columns(2)
    # with col1:
    uploaded_core_file = st.file_uploader("Order File", type = 'xlsx')
    if uploaded_core_file is not None:
            file_name = uploaded_core_file.name
            #st.write(f"Uploaded file: {file_name}")
            df1 = pd.read_excel(uploaded_core_file)
            pd.DataFrame(df1)
            #st.dataframe(df)
            df_core = df1.copy()
    else:
        st.write('')
    #df_core_column = df_core[['TOYNBR','PRTNBR','ALQTOD']]
    # filtered_df_core = df_core_column[df_core_column['TOYNBR'].str.contains('TOYNBR')]
    text_to_remove = 'TOYNBR'
    filtered_df_core = df_core[~df_core['TOYNBR'].str.contains(text_to_remove, na=False)]
    filtered_df_core['unique_code'] = filtered_df_core['TOYNBR'] + filtered_df_core['PRTNBR']
    #df_core_column = filtered_df_core[['unique_code','ALQTOD']]
    filtered_df_core.head(5)
    df_core_cp = filtered_df_core[filtered_df_core['CUSTNM'] == 'TAKEALOT ONLINE CAPE TOWN']  # Filter for CP
    df_core_jhb = filtered_df_core[filtered_df_core['CUSTNM'] == 'TAKEALOT ONLINE JOHANNESBURG']  # Filter for JHB
        #df_core_column
    # with col2:
    uploaded_mapping_file = st.file_uploader("Barcode File", type = 'xlsm')
    if uploaded_mapping_file is not None:
            file_name = uploaded_mapping_file.name
            #st.write(f"Uploaded file: {file_name}")
            df2 = pd.read_excel(uploaded_mapping_file,sheet_name='Data',skiprows=1)
            pd.DataFrame(df2)
            #st.dataframe(df)
            df_map = df2.copy()
    #
    else:
        st.write('')
    # df_map_column = df_map[['TOYNUMBER','PARTPR','CQTYEE']]

    
    df_map['unique_code'] = df_map['TOYNUMBER'] + df_map['PARTPR']
    df_map_column = df_map[['unique_code','CQTYEE']]

    df_core_column_JHB = df_core_jhb[['unique_code','ORDNBR','CUSTNM','ALQTOD']]
    df_core_column_CP = df_core_cp[['unique_code','ORDNBR','CUSTNM','ALQTOD']]
    grouped_df_map = df_map_column.groupby('unique_code').sum()
    df_merge_CP = pd.merge(df_core_column_CP, grouped_df_map, on='unique_code', how='left')
    df_merge_JHB = pd.merge(df_core_column_JHB, grouped_df_map, on='unique_code', how='left')
    df_merge_CP['Flag'] = (df_merge_CP['ALQTOD'] == df_merge_CP['CQTYEE']).astype(int)
    df_merge_JHB['Flag'] = (df_merge_JHB['ALQTOD'] == df_merge_JHB['CQTYEE']).astype(int)
    flagged_merged_rows_CP = df_merge_CP[df_merge_CP['Flag'] == 1]
    flagged_merged_rows_JHB = df_merge_JHB[df_merge_JHB['Flag'] == 1]
    unique_id_list_JHB = list(flagged_merged_rows_JHB.unique_code.unique())
    unique_id_list_CP= list(flagged_merged_rows_CP.unique_code.unique())
    df_map_filtered = df_map[['TOYNUMBER','PARTPR','ITDSAA','STPKQT','CTOYEE','ITDSEE','CQTYEE','UPC1EE','unique_code']]
    # Perform an inner join to filter df_map_filtered based on flagged_merged_rows_CP
    df_map_copy_1_CP = pd.merge(
        df_map_filtered,  # The dataframe to be filtered
        flagged_merged_rows_CP[['unique_code', 'ORDNBR', 'CUSTNM']],  # Columns to bring in
        on='unique_code',  # Match on unique_code
        how='inner'  # Inner join ensures only matching rows are included
    )
    # Perform an inner join to filter df_map_filtered based on flagged_merged_rows_CP
    df_map_copy_1_JHB = pd.merge(
        df_map_filtered,  # The dataframe to be filtered
        flagged_merged_rows_JHB[['unique_code', 'ORDNBR', 'CUSTNM']],  # Columns to bring in
        on='unique_code',  # Match on unique_code
        how='inner'  # Inner join ensures only matching rows are included
    )
    flagged_zero_CP = df_merge_CP[df_merge_CP['Flag'] == 0]
    flagged_zero_JHB = df_merge_JHB[df_merge_JHB['Flag'] == 0]

    df_map_copy_2_CP = pd.merge(
        df_map_filtered,  # Base dataframe to filter
        flagged_zero_CP[['unique_code', 'ORDNBR', 'CUSTNM']],  # Include ORDNBR and CUSTNM during filter
        on='unique_code',
        how='left',  # Left join to allow for the anti join logic
        indicator=True  # Add _merge column to track matching rows
        )

    # Step 3: Keep rows that are NOT in flagged_zero_CP (where _merge == 'left_only')
    df_map_copy_2_CP = df_map_copy_2_CP[df_map_copy_2_CP['_merge'] == 'left_only'].drop('_merge', axis=1)
    #df_map_copy_2_CP
    df_map_copy_2_JHB = pd.merge(
        df_map_filtered,  # Base dataframe to filter
        flagged_zero_JHB[['unique_code', 'ORDNBR', 'CUSTNM']],  # Include ORDNBR and CUSTNM during filter
        on='unique_code',
        how='left',  # Left join to allow for the anti join logic
        indicator=True  # Add _merge column to track matching rows
        )

    # Step 3: Keep rows that are NOT in flagged_zero_CP (where _merge == 'left_only')
    df_map_copy_2_JHB = df_map_copy_2_JHB[df_map_copy_2_JHB['_merge'] == 'left_only'].drop('_merge', axis=1)


    # df_map_copy_1_CP = df_map[df_map['unique_code'].isin(unique_id_list_CP)]
    # df_map_copy_1_JHB = df_map[df_map['unique_code'].isin(unique_id_list_JHB)]
    # df_map_copy_2_CP = df_map[~df_map['unique_code'].isin(unique_id_list_CP)]
    # df_map_copy_2_JHB = df_map[~df_map['unique_code'].isin(unique_id_list_JHB)]
    df_map_copy_1_JHB['Multiply By'] = 1
    df_map_copy_1_JHB['Total'] = df_map_copy_1_JHB['CQTYEE'] * df_map_copy_1_JHB['Multiply By']
    # df_map_copy_1.head(5)
    df_map_copy_1_CP['Multiply By'] = 1
    df_map_copy_1_CP['Total'] = df_map_copy_1_CP['CQTYEE'] * df_map_copy_1_CP['Multiply By']
    cqtyee_index = df_map_copy_1_CP.columns.get_loc('CQTYEE')
    df_map_copy_1_CP.insert(cqtyee_index + 1, 'Multiply By', df_map_copy_1_CP.pop('Multiply By'))
    df_map_copy_1_CP.insert(cqtyee_index + 2, 'Total', df_map_copy_1_CP.pop('Total'))
    cqtyee_index = df_map_copy_1_JHB.columns.get_loc('CQTYEE')
    df_map_copy_1_JHB.insert(cqtyee_index + 1, 'Multiply By', df_map_copy_1_JHB.pop('Multiply By'))
    df_map_copy_1_JHB.insert(cqtyee_index + 2, 'Total', df_map_copy_1_JHB.pop('Total'))
    df_core_copy_JHB = df_core_column_JHB[~df_core_column_JHB['unique_code'].isin(unique_id_list_JHB)]
    df_core_copy_CP = df_core_column_CP[~df_core_column_CP['unique_code'].isin(unique_id_list_CP)]

    df_map_copy_JHB  = df_map_filtered.copy()
    df_map_copy_JHB['Multiply By'] = np.nan
    for _, row in df_core_copy_JHB.iterrows():
        unique_id = row['unique_code']
        total_value = row['ALQTOD']
        
        # Filter df2 for matching UniqueID
        filtered_df_map = df_map_copy_JHB[df_map_copy_JHB['unique_code'] == unique_id]
        
        # Check if all CQTYEE values are identical
        if len(filtered_df_map['CQTYEE'].unique()) == 1:
            cqtyee_value = filtered_df_map['CQTYEE'].iloc[0]  # Get the identical CQTYEE value
            sum_cqtyee = filtered_df_map['CQTYEE'].sum()     # Sum of CQTYEE values
            
            # Calculate Total divided by the sum
            if total_value % sum_cqtyee == 0:  # Check if division is an integer
                multiply_by_value = total_value // sum_cqtyee  # Integer division
                
                # Update 'Multiply By' column for matching rows in df2
                df_map_copy_JHB.loc[df_map_copy_JHB['unique_code'] == unique_id, 'Multiply By'] = multiply_by_value
    df_map_copy_CP  = df_map_filtered.copy()
    df_map_copy_CP['Multiply By'] = np.nan
    for _, row in df_core_copy_CP.iterrows():
        unique_id = row['unique_code']
        total_value = row['ALQTOD']
        
        # Filter df2 for matching UniqueID
        filtered_df_map = df_map_copy_CP[df_map_copy_CP['unique_code'] == unique_id]
        
        # Check if all CQTYEE values are identical
        if len(filtered_df_map['CQTYEE'].unique()) == 1:
            cqtyee_value = filtered_df_map['CQTYEE'].iloc[0]  # Get the identical CQTYEE value
            sum_cqtyee = filtered_df_map['CQTYEE'].sum()     # Sum of CQTYEE values
            
            # Calculate Total divided by the sum
            if total_value % sum_cqtyee == 0:  # Check if division is an integer
                multiply_by_value = total_value // sum_cqtyee  # Integer division
                
                # Update 'Multiply By' column for matching rows in df2
                df_map_copy_CP.loc[df_map_copy_CP['unique_code'] == unique_id, 'Multiply By'] = multiply_by_value

    df_cleaned_map_CP = df_map_copy_CP.dropna(subset=['Multiply By'])
    df_cleaned_map_JHB = df_map_copy_JHB.dropna(subset=['Multiply By'])


    df_cleaned_map_CP['Total'] = df_cleaned_map_CP['CQTYEE'] * df_cleaned_map_CP['Multiply By']
    cqtyee_index = df_cleaned_map_CP.columns.get_loc('CQTYEE')
    # Use insert() to relocate the 'Multiply By' and 'Total' columns next to 'CQTYEE'
    # Insert 'Multiply By' immediately after 'CQTYEE'
    df_cleaned_map_CP.insert(cqtyee_index + 1, 'Multiply By', df_cleaned_map_CP.pop('Multiply By'))
    # Insert 'Total' immediately after 'Multiply By'
    df_cleaned_map_CP.insert(cqtyee_index + 2, 'Total', df_cleaned_map_CP.pop('Total'))
    #df_map_column

    df_cleaned_map_JHB['Total'] = df_cleaned_map_JHB['CQTYEE'] * df_cleaned_map_JHB['Multiply By']
    cqtyee_index = df_cleaned_map_JHB.columns.get_loc('CQTYEE')
    # Use insert() to relocate the 'Multiply By' and 'Total' columns next to 'CQTYEE'
    # Insert 'Multiply By' immediately after 'CQTYEE'
    df_cleaned_map_JHB.insert(cqtyee_index + 1, 'Multiply By', df_cleaned_map_JHB.pop('Multiply By'))
    # Insert 'Total' immediately after 'Multiply By'
    df_cleaned_map_JHB.insert(cqtyee_index + 2, 'Total', df_cleaned_map_JHB.pop('Total'))
        #df_source_file = pd.read_excel("/Fauve_issue/")
    # £pages = ['Merge PDFs','Merge Sheets']
    # selected_page = st.sidebar.selectbox('Select Page', pages)
    df_cleaned_map_JHB = pd.merge(
        df_cleaned_map_JHB,
        df_core_copy_JHB[['unique_code', 'ORDNBR', 'CUSTNM']],
        on='unique_code',
        how='left'
    )
    df_cleaned_map_CP = pd.merge(
        df_cleaned_map_CP,
        df_core_copy_CP[['unique_code', 'ORDNBR', 'CUSTNM']],
        on='unique_code',
        how='left'
    )

    final_df = pd.concat([df_map_copy_1_CP, df_cleaned_map_CP,df_map_copy_1_JHB, df_cleaned_map_JHB], ignore_index=True)

    def find_optimal_solution(target, coefficients):  
# Initialize variables  
        optimal_solution = None  
        min_difference = float('inf')  
        coefficient = len(coefficients)  # Number of variables (a, b, c, ...)  
    
        # Determine range for each variable based on coefficients  
        ranges = [range(target // coefficient + 1) for coefficient in coefficients]  
        MAX_COMBINATIONS = 1_000_000

        # Before loop
        total_combinations = 1
        for r in ranges:
            total_combinations *= len(r)
        if total_combinations > MAX_COMBINATIONS:
            print("Too many combinations, skipping")
            return None
    
        # Brute-force search for all possible values of variables  
        for values in itertools.product(*ranges):  
            # Calculate total using coefficients  
            total = sum(coefficient * value for coefficient, value in zip(coefficients, values))  
    
            if total == target:  
                # Calculate individual contributions  
                contributions = [coefficient * value for coefficient, value in zip(coefficients, values)]  
    
                # Calculate max and min contributions  
                max_val = max(contributions)  
                min_val = min(contributions)  
                difference = max_val - min_val  
    
                # Check the additional constraint: difference < 5  
                if difference < 5:  
                    # Update the optimal solution if this one is better  
                    if difference < min_difference:  
                        min_difference = difference  
                        optimal_solution = values  
    
        return optimal_solution
    unique_solved_list = list(final_df.unique_code.unique())
    filtered_df_core = filtered_df_core[~filtered_df_core['unique_code'].isin(unique_solved_list)]
    df_map_filtered['CQTYEE'] = df_map_filtered['CQTYEE'].fillna(0).astype(int)
    # Prepare to collect results
    # Prepare to collect results
    results = []

    # Get total number of rows for iteration info
    total_iterations = len(filtered_df_core)

    # Iterate over each unique_code in filtered_df_core
    for i, (_, row) in enumerate(filtered_df_core.iterrows(), start=1):
        code_to_solve = row['unique_code']
        ALQTOD = row['ALQTOD']
        ORDNBR = row['ORDNBR']
        CUSTNM = row['CUSTNM']

        # Print iteration progress
        print(f"Processing {i}/{total_iterations}: unique_code={code_to_solve}")

        # Filter df_map to get the rows for the specific unique_code
        filtered_df = df_map_filtered[df_map_filtered['unique_code'] == code_to_solve].copy()
        filtered_df['ORDNBR'] = ORDNBR
        filtered_df['CUSTNM'] = CUSTNM

        # Extract the CQTYEE values
        pairof_items = list(filtered_df['CQTYEE'].values)

        # Find optimal solution
        optimal_solution = find_optimal_solution(ALQTOD, pairof_items)

        # Assign labels and store the results
        if optimal_solution:
            labels = list(string.ascii_lowercase)[:len(optimal_solution)]
            filtered_df['Label'] = labels
            filtered_df['OptimalValue'] = optimal_solution
            filtered_df['Outcome'] = ', '.join(f'{label}={value}' for label, value in zip(labels, optimal_solution))
            results.append(filtered_df)
        else:
            filtered_df['Label'] = list(string.ascii_lowercase)[:len(filtered_df)]
            filtered_df['OptimalValue'] = None
            filtered_df['Outcome'] = 'No solution found'
            results.append(filtered_df)

        # Optional: Print when a solution is found
        if optimal_solution:
            print(f"Solution found for unique_code={code_to_solve}")
        else:
            print(f"No solution for unique_code={code_to_solve}")

    # Concatenate all results into a single DataFrame
    results_df = pd.concat(results, ignore_index=True)

    # Concatenate all results into a single DataFrame
    results_df = pd.concat(results, ignore_index=True)

    results_df_worked = results_df[results_df['Outcome']  !=  'No solution found']
    results_df_not_solved = results_df[results_df['Outcome'] == 'No solution found']
    unique_not_solved = results_df_not_solved.unique_code.unique()
    pd.DataFrame(unique_not_solved)

    results_df_worked_renamed = results_df_worked.rename(columns = {'OptimalValue':'Multiply By'})
    results_df_worked_dropped = results_df_worked_renamed.drop(columns = ['Label','Outcome'])
    results_df_worked_dropped['Total'] = results_df_worked_dropped['CQTYEE'] * results_df_worked_dropped['Multiply By']

    cqtyee_index = results_df_worked_dropped.columns.get_loc('CQTYEE')
    results_df_worked_dropped.insert(cqtyee_index + 1, 'Multiply By', results_df_worked_dropped.pop('Multiply By'))
    results_df_worked_dropped.insert(cqtyee_index + 2, 'Total', results_df_worked_dropped.pop('Total'))

    final_df_with_3 = pd.concat([final_df, results_df_worked_dropped], ignore_index=True)
    st.header("Export populated file to Excel")

    col3,col4 = st.columns(2)

    def convert_df_to_excel(dataframe):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            dataframe.to_excel(writer, index=False, sheet_name='Sheet1')
        processed_data = output.getvalue()
        return processed_data
    

    # Streamlit App Interface
    with col3:
        
        # Button for DataFrame 1
        if st.button("Export correct populated file"):
            excel_data = convert_df_to_excel(final_df_with_3)
            st.download_button(
                label="Download populated file as Excel",
                data=excel_data,
                file_name="final_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with col4:
    # Button for DataFrame 2
        unique_not_solved
        if st.button("Export file where there was no solution"):
            excel_data = convert_df_to_excel(unique_not_solved)
            st.download_button(
                label="Download no solution found file as Excel",
                data=excel_data,
                file_name="no-solution.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

if page == 'Delete PDF Page':
    page_1()
if page == 'Merge PDFs':
    page_2()
if page == 'Merge Excel Sheets':
    page_3()
if page == 'SA_Sales_Test':
    page_4()
