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
page = st.sidebar.radio('Navigation',['Delete PDF Page','Merge PDFs','Merge Excel Sheets',])
# st.sidebar.radio()

def page_1():
    st.title('Delete PDF Pages')
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
        # File uploader for selecting PDF files folder
        text_paragraphs = """
        \tClick Browse files and select the the files you want to merge,\n
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
    st.header('Merge multiple excel sheets')
    st.subheader('Note; if the excel sheets have different columns, the final data may be skewed.')
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
        st.write('please upload excel file')

# £pages = ['Merge PDFs','Merge Sheets']
# selected_page = st.sidebar.selectbox('Select Page', pages)

if page == 'Delete PDF Page':
    page_1()
if page == 'Merge PDFs':
    page_2()
if page == 'Merge Excel Sheets':
    page_3()
