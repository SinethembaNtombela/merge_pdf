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
page = st.sidebar.radio('Navigation',['Merge PDFs','Merge Excel Sheets'])
# st.sidebar.radio()


def page_1():
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

def page_2():
    st.header('Merge multiple excel sheets')
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

if page == 'Merge PDFs':
    page_1()
    
if page == 'Merge Excel Sheets':
    page_2()
