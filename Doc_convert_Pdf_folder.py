from docx2pdf import convert
import os


docx_folder_path = r"E:\Word folder"


pdf_folder_path = r"E:\out"


docx_files = [f for f in os.listdir(docx_folder_path) if f.endswith(".docx")]


for docx_file in docx_files:
    # Build the full path for the input DOCX file
    docx_file_path = os.path.join(docx_folder_path, docx_file)
    pdf_file_path = os.path.join(pdf_folder_path, os.path.splitext(docx_file)[0] + ".pdf")
    convert(docx_file_path, pdf_file_path)

print("Conversion completed.")
