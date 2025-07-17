#!/usr/bin/env python
# coding: utf-8

# Here's a Python script that will:
# 
# Convert all .docx files in a folder to .pdf.
# 
# Move successfully converted .docx files to a separate "processed" folder.
# 
# Remove the original .docx file from the source folder after conversion.
# 
# ✅ Requirements:
# Python 3.x
# 
# docx2pdf library (for conversion)
# 
# Windows system (for docx2pdf; Mac support is limited)
# 
# Microsoft Word installed
# 
# 📜 Script: Bulk Word to PDF with Move on Success
# python
# Copy
# Edit
# import os
# import shutil
# from docx2pdf import convert
# 
# # Configure paths
# source_folder = r"C:\path\to\word_files"
# output_folder = r"C:\path\to\pdf_output"
# processed_folder = r"C:\path\to\processed_docs"
# 
# # Create folders if they don't exist
# os.makedirs(output_folder, exist_ok=True)
# os.makedirs(processed_folder, exist_ok=True)
# 
# # Process each .docx file
# for filename in os.listdir(source_folder):
#     if filename.lower().endswith(".docx"):
#         docx_path = os.path.join(source_folder, filename)
#         pdf_path = os.path.join(output_folder, os.path.splitext(filename)[0] + ".pdf")
#         try:
#             convert(docx_path, pdf_path)
#             # Move .docx to processed folder
#             shutil.move(docx_path, os.path.join(processed_folder, filename))
#             print(f"Converted and moved: {filename}")
#         except Exception as e:
#             print(f"Failed to convert {filename}: {e}")
# 🔧 Install Required Library
# Run this in your terminal or VS Code terminal:
# 
# bash
# Copy
# Edit
# pip install docx2pdf
# ✅ Notes:
# Only .docx files are supported, not .doc.
# 
# The script skips already converted files (those already moved).
# 
# For non-Windows systems, use pypandoc or comtypes alternatives, but setup is more complex.
# 
# Would you like a version that works on Linux/Mac or with .doc files too?
# 
# 
# ⚠️ Notes
# This only works on Windows with Microsoft Word installed.
# 
# The FileFormat=17 argument saves the file as PDF.
# 
# Word automation can hang if there's a corrupt document; handle with care on large batches.
# 

# In[1]:


import os
import shutil
import csv
import datetime
import win32com.client

# Configure folders
source_folder = r"C:\Users\ambar\Projects\PPT to PDFs\doc"
output_folder = r"C:\Users\ambar\Projects\PPT to PDFs\doc-pdf"
processed_folder = r"C:\Users\ambar\Projects\PPT to PDFs\doc-pdf1"
log_file_path = r"C:\Users\ambar\Projects\PPT to PDFs\conversion_failures.csv"

# Create necessary folders if they don't exist
os.makedirs(output_folder, exist_ok=True)
os.makedirs(processed_folder, exist_ok=True)

# Initialize failure log
log_entries = []

# Start Microsoft Word COM instance
word = win32com.client.Dispatch("Word.Application")
word.Visible = False

# Loop over documents
for filename in os.listdir(source_folder):
    if filename.lower().endswith((".doc", ".docx")):
        input_path = os.path.join(source_folder, filename)
        output_pdf_name = os.path.splitext(filename)[0] + ".pdf"
        output_path = os.path.join(output_folder, output_pdf_name)

        try:
            # Open and save as PDF
            doc = word.Documents.Open(input_path)
            doc.SaveAs(output_path, FileFormat=17)  # 17 = wdFormatPDF
            doc.Close()

            # Move original to processed folder
            shutil.move(input_path, os.path.join(processed_folder, filename))
            print(f"✅ Converted and moved: {filename}")

        except Exception as e:
            print(f"❌ Failed to convert {filename}: {e}")
            log_entries.append({
                "File Name": filename,
                "Error": str(e),
                "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

# Quit Word
word.Quit()

# Write failure log if any failures occurred
if log_entries:
    file_exists = os.path.exists(log_file_path)
    with open(log_file_path, mode='a', newline='', encoding='utf-8') as log_file:
        fieldnames = ["File Name", "Error", "Timestamp"]
        writer = csv.DictWriter(log_file, fieldnames=fieldnames)

        if not file_exists:
            writer.writeheader()
        writer.writerows(log_entries)

    print(f"\n Logged {len(log_entries)} failure(s) to: {log_file_path}")
else:
    print("\n All files converted successfully. No errors to log.")


# In[ ]:




