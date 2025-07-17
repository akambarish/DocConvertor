#!/usr/bin/env python
# coding: utf-8

# ✅ Features:
# Supports .csv, .xls, .xlsx, .xlsm, .xlsb
# 
# Combines all sheets into one PDF
# 
# Fits entire sheet to one PDF page (in terms of columns) by scaling
# 
# Reduces font size as needed (via page scaling)
# 
# Moves successfully converted files to a processed folder
# 
# Logs errors to a CSV (conversion_failures.csv)

# In[ ]:


import os
import shutil
import csv
import datetime
import win32com.client

# === CONFIGURE PATHS ===
source_folder = r"C:\path\to\excel_files"
output_folder = r"C:\path\to\pdf_output"
processed_folder = r"C:\path\to\processed_excels"
log_file_path = r"C:\path\to\conversion_failures.csv"

# === CREATE NECESSARY FOLDERS ===
os.makedirs(output_folder, exist_ok=True)
os.makedirs(processed_folder, exist_ok=True)

# === SUPPORTED EXCEL FILE TYPES ===
excel_extensions = (".xls", ".xlsx", ".xlsm", ".xlsb", ".csv")

# === INIT FAILURE LOG LIST ===
log_entries = []

# === START EXCEL COM INSTANCE ===
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

# === CONVERT FILES ===
for filename in os.listdir(source_folder):
    if filename.lower().endswith(excel_extensions):
        input_path = os.path.join(source_folder, filename)
        pdf_name = os.path.splitext(filename)[0] + ".pdf"
        pdf_path = os.path.join(output_folder, pdf_name)

        try:
            # Open Excel Workbook
            workbook = excel.Workbooks.Open(input_path)

            # Set all sheets to fit columns on one page
            for sheet in workbook.Sheets:
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesTall = False  # Allows multiple pages vertically
                sheet.PageSetup.FitToPagesWide = 1       # Shrinks columns to fit one page width

            # Export all sheets to a single PDF
            workbook.ExportAsFixedFormat(0, pdf_path)  # 0 = PDF
            workbook.Close(False)

            # Move original Excel file to processed folder
            shutil.move(input_path, os.path.join(processed_folder, filename))
            print(f"✅ Converted and moved: {filename}")

        except Exception as e:
            print(f"❌ Failed to convert {filename}: {e}")
            log_entries.append({
                "File Name": filename,
                "Error": str(e),
                "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

# === CLOSE EXCEL ===
excel.Quit()

# === WRITE FAILURE LOG (IF ANY) ===
if log_entries:
    file_exists = os.path.exists(log_file_path)
    with open(log_file_path, mode='a', newline='', encoding='utf-8') as log_file:
        fieldnames = ["File Name", "Error", "Timestamp"]
        writer = csv.DictWriter(log_file, fieldnames=fieldnames)

        if not file_exists:
            writer.writeheader()
        writer.writerows(log_entries)

    print(f"\n⚠️ Logged {len(log_entries)} failure(s) to: {log_file_path}")
else:
    print("\n✅ All Excel files converted successfully. No errors to log.")

