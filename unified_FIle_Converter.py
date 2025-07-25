import os
import shutil
import csv
import datetime
import win32com.client

# === CONFIGURE PATHS ===
source_folder = r"C:\path\to\source_files"  # Folder containing all Office files
output_folder = r"C:\path\to\pdf_output"
processed_folder = r"C:\path\to\processed_files"
log_file_path = r"C:\path\to\conversion_failures.csv"

# === SUPPORTED FILE TYPES ===
word_exts = (".doc", ".docx")
excel_exts = (".xls", ".xlsx", ".xlsm", ".xlsb", ".csv")
ppt_exts = (".ppt", ".pptx")

# === CREATE FOLDERS IF NEEDED ===
os.makedirs(output_folder, exist_ok=True)
os.makedirs(processed_folder, exist_ok=True)

# === LOG LIST FOR FAILURES ===
log_entries = []


def log_failure(filename, error):
    """Append failure entry to log list."""
    log_entries.append({
        "File Name": filename,
        "Error": str(error),
        "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    })


def convert_word_to_pdf(word_app, input_file, output_pdf):
    """Convert a Word document to PDF."""
    doc = word_app.Documents.Open(input_file)
    doc.SaveAs(output_pdf, FileFormat=17)  # 17 = PDF
    doc.Close()


def convert_excel_to_pdf(excel_app, input_file, output_pdf):
    """Convert an Excel workbook (all sheets) to PDF, scaled to fit width."""
    workbook = excel_app.Workbooks.Open(input_file)

    for sheet in workbook.Sheets:
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesTall = False
        sheet.PageSetup.FitToPagesWide = 1  # Fit columns to one page width

    workbook.ExportAsFixedFormat(0, output_pdf)  # 0 = PDF
    workbook.Close(False)


def convert_ppt_to_pdf(ppt_app, input_file, output_pdf):
    """Convert a PowerPoint presentation to PDF."""
    presentation = ppt_app.Presentations.Open(input_file, WithWindow=False)
    presentation.SaveAs(output_pdf, 32)  # 32 = PDF
    presentation.Close()


def process_files():
    """Main processing loop for converting all supported files in the source folder."""
    # Start COM apps
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = False

    # Loop through files
    for filename in os.listdir(source_folder):
        filepath = os.path.join(source_folder, filename)
        pdf_name = os.path.splitext(filename)[0] + ".pdf"
        pdf_path = os.path.join(output_folder, pdf_name)

        try:
            if filename.lower().endswith(word_exts):
                convert_word_to_pdf(word, filepath, pdf_path)

            elif filename.lower().endswith(excel_exts):
                convert_excel_to_pdf(excel, filepath, pdf_path)

            elif filename.lower().endswith(ppt_exts):
                convert_ppt_to_pdf(powerpoint, filepath, pdf_path)

            else:
                continue  # Skip unsupported file types

            # Move processed file
            shutil.move(filepath, os.path.join(processed_folder, filename))
            print(f"✅ Converted and moved: {filename}")

        except Exception as e:
            print(f"❌ Failed to convert {filename}: {e}")
            log_failure(filename, e)

    # Quit apps
    word.Quit()
    excel.Quit()
    powerpoint.Quit()

    # Write failure log if needed
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
        print("\n✅ All files converted successfully. No errors to log.")


if __name__ == "__main__":
    process_files()
