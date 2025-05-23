import os
import shutil
import random
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from zipfile import ZipFile
import pandas as pd
import openpyxl
from app.utils import replace_placeholders
from app.utils import (
    replace_placeholders,
    load_global_settings,   # ✅ Newly imported
    load_template_settings  # ✅ Newly imported
)


# Define the Downloads folder path
DOWNLOADS_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")

def download_template():
    """Download the active Excel template to the Downloads folder."""
    global_settings = load_global_settings()
    active_template = global_settings.get("active_template", "Default")
    template_filename = f"{active_template}_template.xlsx"

    templates_folder = os.path.join(os.path.dirname(__file__), "..", "templates")
    source_template_path = os.path.join(templates_folder, template_filename)

    if not os.path.exists(source_template_path):
        messagebox.showerror("Error", f"Template '{active_template}' not found.")
        return

    destination_path = os.path.join(DOWNLOADS_FOLDER, template_filename)

    try:
        shutil.copy(source_template_path, destination_path)
        os.startfile(destination_path)

    except Exception as e:
        messagebox.showerror("Error", f"Failed to download template: {str(e)}")

# Define the correct mapping for worksheet names to header record names
HEADER_NAME_MAPPING = {
    "ACCTMST": "AccountMaster",
    "DISCRET": "DiscretionarySleeve",
    "HOLDING": "Holding",
    "MODALLOC": "ModelAllocation",
    "MSTMODALLOC": "MasterModelAllocation",
    "POSITION": "Position",
    "RESTRC": "Restriction",
    "RESTRCGRPITEM": "RestrictionGroupItem",
    "TAXBUDGET": "TaxBudget",
    "TAXLOT": "Taxlot"
}

def show_message(title, message, error=False):
    """Show a Tkinter message box safely without event loop issues."""
    root = tk.Toplevel()
    root.withdraw()  # Hide the root window
    if error:
        messagebox.showerror(title, message)
    else:
        messagebox.showinfo(title, message)
    root.destroy()  # Ensure Tkinter properly closes

def convert_excel_to_zip():
    """Convert an Excel workbook to a ZIP file with dynamically defined headers & trailers."""
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not excel_path:
        show_message("Info", "No Excel file selected.")
        return

    global_settings = load_global_settings()
    active_template = global_settings.get("active_template", "Default")
    template_settings = load_template_settings(active_template)

    trading_request_id = random.randint(100000, 999999)
    now = datetime.datetime.now()
    timestamp = now.strftime("%Y%m%d%H%M%S")

    zip_folder_name = f"{template_settings['format_name']}.{timestamp}.{trading_request_id}"
    zip_folder_path = os.path.join(DOWNLOADS_FOLDER, zip_folder_name)
    os.makedirs(zip_folder_path, exist_ok=True)

    try:
        with pd.ExcelFile(excel_path) as excel_data:
            for sheet_name in excel_data.sheet_names:
                # Skip FactSet Cache Files
                if "__FDSCACHE__" in sheet_name:
                    continue

                df = pd.read_excel(excel_data, sheet_name=sheet_name, dtype=str)

                # Apply user-defined column mappings
                df.rename(columns=template_settings["column_mappings"], inplace=True)

                # Strip whitespace from cells
                df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

                # Determine the header name dynamically using header_mappings
                header_name = template_settings.get("header_mappings", {}).get(sheet_name, sheet_name)

                # Corrected row count logic (after filtering unwanted rows)
                row_count = len(df)

                # Generate headers/trailers using enhanced placeholder logic
                header_record = replace_placeholders(template_settings["header_format"], header_name, row_count)
                trailer_record = replace_placeholders(template_settings["trailer_format"], header_name, row_count)

                # Apply formatting and delimiter
                df_str = df.to_csv(index=False, sep=template_settings["delimiter"], header=True).strip()

                # Combine all parts
                txt_content = f"{header_record}\n{df_str}\n{trailer_record}"

                # Generate filename dynamically
                txt_filename = template_settings["file_naming_pattern"].format(
                    timestamp=timestamp,
                    format_name=template_settings["format_name"]
                )
                txt_file_path = os.path.join(zip_folder_path, txt_filename)

                with open(txt_file_path, 'w', encoding='utf-8') as txt_file:
                    txt_file.write(txt_content)

        # Create ZIP archive
        zip_file_path = os.path.join(DOWNLOADS_FOLDER, f"{zip_folder_name}.zip")
        with ZipFile(zip_file_path, 'w') as zipf:
            for txt_file in os.listdir(zip_folder_path):
                zipf.write(os.path.join(zip_folder_path, txt_file), txt_file)

        # Clean up: Delete the folder after creating the ZIP
        shutil.rmtree(zip_folder_path)

        show_message("Success", f"ZIP file created at {zip_file_path}")

    except Exception as e:
        show_message("Error", f"Failed to convert Excel to ZIP: {str(e)}", error=True)
