import os
import shutil
import random
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from zipfile import ZipFile
import pandas as pd
import openpyxl

# Define the Downloads folder path
DOWNLOADS_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")

def download_template():
    """Download the existing Excel template to the Downloads folder and prompt the user to open it."""
    templates_folder = os.path.join(os.path.dirname(__file__), "..", "templates")
    source_template_path = os.path.join(templates_folder, "Excel_Template.xlsx")

    if not os.path.exists(source_template_path):
        messagebox.showerror("Error", "Template file not found in templates folder.")
        return

    destination_path = os.path.join(DOWNLOADS_FOLDER, "Excel_Template.xlsx")

    try:
        shutil.copy(source_template_path, destination_path)

        # Show success message with Open option
        root = tk.Toplevel()  # Use Toplevel instead of Tk()
        root.withdraw()  # Hide the window

        result = messagebox.askyesno(
            "Template Downloaded",
            f"Template saved to {destination_path}\n\nWould you like to open it now?"
        )

        if result:
            os.startfile(destination_path)  # Open file in Excel (Windows)
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to download template: {str(e)}")
    finally:
        root.destroy()  # Ensure Tkinter properly closes

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
    root = tk.Toplevel()  # Use Toplevel instead of Tk()
    root.withdraw()  # Hide the root window
    if error:
        messagebox.showerror(title, message)
    else:
        messagebox.showinfo(title, message)
    root.destroy()  # Ensure Tkinter properly closes

def convert_excel_to_zip():
    """Convert an Excel workbook to a ZIP file with properly formatted TXT files."""
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not excel_path:
        show_message("Info", "No Excel file selected.")
        return

    trading_request_id = random.randint(100000, 999999)
    now = datetime.datetime.now()
    timestamp = now.strftime("%Y%m%d:%H:%M:%S:%f")[:-3]

    zip_folder_name = f"TAXOPT.LPB.{timestamp.replace(':', '')}.{trading_request_id}"
    zip_folder_path = os.path.join(DOWNLOADS_FOLDER, zip_folder_name)
    os.makedirs(zip_folder_path, exist_ok=True)

    try:
        # Ensure proper closure of Excel file
        with pd.ExcelFile(excel_path) as excel_data:
            valid_sheets = [sheet for sheet in excel_data.sheet_names if "__FDSCACHE__" not in sheet]
            
            for sheet_name in valid_sheets:
                df = pd.read_excel(excel_data, sheet_name=sheet_name, dtype=str)

                # ✅ Strip whitespace and remove empty rows
                df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
                df = df.dropna(how="all")  # ✅ Remove rows that are completely empty

                # Retrieve correct header name
                header_name = HEADER_NAME_MAPPING.get(sheet_name, sheet_name)

                # Format header and trailer
                header_record = f"HDR|LPB|{trading_request_id}|{timestamp}|{header_name}"
                trailer_record = f"TLR|{len(df)}"

                # ✅ Convert DataFrame to text format without extra blank lines
                df_str = df.to_csv(index=False, sep='|', header=True, lineterminator='\n').strip()

                # Combine header, data, and trailer
                txt_content = f"{header_record}\n{df_str}\n{trailer_record}"

                # Define file name and path
                txt_filename = f"TAXOPT.{timestamp.replace(':', '')}.{trading_request_id}.{sheet_name}.txt"
                txt_file_path = os.path.join(zip_folder_path, txt_filename)

                # Write to text file
                with open(txt_file_path, 'w', encoding='utf-8') as txt_file:
                    txt_file.write(txt_content)

        # Create ZIP archive and exclude `_FDSCACHE__` files
        zip_file_path = os.path.join(DOWNLOADS_FOLDER, f"{zip_folder_name}.zip")
        with ZipFile(zip_file_path, 'w') as zipf:
            for txt_file in os.listdir(zip_folder_path):
                if "_FDSCACHE__" not in txt_file:  # ✅ Exclude FactSet cache files
                    zipf.write(os.path.join(zip_folder_path, txt_file), txt_file)

        # Automatically delete folder after creating ZIP
        shutil.rmtree(zip_folder_path)

        show_message("Success", f"ZIP file created at {zip_file_path}")

    except Exception as e:
        show_message("Error", f"Failed to convert Excel to ZIP: {str(e)}", error=True)
