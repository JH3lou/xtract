import pandas as pd
import os
import shutil
from zipfile import ZipFile
from tkinter import filedialog, messagebox
from app.utils import get_unique_filename
from app.template_manager import TemplateManager

# Define a hidden temp directory inside the app folder
TEMP_UNZIP_DIR = os.path.join(os.path.dirname(__file__), "temp_unzip")

# Initialize Template Manager
template_manager = TemplateManager()

def select_and_process_zip_file(template='TaxOverlayS3Data'):
    zip_path = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")])
    if not zip_path:
        messagebox.showinfo("Info", "No ZIP file selected.")
        return

    # Template Configuration
    template_config = template_manager.get_template(template)

    # Extract Naming Logic via Dynamic Token Resolution
    zip_filename = template_manager.resolve_tokens(
        template_config["zip_naming_convention"]
    )

    output_filename = template_manager.resolve_tokens(
        f"{zip_filename}.xlsx"
    )
    
    output_filepath = os.path.join(os.path.expanduser("~"), "Downloads", output_filename)

    # Ensure the temp directory is clean
    if os.path.exists(TEMP_UNZIP_DIR):
        shutil.rmtree(TEMP_UNZIP_DIR)
    os.makedirs(TEMP_UNZIP_DIR, exist_ok=True)

    with ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(TEMP_UNZIP_DIR)

    txt_files = [f for f in os.listdir(TEMP_UNZIP_DIR) if f.endswith('.txt')]

    if not txt_files:
        messagebox.showinfo("Info", "No .txt files found in the selected ZIP folder.")
        return

    dataframes = {}
    for txt_file in txt_files:
        file_path = os.path.join(TEMP_UNZIP_DIR, txt_file)

        try:
            df = pd.read_csv(
                file_path,
                delimiter=template_config.get("delimiter", "|"),
                skiprows=template_config.get("header_records", 0),
                skipfooter=template_config.get("trailer_records", 0),
                engine='python'
            )

            sheet_name = txt_file.split(".")[3]

            header_record = template_manager.resolve_tokens(
                template_config["header_format"],
                sheet_name=sheet_name
            )

            trailer_record = template_manager.resolve_tokens(
                template_config["trailer_format"],
                row_count=len(df)
            )

            # Combine data with header and trailer
            dataframes[sheet_name] = pd.DataFrame(
                [[header_record]] +
                df.values.tolist() +
                [[trailer_record]]
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file {file_path}: {str(e)}")
            return

    export_to_excel(dataframes, output_filepath)

def export_to_excel(dataframes, output_filepath):
    try:
        unique_filepath = get_unique_filename(output_filepath)
        with pd.ExcelWriter(unique_filepath, engine='openpyxl') as writer:
            for sheet_name, df in dataframes.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        messagebox.showinfo("Success", f"Data exported to {unique_filepath} successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to export to Excel: {str(e)}")
