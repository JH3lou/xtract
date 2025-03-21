import os
import shutil
from zipfile import ZipFile
import pandas as pd
import openpyxl
from app.template_manager import TemplateManager
from tkinter import filedialog, messagebox

# Define the Downloads folder path
DOWNLOADS_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")

# Initialize Template Manager
template_manager = TemplateManager()

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
        result = messagebox.askyesno(
            "Template Downloaded",
            f"Template saved to {destination_path}\n\nWould you like to open it now?"
        )

        if result:
            os.startfile(destination_path)  # Open file in Excel (Windows)
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to download template: {str(e)}")


def convert_excel_to_zip(template='TaxOverlayS3Data'):
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not excel_path:
        messagebox.showinfo("Info", "No Excel file selected.")
        return

    template_config = template_manager.get_template(template)

    # Dynamic Naming Logic
    zip_folder_name = template_manager.resolve_tokens(
        template_config["zip_naming_convention"]
    )

    zip_folder_path = os.path.join(DOWNLOADS_FOLDER, zip_folder_name)
    os.makedirs(zip_folder_path, exist_ok=True)

    try:
        with pd.ExcelFile(excel_path) as excel_data:
            valid_sheets = [sheet for sheet in excel_data.sheet_names if "__FDSCACHE__" not in sheet]

            for sheet_name in valid_sheets:
                df = pd.read_excel(excel_data, sheet_name=sheet_name, dtype=str)

                # Resolve Header and Trailer Formats Using Tokens
                header_record = template_manager.resolve_tokens(
                    template_config["header_format"],
                    sheet_name=sheet_name
                )

                trailer_record = template_manager.resolve_tokens(
                    template_config["trailer_format"],
                    row_count=len(df)
                )

                # Convert DataFrame to text format
                df_str = df.to_csv(index=False, sep=template_config['delimiter'], header=True, lineterminator='\n').strip()

                # TXT File Naming
                txt_filename = template_manager.resolve_tokens(
                    template_config["txt_filename_pattern"],
                    sheet_name=sheet_name
                )

                txt_file_path = os.path.join(zip_folder_path, txt_filename)

                with open(txt_file_path, 'w', encoding='utf-8') as txt_file:
                    txt_file.write(f"{header_record}\n{df_str}\n{trailer_record}")

        # Create ZIP Archive
        zip_file_path = os.path.join(DOWNLOADS_FOLDER, f"{zip_folder_name}.zip")
        with ZipFile(zip_file_path, 'w') as zipf:
            for txt_file in os.listdir(zip_folder_path):
                zipf.write(os.path.join(zip_folder_path, txt_file), txt_file)

        shutil.rmtree(zip_folder_path)

        messagebox.showinfo("Success", f"ZIP file created at {zip_file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert Excel to ZIP: {str(e)}")
