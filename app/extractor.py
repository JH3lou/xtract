import pandas as pd
import os
import shutil
from zipfile import ZipFile
from tkinter import filedialog, messagebox
from app.utils import get_unique_filename

# Define a hidden temp directory inside the app folder
TEMP_UNZIP_DIR = os.path.join(os.path.dirname(__file__), "temp_unzip")

def select_and_process_zip_file():
    zip_path = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")])
    if not zip_path:
        messagebox.showinfo("Info", "No ZIP file selected.")
        return

    zip_filename = os.path.basename(zip_path)
    parts = zip_filename.split(".")

    try:
        name_part = parts[0]
        date_part = parts[2][:8]
        last_number_part = parts[-2]
        formatted_date = f"{date_part[:4]}_{date_part[4:6]}_{date_part[6:8]}"
        output_filename = f"{name_part}.{formatted_date}.{last_number_part}.xlsx"
    except IndexError:
        messagebox.showerror("Error", "Invalid ZIP filename format.")
        return

    output_filepath = os.path.join(os.path.expanduser("~"), "Downloads", output_filename)

    # Ensure the temp directory is clean
    if os.path.exists(TEMP_UNZIP_DIR):
        shutil.rmtree(TEMP_UNZIP_DIR)  # Remove existing temp folder
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
            df = pd.read_csv(file_path, delimiter='|', skiprows=[0], skipfooter=1, engine='python')
            sheet_name = txt_file.split(".")[3]
            dataframes[sheet_name] = df
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file {file_path}: {str(e)}")
            return

    export_to_excel(dataframes, output_filepath)

def export_to_excel(dataframes, output_filepath):
    try:
        unique_filepath = get_unique_filename(output_filepath)
        with pd.ExcelWriter(unique_filepath, engine='openpyxl') as writer:
            for sheet_name, df in dataframes.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        messagebox.showinfo("Success", f"Data exported to {unique_filepath} successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to export to Excel: {str(e)}")

        
