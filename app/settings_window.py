import os
import json
import customtkinter as ctk
from tkinter import messagebox, filedialog

# Define file paths
SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "settings.json")
TEMPLATES_FOLDER = os.path.join(os.path.dirname(__file__), "..", "templates")
SETTINGS_CONFIG_FOLDER = os.path.join(os.path.dirname(__file__), "..", "settings_configs")

# Ensure settings_configs folder exists
if not os.path.exists(SETTINGS_CONFIG_FOLDER):
    os.makedirs(SETTINGS_CONFIG_FOLDER)

# Default Configuration Values
DEFAULT_TEMPLATE_SETTINGS = {
    "header_format": "HDR|{timestamp}|{file_name}|{random_6_digits}",
    "trailer_format": "TLR|{row_count}",
    "header_mappings": {
        "ACCTMST": "AccountMaster",
        "DISCRET": "DiscretionarySleeve"
    }
}

# Load Settings
def load_global_settings():
    if not os.path.exists(SETTINGS_FILE):
        return {
            "active_template": "Default",
            "delimiter_extract": "|",
            "delimiter_zip": "|",
            "enable_header": True,
            "enable_trailer": True
        }

    try:
        with open(SETTINGS_FILE, "r") as file:
            return json.load(file)
    except (json.JSONDecodeError, ValueError):
        return {
            "active_template": "Default",
            "delimiter_extract": "|",
            "delimiter_zip": "|",
            "enable_header": True,
            "enable_trailer": True
        }

def save_global_settings(settings):
    with open(SETTINGS_FILE, "w") as file:
        json.dump(settings, file, indent=4)

def load_template_settings(template_name):
    """Load template settings with fallback logic for missing keys."""
    config_path = os.path.join(SETTINGS_CONFIG_FOLDER, f"{template_name}.json")

    if not os.path.exists(config_path):
        save_template_settings(template_name, DEFAULT_TEMPLATE_SETTINGS)
        return DEFAULT_TEMPLATE_SETTINGS.copy()

    try:
        with open(config_path, "r") as file:
            template_settings = json.load(file)
    except (json.JSONDecodeError, ValueError):
        return DEFAULT_TEMPLATE_SETTINGS.copy()

    for key, value in DEFAULT_TEMPLATE_SETTINGS.items():
        if key not in template_settings:
            template_settings[key] = value

    return template_settings

def save_template_settings(template_name, settings):
    config_path = os.path.join(SETTINGS_CONFIG_FOLDER, f"{template_name}.json")

    # Ensure folder exists before writing
    if not os.path.exists(SETTINGS_CONFIG_FOLDER):
        os.makedirs(SETTINGS_CONFIG_FOLDER)

    with open(config_path, "w") as file:
        json.dump(settings, file, indent=4)

def get_available_templates():
    if not os.path.exists(TEMPLATES_FOLDER):
        os.makedirs(TEMPLATES_FOLDER)

    templates = [
        file.replace("_template.xlsx", "")
        for file in os.listdir(TEMPLATES_FOLDER)
        if file.endswith("_template.xlsx")
    ]
    return templates

# Settings Window Class
class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Settings")
        self.geometry("500x700")

        self.global_settings = load_global_settings()
        self.active_template = self.global_settings.get("active_template", "Default")

        self.global_settings["available_templates"] = get_available_templates() or ["Default"]
        save_global_settings(self.global_settings)

        self.template_settings = load_template_settings(self.active_template)
        self.last_saved_settings = self.template_settings.copy()

        self.setup_ui()

    def setup_ui(self):
        self.columnconfigure(1, weight=1)

        # Template Selection
        ctk.CTkLabel(self, text="Select Active Template:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.template_var = ctk.StringVar(value=self.active_template)
        self.template_dropdown = ctk.CTkComboBox(
            self,
            values=self.global_settings.get("available_templates", ["Default"]),
            variable=self.template_var,
            command=self.change_template
        )
        self.template_dropdown.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # Create New Template Button
        self.create_template_button = ctk.CTkButton(self, text="Create New Template", command=self.create_new_template)
        self.create_template_button.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

        # Excel Template Upload Section
        self.upload_template_button = ctk.CTkButton(self, text="Upload Excel Template", command=self.upload_excel_template)
        self.upload_template_button.grid(row=2, column=0, columnspan=2, padx=10, pady=5)

        # Header & Trailer Entry
        ctk.CTkLabel(self, text="Header Format:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.header_entry = ctk.CTkEntry(self)
        self.header_entry.insert(0, self.template_settings.get("header_format", ""))
        self.header_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(self, text="Trailer Format:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.trailer_entry = ctk.CTkEntry(self)
        self.trailer_entry.insert(0, self.template_settings.get("trailer_format", ""))
        self.trailer_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        # Buttons
        self.save_button = ctk.CTkButton(self, text="Save", command=self.save_changes)
        self.save_button.grid(row=6, column=0, padx=5, pady=5)

        self.reset_button = ctk.CTkButton(self, text="Reset", command=self.reset_to_last_saved)
        self.reset_button.grid(row=6, column=1, padx=5, pady=5)

    def change_template(self, event=None):
        """Change template and load its corresponding data."""
        self.active_template = self.template_var.get()
        self.template_settings = load_template_settings(self.active_template)

        self.header_entry.delete(0, "end")
        self.header_entry.insert(0, self.template_settings.get("header_format", ""))

        self.trailer_entry.delete(0, "end")
        self.trailer_entry.insert(0, self.template_settings.get("trailer_format", ""))

        self.last_saved_settings = self.template_settings.copy()

    def reset_to_last_saved(self):
        self.header_entry.delete(0, "end")
        self.header_entry.insert(0, self.last_saved_settings.get("header_format", ""))
        self.trailer_entry.delete(0, "end")
        self.trailer_entry.insert(0, self.last_saved_settings.get("trailer_format", ""))

    def upload_excel_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            file_name = os.path.basename(file_path)
            destination = os.path.join(TEMPLATES_FOLDER, file_name)
            os.rename(file_path, destination)
            messagebox.showinfo("Success", "Template uploaded successfully.")

    def create_new_template(self):
        template_name = filedialog.asksaveasfilename(
            defaultextension=".json",
            initialdir=SETTINGS_CONFIG_FOLDER,
            title="Create New Template"
        )
        if template_name:
            save_template_settings(os.path.basename(template_name).replace(".json", ""), DEFAULT_TEMPLATE_SETTINGS)
            messagebox.showinfo("Success", "New template created successfully.")

    def save_changes(self):
        self.template_settings.update({
            "header_format": self.header_entry.get(),
            "trailer_format": self.trailer_entry.get()
        })
        save_template_settings(self.active_template, self.template_settings)
        messagebox.showinfo("Success", "Settings saved successfully.")
        self.destroy()
