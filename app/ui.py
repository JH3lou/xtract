import os
import customtkinter as ctk
from tkinter import messagebox
from app.extractor import select_and_process_zip_file
from app.converter import convert_excel_to_zip, download_template
from app.settings_window import SettingsWindow

class TaxOptUI:
    def __init__(self, root):
        self.root = root
        self.root.geometry("450x225")
        self.root.title("Xtract - File Converter")

        self.settings_icon = ctk.CTkButton(self.root, text="⚙", width=30, command=self.open_settings)
        self.settings_icon.place(relx=0.95, rely=0.05, anchor="ne")

        # Dynamically adjust working directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(current_dir)  # Ensures all relative paths work
        project_root = os.path.dirname(current_dir)

        # Convert icon path to absolute path
        icon_path = os.path.abspath(os.path.join(project_root, "assets", "siteIcon__.ico"))

        # Debugging
        # print(f"Using absolute path for icon: {icon_path}")

        try:
            root.iconbitmap(icon_path)
            #print("Icon successfully applied!")
        except Exception as e:
            #print(f"Error loading icon: {e}")
            messagebox.showwarning("Warning", "Icon file not found. Using default icon.")

        # Initialize CTkTabView
        self.tab_view = ctk.CTkTabview(self.root)
        self.tab_view.pack(expand=True, fill="both", padx=20, pady=20)

        # Add tabs
        self.tab_extract = self.tab_view.add("Extract ZIP to Excel")
        self.tab_convert = self.tab_view.add("Convert Excel to ZIP")

        # Extract ZIP to Excel Tab
        self.setup_extract_tab()

        # Convert Excel to ZIP Tab
        self.setup_convert_tab()

        # Bind close event to custom cleanup
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_extract_tab(self):
        """Setup UI for the Extract ZIP to Excel tab."""
        extract_button = ctk.CTkButton(
            master=self.tab_extract, 
            text="Select and Process ZIP Folder", 
            command=select_and_process_zip_file
        )
        extract_button.pack(pady=20, padx=10)

    def open_settings(self):
        """Open the settings window."""
        settings_window = SettingsWindow(self.root)
        settings_window.grab_set()  # Ensure it stays on top


    def setup_convert_tab(self):
        """Setup UI for the Convert Excel to ZIP tab."""
        download_button = ctk.CTkButton(
            master=self.tab_convert, 
            text="Download Excel Template", 
            command=download_template
        )
        download_button.pack(pady=10, padx=10)

        convert_button = ctk.CTkButton(
            master=self.tab_convert, 
            text="Convert Excel to ZIP", 
            command=convert_excel_to_zip
        )
        convert_button.pack(pady=10, padx=10)

    def on_close(self):
        """Handle the application cleanup before exiting."""
        try:
            for task in self.root.tk.call('after', 'info'):
                self.root.after_cancel(task)
        except Exception:
            pass
        self.root.destroy()

