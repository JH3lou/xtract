import os
import customtkinter as ctk
from tkinter import filedialog, messagebox, StringVar
from threading import Thread
from app.extractor import select_and_process_zip_file
from app.converter import convert_excel_to_zip, download_template
from app.template_manager import TemplateManager

class TaxOptUI:
    def __init__(self, root):
        self.root = root
        self.root.geometry("500x350")
        self.root.title("Xtract - File Converter")

        # Initialize TemplateManager
        self.template_manager = TemplateManager()

        # Initialize CTkTabView
        self.tab_view = ctk.CTkTabview(self.root)
        self.tab_view.pack(expand=True, fill="both", padx=20, pady=20)

        # Add tabs
        self.tab_extract = self.tab_view.add("Extract ZIP to Excel")
        self.tab_convert = self.tab_view.add("Convert Excel to ZIP")
        self.tab_settings = self.tab_view.add("Settings")  # NEW TAB for Settings

        # Extract ZIP to Excel Tab
        self.setup_extract_tab()

        # Convert Excel to ZIP Tab
        self.setup_convert_tab()

        # Settings Tab
        self.setup_settings_tab()

        # Bind close event to custom cleanup
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ================================
    # Extract ZIP to Excel Tab
    # ================================
    def setup_extract_tab(self):
        self.extract_template_var = StringVar(self.root)
        self.extract_template_var.set("default")

        template_dropdown = ctk.CTkComboBox(
            master=self.tab_extract,
            values=self.template_manager.list_templates(),
            variable=self.extract_template_var
        )
        template_dropdown.pack(pady=10, padx=10)

        extract_button = ctk.CTkButton(
            master=self.tab_extract,
            text="Select and Process ZIP Folder",
            command=lambda: self.run_with_loading_screen(self.process_zip_with_template)
        )
        extract_button.pack(pady=10, padx=10)

    def process_zip_with_template(self):
        select_and_process_zip_file(template=self.extract_template_var.get())

    # ================================
    # Convert Excel to ZIP Tab
    # ================================
    def setup_convert_tab(self):
        self.convert_template_var = StringVar(self.root)
        self.convert_template_var.set("default")

        template_dropdown = ctk.CTkComboBox(
            master=self.tab_convert,
            values=self.template_manager.list_templates(),
            variable=self.convert_template_var
        )
        template_dropdown.pack(pady=10, padx=10)

        download_button = ctk.CTkButton(
            master=self.tab_convert,
            text="Download Excel Template",
            command=download_template
        )
        download_button.pack(pady=10, padx=10)

        convert_button = ctk.CTkButton(
            master=self.tab_convert,
            text="Convert Excel to ZIP",
            command=lambda: self.run_with_loading_screen(self.convert_excel_with_template)
        )
        convert_button.pack(pady=10, padx=10)

    def convert_excel_with_template(self):
        convert_excel_to_zip(template=self.convert_template_var.get())

    # ================================
    # Settings Tab - Template Management
    # ================================
    def setup_settings_tab(self):
        ctk.CTkLabel(master=self.tab_settings, text="Template Name:").pack(pady=5)
        self.new_template_entry = ctk.CTkEntry(master=self.tab_settings)
        self.new_template_entry.pack(pady=5)

        ctk.CTkLabel(master=self.tab_settings, text="Delimiter:").pack(pady=5)
        self.delimiter_entry = ctk.CTkEntry(master=self.tab_settings)
        self.delimiter_entry.insert(0, "|")  # Default delimiter
        self.delimiter_entry.pack(pady=5)

        ctk.CTkLabel(master=self.tab_settings, text="Header Records:").pack(pady=5)
        self.header_entry = ctk.CTkEntry(master=self.tab_settings)
        self.header_entry.insert(0, "0")
        self.header_entry.pack(pady=5)

        ctk.CTkLabel(master=self.tab_settings, text="Trailer Records:").pack(pady=5)
        self.trailer_entry = ctk.CTkEntry(master=self.tab_settings)
        self.trailer_entry.insert(0, "0")
        self.trailer_entry.pack(pady=5)

        upload_button = ctk.CTkButton(
            master=self.tab_settings,
            text="Upload Excel Template",
            command=self.upload_excel_template
        )
        upload_button.pack(pady=10)

        save_button = ctk.CTkButton(
            master=self.tab_settings,
            text="Save New Template",
            command=self.save_template
        )
        save_button.pack(pady=10)

        delete_button = ctk.CTkButton(
            master=self.tab_settings,
            text="Delete Template",
            command=self.delete_template
        )
        delete_button.pack(pady=10)

    def upload_excel_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        template_name = self.new_template_entry.get().strip()
        if not template_name:
            messagebox.showerror("Error", "Please enter a template name first.")
            return

        self.template_manager.upload_excel_template(file_path, template_name)
        messagebox.showinfo("Success", f"Excel template '{template_name}' uploaded successfully.")

    def save_template(self):
        template_name = self.new_template_entry.get().strip()
        if not template_name:
            messagebox.showerror("Error", "Please enter a template name.")
            return

        self.template_manager.create_template(
            template_name,
            self.delimiter_entry.get(),
            int(self.header_entry.get()),
            int(self.trailer_entry.get())
        )
        messagebox.showinfo("Success", f"Template '{template_name}' saved successfully.")

    def delete_template(self):
        template_name = self.new_template_entry.get().strip()
        if not template_name:
            messagebox.showerror("Error", "Please enter a template name.")
            return

        if self.template_manager.delete_template(template_name):
            messagebox.showinfo("Success", f"Template '{template_name}' deleted successfully.")
        else:
            messagebox.showerror("Error", f"Template '{template_name}' not found.")

    # ================================
    # Loading Screen
    # ================================
    def run_with_loading_screen(self, task_function):
        loading_screen = ctk.CTkToplevel(self.root)
        loading_screen.title("Processing...")
        ctk.CTkLabel(loading_screen, text="Processing... Please wait").pack(padx=20, pady=20)

        def process():
            task_function()
            loading_screen.destroy()

        Thread(target=process).start()

    def on_close(self):
        try:
            for task in self.root.tk.call('after', 'info'):
                self.root.after_cancel(task)
        except Exception:
            pass
        self.root.destroy()
