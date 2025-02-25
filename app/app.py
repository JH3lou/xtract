## This file initializes the application and integrates all modules.

import customtkinter as ctk
from app.ui import TaxOptUI  # Import UI from app package

def main():
    app = ctk.CTk()
    app.geometry("500x300")
    app.title("Xtract - File Converter")

    # Load UI
    TaxOptUI(app)

    # Run application
    app.mainloop()

if __name__ == "__main__":
    main()
