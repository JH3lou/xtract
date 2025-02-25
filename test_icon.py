import tkinter as tk

root = tk.Tk()
root.title("Test Icon")

# Set the icon
icon_path = "assets/siteIcon__.ico"
print(f"Trying to load icon: {icon_path}")

try:
    root.iconbitmap(icon_path)
    print("Icon successfully loaded.")
except Exception as e:
    print(f"Failed to load icon: {e}")

root.mainloop()
