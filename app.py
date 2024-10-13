import tkinter as tk
from tkinter import filedialog, messagebox
from convert_to_json import convert_excel_to_json

def browse_file():
    # Open file dialog to select an Excel file
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_excel_file.delete(0, tk.END)
        entry_excel_file.insert(0, file_path)

def save_file():
    # Open file dialog to select a location to save JSON file
    file_path = filedialog.asksaveasfilename(
        defaultextension=".json",
        filetypes=[("JSON files", "*.json")])
    if file_path:
        entry_json_file.delete(0, tk.END)
        entry_json_file.insert(0, file_path)

def convert_file():
    excel_file_path = entry_excel_file.get()
    json_file_path = entry_json_file.get()

    if not excel_file_path or not json_file_path:
        messagebox.showwarning("Warning", "Please select both files.")
        return

    try:
        convert_excel_to_json(excel_file_path, json_file_path)
        messagebox.showinfo("Success", "Conversion complete!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Set up the main window
root = tk.Tk()
root.title("Excel to JSON Converter")

# Create and place widgets
tk.Label(root, text="Excel File:").grid(row=0, column=0, padx=10, pady=10)
entry_excel_file = tk.Entry(root, width=50)
entry_excel_file.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=browse_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Save As JSON:").grid(row=1, column=0, padx=10, pady=10)
entry_json_file = tk.Entry(root, width=50)
entry_json_file.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Save As", command=save_file).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Convert", command=convert_file).grid(row=2, column=1, padx=10, pady=20)

root.mainloop()
