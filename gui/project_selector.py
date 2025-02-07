import os
import tkinter as tk
from tkinter import filedialog, messagebox

def select_project():
    file_path = filedialog.askopenfilename(
        title="Select Project Config File",
        filetypes=[("INI files", "*.ini")]
    )
    return file_path

def delete_project():
    file_path = filedialog.askopenfilename(
        title="Select Project Config File to Delete",
        filetypes=[("INI files", "*.ini")]
    )
    if file_path:
        confirm = tk.messagebox.askyesno("Confirm Deletion", f"Delete {os.path.basename(file_path)}?")
        if confirm:
            try:
                os.remove(file_path)
                tk.messagebox.showinfo("Deleted", "Project deleted successfully.")
            except Exception as e:
                tk.messagebox.showerror("Error", str(e))