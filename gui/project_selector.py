import os
import tkinter as tk
from tkinter import ttk, messagebox
from config import config_manager

def load_settings(parent):
    top = tk.Toplevel(parent)
    top.title("Load Settings")
    
    saved_folder = os.path.join(os.getcwd(), "saved_settings")
    files = [f for f in os.listdir(saved_folder) if f.endswith(".ini")]
    
    tk.Label(top, text="Select a project:").pack(pady=5)
    combo = ttk.Combobox(top, values=files, state="readonly")
    combo.pack(pady=5)
    if files:
        combo.current(0)
    
    selected = {"config": None}

    def load_action():
        selection = combo.get()
        if selection:
            file_path = os.path.join(saved_folder, selection)
            config = config_manager.load_config(file_path)
            selected["config"] = config
            top.destroy()

    def delete_action():
        selection = combo.get()
        if selection:
            confirm = messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{selection}'?")
            if confirm:
                file_path = os.path.join(saved_folder, selection)
                try:
                    os.remove(file_path)
                    messagebox.showinfo("Deleted", f"{selection} deleted.")
                    # Refresh the dropdown list.
                    new_files = [f for f in os.listdir(saved_folder) if f.endswith(".ini")]
                    combo['values'] = new_files
                    if new_files:
                        combo.current(0)
                    else:
                        combo.set("")
                except Exception as e:
                    messagebox.showerror("Error", str(e))
    
    btn_frame = tk.Frame(top)
    btn_frame.pack(pady=10)
    btn_load = tk.Button(btn_frame, text="Load", command=load_action)
    btn_load.grid(row=0, column=0, padx=5)
    btn_delete = tk.Button(btn_frame, text="Delete", command=delete_action)
    btn_delete.grid(row=0, column=1, padx=5)
    btn_cancel = tk.Button(btn_frame, text="Cancel", command=top.destroy)
    btn_cancel.grid(row=0, column=2, padx=5)

    top.grab_set()
    top.wait_window()
    return selected["config"]