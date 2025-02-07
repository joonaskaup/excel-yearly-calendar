import tkinter as tk
from tkinter import messagebox
from config import config_manager
import os

def edit_settings(parent, settings):
    editor = tk.Toplevel(parent)
    editor.title("Edit Settings")

    # Dictionary to hold entry widgets by section and key.
    entries = {}
    row = 0

    for section in settings:
        section_label = tk.Label(editor, text=f"[{section}]", font=("Arial", 10, "bold"))
        section_label.grid(row=row, column=0, columnspan=2, pady=(10, 0), sticky="w")
        row += 1
        entries[section] = {}
        for key, value in settings[section].items():
            lbl = tk.Label(editor, text=key)
            lbl.grid(row=row, column=0, padx=5, pady=2, sticky="e")
            ent = tk.Entry(editor)
            ent.insert(0, value)
            ent.grid(row=row, column=1, padx=5, pady=2, sticky="w")
            entries[section][key] = ent
            row += 1

    def save_and_close():
        new_settings = {}
        for section, opts in entries.items():
            new_settings[section] = {}
            for key, entry in opts.items():
                new_settings[section][key] = entry.get()
        # Save the settings to the saved_settings folder using the TITLE as project name.
        saved_folder = os.path.join(os.getcwd(), "saved_settings")
        if not os.path.exists(saved_folder):
            os.makedirs(saved_folder)
        file_path = config_manager.save_project_config(new_settings, saved_folder)
        messagebox.showinfo("Saved", f"Settings saved to {file_path}")
        editor.new_settings = new_settings
        editor.destroy()

    def revert_to_default():
        default = config_manager.default_config()
        for section, opts in entries.items():
            if section in default:
                for key, entry in opts.items():
                    if key in default[section]:
                        entry.delete(0, tk.END)
                        entry.insert(0, default[section][key])

    btn_frame = tk.Frame(editor)
    btn_frame.grid(row=row, column=0, columnspan=2, pady=10)
    btn_save = tk.Button(btn_frame, text="Save", command=save_and_close)
    btn_save.grid(row=0, column=0, padx=5)
    btn_default = tk.Button(btn_frame, text="Default", command=revert_to_default)
    btn_default.grid(row=0, column=1, padx=5)
    btn_cancel = tk.Button(btn_frame, text="Cancel", command=editor.destroy)
    btn_cancel.grid(row=0, column=2, padx=5)

    editor.wait_window()
    return getattr(editor, "new_settings", None)