import os
import tkinter as tk
from tkinter import filedialog, messagebox
from config import config_manager
from gui import project_selector, settings_editor
from calgenerator import generator

def main():
    root = tk.Tk()
    root.title("Excel Yearly Calendar")

    # Ensure the saved_settings folder exists
    saved_settings_folder = os.path.join(os.getcwd(), "saved_settings")
    if not os.path.exists(saved_settings_folder):
        os.makedirs(saved_settings_folder)

    # Load default config from resources/default_settings.ini if available;
    # otherwise use the default config.
    default_config_path = os.path.join("resources", "default_settings.ini")
    if os.path.exists(default_config_path):
        config = config_manager.load_config(default_config_path)
    else:
        config = config_manager.default_config()

    def load_settings_action():
        nonlocal config
        new_config = project_selector.load_settings(root)
        if new_config:
            config = new_config
            messagebox.showinfo("Settings Loaded", "Settings have been loaded.")

    def edit_settings_action():
        nonlocal config
        new_config = settings_editor.edit_settings(root, config)
        if new_config:
            config = new_config
            messagebox.showinfo("Settings Saved", "Settings have been saved.")

    def generate_calendar_action():
        input_file = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not input_file:
            return
        output_file = filedialog.asksaveasfilename(
            title="Select Output Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not output_file:
            return
        try:
            generator.generate_calendar(input_file, output_file, config)
            messagebox.showinfo("Success", f"Calendar generated and saved to:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    btn_load = tk.Button(root, text="Load Settings", command=load_settings_action)
    btn_load.pack(pady=5)

    btn_edit = tk.Button(root, text="Edit Settings", command=edit_settings_action)
    btn_edit.pack(pady=5)

    btn_generate = tk.Button(root, text="Generate Calendar", command=generate_calendar_action)
    btn_generate.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()