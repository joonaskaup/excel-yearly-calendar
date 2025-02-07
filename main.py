import os
import tkinter as tk
from tkinter import filedialog, messagebox
from config import config_manager
from gui import project_selector, settings_editor
from calgenerator import generator

def main():
    root = tk.Tk()
    root.title("Excel Yearly Calendar")

    # Load default configuration from resources/default_settings.ini
    default_config_path = os.path.join("resources", "default_settings.ini")
    if os.path.exists(default_config_path):
        config = config_manager.load_config(default_config_path)
    else:
        config = config_manager.default_config()

    def select_project():
        nonlocal config
        project_path = project_selector.select_project()
        if project_path:
            config = config_manager.load_config(project_path)
            messagebox.showinfo("Project Loaded", f"Loaded project from {project_path}")

    def edit_settings():
        nonlocal config
        new_config = settings_editor.edit_settings(root, config)
        if new_config:
            config = new_config
            messagebox.showinfo("Settings Saved", "Settings have been updated.")

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

    # Main window buttons
    btn_select_project = tk.Button(root, text="Select Project", command=select_project)
    btn_select_project.pack(pady=5)

    btn_edit_settings = tk.Button(root, text="Edit Settings", command=edit_settings)
    btn_edit_settings.pack(pady=5)

    btn_generate = tk.Button(root, text="Generate Calendar", command=generate_calendar_action)
    btn_generate.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()