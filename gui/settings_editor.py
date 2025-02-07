import tkinter as tk

def edit_settings(parent, settings):
    editor = tk.Toplevel(parent)
    editor.title("Edit Settings")

    # Dictionary to hold entry widgets by section and key
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
        editor.new_settings = new_settings
        editor.destroy()

    btn_save = tk.Button(editor, text="Save", command=save_and_close)
    btn_save.grid(row=row, column=0, columnspan=2, pady=10)

    editor.wait_window()
    return getattr(editor, "new_settings", None)