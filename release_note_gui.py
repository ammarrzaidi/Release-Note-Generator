"""
release_note_gui.py

Desktop GUI (Tkinter, ships with Python - no extra install) for turning a
excel file ticket export into a formatted .docx release note using pre-existing
templates and placeholders.

Run with:
    python release_note_gui.py

Requires: python-docx, pandas, openpyxl
    pip install python-docx pandas openpyxl
"""

import json
import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog

from release_note_core import generate_release_note


class ReleaseNoteApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Release Note Generator")
        self.geometry("640x630")
        self.resizable(False, False)

        self.config_path = Path.home() / ".release_note_config.json"
        self.input_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.include_rejected = tk.BooleanVar(value=False)

        self.fields = {
            "project_id": tk.StringVar(),
            "project_name": tk.StringVar(),
            "client_name": tk.StringVar(),
            "version_no": tk.StringVar(),
            "release_date": tk.StringVar(),
            "prepared_by": tk.StringVar(),
            "reviewed_by": tk.StringVar(),
            "document_id": tk.StringVar(),
            "revision_no": tk.StringVar(value="1.0"),
            "list_of_changes": tk.StringVar(value="Document release"),
        }

        self._load_config()
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 4}

        # --- file pickers ---
        file_frame = ttk.Frame(self)
        file_frame.pack(fill="x", **pad)

        ttk.Label(file_frame, text="Ticket export (CSV/XLSX):").pack(anchor="w")
        row1 = ttk.Frame(file_frame)
        row1.pack(fill="x", pady=(2, 6))
        ttk.Entry(row1, textvariable=self.input_path).pack(side="left", fill="x", expand=True)
        ttk.Button(row1, text="Browse...", command=self._browse_input).pack(side="left", padx=(6, 0))

        ttk.Label(file_frame, text="Word Template (.docx) [Optional]:").pack(anchor="w")
        row2 = ttk.Frame(file_frame)
        row2.pack(fill="x", pady=(2, 0))
        ttk.Entry(row2, textvariable=self.template_path).pack(side="left", fill="x", expand=True)
        ttk.Button(row2, text="Browse...", command=self._browse_template).pack(side="left", padx=(6, 0))

        # --- header fields ---
        form = ttk.Frame(self)
        form.pack(fill="x", **pad)
        labels = [
            ("Project ID:", "project_id"),
            ("Project Name:", "project_name"),
            ("Client Name:", "client_name"),
            ("Version No", "version_no"),
            ("Release Date", "release_date"),
            ("Prepared By ", "prepared_by"),
            ("Reviewed / Verified By", "reviewed_by"),
            ("Document ID", "document_id"),
            ("Revision No:", "revision_no"),
            ("List of Changes:", "list_of_changes"),
        ]
        for i, (label, key) in enumerate(labels):
            ttk.Label(form, text=label).grid(row=i, column=0, sticky="w", pady=3)
            ttk.Entry(form, textvariable=self.fields[key], width=42).grid(
                row=i, column=1, sticky="ew", pady=3, padx=(8, 0)
            )
        form.columnconfigure(1, weight=1)

        # --- checkbox ---
        ttk.Checkbutton(
            self, text="Include QC Rejected tickets", variable=self.include_rejected
        ).pack(anchor="w", **pad)

        # --- generate button ---
        self.generate_btn = ttk.Button(self, text="Generate Release Note", command=self._on_generate)
        self.generate_btn.pack(fill="x", **pad)

        # --- status ---
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(self, textvariable=self.status_var, wraplength=580)
        self.status_label.pack(fill="x", **pad)

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select the excel file export",
            filetypes=[("Excel files", "*.csv *.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if path:
            self.input_path.set(path)

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="Select Word Template (.docx)",
            filetypes=[("Word Documents", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.template_path.set(path)

    def _set_status(self, text, error=False):
        self.status_var.set(text)
        self.status_label.configure(foreground="red" if error else "green")

    def _load_config(self):
        if not self.config_path.exists():
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            self.input_path.set(config.get("input_path", ""))
            self.template_path.set(config.get("template_path", ""))
            self.include_rejected.set(config.get("include_rejected", False))
            for key, var in self.fields.items():
                var.set(config.get(key, var.get()))
        except Exception:
            pass

    def _save_config(self):
        try:
            config = {
                "input_path": self.input_path.get(),
                "template_path": self.template_path.get(),
                "include_rejected": self.include_rejected.get(),
            }
            config.update({key: var.get() for key, var in self.fields.items()})
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2)
        except Exception:
            pass

    def _on_generate(self):
        input_path = self.input_path.get().strip()
        if not input_path or not os.path.exists(input_path):
            self._set_status("Please select a valid ticket export file first.", error=True)
            return

        default_name = f"Release_Note_{self.fields['version_no'].get().replace('.', '_') or 'output'}.docx"
        output_path = filedialog.asksaveasfilename(
            title="Save release note as",
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Word Document", "*.docx")],
        )
        if not output_path:
            return

        self.generate_btn.configure(state="disabled")
        self._set_status("Generating...")

        header = {k: v.get() for k, v in self.fields.items()}
        include_rejected = self.include_rejected.get()
        template_path = self.template_path.get().strip() or None

        def worker():
            try:
                self._save_config()
                result = generate_release_note(
                    input_path, output_path, header, include_rejected, template_path
                )
                self.after(0, lambda: self._on_success(output_path, result))
            except Exception as e:
                self.after(0, lambda err=str(e): self._on_error(err))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, output_path, result):
        self.generate_btn.configure(state="normal")
        self._set_status(
            f"Done. Saved to: {output_path}\n"
            f"({result['enhancements_count']} enhancements, {result['bugs_count']} bugs included)"
        )

    def _on_error(self, message):
        self.generate_btn.configure(state="normal")
        self._set_status(f"Failed: {message}", error=True)


if __name__ == "__main__":
    app = ReleaseNoteApp()
    app.mainloop()