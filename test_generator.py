"""
test_generator.py

Demo helper for file_manager.py.
Reads config.json, lets you pick one of your configured source folders,
and creates a small batch of empty .txt files inside it so you can
immediately click GetFiles and watch the copy happen.
"""

import json
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

# Must live in the same folder as file_manager.py and config.json
CONFIG_FILE = Path(__file__).parent / "config.json"


def load_config() -> dict:
    """Read config.json; raise a friendly error if it's missing or broken."""
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(f"config.json not found at {CONFIG_FILE}")
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)


def create_test_files(source_dir: Path, folder_name: str, count: int) -> list[Path]:
    """Create <count> empty .txt files inside source_dir / folder_name.

    File names include a timestamp so repeated runs produce unique files.
    Returns the list of paths that were created.
    """
    target = source_dir / folder_name
    target.mkdir(parents=True, exist_ok=True)  # create the folder if it doesn't exist yet

    stamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    created = []

    for i in range(1, count + 1):
        file_path = target / f"test_{stamp}_{i:02d}.txt"
        file_path.write_text(
            f"Test file {i} of {count}\n"
            f"Created by test_generator.py at {datetime.now()}\n"
            f"Source folder : {folder_name}\n"
        )
        created.append(file_path)

    return created


class TestGeneratorApp(tk.Tk):
    """Small demo window for generating test files."""

    def __init__(self):
        super().__init__()
        self.title("Test File Generator")
        self.resizable(False, False)

        try:
            self.config_data = load_config()
        except Exception as e:
            messagebox.showerror("Config Error", str(e))
            self.destroy()
            return

        self._build_ui()

    def _build_ui(self):
        """Lay out all widgets."""
        pad = {"padx": 12, "pady": 6}

        # ── Info label ────────────────────────────────────────────────────────
        tk.Label(
            self,
            text="Create dummy files in a source folder to demo GetFiles.",
            font=("Segoe UI", 9), fg="#444",
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, padx=12, pady=(10, 2))

        source_dir = self.config_data.get("source_dir", "(not set)")
        tk.Label(
            self,
            text=f"Source directory: {source_dir}",
            font=("Segoe UI", 8), fg="#888",
        ).grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=12, pady=(0, 8))

        ttk.Separator(self, orient=tk.HORIZONTAL).grid(
            row=2, column=0, columnspan=3, sticky=tk.EW, padx=12, pady=(0, 8)
        )

        # ── Folder picker ─────────────────────────────────────────────────────
        ttk.Label(self, text="Drop files into folder:").grid(
            row=3, column=0, sticky=tk.W, **pad
        )

        folder_names = list(self.config_data.get("folder_map", {}).keys())
        if not folder_names:
            messagebox.showwarning(
                "No Mappings",
                "No folder mappings found in config.json.\n"
                "Add at least one mapping in the File Manager Settings tab first.",
            )
            self.destroy()
            return

        self.folder_var = tk.StringVar(value=folder_names[0])
        ttk.Combobox(
            self,
            textvariable=self.folder_var,
            values=folder_names,
            state="readonly",   # only allow values from the list
            width=20,
        ).grid(row=3, column=1, sticky=tk.W, padx=(0, 12), pady=6)

        # ── File count slider ─────────────────────────────────────────────────
        ttk.Label(self, text="Number of files:").grid(
            row=4, column=0, sticky=tk.W, **pad
        )

        self.count_var = tk.IntVar(value=3)
        count_frame = tk.Frame(self)
        count_frame.grid(row=4, column=1, sticky=tk.W, padx=(0, 12), pady=6)

        # Slider (1–10 files)
        ttk.Scale(
            count_frame,
            from_=1, to=10,
            orient=tk.HORIZONTAL,
            variable=self.count_var,
            length=140,
            command=lambda v: self.count_var.set(int(float(v))),  # snap to integer
        ).pack(side=tk.LEFT)

        # Live label showing the current count
        tk.Label(count_frame, textvariable=self.count_var, width=3,
                 font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(6, 0))

        # ── Create button ─────────────────────────────────────────────────────
        ttk.Separator(self, orient=tk.HORIZONTAL).grid(
            row=5, column=0, columnspan=3, sticky=tk.EW, padx=12, pady=(8, 4)
        )

        tk.Button(
            self,
            text="Create Test Files",
            bg="#2d7dd2", fg="white",
            font=("Segoe UI", 10, "bold"),
            width=18,
            command=self._on_create,
        ).grid(row=6, column=0, columnspan=3, pady=(4, 4))

        # ── Output area ───────────────────────────────────────────────────────
        self.output = tk.Text(
            self, height=8, width=56,
            state=tk.DISABLED,
            font=("Consolas", 9),
            bg="#1e1e1e", fg="#d4d4d4",
        )
        self.output.grid(row=7, column=0, columnspan=3, padx=12, pady=(4, 12))

    def _on_create(self):
        """Create the test files and report results in the output area."""
        source_dir  = Path(self.config_data["source_dir"])
        folder_name = self.folder_var.get()
        count       = self.count_var.get()

        try:
            created = create_test_files(source_dir, folder_name, count)
        except Exception as e:
            messagebox.showerror("Error", f"Could not create files:\n{e}")
            return

        # Show results in the output box
        self.output.config(state=tk.NORMAL)
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, f"Created {len(created)} file(s) in:\n")
        self.output.insert(tk.END, f"  {source_dir / folder_name}\n\n")
        for path in created:
            self.output.insert(tk.END, f"  ✓  {path.name}\n")
        self.output.insert(tk.END, "\nNow open File Manager and click GetFiles!")
        self.output.config(state=tk.DISABLED)


if __name__ == "__main__":
    app = TestGeneratorApp()
    app.mainloop()
