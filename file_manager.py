import ctypes
import ctypes.wintypes
import json
import logging
import shutil
import threading
import tkinter as tk
from datetime import datetime, timezone
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

# pyodbc is only needed for the SQL Import tab; the rest of the app works without it
try:
    import pyodbc
    PYODBC_AVAILABLE = True
except ImportError:
    PYODBC_AVAILABLE = False

# Paths resolved relative to this script so the program works from any working directory
BASE_DIR          = Path(__file__).parent
CONFIG_FILE       = BASE_DIR / "config.json"
STATE_FILE        = BASE_DIR / "last_run.json"
LOG_FILE          = BASE_DIR / "file_manager.log"
RUN_HISTORY_FILE  = BASE_DIR / "run_history.json"   # one record per GetFiles run
FILE_HISTORY_FILE = BASE_DIR / "file_history.json"  # one record per file copied

# File logger — appends one entry per action
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


# ── File copy with full metadata preservation ────────────────────────────────

def _set_windows_creation_time(path: Path, ctime: float) -> None:
    """Set the Windows filesystem creation time on a file via the Win32 API.

    shutil.copy2 preserves mtime and atime but resets the Windows creation
    time (ctime) to now. This restores the original value using ctypes so no
    third-party packages are required.
    """
    # Windows FILETIME counts 100-nanosecond intervals since 1601-01-01.
    # Unix timestamps count seconds since 1970-01-01.
    # The offset between the two epochs is 116,444,736,000,000,000 intervals.
    EPOCH_DIFF  = 116_444_736_000_000_000
    win_time    = int(ctime * 10_000_000) + EPOCH_DIFF
    filetime    = ctypes.wintypes.FILETIME(win_time & 0xFFFFFFFF, win_time >> 32)

    # Open a file handle with write-attributes permission only
    handle = ctypes.windll.kernel32.CreateFileW(
        str(path),
        0x0100,     # FILE_WRITE_ATTRIBUTES
        0, None,
        3,          # OPEN_EXISTING
        0x80,       # FILE_ATTRIBUTE_NORMAL
        None,
    )

    if handle == ctypes.wintypes.HANDLE(-1).value:
        return  # silently skip if the handle couldn't be opened (e.g. permissions)

    try:
        # SetFileTime(handle, lpCreationTime, lpLastAccessTime, lpLastWriteTime)
        # Passing None for access/write leaves those timestamps unchanged.
        ctypes.windll.kernel32.SetFileTime(handle, ctypes.byref(filetime), None, None)
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def copy_with_metadata(src: Path, dst: Path) -> None:
    """Copy src to dst preserving mtime, atime, and Windows creation time.

    shutil.copy2 handles mtime and atime; _set_windows_creation_time
    restores the original creation time that copy2 resets.
    """
    src_stat = src.stat()
    shutil.copy2(src, dst)                                  # copies content + mtime/atime
    _set_windows_creation_time(dst, src_stat.st_birthtime)  # restore original creation time


# ── Demo / test file generation ──────────────────────────────────────────────

def create_test_files(source_dir: Path, folder_name: str, count: int) -> list:
    """Create <count> small text files inside source_dir / folder_name.

    File names include a timestamp so repeated calls always produce unique files.
    Returns the list of Path objects that were created.
    """
    target = source_dir / folder_name
    target.mkdir(parents=True, exist_ok=True)  # create the folder if it doesn't exist yet

    stamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    created = []

    for i in range(1, count + 1):
        file_path = target / f"test_{stamp}_{i:02d}.txt"
        file_path.write_text(
            f"Test file {i} of {count}\n"
            f"Created by File Manager demo at {datetime.now()}\n"
            f"Source folder : {folder_name}\n"
        )
        created.append(file_path)

    return created


# ── Config / state helpers ────────────────────────────────────────────────────

def load_config() -> dict:
    """Read config.json and return its contents as a dict."""
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)


def load_last_run() -> datetime:
    """Return the timestamp of the last successful GetFiles run.

    Returns epoch 0 on first run so every existing file is treated as new.
    """
    if STATE_FILE.exists():
        with open(STATE_FILE, "r") as f:
            raw = json.load(f).get("last_run", "")
        if raw:
            return datetime.fromisoformat(raw)
    return datetime.fromtimestamp(0, tz=timezone.utc).replace(tzinfo=None)


def save_last_run(ts: datetime) -> None:
    """Persist the run timestamp to last_run.json."""
    with open(STATE_FILE, "w") as f:
        json.dump({"last_run": ts.isoformat()}, f, indent=2)


def load_run_history() -> list:
    """Return the list of past run summary records."""
    if RUN_HISTORY_FILE.exists():
        with open(RUN_HISTORY_FILE, "r") as f:
            return json.load(f)
    return []


def append_run_record(record: dict) -> None:
    """Append a single run summary to run_history.json."""
    history = load_run_history()
    history.append(record)
    with open(RUN_HISTORY_FILE, "w") as f:
        json.dump(history, f, indent=2)


def load_file_history() -> list:
    """Return the list of all file-copy records."""
    if FILE_HISTORY_FILE.exists():
        with open(FILE_HISTORY_FILE, "r") as f:
            return json.load(f)
    return []


def append_file_records(records: list) -> None:
    """Append a batch of file-copy records to file_history.json."""
    history = load_file_history()
    history.extend(records)
    with open(FILE_HISTORY_FILE, "w") as f:
        json.dump(history, f, indent=2)


# ── Core logic ────────────────────────────────────────────────────────────────

def get_files(log_fn) -> int:
    """Scan source_dir for new files, copy each to its mapped backup folder.

    Args:
        log_fn: callable(str) — emits progress messages to the GUI and log file.

    Returns:
        Number of files successfully copied.
    """
    try:
        config = load_config()
    except Exception as e:
        log_fn(f"ERROR: Could not read config.json — {e}")
        return 0

    source_dir  = Path(config["source_dir"])
    backup_dir  = Path(config["backup_dir"])
    folder_map  = config.get("folder_map", {})
    last_run    = load_last_run()
    run_time    = datetime.now()

    log_fn(f"── GetFiles started at {run_time.strftime('%Y-%m-%d %H:%M:%S')} ──")
    log_fn(f"Source : {source_dir}")
    log_fn(f"Backup : {backup_dir}")
    log_fn(f"Checking for files newer than: {last_run.strftime('%Y-%m-%d %H:%M:%S')}")
    log_fn("")

    if not source_dir.exists():
        log_fn(f"ERROR: Source directory not found: {source_dir}")
        return 0

    copied       = 0
    skipped      = 0
    file_records = []  # accumulate per-file history entries for this run

    # Walk every immediate subdirectory of source_dir
    for folder in sorted(source_dir.iterdir()):
        if not folder.is_dir():
            continue  # ignore loose files at the top level

        for file in sorted(folder.rglob("*")):
            if not file.is_file():
                continue

            # Use whichever timestamp is newer: mtime (last modified) or
            # st_birthtime (when the file appeared in this location on Windows).
            # This catches files copied from elsewhere that carry an old mtime.
            stat  = file.stat()
            mtime = datetime.fromtimestamp(max(stat.st_mtime, stat.st_birthtime))
            if mtime <= last_run:
                continue  # file predates last run; skip

            log_fn(f"  Found : {file.relative_to(source_dir)}  ({mtime.strftime('%Y-%m-%d %H:%M:%S')})")
            logging.info("Found: %s", file)

            # Look up the top-level folder name in the mapping
            top_folder = folder.name
            if top_folder not in folder_map:
                log_fn(f"  SKIP  : '{top_folder}' has no mapping in config.json")
                logging.warning("Skipped (no mapping): %s", file)
                skipped += 1
                continue

            # Build destination path and copy
            dest_folder = backup_dir / folder_map[top_folder]
            dest_file   = dest_folder / file.name

            try:
                dest_folder.mkdir(parents=True, exist_ok=True)
                copy_with_metadata(file, dest_file)  # preserves mtime, atime, and creation time
                log_fn(f"  Copied → {dest_file}")
                logging.info("Copied: %s → %s", file, dest_file)
                copied += 1

                # Record this file for history
                file_records.append({
                    "run_time":    run_time.isoformat(),
                    "file_name":   file.name,
                    "source_path": str(file),
                    "dest_path":   str(dest_file),
                    "folder":      top_folder,
                })

            except Exception as e:
                log_fn(f"  ERROR : {file.name} — {e}")
                logging.error("Failed to copy %s: %s", file, e)

    # Summary line
    log_fn("")
    log_fn(f"Done — {copied} file(s) copied, {skipped} skipped (unmapped folders).")
    log_fn("")
    logging.info("Run complete: %d copied, %d skipped.", copied, skipped)

    # Persist history and update last-run timestamp
    append_run_record({
        "run_time":    run_time.isoformat(),
        "files_found": copied + skipped,
        "copied":      copied,
        "skipped":     skipped,
    })
    if file_records:
        append_file_records(file_records)

    save_last_run(run_time)
    return copied


def _sniff_csv(file: Path) -> tuple[str, int]:
    """Return (delimiter, skip_header) by inspecting the first two lines of a file.

    Delimiter is chosen as whichever of |  ,  \\t  ;  appears most in line 1.
    skip_header is 1 if every field in line 1 looks like a column name (no digits),
    0 otherwise.
    """
    try:
        with open(file, encoding="utf-8", errors="replace") as fh:
            first  = fh.readline().rstrip("\r\n")
    except Exception:
        return ",", 1   # safe default

    # Pick delimiter with the most occurrences in the first line
    candidates = ["|", ",", "\t", ";"]
    delimiter  = max(candidates, key=lambda d: first.count(d))

    # If the winning delimiter has 0 occurrences fall back to comma
    if first.count(delimiter) == 0:
        delimiter = ","

    # Detect header: every field is purely letters / underscores / spaces
    import re
    fields      = first.split(delimiter)
    looks_like_header = all(re.fullmatch(r"[A-Za-z_\s]+", f.strip()) for f in fields if f.strip())
    skip_header = 1 if looks_like_header else 0

    return delimiter, skip_header


# ── GUI ───────────────────────────────────────────────────────────────────────

class FileManagerApp(tk.Tk):
    """Main application window with tabbed layout."""

    def __init__(self):
        super().__init__()
        self.title("File Manager")
        self.resizable(True, True)
        self.minsize(700, 480)
        self._build_ui()
        self._update_status()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        """Construct and lay out all widgets."""

        # Shared status bar at the very bottom (outside the notebook)
        self.status_var = tk.StringVar()
        tk.Label(
            self, textvariable=self.status_var,
            anchor=tk.W, relief=tk.SUNKEN,
            font=("Segoe UI", 8), fg="#555",
        ).pack(fill=tk.X, side=tk.BOTTOM, ipady=2)

        # Tabbed notebook — fills the rest of the window
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        self._build_main_tab()
        self._build_settings_tab()
        self._build_demo_tab()
        self._build_run_history_tab()
        self._build_file_history_tab()
        self._build_log_tab()
        self._build_sql_import_tab()

        # Refresh tabs whenever they are switched to
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _build_main_tab(self):
        """Tab 1 — GetFiles button and live log output."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  GetFiles  ")

        # Button bar
        btn_frame = tk.Frame(frame, pady=6)
        btn_frame.pack(fill=tk.X, padx=8)

        self.btn_get = tk.Button(
            btn_frame, text="GetFiles",
            width=14, bg="#2d7dd2", fg="white",
            font=("Segoe UI", 10, "bold"),
            command=self._on_get_files,
        )
        self.btn_get.pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            btn_frame, text="Clear Log",
            width=10,
            command=self._clear_log,
        ).pack(side=tk.LEFT)

        # Live log output area
        self.log_area = scrolledtext.ScrolledText(
            frame, wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Consolas", 9),
            bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white",
        )
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 6))

    def _build_settings_tab(self):
        """Tab 2 — edit backup directory and folder mappings without touching config.json."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Settings  ")

        pad = {"padx": 10, "pady": 4}

        # ── Backup root directory ─────────────────────────────────────────────
        ttk.Label(frame, text="Backup Root Directory:", font=("Segoe UI", 9, "bold")
                  ).grid(row=0, column=0, sticky=tk.W, **pad)

        self.backup_dir_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.backup_dir_var, width=48
                  ).grid(row=0, column=1, sticky=tk.EW, padx=(0, 4), pady=4)

        ttk.Button(frame, text="Browse…", command=self._browse_backup_dir
                   ).grid(row=0, column=2, padx=(0, 10), pady=4)

        # ── Folder mappings table ─────────────────────────────────────────────
        ttk.Label(frame, text="Folder Mappings:", font=("Segoe UI", 9, "bold")
                  ).grid(row=1, column=0, sticky=tk.W, **pad)

        ttk.Label(
            frame,
            text='Each row: when a file arrives in "Arriving In", copy it to the matching "Backup Subfolder".',
            foreground="#666", font=("Segoe UI", 8),
        ).grid(row=2, column=0, columnspan=3, sticky=tk.W, padx=10, pady=(0, 4))

        # Treeview showing current mappings
        tree_frame = ttk.Frame(frame)
        tree_frame.grid(row=3, column=0, columnspan=3, sticky=tk.NSEW, padx=10, pady=(0, 4))

        cols = ("source", "dest")
        self.map_tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                     selectmode="browse", height=7)
        self.map_tree.heading("source", text="Arriving In (source folder name)")
        self.map_tree.heading("dest",   text="Backup Subfolder")
        self.map_tree.column("source",  width=240, anchor=tk.W)
        self.map_tree.column("dest",    width=240, anchor=tk.W)

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.map_tree.yview)
        self.map_tree.configure(yscrollcommand=vsb.set)
        self.map_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.LEFT, fill=tk.Y)

        # ── Add new mapping ───────────────────────────────────────────────────
        add_frame = ttk.LabelFrame(frame, text=" Add New Mapping ", padding=6)
        add_frame.grid(row=4, column=0, columnspan=3, sticky=tk.EW, padx=10, pady=(4, 2))

        ttk.Label(add_frame, text="Arriving In:").pack(side=tk.LEFT)
        self.new_source_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self.new_source_var, width=18
                  ).pack(side=tk.LEFT, padx=(4, 8))

        ttk.Label(add_frame, text="→   Backup Subfolder:").pack(side=tk.LEFT)
        self.new_dest_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self.new_dest_var, width=18
                  ).pack(side=tk.LEFT, padx=(4, 8))

        ttk.Button(add_frame, text="Add", command=self._add_mapping
                   ).pack(side=tk.LEFT)

        # ── Bottom action buttons ─────────────────────────────────────────────
        action_frame = tk.Frame(frame, pady=6)
        action_frame.grid(row=5, column=0, columnspan=3, sticky=tk.EW, padx=10)

        ttk.Button(action_frame, text="Remove Selected",
                   command=self._remove_mapping).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Button(action_frame, text="Create Backup Folders",
                   command=self._create_backup_folders).pack(side=tk.LEFT, padx=(0, 8))

        tk.Button(
            action_frame, text="Save Settings",
            bg="#2d7dd2", fg="white", font=("Segoe UI", 9, "bold"),
            command=self._save_settings,
        ).pack(side=tk.RIGHT)

        # Allow the treeview column to stretch when the window is resized
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(3, weight=1)

    def _build_demo_tab(self):
        """Demo tab — create test files in a source folder to try out GetFiles."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Demo  ")

        pad = {"padx": 12, "pady": 5}

        # ── Description ───────────────────────────────────────────────────────
        ttk.Label(
            frame,
            text="Create dummy files in a source folder to demo or test GetFiles.",
            font=("Segoe UI", 9),
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, padx=12, pady=(10, 0))

        # Dynamic label showing the current source directory from config
        self.demo_source_label = ttk.Label(frame, text="", foreground="#888",
                                           font=("Segoe UI", 8))
        self.demo_source_label.grid(row=1, column=0, columnspan=3, sticky=tk.W,
                                    padx=12, pady=(0, 8))

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(
            row=2, column=0, columnspan=3, sticky=tk.EW, padx=12, pady=(0, 8)
        )

        # ── Folder picker ─────────────────────────────────────────────────────
        ttk.Label(frame, text="Drop files into folder:").grid(
            row=3, column=0, sticky=tk.W, **pad
        )

        self.demo_folder_var = tk.StringVar()
        self.demo_folder_combo = ttk.Combobox(
            frame,
            textvariable=self.demo_folder_var,
            state="readonly",   # only allow values from the configured list
            width=22,
        )
        self.demo_folder_combo.grid(row=3, column=1, sticky=tk.W, padx=(0, 12), pady=5)

        # ── File count slider ─────────────────────────────────────────────────
        ttk.Label(frame, text="Number of files:").grid(
            row=4, column=0, sticky=tk.W, **pad
        )

        count_frame = tk.Frame(frame)
        count_frame.grid(row=4, column=1, sticky=tk.W, padx=(0, 12), pady=5)

        self.demo_count_var = tk.IntVar(value=3)
        ttk.Scale(
            count_frame,
            from_=1, to=10,
            orient=tk.HORIZONTAL,
            variable=self.demo_count_var,
            length=140,
            # Snap the float value from the slider to a whole number
            command=lambda v: self.demo_count_var.set(int(float(v))),
        ).pack(side=tk.LEFT)

        tk.Label(count_frame, textvariable=self.demo_count_var, width=3,
                 font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(6, 0))

        # ── Create button ─────────────────────────────────────────────────────
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(
            row=5, column=0, columnspan=3, sticky=tk.EW, padx=12, pady=(8, 4)
        )

        tk.Button(
            frame,
            text="Create Test Files",
            bg="#2d7dd2", fg="white",
            font=("Segoe UI", 10, "bold"),
            width=18,
            command=self._on_demo_create,
        ).grid(row=6, column=0, columnspan=3, pady=(4, 4))

        # ── Output area ───────────────────────────────────────────────────────
        self.demo_output = tk.Text(
            frame, height=9, width=58,
            state=tk.DISABLED,
            font=("Consolas", 9),
            bg="#1e1e1e", fg="#d4d4d4",
        )
        self.demo_output.grid(row=7, column=0, columnspan=3, padx=12, pady=(4, 12))

    def _build_run_history_tab(self):
        """Tab 3 — one row per GetFiles run."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Run History  ")

        cols = ("run_time", "found", "copied", "skipped")

        self.run_tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")

        # Column headings and widths
        self.run_tree.heading("run_time", text="Date / Time")
        self.run_tree.heading("found",    text="Found")
        self.run_tree.heading("copied",   text="Copied")
        self.run_tree.heading("skipped",  text="Skipped")

        self.run_tree.column("run_time", width=180, anchor=tk.W)
        self.run_tree.column("found",    width=70,  anchor=tk.CENTER)
        self.run_tree.column("copied",   width=70,  anchor=tk.CENTER)
        self.run_tree.column("skipped",  width=70,  anchor=tk.CENTER)

        # Vertical scrollbar
        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.run_tree.yview)
        self.run_tree.configure(yscrollcommand=vsb.set)

        self.run_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=8)
        vsb.pack(side=tk.LEFT, fill=tk.Y, pady=8, padx=(0, 8))

    def _build_file_history_tab(self):
        """Tab 3 — every file ever copied, with a live search box."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  File History  ")

        # Search bar
        search_frame = tk.Frame(frame, pady=6)
        search_frame.pack(fill=tk.X, padx=8)

        tk.Label(search_frame, text="Search:").pack(side=tk.LEFT, padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_changed)  # filter as you type
        tk.Entry(search_frame, textvariable=self.search_var, width=35).pack(side=tk.LEFT)

        tk.Label(search_frame, text="  (searches file name and folder)", fg="#888",
                 font=("Segoe UI", 8)).pack(side=tk.LEFT)

        # File history table
        cols = ("run_time", "file_name", "folder", "dest_path")
        self.file_tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")

        self.file_tree.heading("run_time",  text="Date / Time")
        self.file_tree.heading("file_name", text="File Name")
        self.file_tree.heading("folder",    text="Source Folder")
        self.file_tree.heading("dest_path", text="Destination")

        self.file_tree.column("run_time",  width=155, anchor=tk.W)
        self.file_tree.column("file_name", width=180, anchor=tk.W)
        self.file_tree.column("folder",    width=100, anchor=tk.W)
        self.file_tree.column("dest_path", width=280, anchor=tk.W)

        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=vsb.set)

        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=(0, 8))
        vsb.pack(side=tk.LEFT, fill=tk.Y, pady=(0, 8), padx=(0, 8))

        # Cache the full list so search doesn't re-read the file on every keystroke
        self._all_file_records = []

    def _build_log_tab(self):
        """Tab 4 — raw contents of file_manager.log."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Log File  ")

        # Refresh button
        btn_frame = tk.Frame(frame, pady=6)
        btn_frame.pack(fill=tk.X, padx=8)
        tk.Button(btn_frame, text="Refresh", width=10,
                  command=self._refresh_log_tab).pack(side=tk.LEFT)

        self.log_file_area = scrolledtext.ScrolledText(
            frame, wrap=tk.NONE,      # no-wrap so long lines stay readable
            state=tk.DISABLED,
            font=("Consolas", 9),
            bg="#1e1e1e", fg="#d4d4d4",
        )
        # Horizontal scrollbar for wide log lines
        hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL,
                             command=self.log_file_area.xview)
        self.log_file_area.configure(xscrollcommand=hsb.set)

        hsb.pack(side=tk.BOTTOM, fill=tk.X, padx=8)
        self.log_file_area.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))

    # ── Tab switching ─────────────────────────────────────────────────────────

    def _on_tab_changed(self, _event):
        """Reload the active tab whenever the user switches to it."""
        tab_name = self.notebook.tab(self.notebook.select(), "text").strip()
        if tab_name == "Settings":
            self._load_settings_form()
        elif tab_name == "Demo":
            self._refresh_demo_tab()
        elif tab_name == "Run History":
            self._refresh_run_history()
        elif tab_name == "File History":
            self._refresh_file_history()
        elif tab_name == "Log File":
            self._refresh_log_tab()
        elif tab_name == "SQL Import":
            self._refresh_sql_import_log()
            self._refresh_sql_import_folders()

    # ── Settings helpers ──────────────────────────────────────────────────────

    def _load_settings_form(self):
        """Populate the Settings tab from the current config.json."""
        try:
            config = load_config()
        except Exception as e:
            messagebox.showerror("Error", f"Could not read config.json:\n{e}")
            return

        # Fill the backup directory field
        self.backup_dir_var.set(config.get("backup_dir", ""))

        # Fill the mappings table
        self.map_tree.delete(*self.map_tree.get_children())
        for source, dest in config.get("folder_map", {}).items():
            self.map_tree.insert("", tk.END, values=(source, dest))

    def _browse_backup_dir(self):
        """Open a folder-picker dialog and put the chosen path in the backup dir field."""
        chosen = filedialog.askdirectory(title="Select Backup Root Directory")
        if chosen:
            self.backup_dir_var.set(chosen)

    def _add_mapping(self):
        """Validate and add a new row to the mappings table."""
        source = self.new_source_var.get().strip()
        dest   = self.new_dest_var.get().strip()

        if not source or not dest:
            messagebox.showwarning("Missing Input", "Both 'Arriving In' and 'Backup Subfolder' are required.")
            return

        # Check for duplicate source folder name
        existing = [self.map_tree.item(row)["values"][0] for row in self.map_tree.get_children()]
        if source in existing:
            messagebox.showwarning("Duplicate", f"'{source}' already has a mapping. Remove it first to replace it.")
            return

        self.map_tree.insert("", tk.END, values=(source, dest))
        # Clear the input fields after a successful add
        self.new_source_var.set("")
        self.new_dest_var.set("")

    def _remove_mapping(self):
        """Delete the currently selected row from the mappings table."""
        selected = self.map_tree.selection()
        if not selected:
            messagebox.showinfo("Nothing Selected", "Click a row in the table to select it, then click Remove.")
            return
        self.map_tree.delete(*selected)

    def _save_settings(self):
        """Write the current form values back to config.json."""
        backup_dir = self.backup_dir_var.get().strip()
        if not backup_dir:
            messagebox.showwarning("Missing Value", "Please enter a Backup Root Directory before saving.")
            return

        # Rebuild folder_map from the treeview rows
        folder_map = {}
        for row in self.map_tree.get_children():
            source, dest = self.map_tree.item(row)["values"]
            folder_map[source] = dest

        # Load existing config to preserve any keys we don't manage (e.g. source_dir)
        try:
            config = load_config()
        except Exception:
            config = {}

        config["backup_dir"]  = backup_dir
        config["folder_map"]  = folder_map

        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f, indent=4)

        messagebox.showinfo("Saved", "Settings saved to config.json.")

    def _create_backup_folders(self):
        """Create each mapped destination folder under the backup root directory.

        Reads the current form values (not the saved config) so unsaved changes
        are respected. Source/arrival folders are intentionally NOT created here —
        those are managed by an external process.
        """
        backup_dir = self.backup_dir_var.get().strip()
        if not backup_dir:
            messagebox.showwarning("Missing Value", "Enter a Backup Root Directory first.")
            return

        rows = self.map_tree.get_children()
        if not rows:
            messagebox.showwarning("No Mappings", "Add at least one folder mapping first.")
            return

        root      = Path(backup_dir)
        created   = []
        existing  = []

        for row in rows:
            dest_name   = self.map_tree.item(row)["values"][1]  # backup subfolder name
            dest_folder = root / dest_name

            if dest_folder.exists():
                existing.append(str(dest_folder))
            else:
                dest_folder.mkdir(parents=True, exist_ok=True)
                created.append(str(dest_folder))
                logging.info("Created backup folder: %s", dest_folder)

        # Build a summary message
        lines = []
        if created:
            lines.append(f"Created {len(created)} folder(s):")
            lines.extend(f"  \u2713  {p}" for p in created)
        if existing:
            if lines:
                lines.append("")
            lines.append(f"Already existed ({len(existing)}):")
            lines.extend(f"  –  {p}" for p in existing)

        messagebox.showinfo("Backup Folders", "\n".join(lines))

    # ── Demo tab helpers ──────────────────────────────────────────────────────

    def _refresh_demo_tab(self):
        """Reload folder list from config so the Demo tab stays in sync with Settings."""
        try:
            config = load_config()
        except Exception:
            return

        source_dir   = config.get("source_dir", "(not set)")
        folder_names = list(config.get("folder_map", {}).keys())

        # Update the info label and folder dropdown
        self.demo_source_label.config(text=f"Source directory: {source_dir}")
        self.demo_folder_combo.config(values=folder_names)

        # Keep the current selection if it's still valid; otherwise pick the first entry
        if folder_names and self.demo_folder_var.get() not in folder_names:
            self.demo_folder_var.set(folder_names[0])

    def _on_demo_create(self):
        """Create the test files and report results in the demo output area."""
        try:
            config = load_config()
        except Exception as e:
            messagebox.showerror("Config Error", f"Could not read config.json:\n{e}")
            return

        source_dir  = Path(config["source_dir"])
        folder_name = self.demo_folder_var.get()
        count       = self.demo_count_var.get()

        if not folder_name:
            messagebox.showwarning("No Folder", "No folder selected — add a mapping in Settings first.")
            return

        try:
            created = create_test_files(source_dir, folder_name, count)
        except Exception as e:
            messagebox.showerror("Error", f"Could not create files:\n{e}")
            return

        # Display results in the output area
        self.demo_output.config(state=tk.NORMAL)
        self.demo_output.delete("1.0", tk.END)
        self.demo_output.insert(tk.END, f"Created {len(created)} file(s) in:\n")
        self.demo_output.insert(tk.END, f"  {source_dir / folder_name}\n\n")
        for path in created:
            self.demo_output.insert(tk.END, f"  \u2713  {path.name}\n")
        self.demo_output.insert(tk.END, "\nSwitch to the GetFiles tab and click GetFiles!")
        self.demo_output.config(state=tk.DISABLED)

    # ── Refresh helpers ───────────────────────────────────────────────────────

    def _refresh_run_history(self):
        """Reload run_history.json into the Run History treeview, newest first."""
        self.run_tree.delete(*self.run_tree.get_children())
        for record in reversed(load_run_history()):
            dt = datetime.fromisoformat(record["run_time"]).strftime("%Y-%m-%d %H:%M:%S")
            self.run_tree.insert("", tk.END, values=(
                dt,
                record.get("files_found", 0),
                record.get("copied", 0),
                record.get("skipped", 0),
            ))

    def _refresh_file_history(self):
        """Reload file_history.json and repopulate the File History treeview."""
        self._all_file_records = list(reversed(load_file_history()))  # newest first
        self._populate_file_tree(self._all_file_records)

    def _populate_file_tree(self, records: list):
        """Fill the file history treeview from a (possibly filtered) record list."""
        self.file_tree.delete(*self.file_tree.get_children())
        for rec in records:
            dt = datetime.fromisoformat(rec["run_time"]).strftime("%Y-%m-%d %H:%M:%S")
            self.file_tree.insert("", tk.END, values=(
                dt,
                rec.get("file_name", ""),
                rec.get("folder", ""),
                rec.get("dest_path", ""),
            ))

    def _on_search_changed(self, *_args):
        """Filter the file history table as the user types in the search box."""
        query = self.search_var.get().lower()
        if not query:
            self._populate_file_tree(self._all_file_records)
            return
        filtered = [
            r for r in self._all_file_records
            if query in r.get("file_name", "").lower()
            or query in r.get("folder", "").lower()
        ]
        self._populate_file_tree(filtered)

    def _refresh_log_tab(self):
        """Read file_manager.log and display its contents, scrolled to the bottom."""
        self.log_file_area.config(state=tk.NORMAL)
        self.log_file_area.delete("1.0", tk.END)
        if LOG_FILE.exists():
            self.log_file_area.insert(tk.END, LOG_FILE.read_text(encoding="utf-8", errors="replace"))
            self.log_file_area.see(tk.END)   # jump to latest entry
        else:
            self.log_file_area.insert(tk.END, "(Log file not found — run GetFiles first)")
        self.log_file_area.config(state=tk.DISABLED)

    # ── GetFiles handlers ─────────────────────────────────────────────────────

    def _on_get_files(self):
        """Disable the button and run get_files in a background thread."""
        self.btn_get.config(state=tk.DISABLED, text="Running…")

        def run():
            get_files(self._log)
            self.after(0, self._on_done)

        threading.Thread(target=run, daemon=True).start()

    def _on_done(self):
        """Re-enable the button and refresh the status bar after a run completes."""
        self.btn_get.config(state=tk.NORMAL, text="GetFiles")
        self._update_status()

    def _log(self, message: str):
        """Append a line to the live log area. Safe to call from a background thread."""
        def _insert():
            self.log_area.config(state=tk.NORMAL)
            self.log_area.insert(tk.END, message + "\n")
            self.log_area.see(tk.END)
            self.log_area.config(state=tk.DISABLED)

        self.after(0, _insert)

    def _clear_log(self):
        """Wipe the live log display (does not touch the log file on disk)."""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete("1.0", tk.END)
        self.log_area.config(state=tk.DISABLED)

    def _update_status(self):
        """Refresh the status bar with the last run timestamp."""
        last = load_last_run()
        if last.year == 1970:
            self.status_var.set("Never run  |  Log: " + str(LOG_FILE))
        else:
            self.status_var.set(
                f"Last run: {last.strftime('%Y-%m-%d %H:%M:%S')}  |  Log: {LOG_FILE}"
            )


    # ── SQL Import tab ────────────────────────────────────────────────────────

    def _build_sql_import_tab(self):
        """SQL Import tab — connect to SQL Server, import files, view import log."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  SQL Import  ")

        # Show a friendly message if pyodbc is not installed
        if not PYODBC_AVAILABLE:
            msg_frame = ttk.Frame(frame)
            msg_frame.pack(expand=True)
            ttk.Label(
                msg_frame,
                text="pyodbc is not installed.\n\nRun this in a terminal to fix it:\n\n"
                     "    pip install pyodbc",
                font=("Consolas", 10), foreground="#cc0000", justify=tk.CENTER,
            ).pack(pady=40)
            return

        # ── Connection settings ───────────────────────────────────────────────
        conn_frame = ttk.LabelFrame(frame, text=" Database Connection ", padding=6)
        conn_frame.pack(fill=tk.X, padx=8, pady=(6, 4))

        # Row 1: server, database, buttons
        row1 = tk.Frame(conn_frame)
        row1.pack(fill=tk.X)

        ttk.Label(row1, text="Server:").pack(side=tk.LEFT, padx=(0, 4))
        self.sql_server_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.sql_server_var, width=22
                  ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row1, text="Database:").pack(side=tk.LEFT, padx=(0, 4))
        self.sql_db_var   = tk.StringVar()
        self.sql_db_combo = ttk.Combobox(
            row1, textvariable=self.sql_db_var, width=22,
        )
        self.sql_db_combo.pack(side=tk.LEFT, padx=(0, 12))
        # Refresh the SP list whenever the user selects a different database
        self.sql_db_combo.bind("<<ComboboxSelected>>",
                               lambda _e: self._load_stored_procedures())

        ttk.Button(row1, text="Test Connection",
                   command=self._test_sql_connection).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(row1, text="Save",
                   command=self._save_sql_connection).pack(side=tk.LEFT)

        # Row 2: connection status label (green = ok, red = error)
        self.sql_conn_status = ttk.Label(
            conn_frame, text="Not connected",
            foreground="#888", font=("Segoe UI", 8),
        )
        self.sql_conn_status.pack(anchor=tk.W, pady=(4, 0))

        # ── Import controls ───────────────────────────────────────────────────
        import_frame = ttk.LabelFrame(frame, text=" Import Files to SQL Staging ", padding=6)
        import_frame.pack(fill=tk.X, padx=8, pady=(0, 4))

        ctrl_row = tk.Frame(import_frame)
        ctrl_row.pack(fill=tk.X)

        # Row 1: folder + file type
        ttk.Label(ctrl_row, text="Folder:").pack(side=tk.LEFT, padx=(0, 4))
        self.sql_folder_var = tk.StringVar(value="All Folders")
        self.sql_folder_combo = ttk.Combobox(
            ctrl_row, textvariable=self.sql_folder_var,
            state="readonly", width=18,
        )
        self.sql_folder_combo.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(ctrl_row, text="File type:").pack(side=tk.LEFT, padx=(0, 4))
        self.sql_filetype_var = tk.StringVar(value="Auto-detect")
        ttk.Combobox(
            ctrl_row, textvariable=self.sql_filetype_var,
            values=["Auto-detect", "CSV", "NDJSON"],
            state="readonly", width=12,
        ).pack(side=tk.LEFT, padx=(0, 12))

        # Row 2: stored procedure picker + run button
        sp_row = tk.Frame(import_frame)
        sp_row.pack(fill=tk.X, pady=(4, 0))

        ttk.Label(sp_row, text="Stored Procedure:").pack(side=tk.LEFT, padx=(0, 4))
        self.sql_sp_var   = tk.StringVar(value="Auto-detect")
        self.sql_sp_combo = ttk.Combobox(
            sp_row, textvariable=self.sql_sp_var,
            values=["Auto-detect"], state="readonly", width=36,
        )
        self.sql_sp_combo.pack(side=tk.LEFT, padx=(0, 12))

        self.btn_sql_import = tk.Button(
            sp_row,
            text="Run SQL Import",
            bg="#2d7dd2", fg="white", font=("Segoe UI", 10, "bold"),
            width=16, command=self._on_sql_import,
        )
        self.btn_sql_import.pack(side=tk.LEFT)

        ttk.Label(
            import_frame,
            text="Scans backup folders and imports files into staging.RawCSV / staging.RawJSON. "
                 "Files already in the import log are skipped.",
            foreground="#666", font=("Segoe UI", 8),
        ).pack(anchor=tk.W, pady=(4, 0))

        # ── Live log output ───────────────────────────────────────────────────
        self.sql_log_area = scrolledtext.ScrolledText(
            frame, height=7, wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Consolas", 9),
            bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white",
        )
        self.sql_log_area.pack(fill=tk.X, padx=8, pady=(0, 4))

        # ── Import log grid header ────────────────────────────────────────────
        log_hdr = tk.Frame(frame)
        log_hdr.pack(fill=tk.X, padx=8, pady=(0, 2))

        ttk.Label(
            log_hdr, text="Import Log  (from database)",
            font=("Segoe UI", 9, "bold"),
        ).pack(side=tk.LEFT)

        ttk.Button(
            log_hdr, text="Refresh",
            command=self._refresh_sql_import_log,
        ).pack(side=tk.RIGHT)

        # ── Import log treeview ───────────────────────────────────────────────
        cols = ("import_id", "imported_at", "source_folder", "file_name",
                "type", "destination", "loaded", "valid", "rejected", "status")

        self.sql_log_tree = ttk.Treeview(
            frame, columns=cols, show="headings",
            selectmode="browse", height=6,
        )

        self.sql_log_tree.heading("import_id",     text="ID")
        self.sql_log_tree.heading("imported_at",   text="Imported At")
        self.sql_log_tree.heading("source_folder", text="Source")
        self.sql_log_tree.heading("file_name",     text="File")
        self.sql_log_tree.heading("type",          text="Type")
        self.sql_log_tree.heading("destination",   text="Destination")
        self.sql_log_tree.heading("loaded",        text="Loaded")
        self.sql_log_tree.heading("valid",         text="Valid")
        self.sql_log_tree.heading("rejected",      text="Rejected")
        self.sql_log_tree.heading("status",        text="Status")

        self.sql_log_tree.column("import_id",     width=40,  anchor=tk.CENTER)
        self.sql_log_tree.column("imported_at",   width=145, anchor=tk.W)
        self.sql_log_tree.column("source_folder", width=70,  anchor=tk.CENTER)
        self.sql_log_tree.column("file_name",     width=180, anchor=tk.W)
        self.sql_log_tree.column("type",          width=50,  anchor=tk.CENTER)
        self.sql_log_tree.column("destination",   width=120, anchor=tk.W)
        self.sql_log_tree.column("loaded",        width=55,  anchor=tk.CENTER)
        self.sql_log_tree.column("valid",         width=50,  anchor=tk.CENTER)
        self.sql_log_tree.column("rejected",      width=65,  anchor=tk.CENTER)
        self.sql_log_tree.column("status",        width=90,  anchor=tk.CENTER)

        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL,   command=self.sql_log_tree.yview)
        hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=self.sql_log_tree.xview)
        self.sql_log_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.sql_log_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True,
                               padx=(8, 0), pady=(0, 0))
        vsb.pack(side=tk.LEFT, fill=tk.Y, pady=(0, 0), padx=(0, 8))
        hsb.pack(side=tk.BOTTOM, fill=tk.X, padx=(8, 8), pady=(0, 8))

        # Load any saved connection settings into the form
        self._load_sql_connection_settings()

    # ── SQL connection helpers ─────────────────────────────────────────────────

    def _load_sql_connection_settings(self):
        """Populate Server and Database fields from config.json if saved previously."""
        try:
            config = load_config()
            sql_cfg = config.get("sql_connection", {})
            self.sql_server_var.set(sql_cfg.get("server", ""))
            self.sql_db_var.set(sql_cfg.get("database", ""))
        except Exception:
            pass  # first run or bad config; fields stay blank

    def _get_sql_conn_str(self) -> str:
        """Build a Windows-auth pyodbc connection string using the best available driver.

        Tries ODBC Driver 17 / 18 for SQL Server first (modern drivers), then falls
        back to the older generic 'SQL Server' driver that ships with Windows.
        """
        server   = self.sql_server_var.get().strip()
        database = self.sql_db_var.get().strip()

        if not server or not database:
            raise ValueError("Server and Database must both be filled in.")

        # Pick the most capable ODBC driver available on this machine
        available = [d for d in pyodbc.drivers() if "SQL Server" in d]
        if not available:
            raise RuntimeError(
                "No SQL Server ODBC driver found.\n"
                "Install 'ODBC Driver 17 for SQL Server' from Microsoft."
            )

        # Prefer numbered drivers (17, 18) over the plain 'SQL Server' fallback
        numbered = sorted(
            [d for d in available if any(c.isdigit() for c in d)],
            reverse=True,   # highest version number first
        )
        driver = numbered[0] if numbered else available[0]

        return (
            f"DRIVER={{{driver}}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            "Trusted_Connection=yes;"
            "TrustServerCertificate=yes;"
        )

    def _test_sql_connection(self):
        """Try to connect and run a trivial query; populate the DB and SP dropdowns on success."""
        try:
            conn_str = self._get_sql_conn_str()
            conn     = pyodbc.connect(conn_str, timeout=5)
            conn.execute("SELECT 1")
            self._load_databases(conn)   # populate DB dropdown before closing
            conn.close()
            self.sql_conn_status.config(text="Connected \u2713", foreground="#2a7a2a")
            self._load_stored_procedures()   # populate SP dropdown for selected DB
        except Exception as e:
            self.sql_conn_status.config(text=f"Error: {e}", foreground="#cc0000")

    def _load_databases(self, conn) -> None:
        """Populate the Database dropdown from sys.databases on the already-open connection.

        Excludes the four built-in system databases (master, tempdb, model, msdb)
        which have database_id <= 4.  Preserves the current selection if it is still
        in the returned list.
        """
        try:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT name FROM sys.databases "
                "WHERE database_id > 4 ORDER BY name"
            )
            db_names = [row[0] for row in cursor.fetchall()]
            cursor.close()
        except Exception:
            return  # leave dropdown unchanged if the query fails

        current = self.sql_db_var.get()
        self.sql_db_combo.config(values=db_names)

        # Keep the current selection if it still exists, otherwise leave it
        if current not in db_names and db_names:
            self.sql_db_var.set(db_names[0])

    def _load_stored_procedures(self) -> None:
        """Populate the SP dropdown with all stored procedures in the selected database.

        Opens a fresh connection so this can be called independently of Test Connection.
        Silently no-ops if the Server or Database fields are empty or the connection fails.
        Resets the selection to 'Auto-detect' if the previously selected SP is gone.
        """
        if not PYODBC_AVAILABLE:
            return
        try:
            conn_str = self._get_sql_conn_str()
        except ValueError:
            return  # fields not filled in yet

        try:
            conn   = pyodbc.connect(conn_str, timeout=5)
            cursor = conn.cursor()
            cursor.execute(
                "SELECT SCHEMA_NAME(schema_id) + '.' + name "
                "FROM sys.procedures "
                "ORDER BY SCHEMA_NAME(schema_id), name"
            )
            sp_names = [row[0] for row in cursor.fetchall()]
            cursor.close()
            conn.close()
        except Exception:
            return  # leave dropdown unchanged on any error

        values  = ["Auto-detect"] + sp_names
        current = self.sql_sp_var.get()
        self.sql_sp_combo.config(values=values)

        if current not in values:
            self.sql_sp_var.set("Auto-detect")

    def _save_sql_connection(self):
        """Persist server and database names to config.json under sql_connection."""
        server   = self.sql_server_var.get().strip()
        database = self.sql_db_var.get().strip()

        if not server or not database:
            messagebox.showwarning("Missing Values",
                                   "Enter both Server and Database before saving.")
            return

        try:
            config = load_config()
        except Exception:
            config = {}

        config["sql_connection"] = {"server": server, "database": database}

        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f, indent=4)

        messagebox.showinfo("Saved", "SQL connection settings saved to config.json.")

    # ── SQL Import helpers ─────────────────────────────────────────────────────

    def _refresh_sql_import_folders(self):
        """Rebuild the folder dropdown from the backup subfolder names in config."""
        try:
            config = load_config()
        except Exception:
            return

        # Use the destination (backup) subfolder names, not the source names
        dest_folders = list(config.get("folder_map", {}).values())
        values = ["All Folders"] + dest_folders

        self.sql_folder_combo.config(values=values)
        if self.sql_folder_var.get() not in values:
            self.sql_folder_var.set("All Folders")

    def _on_sql_import(self):
        """Disable the import button and run the import in a background thread."""
        self.btn_sql_import.config(state=tk.DISABLED, text="Importing…")

        # Clear the live log before starting
        self.sql_log_area.config(state=tk.NORMAL)
        self.sql_log_area.delete("1.0", tk.END)
        self.sql_log_area.config(state=tk.DISABLED)

        def run():
            try:
                self._do_sql_import()
            finally:
                # Re-enable the button and refresh the log grid when done
                self.after(0, lambda: self.btn_sql_import.config(
                    state=tk.NORMAL, text="Run SQL Import"
                ))
                self.after(0, self._refresh_sql_import_log)

        threading.Thread(target=run, daemon=True).start()

    def _do_sql_import(self):
        """Core import logic — runs in a background thread.

        Scans the selected backup folder(s), skips files already in the import log,
        and calls the appropriate stored procedure for each new file.
        """
        try:
            config     = load_config()
            backup_dir = Path(config["backup_dir"])
            folder_map = config.get("folder_map", {})
        except Exception as e:
            self._sql_log(f"ERROR: Could not read config — {e}")
            return

        try:
            conn_str = self._get_sql_conn_str()
            conn     = pyodbc.connect(conn_str, timeout=10)
        except Exception as e:
            self._sql_log(f"ERROR: Could not connect to SQL Server — {e}")
            return

        selected_folder = self.sql_folder_var.get()
        force_type      = self.sql_filetype_var.get()   # 'Auto-detect', 'CSV', 'NDJSON'
        selected_sp     = self.sql_sp_var.get().strip()
        use_auto_sp     = (not selected_sp or selected_sp == "Auto-detect")

        # Build the list of backup subfolders to scan
        if selected_folder == "All Folders":
            subfolders = list(folder_map.values())
        else:
            subfolders = [selected_folder]

        self._sql_log(f"── SQL Import started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ──")
        self._sql_log(f"Backup dir : {backup_dir}")
        self._sql_log(f"Folders    : {', '.join(subfolders)}")
        self._sql_log("")

        total_imported = 0
        total_skipped  = 0

        cursor = conn.cursor()

        for subfolder in subfolders:
            folder_path = backup_dir / subfolder
            if not folder_path.exists():
                self._sql_log(f"  SKIP folder '{subfolder}' — not found on disk")
                continue

            # Collect all files in this backup subfolder
            files = sorted(p for p in folder_path.rglob("*") if p.is_file())
            if not files:
                self._sql_log(f"  [{subfolder}]  No files found")
                continue

            self._sql_log(f"  [{subfolder}]  {len(files)} file(s) found")

            for file in files:
                file_str = str(file)

                # Check whether this file is already in the import log
                cursor.execute(
                    "SELECT 1 FROM import.FileImportLog "
                    "WHERE FilePath = ? AND Status <> 'ERROR'",
                    file_str,
                )
                if cursor.fetchone():
                    self._sql_log(f"    SKIP  {file.name}  (already imported)")
                    total_skipped += 1
                    continue

                # Determine file type
                ext = file.suffix.lower()
                if force_type == "CSV":
                    file_type = "csv"
                elif force_type == "NDJSON":
                    file_type = "json"
                else:
                    # Auto-detect by extension
                    if ext in (".csv", ".txt"):
                        file_type = "csv"
                    elif ext in (".json", ".ndjson"):
                        file_type = "json"
                    else:
                        self._sql_log(f"    SKIP  {file.name}  (unsupported extension {ext})")
                        total_skipped += 1
                        continue

                try:
                    if not use_auto_sp:
                        # User picked a specific SP — call it with just @FilePath
                        self._sql_log(f"    INFO  {file.name}  (SP: {selected_sp})")
                        cursor.execute(f"EXEC {selected_sp} @FilePath=?", file_str)
                    elif file_type == "csv":
                        delimiter, skip_header = _sniff_csv(file)
                        self._sql_log(
                            f"    INFO  {file.name}  "
                            f"(delimiter={repr(delimiter)}, header={bool(skip_header)})"
                        )
                        cursor.execute(
                            "EXEC import.usp_ImportCSVFile "
                            "@FilePath=?, @Delimiter=?, @SkipHeader=?",
                            file_str, delimiter, skip_header,
                        )
                    else:
                        cursor.execute(
                            "EXEC import.usp_ImportJSONFile @FilePath=?",
                            file_str,
                        )

                    row = cursor.fetchone()
                    import_id   = row[0] if row else "?"
                    rows_loaded = row[1] if row else "?"
                    conn.commit()

                    self._sql_log(
                        f"    OK    {file.name}  "
                        f"(ImportId={import_id}, {rows_loaded} row(s) loaded)"
                    )
                    total_imported += 1

                except Exception as e:
                    conn.rollback()
                    self._sql_log(f"    ERROR {file.name}  — {e}")

        self._sql_log("")
        self._sql_log(
            f"Done — {total_imported} file(s) imported, "
            f"{total_skipped} skipped (already done or unsupported)."
        )

        cursor.close()
        conn.close()

    def _sql_log(self, message: str):
        """Append a line to the SQL Import live log. Safe to call from a background thread."""
        def _insert():
            self.sql_log_area.config(state=tk.NORMAL)
            self.sql_log_area.insert(tk.END, message + "\n")
            self.sql_log_area.see(tk.END)
            self.sql_log_area.config(state=tk.DISABLED)

        self.after(0, _insert)

    # ── SQL Import log grid ────────────────────────────────────────────────────

    def _refresh_sql_import_log(self):
        """Query import.FileImportLog and populate the grid. Silently no-ops if
        there is no connection configured yet."""
        if not PYODBC_AVAILABLE:
            return

        try:
            conn_str = self._get_sql_conn_str()
        except ValueError:
            # Server or database not filled in yet — leave the grid empty
            return

        try:
            conn   = pyodbc.connect(conn_str, timeout=5)
            cursor = conn.cursor()
            cursor.execute("""
                SELECT TOP 100
                    ImportId,
                    ImportedAt,
                    FilePath,
                    FileType,
                    ISNULL(RowsLoaded,   0),
                    ISNULL(RowsValid,    0),
                    ISNULL(RowsRejected, 0),
                    Status
                FROM import.FileImportLog
                ORDER BY ImportedAt DESC
            """)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
        except Exception as e:
            # Connection works for test but fails here — surface the error
            self.sql_conn_status.config(
                text=f"Log refresh error: {e}", foreground="#cc0000"
            )
            return

        self.sql_log_tree.delete(*self.sql_log_tree.get_children())

        for row in rows:
            import_id, imported_at, file_path, file_type, loaded, valid, rejected, status = row

            p             = Path(file_path) if file_path else Path("")
            file_name     = p.name
            source_folder = p.parent.name   # e.g. "ABC", "DEF"

            # Map file type to its staging table destination
            destination = {
                "CSV":  "staging.RawCSV",
                "JSON": "staging.RawJSON",
            }.get((file_type or "").upper(), "staging.RawCSV")

            # Format the timestamp if it's a datetime object
            if isinstance(imported_at, datetime):
                imported_at = imported_at.strftime("%Y-%m-%d %H:%M:%S")

            self.sql_log_tree.insert("", tk.END, values=(
                import_id,
                imported_at,
                source_folder,
                file_name,
                file_type,
                destination,
                loaded,
                valid,
                rejected,
                status,
            ))


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = FileManagerApp()
    app.mainloop()
