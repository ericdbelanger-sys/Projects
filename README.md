# File Manager

A Windows desktop application for monitoring source folders, copying files to mapped backup destinations, and importing file data into SQL Server staging tables.

Built with Python + tkinter. No external dependencies except `pyodbc` (optional, for the SQL Import tab).

---

## Running the App

```bash
py file_manager.py
```

---

## Configuration

All settings are stored in `config.json` next to the script.

| Key | Type | Example | Purpose |
|---|---|---|---|
| `source_dir` | string | `"C:/FileLanding"` | Root folder where incoming files arrive in subfolders |
| `backup_dir` | string | `"C:/FileDest"` | Root folder where backed-up files are written |
| `folder_map` | object | `{"ABC": "ABC"}` | Maps each source subfolder name to a destination subfolder name under `backup_dir` |
| `sql_connection` | object | `{"server": "localhost", "database": "FileImportDB"}` | SQL Server connection (saved by the SQL Import tab) |

Example:
```json
{
    "source_dir": "C:/FileLanding",
    "backup_dir": "C:/FileDest",
    "folder_map": {
        "ABC": "ABC",
        "DEF": "DEF",
        "GHI": "GHI"
    },
    "sql_connection": {
        "server": "DESKTOP-GAN774J",
        "database": "FileImportDB"
    }
}
```

---

## Tabs

### GetFiles
The main action tab. Click **Run GetFiles** to scan `source_dir` for new files and copy them to their mapped backup folders.

- Files are only copied if they are newer than the last successful run timestamp
- Each file is copied with full metadata preserved (mtime, atime, Windows creation time)
- Live log shows what was copied, skipped, or flagged as having no mapping
- **Clear Log** wipes the on-screen log without affecting the file log

### Settings
Configure source/backup directories and folder mappings.

- **Backup Directory** — path field with a Browse button
- **Folder Mappings** — table of source folder name → destination subfolder name pairs
  - Add Row / Remove Row buttons manage entries
  - **Create Backup Folders** — creates any missing destination folders on disk
- **Save Settings** — writes all changes to `config.json`

### Demo
Generate test files in a source folder to try out the GetFiles workflow.

- Pick a source folder from the folder map
- Choose how many files to create (slider)
- Click **Create Test Files** — files are randomly `.csv`, `.txt` (pipe-delimited), or `.ndjson`, each with realistic mock data (5–15 rows)
- Switch to GetFiles and click Run to see them get picked up and copied

### Run History
Read-only table of every past GetFiles execution.

Columns: **Date/Time**, **Found**, **Copied**, **Skipped**

Data source: `run_history.json`

### File History
Read-only table of every individual file that has been copied.

- Live search box filters by file name or folder name
- Columns: **Date/Time**, **File Name**, **Source Folder**, **Destination**

Data source: `file_history.json`

### Log File
Displays the raw contents of `file_manager.log`. Click **Refresh** to reload.

### SQL Import
Connect to SQL Server and import files from backup folders into staging tables.

#### Connection Settings
| Field | Purpose |
|---|---|
| Server | SQL Server instance name (e.g. `DESKTOP-GAN774J` or `localhost\SQLEXPRESS`) |
| Database | Dropdown populated from server after connecting — pick the target database |
| **Test Connection** | Verifies connection, auto-populates the Database and Stored Procedure dropdowns |
| **Save** | Persists Server + Database to `config.json` |

#### Import Controls
| Control | Purpose |
|---|---|
| Folder | Which backup subfolder to scan (`All Folders` or a specific one) |
| File type | Force CSV or NDJSON, or leave as `Auto-detect` |
| Stored Procedure | Pick any SP in the database, or `Auto-detect` (uses `usp_ImportCSVFile` / `usp_ImportJSONFile`) |
| **Run SQL Import** | Scans the selected folder(s), skips already-imported files, calls the SP for each new file |

#### Import Log Grid
Shows the last 100 rows from `import.FileImportLog`.

Columns: **ID**, **Imported At**, **Source** (folder name), **File**, **Type**, **Destination** (staging table), **Loaded**, **Valid**, **Rejected**, **Status**

> **Note:** SQL Server's service account must have read access to the backup folder path. Keep files outside `C:\Users\` (e.g. `C:\FileDest\`) to avoid permission errors.

### File Inspector
Pick any file from a backup subfolder and view a statistical report.

- Select a **Folder** → File dropdown auto-populates
- Select a **File** → click **Inspect**

**CSV / TXT output:**
```
FILE SUMMARY
  Path      : C:\FileDest\ABC\employees2.csv
  Size      : 312 B
  Rows      : 5  (excluding header)
  Columns   : 8
  Delimiter : ','
  Header    : yes

COLUMN ANALYSIS
  #   Name                   MaxLen  MinLen  AvgLen  NonEmpty  Samples
  1   EmployeeId                  4       3     3.4         5  1042, 872, 315
  2   FirstName                   6       3     5.0         5  Alice, Bob, Carol
  ...
```

**NDJSON output:**
```
FILE SUMMARY
  Path    : C:\FileDest\GHI\inventory.ndjson
  Size    : 512 B
  Records : 6  (valid: 6, invalid: 0)

KEY ANALYSIS
  Key                    Types           MaxLen  NonNull  Samples
  accountType            str                 10        6  Standard, Premium, Enterprise
  balance                float                -        6  1234.56, -42.0, 9999.99
  customerId             str                  8        6  CUST1042, CUST7731
  ...
```

---

## Data Files

| File | Purpose |
|---|---|
| `config.json` | Application settings (see Configuration above) |
| `last_run.json` | Timestamp of the last successful GetFiles run |
| `run_history.json` | One record per GetFiles execution |
| `file_history.json` | One record per file copied |
| `file_manager.log` | Rolling activity log (appended on every run) |

---

## Module Reference

### Module-Level Functions

| Function | Purpose |
|---|---|
| `copy_with_metadata(src, dst)` | Copy a file preserving mtime, atime, and Windows creation time |
| `create_test_files(source_dir, folder_name, count)` | Create `count` random demo files (CSV / TXT / NDJSON) with mock data |
| `get_files(log_fn)` | Core copy logic — scans source folders, copies new files, appends history |
| `load_config()` | Read and parse `config.json` |
| `load_last_run()` | Return timestamp of last run (epoch 0 if none) |
| `save_last_run(ts)` | Persist run timestamp to `last_run.json` |
| `load_run_history()` | Load all run summary records from `run_history.json` |
| `append_run_record(record)` | Append a run summary record to `run_history.json` |
| `load_file_history()` | Load all file-copy records from `file_history.json` |
| `append_file_records(records)` | Append a batch of file-copy records to `file_history.json` |
| `_sniff_csv(file)` | Auto-detect delimiter and whether a header row is present |

### FileManagerApp Methods

#### GetFiles Tab
| Method | Purpose |
|---|---|
| `_on_get_files()` | Disable button and run `get_files()` in a background thread |
| `_on_done()` | Re-enable button after background thread completes |
| `_log(message)` | Thread-safe append to the live log area |
| `_clear_log()` | Clear the on-screen log area |
| `_update_status()` | Refresh the bottom status bar |

#### Settings Tab
| Method | Purpose |
|---|---|
| `_load_settings_form()` | Populate form fields from `config.json` |
| `_browse_backup_dir()` | Open a folder picker dialog for the backup directory |
| `_add_mapping()` | Validate and add a new folder mapping row |
| `_remove_mapping()` | Delete the selected mapping row |
| `_save_settings()` | Write form values to `config.json` |
| `_create_backup_folders()` | Create any missing destination folders on disk |

#### Demo Tab
| Method | Purpose |
|---|---|
| `_refresh_demo_tab()` | Reload folder list from config |
| `_on_demo_create()` | Create test files and display results |

#### History Tabs
| Method | Purpose |
|---|---|
| `_refresh_run_history()` | Reload `run_history.json` into the treeview |
| `_refresh_file_history()` | Reload `file_history.json` into the treeview |
| `_populate_file_tree(records)` | Fill the file history grid from a record list |
| `_on_search_changed(*args)` | Filter file history as the user types in the search box |
| `_refresh_log_tab()` | Read and display `file_manager.log` |

#### SQL Import Tab
| Method | Purpose |
|---|---|
| `_load_sql_connection_settings()` | Populate Server/Database fields from `config.json` |
| `_get_sql_conn_str()` | Build Windows-auth pyodbc connection string |
| `_test_sql_connection()` | Test connection, populate DB and SP dropdowns |
| `_load_databases(conn)` | Query `sys.databases` and populate the Database dropdown |
| `_load_stored_procedures()` | Query `sys.procedures` and populate the SP dropdown |
| `_save_sql_connection()` | Persist Server + Database to `config.json` |
| `_refresh_sql_import_folders()` | Rebuild the folder dropdown from `folder_map` |
| `_on_sql_import()` | Disable button and run `_do_sql_import()` in a background thread |
| `_do_sql_import()` | Scan backup folders, skip already-imported files, call SPs |
| `_sql_log(message)` | Thread-safe append to the SQL Import live log |
| `_refresh_sql_import_log()` | Query `import.FileImportLog` and populate the grid |

#### File Inspector Tab
| Method | Purpose |
|---|---|
| `_refresh_inspector_folders()` | Populate the Folder dropdown from `folder_map` |
| `_refresh_inspector_files()` | Populate the File dropdown from the selected folder |
| `_on_inspect()` | Validate selection and run `_do_inspect()` in a background thread |
| `_do_inspect(file_path)` | Analyse the file and write a formatted report to the output area |
| `_insp_write(text)` | Thread-safe append to the inspector output area |

---

## Dependencies

| Package | Required | Install |
|---|---|---|
| Python 3.12+ | Yes | — |
| tkinter | Yes | Included with Python on Windows |
| pyodbc | No (SQL Import tab only) | `pip install pyodbc` |

All other imports (`csv`, `json`, `logging`, `random`, `shutil`, `threading`, `ctypes`, `pathlib`, `datetime`) are Python standard library.
