# DOC to DOCX Batch Converter

A Windows desktop application that converts `.doc` files to `.docx` format in bulk — with a simple graphical interface, an input folder, an output folder, and a real-time log.

---

## Features

- **Batch conversion** — converts every `.doc` file found in the chosen input folder
- **Two conversion engines** — uses Microsoft Word (via COM automation) or LibreOffice (headless), auto-detected or manually selected
- **Input / Output folders** — choose separate source and destination folders
- **Progress bar & log** — real-time feedback for each file
- **Skip existing files** — skips `.docx` files that already exist in the output folder
- **No internet required** — runs entirely offline

---

## Requirements

### Python

- **Python 3.8 or higher** — download from https://www.python.org/downloads/

Python's `tkinter` module is included with the standard Python installer on Windows.

### Conversion Engine (one of the following)

#### Option A — Microsoft Word (recommended if you have Office installed)

1. Microsoft Word must be installed on the machine.
2. Install the `comtypes` Python package:

   ```
   pip install comtypes
   ```

#### Option B — LibreOffice (free, no Microsoft Office needed)

1. Download and install LibreOffice from https://www.libreoffice.org
2. No extra Python packages needed — LibreOffice is invoked as a command-line tool.

---

## Installation

1. Make sure Python 3.8+ is installed and added to your PATH.
2. Open a Command Prompt and run:

   ```
   pip install comtypes
   ```

   *(Only needed for the Microsoft Word method. Skip this if you plan to use LibreOffice.)*

3. Download or copy `doc_to_docx_converter.py` to a folder of your choice.

---

## How to Run

Double-click `doc_to_docx_converter.py` if Python is associated with `.py` files, **or** open a Command Prompt in the folder and run:

```
python doc_to_docx_converter.py
```

---

## Usage

1. **Input Folder** — click **Browse…** and select the folder that contains your `.doc` files.
2. **Output Folder** — click **Browse…** and select where the converted `.docx` files should be saved.  
   *(Defaults to the input folder if left empty.)*
3. **Conversion Method** — leave on **Auto-detect** or pick manually:
   - `Microsoft Word (COM)` — requires Word installed + `comtypes` package
   - `LibreOffice (headless)` — requires LibreOffice installed
4. Click **Convert** — the progress bar and log will update as each file is processed.
5. When done, the log shows how many files succeeded or failed.

---

## Troubleshooting

| Problem | Solution |
|---|---|
| `ModuleNotFoundError: No module named 'comtypes'` | Run `pip install comtypes` in Command Prompt |
| LibreOffice not found | Install LibreOffice from https://www.libreoffice.org, then restart the app |
| Word COM error: "Word not found" | Ensure Microsoft Word is installed and licensed |
| File conversion fails silently | Check the log for the specific error message on that file |
| App won't start | Make sure Python 3.8+ is installed and `python` is in your system PATH |

---

## File Structure

```
doc-converter/
├── doc_to_docx_converter.py   # Main application
├── requirements.txt           # Optional Python dependencies
└── README.md                  # This file
```

---

## Notes

- Only `.doc` files (legacy Word 97–2003 format) in the **top level** of the input folder are converted. Sub-folders are not scanned.
- Files that already have a matching `.docx` in the output folder are skipped automatically.
- The original `.doc` files are never modified or deleted.
- For best results with complex documents (tables, embedded objects, tracked changes), use the Microsoft Word COM method.

---

## License

Free to use and modify for personal or commercial purposes.
