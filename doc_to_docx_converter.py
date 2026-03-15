import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path
import glob


def find_libreoffice():
    """Find LibreOffice executable on Windows."""
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"C:\Program Files\LibreOffice 7\program\soffice.exe",
        r"C:\Program Files\LibreOffice 6\program\soffice.exe",
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    # Try PATH
    try:
        result = subprocess.run(
            ["where", "soffice"],
            capture_output=True, text=True, check=True
        )
        found = result.stdout.strip().splitlines()
        if found:
            return found[0]
    except Exception:
        pass
    return None


def convert_with_libreoffice(soffice_path, input_file, output_dir, log_callback):
    """Convert a .doc file to .docx using LibreOffice headless."""
    input_path = Path(input_file)
    output_path = Path(output_dir) / (input_path.stem + ".docx")

    if output_path.exists():
        log_callback(f"  [SKIP] Already exists: {output_path.name}")
        return True

    cmd = [
        soffice_path,
        "--headless",
        "--convert-to", "docx",
        "--outdir", str(output_dir),
        str(input_file),
    ]

    try:
        result = subprocess.run(
            cmd,
            capture_output=True, text=True, timeout=60
        )
        if result.returncode == 0:
            log_callback(f"  [OK]   {input_path.name} -> {output_path.name}")
            return True
        else:
            log_callback(f"  [FAIL] {input_path.name}: {result.stderr.strip()}")
            return False
    except subprocess.TimeoutExpired:
        log_callback(f"  [FAIL] {input_path.name}: Conversion timed out.")
        return False
    except Exception as e:
        log_callback(f"  [FAIL] {input_path.name}: {e}")
        return False


def convert_with_word_com(input_file, output_dir, log_callback):
    """Convert a .doc file to .docx using Microsoft Word COM automation."""
    try:
        import comtypes.client
    except ImportError:
        log_callback("  [ERROR] comtypes not installed. Install via: pip install comtypes")
        return False

    input_path = Path(input_file).resolve()
    output_path = (Path(output_dir) / (input_path.stem + ".docx")).resolve()

    if output_path.exists():
        log_callback(f"  [SKIP] Already exists: {output_path.name}")
        return True

    word = None
    doc = None
    try:
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(input_path))
        # wdFormatXMLDocument = 12
        doc.SaveAs(str(output_path), FileFormat=12)
        log_callback(f"  [OK]   {input_path.name} -> {output_path.name}")
        return True
    except Exception as e:
        log_callback(f"  [FAIL] {input_path.name}: {e}")
        return False
    finally:
        if doc:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass


class ConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DOC to DOCX Batch Converter")
        self.resizable(True, True)
        self.minsize(620, 480)

        self._input_folder = tk.StringVar()
        self._output_folder = tk.StringVar()
        self._method = tk.StringVar(value="auto")
        self._running = False

        self._build_ui()
        self._detect_tools()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # Title
        title_frame = tk.Frame(self, bg="#1e3a5f")
        title_frame.pack(fill=tk.X)
        tk.Label(
            title_frame,
            text="DOC  ->  DOCX  Batch Converter",
            font=("Segoe UI", 16, "bold"),
            bg="#1e3a5f",
            fg="white",
            pady=12,
        ).pack()

        # Folders
        folder_frame = tk.LabelFrame(self, text="Folders", font=("Segoe UI", 10, "bold"), padx=10, pady=8)
        folder_frame.pack(fill=tk.X, **pad)

        tk.Label(folder_frame, text="Input Folder:", anchor="w", width=14).grid(row=0, column=0, sticky="w")
        tk.Entry(folder_frame, textvariable=self._input_folder, width=50).grid(row=0, column=1, padx=6, sticky="ew")
        tk.Button(folder_frame, text="Browse…", command=self._browse_input, width=10).grid(row=0, column=2)

        tk.Label(folder_frame, text="Output Folder:", anchor="w", width=14).grid(row=1, column=0, sticky="w", pady=(6, 0))
        tk.Entry(folder_frame, textvariable=self._output_folder, width=50).grid(row=1, column=1, padx=6, sticky="ew", pady=(6, 0))
        tk.Button(folder_frame, text="Browse…", command=self._browse_output, width=10).grid(row=1, column=2, pady=(6, 0))

        folder_frame.columnconfigure(1, weight=1)

        # Method
        method_frame = tk.LabelFrame(self, text="Conversion Method", font=("Segoe UI", 10, "bold"), padx=10, pady=8)
        method_frame.pack(fill=tk.X, **pad)

        methods = [
            ("Auto-detect (recommended)", "auto"),
            ("Microsoft Word (COM)", "word"),
            ("LibreOffice (headless)", "libreoffice"),
        ]
        for label, value in methods:
            tk.Radiobutton(
                method_frame, text=label,
                variable=self._method, value=value,
                font=("Segoe UI", 9),
            ).pack(anchor="w")

        self._tool_status = tk.Label(method_frame, text="", font=("Segoe UI", 8, "italic"), fg="#555")
        self._tool_status.pack(anchor="w", pady=(4, 0))

        # Progress
        progress_frame = tk.LabelFrame(self, text="Progress", font=("Segoe UI", 10, "bold"), padx=10, pady=8)
        progress_frame.pack(fill=tk.X, **pad)

        self._progress_var = tk.DoubleVar(value=0)
        self._progress_bar = ttk.Progressbar(progress_frame, variable=self._progress_var, maximum=100)
        self._progress_bar.pack(fill=tk.X)

        self._status_label = tk.Label(progress_frame, text="Ready.", font=("Segoe UI", 9), anchor="w")
        self._status_label.pack(fill=tk.X, pady=(4, 0))

        # Log
        log_frame = tk.LabelFrame(self, text="Log", font=("Segoe UI", 10, "bold"), padx=10, pady=8)
        log_frame.pack(fill=tk.BOTH, expand=True, **pad)

        self._log = scrolledtext.ScrolledText(log_frame, height=10, font=("Consolas", 9), state=tk.DISABLED, wrap=tk.WORD)
        self._log.pack(fill=tk.BOTH, expand=True)

        # Buttons
        btn_frame = tk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 12))

        self._convert_btn = tk.Button(
            btn_frame,
            text="Convert",
            font=("Segoe UI", 11, "bold"),
            bg="#1e3a5f",
            fg="white",
            activebackground="#2a4f7c",
            activeforeground="white",
            relief=tk.FLAT,
            padx=20, pady=6,
            command=self._start_conversion,
        )
        self._convert_btn.pack(side=tk.LEFT, padx=(0, 8))

        tk.Button(
            btn_frame,
            text="Clear Log",
            font=("Segoe UI", 10),
            padx=14, pady=4,
            command=self._clear_log,
        ).pack(side=tk.LEFT)

        tk.Button(
            btn_frame,
            text="Quit",
            font=("Segoe UI", 10),
            padx=14, pady=4,
            command=self.destroy,
        ).pack(side=tk.RIGHT)

    def _detect_tools(self):
        parts = []
        self._libreoffice_path = find_libreoffice()
        if self._libreoffice_path:
            parts.append("LibreOffice: found")
        else:
            parts.append("LibreOffice: not found")

        # Check Word COM availability
        word_available = False
        try:
            import comtypes
            word_available = True
        except ImportError:
            pass
        parts.append("Microsoft Word COM: " + ("available" if word_available else "comtypes not installed"))

        self._tool_status.config(text="  |  ".join(parts))

    def _browse_input(self):
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder:
            self._input_folder.set(folder)
            if not self._output_folder.get():
                self._output_folder.set(folder)

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self._output_folder.set(folder)

    def _log_message(self, msg):
        self._log.config(state=tk.NORMAL)
        self._log.insert(tk.END, msg + "\n")
        self._log.see(tk.END)
        self._log.config(state=tk.DISABLED)

    def _clear_log(self):
        self._log.config(state=tk.NORMAL)
        self._log.delete("1.0", tk.END)
        self._log.config(state=tk.DISABLED)

    def _set_status(self, msg):
        self._status_label.config(text=msg)

    def _start_conversion(self):
        if self._running:
            return

        input_folder = self._input_folder.get().strip()
        output_folder = self._output_folder.get().strip()

        if not input_folder:
            messagebox.showwarning("Missing Input", "Please select an input folder.")
            return
        if not os.path.isdir(input_folder):
            messagebox.showerror("Invalid Input Folder", f"Folder not found:\n{input_folder}")
            return
        if not output_folder:
            messagebox.showwarning("Missing Output", "Please select an output folder.")
            return

        os.makedirs(output_folder, exist_ok=True)

        doc_files = glob.glob(os.path.join(input_folder, "*.doc"))
        doc_files += glob.glob(os.path.join(input_folder, "*.DOC"))
        doc_files = sorted(set(doc_files))

        if not doc_files:
            messagebox.showinfo("No Files Found", "No .doc files found in the input folder.")
            return

        method = self._method.get()
        if method == "auto":
            try:
                import comtypes
                method = "word"
                self._log_message("[INFO] Auto-detected: using Microsoft Word COM.")
            except ImportError:
                if self._libreoffice_path:
                    method = "libreoffice"
                    self._log_message("[INFO] Auto-detected: using LibreOffice.")
                else:
                    messagebox.showerror(
                        "No Conversion Tool Found",
                        "Neither Microsoft Word (comtypes) nor LibreOffice was found.\n\n"
                        "Please install one of the following:\n"
                        "  • Microsoft Word + comtypes (pip install comtypes)\n"
                        "  • LibreOffice (https://www.libreoffice.org)"
                    )
                    return
        elif method == "word":
            try:
                import comtypes
            except ImportError:
                messagebox.showerror(
                    "comtypes Not Installed",
                    "Install comtypes to use Word COM:\n  pip install comtypes"
                )
                return
        elif method == "libreoffice":
            if not self._libreoffice_path:
                messagebox.showerror(
                    "LibreOffice Not Found",
                    "LibreOffice was not found on this system.\n"
                    "Download from: https://www.libreoffice.org"
                )
                return

        self._running = True
        self._convert_btn.config(state=tk.DISABLED)
        self._progress_var.set(0)

        thread = threading.Thread(
            target=self._run_conversion,
            args=(doc_files, output_folder, method),
            daemon=True,
        )
        thread.start()

    def _run_conversion(self, doc_files, output_folder, method):
        total = len(doc_files)
        success_count = 0
        fail_count = 0

        self.after(0, self._log_message, f"Starting conversion of {total} file(s)...")
        self.after(0, self._log_message, f"Output folder: {output_folder}")
        self.after(0, self._log_message, "-" * 60)

        for i, doc_file in enumerate(doc_files, start=1):
            self.after(0, self._set_status, f"Converting {i}/{total}: {Path(doc_file).name}")
            self.after(0, self._progress_var.set, (i - 1) / total * 100)

            if method == "word":
                ok = convert_with_word_com(doc_file, output_folder, lambda m: self.after(0, self._log_message, m))
            else:
                ok = convert_with_libreoffice(self._libreoffice_path, doc_file, output_folder, lambda m: self.after(0, self._log_message, m))

            if ok:
                success_count += 1
            else:
                fail_count += 1

        self.after(0, self._progress_var.set, 100)
        self.after(0, self._log_message, "-" * 60)
        self.after(0, self._log_message, f"Done. {success_count} succeeded, {fail_count} failed out of {total}.")
        self.after(0, self._set_status, f"Finished: {success_count}/{total} converted successfully.")
        self.after(0, self._convert_btn.config, {"state": tk.NORMAL})
        self._running = False


def main():
    app = ConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
