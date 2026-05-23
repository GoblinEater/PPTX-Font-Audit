"""
font_audit_app.pyw — Windows drag-and-drop front-end for font_audit.py

Usage:
  1. Drag a .pptx file onto this script (or its shortcut)
  2. Or double-click to open a file picker
  3. Choose where to save the CSV
  4. Done — results appear in a summary dialog

To create a desktop shortcut:
  Right-click font_audit_app.pyw → Send to → Desktop (create shortcut)
  Then drag .pptx files onto the shortcut icon.

Prerequisites:
  pip install python-pptx
"""

import sys
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox


def find_font_audit_script():
    """Look for font_audit.py next to this script."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(script_dir, "font_audit.py")
    if os.path.isfile(candidate):
        return candidate
    return None


def check_dependencies():
    """Verify python-pptx is installed."""
    try:
        import pptx  # noqa: F401
        return True
    except ImportError:
        return False


def run_audit(pptx_path, csv_path, font_audit_script):
    """Run font_audit.py and capture output."""
    result = subprocess.run(
        [sys.executable, font_audit_script, pptx_path, "--csv", csv_path],
        capture_output=True,
        text=True,
    )
    return result.returncode, result.stdout, result.stderr


def main():
    # Hide the root Tk window
    root = tk.Tk()
    root.withdraw()

    # ─── Check dependencies ───
    if not check_dependencies():
        messagebox.showerror(
            "Font Audit",
            "The python-pptx library is not installed.\n\n"
            "Open Command Prompt and run:\n"
            "  pip install python-pptx",
        )
        sys.exit(1)

    # ─── Find font_audit.py ───
    font_audit_script = find_font_audit_script()
    if not font_audit_script:
        messagebox.showerror(
            "Font Audit",
            "font_audit.py not found.\n\n"
            "Place font_audit.py in the same folder as this app.",
        )
        sys.exit(1)

    # ─── Get the PPTX file ───
    if len(sys.argv) > 1:
        pptx_path = sys.argv[1]
    else:
        pptx_path = filedialog.askopenfilename(
            title="Select a PowerPoint file",
            filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")],
        )
        if not pptx_path:
            sys.exit(0)

    # Validate extension
    if not pptx_path.lower().endswith(".pptx"):
        messagebox.showerror(
            "Font Audit",
            "Please select a .pptx file.",
        )
        sys.exit(1)

    # Verify file exists
    if not os.path.isfile(pptx_path):
        messagebox.showerror(
            "Font Audit",
            f"File not found:\n{pptx_path}",
        )
        sys.exit(1)

    # ─── Ask where to save ───
    basename = os.path.splitext(os.path.basename(pptx_path))[0]
    default_name = f"{basename}_font_audit.csv"

    csv_path = filedialog.asksaveasfilename(
        title="Save font audit CSV as",
        defaultextension=".csv",
        initialfile=default_name,
        initialdir=os.path.expanduser("~/Desktop"),
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
    )
    if not csv_path:
        sys.exit(0)

    # ─── Run the audit ───
    exit_code, stdout, stderr = run_audit(pptx_path, csv_path, font_audit_script)

    if exit_code == 0:
        # Parse summary from stdout
        font_count = "?"
        run_count = "?"
        for line in stdout.splitlines():
            if "Unique fonts:" in line:
                font_count = line.strip().split()[-1]
            if "Total text runs:" in line:
                run_count = line.strip().split()[-1]

        result = messagebox.askyesno(
            "Font Audit — Complete",
            f"Font audit complete!\n\n"
            f"{font_count} unique fonts found across {run_count} text runs.\n\n"
            f"CSV saved to:\n{csv_path}\n\n"
            f"Open the CSV now?",
        )
        if result:
            os.startfile(csv_path)
    else:
        error_msg = stderr if stderr else stdout
        messagebox.showerror(
            "Font Audit — Error",
            f"Error running font audit:\n\n{error_msg}",
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
