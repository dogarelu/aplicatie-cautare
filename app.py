import tkinter as tk
from tkinter import messagebox, filedialog
import subprocess
import shutil
import sys
from pathlib import Path


# Global variable to store selected root folder
selected_root_folder = None

similar_letters = {
    'a': ['ă', 'â', 'a'],
    'e': ['ĕ', 'e'],
    'i': ['î', 'i'],
    'o': ['ô', 'o'],
    'u': ['û', 'u'],
    's': ['ş', 's'],
    't': ['ţ', 't'],
    'z': ['ţ', 'z'],
    'c': ['ţ', 'c'],
    'd': ['ţ', 'd'],
    'ă': ['î', 'â', 'ă'],
    'î': ['î', 'â', 'ă'],
}


def get_base_path():
    """
    Get the base directory path that works with both script and PyInstaller executable.
    Returns the directory where the executable/script is located.
    For PyInstaller --onefile mode, returns the temp extraction directory (for finding bundled files).
    For PyInstaller --onedir mode, returns the directory containing the executable.
    """
    if getattr(sys, 'frozen', False):
        # Running as compiled executable (PyInstaller)
        # Check if we're in onefile mode (temp extraction directory exists)
        if hasattr(sys, '_MEIPASS'):
            # --onefile mode: bundled files are in temp directory
            return Path(sys._MEIPASS)
        else:
            # --onedir mode: files are next to executable
            return Path(sys.executable).parent
    else:
        # Running as script
        return Path(__file__).parent


def get_output_path():
    """
    Get the directory where output files should be saved.
    Always returns a user-accessible location (executable directory, not temp).
    """
    if getattr(sys, 'frozen', False):
        # Always save to executable directory (user-accessible)
        return Path(sys.executable).parent
    else:
        # Running as script
        return Path(__file__).parent


def on_select_folder():
    """Handle the Select Root Folder button click."""
    global selected_root_folder
    folder = filedialog.askdirectory(title="Select Root Folder")
    if folder:
        selected_root_folder = folder
        folder_label.config(text=f"Selected: {folder}")
    else:
        folder_label.config(text="No folder selected")


def generate_filename_from_search(search_term: str) -> str:
    """Generate filename from search term: first 3 words, hyphenated, with '...' if more."""
    words = search_term.split()
    if len(words) <= 3:
        filename = "-".join(words)
    else:
        filename = "-".join(words[:3]) + "..."
    # Remove any invalid filename characters
    filename = "".join(c for c in filename if c.isalnum() or c in "-._")
    return filename + ".docx"


def on_generate():
    """Handle the Generate Report button click."""
    search_term = entry.get().strip()
    words = search_term.split()

    # Validate search term length (2-12 words)
    if not (2 <= len(words) <= 12):
        messagebox.showerror(
            "Invalid input",
            "Please enter between 2 and 12 words."
        )
        return

    # Check if root folder is selected
    if not selected_root_folder:
        messagebox.showerror(
            "Missing folder",
            "Please select a root folder first."
        )
        return

    try:
        # Disable button during processing
        button.config(state="disabled")
        root.update()  # Update UI to show disabled state
        
        # Get directories (works with both script and PyInstaller)
        base_dir = get_base_path()  # Where to find search_text.py (may be temp dir for --onefile)
        output_dir = get_output_path()  # Where to save output files (always user-accessible)
        
        # Call search_text.py with the search term, root folder, and output directory
        script_path = base_dir / "search_text.py"
        # Also check in output_dir as fallback (for --onedir mode)
        if not script_path.exists():
            script_path = output_dir / "search_text.py"
        if not script_path.exists():
            raise Exception(f"search_text.py not found in: {base_dir} or {output_dir}\nPlease ensure search_text.py is included in the application bundle.")
        
        # Get line span value (0 or empty = entire page, otherwise use the value)
        line_span_value = line_span_entry.get().strip()
        if line_span_value and line_span_value.isdigit() and int(line_span_value) > 0:
            line_span_arg = line_span_value
        else:
            line_span_arg = "0"  # 0 means entire page
        
        result = subprocess.run(
            [sys.executable, str(script_path), search_term, selected_root_folder, str(output_dir), line_span_arg],
            capture_output=True,
            text=True,
            encoding="utf-8",
            cwd=str(output_dir)  # Set working directory to output directory
        )
        
        if result.returncode != 0:
            error_msg = result.stderr or result.stdout or "Unknown error"
            raise Exception(f"Search failed:\n{error_msg}")
        
        # Check if default.docx was created in output directory
        default_docx = output_dir / "default.docx"
        if not default_docx.exists():
            raise Exception("default.docx was not created by search_text.py")
        
        # Generate filename from search term
        suggested_filename = generate_filename_from_search(search_term)
        
        # Open save dialog for user to choose location and filename
        save_path = filedialog.asksaveasfilename(
            title="Save Report",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
            initialfile=suggested_filename
        )
        
        if not save_path:
            # User cancelled
            messagebox.showinfo("Cancelled", "Report generation cancelled.")
            return
        
        # Copy default.docx to the chosen location
        shutil.copy2(default_docx, save_path)
        
        # Show success message
        messagebox.showinfo(
            "Success",
            f"Report created:\n{save_path}"
        )
    except Exception as e:
        # Show user-friendly error dialog
        messagebox.showerror(
            "Error",
            f"Failed to generate report:\n{str(e)}"
        )
    finally:
        # Re-enable button
        button.config(state="normal")


# Create main window
root = tk.Tk()
root.title("Report Generator")
root.geometry("500x280")
root.resizable(False, False)

# Create and pack UI elements
# Root folder selection section
tk.Label(root, text="Root Folder:").pack(padx=10, pady=(10, 5))

folder_frame = tk.Frame(root)
folder_frame.pack(padx=10, pady=5, fill=tk.X)

select_folder_button = tk.Button(
    folder_frame, 
    text="Select Root Folder", 
    command=on_select_folder
)
select_folder_button.pack(side=tk.LEFT, padx=(0, 10))

folder_label = tk.Label(
    folder_frame, 
    text="No folder selected", 
    fg="gray",
    anchor="w"
)
folder_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

# Search term section
tk.Label(root, text="Enter search term:").pack(padx=10, pady=(10, 5))

entry = tk.Entry(root, width=50)
entry.pack(padx=10, pady=5)

# Line span section
line_span_frame = tk.Frame(root)
line_span_frame.pack(padx=10, pady=(5, 10), fill=tk.X)

tk.Label(line_span_frame, text="Line span (0 = entire page):").pack(side=tk.LEFT, padx=(0, 10))

line_span_entry = tk.Spinbox(
    line_span_frame,
    from_=0,
    to=20,
    width=10,
    value=0
)
line_span_entry.pack(side=tk.LEFT)

tk.Label(line_span_frame, text="lines", fg="gray").pack(side=tk.LEFT, padx=(5, 0))

button = tk.Button(root, text="Generate Report", command=on_generate)
button.pack(pady=15)

if __name__ == "__main__":
    root.mainloop()

