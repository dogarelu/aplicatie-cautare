import tkinter as tk
from tkinter import messagebox, filedialog
import shutil
import sys
import re
from pathlib import Path
from rapidfuzz import fuzz
from docx import Document


# Global variable to store selected root folder
selected_root_folder = None

# ---------- OCR Constants ----------
# Header de pagină în fișierele OCR (*.txt)
# ex: === PAGE 151 (pag151.jpg) ===
PAGE_HEADER_RE = re.compile(r"^=== PAGE (.+?) \((.+?)\) ===")

# Pagini cu foarte puține cuvinte sunt ignorate (zgomot)
MIN_PAGE_WORDS = 6

# Default threshold pentru matching
DEFAULT_THRESHOLD = 80

# ---------- Text Normalization ----------
# Mapping pentru litere similare (pentru căutare OCR)
# Mapează diacritice la forma de bază pentru matching flexibil
SIMILAR_LETTERS_MAP = {
    # Diacritice românești
    'ș': 's',
    'ț': 't',
    'â': 'a',
    'î': 'a',  # î și â sunt același sunet în română, mapează la 'a' pentru matching (ex: cîmp/câmp -> camp)
    # Diacritice din alte limbi (pentru compatibilitate)
    'á': 'a', 'à': 'a', 'ä': 'a',
    'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
    'í': 'i', 'ì': 'i', 'ï': 'i',
    'ó': 'o', 'ò': 'o', 'ô': 'o', 'ö': 'o',
    'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u',
}

# Reverse mapping: pentru fiecare literă de bază, toate variantele posibile
LETTER_VARIANTS = {
    'a': ['a', 'â', 'î', 'á', 'à', 'ä'],
    'e': ['e', 'é', 'è', 'ê', 'ë'],
    'i': ['i', 'í', 'ì', 'ï'],
    'o': ['o', 'ó', 'ò', 'ô', 'ö'],
    'u': ['u', 'ú', 'ù', 'û', 'ü'],
    's': ['s', 'ș'],
    't': ['t', 'ț'],
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


# ---------- Text Normalization Functions ----------

def normalize_similar_letters(s: str) -> str:
    """
    Normalizează litere similare la forma de bază.
    Ex: ochișorii -> ochisorii, cîmp -> camp, câmp -> camp
    """
    result = []
    for char in s:
        result.append(SIMILAR_LETTERS_MAP.get(char, char))
    return ''.join(result)


def normalize_text(s: str) -> str:
    """
    Normalizare simplă: litere mici, fără punctuație, spații compacte.
    Aplică și normalizarea literelor similare pentru matching flexibil.
    """
    s = s.lower()
    s = re.sub(r"[^\w\săâîșțáéíóúàèìòùâêîôûäëïöü-]", " ", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    s = normalize_similar_letters(s)
    return s


def extract_page_number_from_image(img_name: str) -> str:
    """
    Extract the first full number from image name, removing leading zeros.
    Examples:
        "PPR-102.jpg" -> "102"
        "pag235.jpg" -> "235"
        "IMG_0520.jpg" -> "520"
        "pag172-173.jpg" -> "172"
    """
    if not img_name:
        return ""
    match = re.search(r'\d+', img_name)
    if match:
        number = match.group(0)
        return number.lstrip('0') or '0'
    return ""


# ---------- OCR Loading Functions ----------

def load_ocr_pages(ocr_root: Path):
    """
    Din toate *.txt din subfolderele lui ocr_root extrage paginile OCR.
    Returnează listă de dict cu informații despre pagini.
    """
    pages = []

    for folder in sorted(p for p in ocr_root.iterdir() if p.is_dir()):
        for txt in sorted(folder.glob("*.txt")):
            with txt.open("r", encoding="utf-8") as f:
                current_page_num = None
                current_page_img = None
                buffer = []
                lines_buffer = []

                def flush_page():
                    nonlocal buffer, current_page_num, current_page_img, lines_buffer
                    if current_page_num is None or current_page_img is None:
                        return
                    content = "".join(buffer).strip()
                    norm = normalize_text(content)
                    word_count = len(norm.split())
                    if word_count < MIN_PAGE_WORDS:
                        return
                    clean_lines = [line.strip() for line in lines_buffer if line.strip() and not PAGE_HEADER_RE.match(line.strip())]
                    page_num_from_img = extract_page_number_from_image(current_page_img)
                    pages.append({
                        "folder": folder.name,
                        "txt_file": txt.name,
                        "page_num": page_num_from_img if page_num_from_img else current_page_num,
                        "page_img": current_page_img,
                        "text": content,
                        "lines": clean_lines,
                    })

                for line in f:
                    m = PAGE_HEADER_RE.match(line)
                    if m:
                        flush_page()
                        current_page_num = m.group(1)
                        current_page_img = m.group(2)
                        buffer = []
                        lines_buffer = []
                    else:
                        buffer.append(line)
                        lines_buffer.append(line)
                flush_page()

    return pages


# ---------- Search Functions ----------

def find_match_context(query_text: str, page_lines: list, before: int = 2, after: int = 2):
    """
    Găsește poziția match-ului în linii și returnează contextul.
    """
    norm_query = normalize_text(query_text)
    query_words = norm_query.split()
    
    best_line_idx = None
    best_match_score = 0
    
    for i, line in enumerate(page_lines):
        norm_line = normalize_text(line)
        if not norm_line:
            continue
        words_found = sum(1 for word in query_words if word in norm_line)
        if words_found > 0:
            score = fuzz.partial_ratio(norm_query, norm_line)
            if score > best_match_score:
                best_match_score = score
                best_line_idx = i
    
    if best_line_idx is None:
        return page_lines[:before + after + 1] if page_lines else []
    
    start_idx = max(0, best_line_idx - before)
    end_idx = min(len(page_lines), best_line_idx + after + 1)
    
    return page_lines[start_idx:end_idx]


def check_words_in_line_span(query_words: list, page_lines: list, line_span: int) -> tuple:
    """
    Verifică dacă toate cuvintele din query apar într-un span de linii.
    Returnează (found, score, start_idx, end_idx).
    """
    if not page_lines or not query_words:
        return (False, 0, 0, 0)
    
    norm_lines = [normalize_text(line) for line in page_lines if line.strip()]
    if not norm_lines:
        return (False, 0, 0, 0)
    
    best_score = 0
    best_start = 0
    best_end = 0
    found_all = False
    
    for start_idx in range(len(norm_lines)):
        end_idx = min(start_idx + line_span, len(norm_lines))
        span_lines = norm_lines[start_idx:end_idx]
        span_text = " ".join(span_lines)
        
        words_found = sum(1 for word in query_words if word in span_text)
        word_coverage = (words_found / len(query_words)) * 100 if query_words else 0
        query_text = " ".join(query_words)
        fuzz_score = fuzz.partial_ratio(query_text, span_text)
        
        if words_found == len(query_words):
            score = max(fuzz_score, word_coverage * 0.9)
            if score > best_score:
                best_score = score
                best_start = start_idx
                best_end = end_idx
                found_all = True
        elif words_found >= len(query_words) * 0.8:
            score = (fuzz_score + word_coverage) / 2
            if score > best_score:
                best_score = score
                best_start = start_idx
                best_end = end_idx
        else:
            score = fuzz_score * (words_found / len(query_words))
            if score > best_score and not found_all:
                best_score = score
                best_start = start_idx
                best_end = end_idx
    
    return (found_all, best_score, best_start, best_end)


def search_text_in_pages(query_text: str, pages, threshold: int = DEFAULT_THRESHOLD, line_span: int = None):
    """
    Caută query_text în toate paginile și returnează toate match-urile cu score >= threshold.
    """
    norm_query = normalize_text(query_text)
    query_words = norm_query.split()
    
    if len(query_words) < 2:
        raise ValueError("Query must have at least 2 words")
    if len(query_words) > 12:
        raise ValueError("Query must have at most 12 words")
    
    matches = []
    
    for page in pages:
        page_lines = page.get("lines", [])
        
        if line_span is not None and line_span > 0 and page_lines:
            found, score, start_idx, end_idx = check_words_in_line_span(query_words, page_lines, line_span)
            
            if found and score >= threshold:
                snippet_start = max(0, start_idx - 2)
                snippet_end = min(len(page_lines), end_idx + 2)
                snippet = page_lines[snippet_start:snippet_end]
                
                matches.append({
                    "folder": page["folder"],
                    "image": page["page_img"],
                    "score": round(score, 1),
                    "page_num": page["page_num"],
                    "snippet": snippet
                })
            continue
        
        norm_page = normalize_text(page["text"])
        if not norm_page:
            continue
        
        score = fuzz.partial_ratio(norm_query, norm_page)
        words_found = sum(1 for word in query_words if word in norm_page)
        word_coverage = (words_found / len(query_words)) * 100 if query_words else 0
        
        if words_found == len(query_words):
            final_score = max(score, word_coverage * 0.9)
        elif words_found >= len(query_words) * 0.8:
            final_score = (score + word_coverage) / 2
        else:
            final_score = score * (words_found / len(query_words))
        
        if final_score >= threshold:
            snippet = find_match_context(query_text, page_lines)
            
            matches.append({
                "folder": page["folder"],
                "image": page["page_img"],
                "score": round(final_score, 1),
                "page_num": page["page_num"],
                "snippet": snippet
            })
    
    matches.sort(key=lambda x: x["score"], reverse=True)
    return matches


def run_search_and_generate_report(query_text: str, ocr_root: Path, output_dir: Path, line_span: int = None, threshold: int = DEFAULT_THRESHOLD):
    """
    Run the search and generate the report document.
    This replaces the subprocess call to search_text.py
    """
    if not ocr_root.is_dir():
        raise Exception(f"OCR root folder not found: {ocr_root}")
    
    if threshold < 0 or threshold > 100:
        raise Exception(f"Threshold must be between 0 and 100, got: {threshold}")
    
    # Load OCR pages
    pages = load_ocr_pages(ocr_root)
    
    # Search
    matches = search_text_in_pages(query_text, pages, threshold, line_span=line_span)
    
    # Write results to text file
    output_file = output_dir / "search-result.txt"
    with output_file.open("w", encoding="utf-8") as f:
        if not matches:
            f.write(f"No matches found above threshold {threshold}\n")
        else:
            for match in matches:
                for snippet_line in match['snippet']:
                    if snippet_line.strip():
                        f.write(snippet_line + "\n")
                f.write(f"({match['folder']}, p. {match['page_num']})\n")
                f.write("\n")
    
    # Write results to Word document
    docx_file = None
    
    # Check if running as PyInstaller executable and look for bundled default.docx
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        bundled_docx = Path(sys._MEIPASS) / "default.docx"
        if bundled_docx.exists():
            docx_file = bundled_docx
    
    # If not found in bundle, check output_dir
    if docx_file is None or not docx_file.exists():
        docx_file = output_dir / "default.docx"
    
    # Load template if it exists, otherwise create new document
    if docx_file.exists():
        doc = Document(str(docx_file))
        # Clear existing content (keep styles)
        for para in list(doc.paragraphs):
            p = para._element
            p.getparent().remove(p)
    else:
        doc = Document()
    
    # Add title
    title_para = doc.add_paragraph()
    title_run = title_para.add_run('Search results for "')
    search_run = title_para.add_run(query_text)
    search_run.italic = True
    title_para.add_run('"')
    doc.add_paragraph()
    
    # Add matches
    if not matches:
        doc.add_paragraph(f"No matches found above threshold {threshold}")
    else:
        for match in matches:
            for snippet_line in match['snippet']:
                if snippet_line.strip():
                    para = doc.add_paragraph(snippet_line.strip())
                    try:
                        para.style = "2-Versuri-centru"
                    except KeyError:
                        pass
            source_text = f"({match['folder']}, p. {match['page_num']})"
            para = doc.add_paragraph(source_text)
            try:
                para.style = "4-Sursa text"
            except KeyError:
                pass
            doc.add_paragraph()
    
    # Save to output_dir
    doc.save(str(output_dir / "default.docx"))


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
        
        # Get output directory
        output_dir = get_output_path()
        
        # Get line span value (0 or empty = entire page, otherwise use the value)
        line_span_value = line_span_entry.get().strip()
        if line_span_value and line_span_value.isdigit() and int(line_span_value) > 0:
            line_span = int(line_span_value)
        else:
            line_span = None  # None means entire page
        
        # Run search and generate report directly
        run_search_and_generate_report(
            query_text=search_term,
            ocr_root=Path(selected_root_folder),
            output_dir=output_dir,
            line_span=line_span,
            threshold=DEFAULT_THRESHOLD
        )
        
        # Check if default.docx was created
        default_docx = output_dir / "default.docx"
        if not default_docx.exists():
            raise Exception("default.docx was not created")
        
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

