import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import shutil
import sys
import re
import subprocess
import platform
import requests
import time
from typing import List, Set, Dict
from pathlib import Path
from rapidfuzz import fuzz
from rapidfuzz.distance import Levenshtein
from docx import Document


# Global variable to store selected root folder and selected paths
selected_root_folder = None
selected_paths = []  # List of selected paths (Biblioteca, Rafturi, or Volume paths)
selected_volumes = []  # List of selected "volumes" (final folders with ocr.txt)

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


def words_match_simple(query_word: str, text_word: str) -> bool:
    """
    Simple word matching without hyphen handling (used internally).
    """
    if not query_word or not text_word:
        return False
    
    # Exact match always works
    if query_word == text_word:
        return True
    
    # For words 3 characters or less, require exact match
    if len(query_word) <= 3 or len(text_word) <= 3:
        return False
    
    # Calculate Levenshtein distance using rapidfuzz
    lev_distance = Levenshtein.distance(query_word, text_word)
    
    # Determine allowed distance based on query word length
    # Use the longer of the two words to determine the threshold
    max_length = max(len(query_word), len(text_word))
    
    # For words 4-7 characters, allow 1 character difference
    if max_length >= 4:
        return lev_distance <= 1
    
    # Fallback (shouldn't reach here, but handle it)
    return lev_distance <= 1


def words_match(query_word: str, text_word: str) -> bool:
    """
    Check if two words match according to fuzzy matching rules:
    - If word length <= 3: must be exact match
    - If word length 4-7: allow 1 character difference (Levenshtein distance <= 1)
    - If word length >= 8: allow 2 character differences (Levenshtein distance <= 2)
    - Similar letters/variants don't count as different (already normalized)
    - Handles hyphenated words: matches parts before/after hyphens
    
    Args:
        query_word: The search word (normalized)
        text_word: The word from text to match against (normalized)
    
    Returns:
        True if words match according to the rules
    """
    # Both words should already be normalized (similar letters handled)
    if not query_word or not text_word:
        return False
    
    # Exact match always works
    if query_word == text_word:
        return True
    
    # Handle hyphenated words: check if query matches any part of hyphenated text word
    # e.g., "calu" should match "calu-m" (part before hyphen)
    if '-' in text_word:
        text_parts = text_word.split('-')
        for part in text_parts:
            if part and words_match_simple(query_word, part):
                return True
    
    # Handle hyphenated query: check if any part of hyphenated query matches text word
    # e.g., "intr-o" should match "intro"
    if '-' in query_word:
        query_parts = query_word.split('-')
        for part in query_parts:
            if part and words_match_simple(part, text_word):
                return True
    
    # Also check if removing hyphens makes them match
    # e.g., "intro" should match "intr-o"
    text_word_no_hyphen = text_word.replace('-', '')
    query_word_no_hyphen = query_word.replace('-', '')
    if text_word_no_hyphen and query_word_no_hyphen:
        if words_match_simple(query_word_no_hyphen, text_word_no_hyphen):
            return True
    
    # Standard matching for non-hyphenated words
    return words_match_simple(query_word, text_word)


def get_inflected_forms_from_dex(word: str, cache: Dict[str, Set[str]] = None) -> Set[str]:
    """
    Obține formele flexionare ale unui cuvânt folosind API-ul DEX online.
    
    Args:
        word: Cuvântul pentru care se caută forme flexionare
        cache: Dicționar pentru cache (opțional, pentru a evita cereri duplicate)
    
    Returns:
        Set de forme flexionare (include și cuvântul original)
    """
    if cache is None:
        cache = {}
    
    word_lower = word.lower().strip()
    
    # Verifică cache-ul
    if word_lower in cache:
        return cache[word_lower]
    
    forms = {word_lower}  # Include forma originală
    
    try:
        # Endpoint API DEX online pentru lexem
        url = f"https://dexonline.ro/api/lexem/{word_lower}"
        
        response = requests.get(url, timeout=5)
        
        if response.status_code == 200:
            data = response.json()
            
            # Extrage formele flexionare din răspuns
            # Structura poate varia - ajustează în funcție de răspunsul real
            if isinstance(data, dict):
                # Caută câmpuri care conțin forme flexionare
                if 'flexiuni' in data:
                    flexiuni = data['flexiuni']
                    if isinstance(flexiuni, list):
                        for form in flexiuni:
                            if isinstance(form, str):
                                forms.add(form.lower())
                            elif isinstance(form, dict) and 'form' in form:
                                forms.add(form['form'].lower())
                
                # Poate fi și în alt format
                if 'forms' in data:
                    forms_data = data['forms']
                    if isinstance(forms_data, list):
                        for form in forms_data:
                            if isinstance(form, str):
                                forms.add(form.lower())
                            elif isinstance(form, dict) and 'form' in form:
                                forms.add(form['form'].lower())
                
                # Sau în definiții
                if 'definitions' in data:
                    for def_item in data['definitions']:
                        if isinstance(def_item, dict) and 'flexiuni' in def_item:
                            flexiuni = def_item['flexiuni']
                            if isinstance(flexiuni, list):
                                for form in flexiuni:
                                    if isinstance(form, str):
                                        forms.add(form.lower())
                
                # Caută și în structura de lexem
                if 'lexem' in data:
                    lexem_data = data['lexem']
                    if isinstance(lexem_data, dict) and 'flexiuni' in lexem_data:
                        flexiuni = lexem_data['flexiuni']
                        if isinstance(flexiuni, list):
                            for form in flexiuni:
                                if isinstance(form, str):
                                    forms.add(form.lower())
            
        # Rate limiting - așteaptă puțin între cereri
        time.sleep(0.1)
        
    except requests.exceptions.RequestException as e:
        # Dacă API-ul nu funcționează, returnează doar forma originală
        pass  # Fail silently - use original word only
    except Exception as e:
        pass  # Fail silently - use original word only
    
    # Salvează în cache
    cache[word_lower] = forms
    
    return forms


def expand_search_terms_with_inflections(query_text: str, cache: Dict[str, Set[str]] = None) -> List[str]:
    """
    Extinde termenii de căutare cu formele lor flexionare.
    
    Args:
        query_text: Textul de căutare original
        cache: Cache pentru forme flexionare (opțional)
    
    Returns:
        Listă de cuvinte extinse (original + forme flexionare)
    """
    words = query_text.split()
    expanded_words = []
    
    for word in words:
        # Obține forme flexionare
        inflected_forms = get_inflected_forms_from_dex(word, cache)
        expanded_words.extend(inflected_forms)
    
    # Returnează lista unică, păstrând ordinea
    seen = set()
    result = []
    for word in expanded_words:
        if word not in seen:
            seen.add(word)
            result.append(word)
    
    return result


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

def load_ocr_pages(ocr_root: Path, selected_volumes: list = None):
    """
    Din toate *.txt din folderele selectate extrage paginile OCR.
    Returnează listă de dict cu informații despre pagini.
    selected_volumes: listă de căi către folderele finale (volumes) care conțin ocr.txt
    """
    pages = []

    # Dacă nu sunt volume selectate, nu căuta nimic
    if not selected_volumes:
        return pages

    # Încarcă paginile din fiecare volume selectat
    for volume_path_str in selected_volumes:
        volume_path = Path(volume_path_str)
        if volume_path.exists() and folder_has_ocr(volume_path):
            _load_pages_from_folder(volume_path, pages)

    return pages


def _search_recursive(folder: Path, pages: list):
    """Recursively search for OCR files in folder and subfolders.
    If a folder has OCR, it's an end folder - load from it and don't search subfolders."""
    # Check if this folder has OCR files
    if folder_has_ocr(folder):
        _load_pages_from_folder(folder, pages)
        # Don't search in subfolders if this folder has OCR (it's an end folder)
        return
    
    # No OCR in this folder, search in subfolders
    for subfolder in sorted(p for p in folder.iterdir() if p.is_dir()):
        _search_recursive(subfolder, pages)


def folder_has_ocr(folder: Path) -> bool:
    """Check if a folder contains ocr.txt files (marks it as end folder)."""
    # Check for ocr.txt specifically, or any .txt files
    return (folder / "ocr.txt").exists() or any(folder.glob("*.txt"))


def _load_pages_from_folder(folder: Path, pages: list):
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
                    
                    # Check if image file exists
                    image_path = folder / current_page_img
                    image_exists = image_path.exists()
                    if not image_exists:
                        # Try to find it in subdirectories
                        for img_file in folder.rglob(current_page_img):
                            if img_file.exists():
                                image_exists = True
                                break
                    
                    pages.append({
                        "folder": folder.name,
                        "volume_path": str(folder),  # Full path to volume folder
                        "txt_file": txt.name,
                        "page_num": page_num_from_img if page_num_from_img else current_page_num,
                        "page_img": current_page_img,
                        "image_exists": image_exists,  # Store whether image exists
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
        norm_line_words = norm_line.split()
        # Ensure each query word matches a different text word
        matched_word_indices = set()
        words_found = 0
        for query_word in query_words:
            for idx, text_word in enumerate(norm_line_words):
                if idx not in matched_word_indices and words_match(query_word, text_word):
                    matched_word_indices.add(idx)
                    words_found += 1
                    break
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


def check_words_in_word_span(query_words: list, page_lines: list, word_span: int, exact_order: bool = True) -> tuple:
    """
    Verifică dacă toate cuvintele din query apar într-un span de cuvinte.
    exact_order: True = cuvintele trebuie să apară în ordinea exactă, False = orice ordine
    Returnează (found, score, start_line_idx, end_line_idx, matched_words).
    matched_words: lista cu cuvintele exacte din fragmentul care s-a potrivit (din liniile originale, nu normalizate)
    """
    if not page_lines or not query_words:
        return (False, 0, 0, 0, [])
    
    # Keep original lines for the matched fragment
    original_lines = [line.strip() for line in page_lines if line.strip()]
    norm_lines = [normalize_text(line) for line in original_lines]
    if not norm_lines:
        return (False, 0, 0, 0, [])
    
    # Convert lines to a list of words with their line indices and original word text
    all_words = []  # (normalized_word, line_idx, original_word, original_line)
    for line_idx, (norm_line, orig_line) in enumerate(zip(norm_lines, original_lines)):
        norm_words = norm_line.split()
        orig_words = orig_line.split()
        # Match normalized words with original words (they should have same count)
        for i, norm_word in enumerate(norm_words):
            orig_word = orig_words[i] if i < len(orig_words) else norm_word
            all_words.append((norm_word, line_idx, orig_word, orig_line))
    
    if not all_words:
        return (False, 0, 0, 0, [])
    
    best_score = 0
    best_start_line = 0
    best_end_line = 0
    best_matched_words = []
    found_all = False
    
    # Slide a window of word_span words across the text
    for start_word_idx in range(len(all_words)):
        end_word_idx = min(start_word_idx + word_span, len(all_words))
        span_words_data = all_words[start_word_idx:end_word_idx]
        span_words = [w[0] for w in span_words_data]  # normalized words
        span_text = " ".join(span_words)
        
        # Get line indices for the span
        if span_words_data:
            start_line_idx = span_words_data[0][1]
            end_line_idx = span_words_data[-1][1] + 1
        else:
            start_line_idx = 0
            end_line_idx = 0
        
        # Check word order if exact_order is True
        if exact_order:
            # Check if words appear in exact order
            # Track which span word indices have been matched to ensure each query word matches a different text word
            matched_indices = set()
            word_indices = []
            for query_word in query_words:
                # Find matching word in span_words using fuzzy matching
                found_idx = None
                for idx, span_word in enumerate(span_words):
                    # Only match if this word hasn't been matched yet and it matches the query word
                    if idx not in matched_indices and words_match(query_word, span_word):
                        found_idx = idx
                        matched_indices.add(idx)  # Mark this word as used
                        break
                
                if found_idx is not None:
                    word_indices.append(found_idx)
                else:
                    # Word not found
                    word_indices = None
                    break
            
            if word_indices and len(word_indices) == len(query_words):
                # Check if indices are in ascending order (exact order)
                if all(word_indices[i] <= word_indices[i+1] for i in range(len(word_indices)-1)):
                    words_found = len(query_words)
                else:
                    words_found = 0
            else:
                words_found = 0
        else:
            # Random order: check if all words are present, but each query word must match a different text word
            matched_indices = set()
            words_found = 0
            for query_word in query_words:
                found = False
                for idx, span_word in enumerate(span_words):
                    # Only match if this word hasn't been matched yet and it matches the query word
                    if idx not in matched_indices and words_match(query_word, span_word):
                        matched_indices.add(idx)  # Mark this word as used
                        words_found += 1
                        found = True
                        break
                if not found:
                    # This query word couldn't be matched to a unique text word
                    break
        
        if words_found == 0:
            continue
        
        word_coverage = (words_found / len(query_words)) * 100 if query_words else 0
        query_text = " ".join(query_words)
        fuzz_score = fuzz.partial_ratio(query_text, span_text)
        
        if words_found == len(query_words):
            score = max(fuzz_score, word_coverage * 0.9)
            if score > best_score:
                best_score = score
                best_start_line = start_line_idx
                best_end_line = end_line_idx
                # Store the exact matched words (original, not normalized)
                best_matched_words = [w[2] for w in span_words_data]
                found_all = True
        elif words_found >= len(query_words) * 0.8:
            score = (fuzz_score + word_coverage) / 2
            if score > best_score:
                best_score = score
                best_start_line = start_line_idx
                best_end_line = end_line_idx
                best_matched_words = [w[2] for w in span_words_data]
        else:
            score = fuzz_score * (words_found / len(query_words))
            if score > best_score and not found_all:
                best_score = score
                best_start_line = start_line_idx
                best_end_line = end_line_idx
                best_matched_words = [w[2] for w in span_words_data]
    
    return (found_all, best_score, best_start_line, best_end_line, best_matched_words)


def search_text_in_pages(query_text: str, pages, threshold: int = DEFAULT_THRESHOLD, word_span: int = None, exact_order: bool = True, use_inflections: bool = True):
    """
    Caută query_text în toate paginile și returnează toate match-urile cu score >= threshold.
    word_span: numărul de cuvinte în care să caute (None = întreaga pagină)
    exact_order: True = cuvintele trebuie să apară în ordinea exactă, False = orice ordine
    use_inflections: Dacă True, extinde căutarea cu forme flexionare din DEX online
    """
    # Cache pentru forme flexionare (partajat între apeluri)
    if not hasattr(search_text_in_pages, 'inflection_cache'):
        search_text_in_pages.inflection_cache = {}
    
    # Extinde cuvintele cu forme flexionare dacă este activat
    if use_inflections:
        expanded_words = expand_search_terms_with_inflections(
            query_text, 
            search_text_in_pages.inflection_cache
        )
        # Creează un query extins pentru normalizare
        expanded_query = " ".join(expanded_words)
        norm_query = normalize_text(expanded_query)
        query_words = norm_query.split()
    else:
        norm_query = normalize_text(query_text)
        query_words = norm_query.split()
    
    if len(query_words) < 1:
        raise ValueError("Query must have at least 1 word")
    if len(query_words) > 10:
        raise ValueError("Query must have at most 10 words")
    
    matches = []
    
    for page in pages:
        page_lines = page.get("lines", [])
        
        if word_span is not None and word_span > 0 and page_lines:
            found, score, start_line_idx, end_line_idx, matched_words = check_words_in_word_span(query_words, page_lines, word_span, exact_order)
            
            if found and score >= threshold:
                # Use the exact matched fragment
                matched_fragment = " ".join(matched_words) if matched_words else ""
                
                matches.append({
                    "folder": page["folder"],
                    "volume_path": page.get("volume_path", ""),
                    "image": page["page_img"],
                    "score": round(score, 1),
                    "page_num": page["page_num"],
                    "snippet": [matched_fragment],  # Store as list for consistency
                    "matched_fragment": matched_fragment,  # Exact matched fragment
                    "page_lines": page_lines,  # Store full page lines for context
                    "match_start_line_idx": start_line_idx,  # Line index where match starts
                    "match_end_line_idx": end_line_idx  # Line index where match ends
                })
            continue
        
        norm_page = normalize_text(page["text"])
        if not norm_page:
            continue
        
        score = fuzz.partial_ratio(norm_query, norm_page)
        norm_page_words = norm_page.split()
        # Ensure each query word matches a different text word
        matched_word_indices = set()
        words_found = 0
        for query_word in query_words:
            for idx, text_word in enumerate(norm_page_words):
                if idx not in matched_word_indices and words_match(query_word, text_word):
                    matched_word_indices.add(idx)
                    words_found += 1
                    break
        word_coverage = (words_found / len(query_words)) * 100 if query_words else 0
        
        if words_found == len(query_words):
            final_score = max(score, word_coverage * 0.9)
        elif words_found >= len(query_words) * 0.8:
            final_score = (score + word_coverage) / 2
        else:
            final_score = score * (words_found / len(query_words))
        
        if final_score >= threshold:
            snippet = find_match_context(query_text, page_lines)
            # For full page search, extract the best matching fragment
            # Find the line with the best match and extract words around it
            snippet_text = " ".join(snippet)
            snippet_words = snippet_text.split()
            
            # Find where search words appear in snippet
            norm_snippet_words = [normalize_text(w) for w in snippet_words]
            norm_query_words = [normalize_text(w) for w in query_words]
            
            match_start_idx = None
            for i, norm_word in enumerate(norm_snippet_words):
                for norm_query_word in norm_query_words:
                    if words_match(norm_query_word, norm_word):
                        match_start_idx = i
                        break
                if match_start_idx is not None:
                    break
            
            # Find the line index in page_lines where the match occurs
            best_line_idx = None
            best_match_score = 0
            for i, line in enumerate(page_lines):
                norm_line = normalize_text(line)
                if not norm_line:
                    continue
                norm_line_words = norm_line.split()
                # Ensure each query word matches a different text word
                matched_word_indices = set()
                words_found = 0
                for query_word in query_words:
                    for idx, text_word in enumerate(norm_line_words):
                        if idx not in matched_word_indices and words_match(query_word, text_word):
                            matched_word_indices.add(idx)
                            words_found += 1
                            break
                if words_found > 0:
                    score = fuzz.partial_ratio(norm_query, norm_line)
                    if score > best_match_score:
                        best_match_score = score
                        best_line_idx = i
            
            if match_start_idx is not None:
                # Extract a window around the match (use word_span if available, otherwise 20 words)
                fragment_size = word_span if word_span else 20
                half_size = fragment_size // 2
                start_idx = max(0, match_start_idx - half_size)
                end_idx = min(len(snippet_words), start_idx + fragment_size)
                if end_idx - start_idx < fragment_size and start_idx > 0:
                    start_idx = max(0, end_idx - fragment_size)
                matched_fragment = " ".join(snippet_words[start_idx:end_idx])
            else:
                # Fallback: use first fragment_size words
                fragment_size = word_span if word_span else 20
                matched_fragment = " ".join(snippet_words[:fragment_size])
            
            matches.append({
                "folder": page["folder"],
                "volume_path": page.get("volume_path", ""),
                "image": page["page_img"],
                "score": round(final_score, 1),
                "page_num": page["page_num"],
                "snippet": snippet,
                "matched_fragment": matched_fragment,  # Exact matched fragment
                "page_lines": page_lines,  # Store full page lines for context
                "match_start_line_idx": best_line_idx if best_line_idx is not None else 0,  # Line index where match occurs
                "match_end_line_idx": best_line_idx + 1 if best_line_idx is not None else 1  # Line index where match ends
            })
    
    matches.sort(key=lambda x: x["score"], reverse=True)
    return matches


def run_search_only(query_text: str, ocr_root: Path, word_span: int = None, threshold: int = DEFAULT_THRESHOLD, selected_volumes: list = None, exact_order: bool = True):
    """
    Run the search and return matches without generating documents.
    selected_volumes: list of paths to volumes (folders with ocr.txt) to search
    word_span: number of words to search within (None = entire page)
    exact_order: True = words must appear in exact order, False = any order
    Returns: list of match dictionaries
    """
    if not ocr_root.is_dir():
        raise Exception(f"OCR root folder not found: {ocr_root}")
    
    if threshold < 0 or threshold > 100:
        raise Exception(f"Threshold must be between 0 and 100, got: {threshold}")
    
    # Load OCR pages (only from selected volumes)
    pages = load_ocr_pages(ocr_root, selected_volumes)
    
    if not pages:
        raise Exception("No pages found in selected folders. Please check your selection.")
    
    # Search
    matches = search_text_in_pages(query_text, pages, threshold, word_span=word_span, exact_order=exact_order)
    
    return matches


def run_search_and_generate_report(query_text: str, ocr_root: Path, output_dir: Path, word_span: int = None, threshold: int = DEFAULT_THRESHOLD, selected_volumes: list = None, exact_order: bool = True):
    """
    Run the search and generate the report document.
    This replaces the subprocess call to search_text.py
    NOTE: This function is kept for future use but is not currently called.
    selected_volumes: list of paths to volumes (folders with ocr.txt) to search
    word_span: number of words to search within (None = entire page)
    exact_order: True = words must appear in exact order, False = any order
    """
    if not ocr_root.is_dir():
        raise Exception(f"OCR root folder not found: {ocr_root}")
    
    if threshold < 0 or threshold > 100:
        raise Exception(f"Threshold must be between 0 and 100, got: {threshold}")
    
    # Load OCR pages (only from selected volumes)
    pages = load_ocr_pages(ocr_root, selected_volumes)
    
    if not pages:
        raise Exception("No pages found in selected folders. Please check your selection.")
    
    # Search
    matches = search_text_in_pages(query_text, pages, threshold, word_span=word_span, exact_order=exact_order)
    
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


def open_image(image_path: Path):
    """Open an image file using the system's default image viewer."""
    if not image_path.exists():
        messagebox.showerror("Error", f"Image not found: {image_path}")
        return
    
    try:
        system = platform.system()
        if system == "Darwin":  # macOS
            subprocess.run(["open", str(image_path)])
        elif system == "Windows":
            import os
            os.startfile(str(image_path))
        else:  # Linux and others
            subprocess.run(["xdg-open", str(image_path)])
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open image: {str(e)}")


def format_quote_with_bold(quote_text: str, search_words: list) -> str:
    """
    Format quote text with search words marked for display.
    Returns the text with search words wrapped in markers for bold formatting.
    For now, we'll use a simple approach and mark words with ** for display.
    """
    if not quote_text or not search_words:
        return quote_text
    
    # Normalize search words
    norm_search_words = [normalize_text(word) for word in search_words]
    
    # Split quote into words and mark matches
    words = quote_text.split()
    result_words = []
    
    for word in words:
        norm_word = normalize_text(word)
        # Check if this word matches any search word
        is_match = any(search_word in norm_word or norm_word in search_word for search_word in norm_search_words)
        
        if is_match:
            # Mark for bold (we'll use a special marker that we'll replace in the UI)
            result_words.append(f"**{word}**")
        else:
            result_words.append(word)
    
    return " ".join(result_words)


def find_image_path(volume_path: Path, image_filename: str) -> Path:
    """Find the full path to an image file given the volume path and image filename."""
    # Image should be in the same folder as the OCR txt file
    image_path = volume_path / image_filename
    if image_path.exists():
        return image_path
    
    # Try to find it in subdirectories
    for img_file in volume_path.rglob(image_filename):
        if img_file.exists():
            return img_file
    
    return image_path  # Return even if not found, let open_image handle the error


def on_select_folder():
    """Handle the Select Biblioteca button click."""
    global selected_root_folder, selected_paths, selected_volumes
    folder = filedialog.askdirectory(title="Select Biblioteca Folder")
    if folder:
        selected_root_folder = folder
        selected_paths = []  # Reset selections
        selected_volumes = []  # Reset volumes
        folder_label.config(text=f"Biblioteca: {Path(folder).name}")
        populate_library_tree(folder)
    else:
        folder_label.config(text="No Biblioteca selected")
        library_tree.delete(*library_tree.get_children())
        selected_volumes = []  # Clear volumes


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


def create_results_window():
    """Create or show the results window."""
    global results_window, results_scrollable_frame, results_canvas, export_button_frame, results_title_label
    
    # If window exists, just bring it to front
    if results_window is not None and results_window.winfo_exists():
        results_window.lift()
        results_window.focus()
        return
    
    # Create new results window
    results_window = tk.Toplevel(root)
    results_window.title("Rezultate Căutare")
    results_window.geometry("1400x800")
    results_window.configure(bg=COLOR_BACKGROUND)
    
    # Header
    header_frame = tk.Frame(results_window, bg=COLOR_BACKGROUND, pady=10)
    header_frame.pack(fill=tk.X, padx=15)
    
    results_title_label = tk.Label(
        header_frame,
        text="📋 Rezultate Căutare",
        font=("Arial", 44, "bold"),
        bg=COLOR_BACKGROUND,
        fg=COLOR_TEXT
    )
    results_title_label.pack(anchor="w")
    
    # Results container
    results_container = tk.Frame(results_window, bg=COLOR_BACKGROUND)
    results_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
    
    # Create scrollable frame for results table
    results_canvas = tk.Canvas(results_container, bg=COLOR_BACKGROUND, highlightthickness=0)
    results_scrollbar = ttk.Scrollbar(results_container, orient="vertical", command=results_canvas.yview)
    results_scrollable_frame = tk.Frame(results_canvas, bg=COLOR_BACKGROUND)
    
    results_scrollable_frame.bind(
        "<Configure>",
        lambda e: results_canvas.configure(scrollregion=results_canvas.bbox("all"))
    )
    
    results_canvas.create_window((0, 0), window=results_scrollable_frame, anchor="nw")
    results_canvas.configure(yscrollcommand=results_scrollbar.set)
    
    results_canvas.pack(side="left", fill="both", expand=True)
    results_scrollbar.pack(side="right", fill="y")
    
    # Bind mousewheel to canvas
    def _on_mousewheel(event):
        results_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    results_canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    # Update canvas width when scrollable frame changes
    def configure_canvas_width(event):
        canvas_width = event.width
        results_canvas.itemconfig(results_canvas.find_all()[0], width=canvas_width)
    
    results_scrollable_frame.bind("<Configure>", lambda e: results_canvas.configure(scrollregion=results_canvas.bbox("all")))
    results_canvas.bind("<Configure>", configure_canvas_width)
    
    # Export button frame at the bottom
    export_button_frame = tk.Frame(results_window, bg=COLOR_BACKGROUND, pady=10)
    export_button_frame.pack(fill=tk.X, padx=15, pady=(0, 10))


def display_results_in_table(matches: list, search_term: str, word_span: int = 10):
    """Display search results in a table format.
    word_span: number of words to show in the quote fragment."""
    global results_title_label
    
    # Create or show results window
    create_results_window()
    
    # Update title with result count
    num_results = len(matches)
    results_title_label.config(text=f"📋 Rezultate Căutare ({num_results})")
    
    # Clear existing results
    for widget in results_scrollable_frame.winfo_children():
        widget.destroy()
    
    if not matches:
        no_results_label = tk.Label(
            results_scrollable_frame,
            text="Nu s-au găsit rezultate.",
            font=("Arial", 30),
            bg=COLOR_BACKGROUND,
            fg=COLOR_TEXT
        )
        no_results_label.pack(pady=10)
        return
    
    # Clear match checkboxes list
    global match_checkboxes
    match_checkboxes = []
    
    # Create header row
    header_frame = tk.Frame(results_scrollable_frame, bg=COLOR_SHELF, relief=tk.RAISED, bd=2)
    header_frame.pack(fill=tk.X, pady=(0, 5), padx=5)
    
    # Checkbox column header
    tk.Label(header_frame, text="", font=("Arial", 32, "bold"), bg=COLOR_SHELF, fg="white", width=3).pack(side=tk.LEFT, padx=5, pady=5)
    tk.Label(header_frame, text="Volume", font=("Arial", 32, "bold"), bg=COLOR_SHELF, fg="white", width=15).pack(side=tk.LEFT, padx=5, pady=5)
    tk.Label(header_frame, text="", font=("Arial", 32, "bold"), bg=COLOR_SHELF, fg="white", width=8).pack(side=tk.LEFT, padx=5, pady=5)  # Space for button
    tk.Label(header_frame, text="Pagina", font=("Arial", 32, "bold"), bg=COLOR_SHELF, fg="white", width=10).pack(side=tk.LEFT, padx=5, pady=5)
    tk.Label(header_frame, text="Citat", font=("Arial", 32, "bold"), bg=COLOR_SHELF, fg="white", width=60).pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
    
    # Get search words for highlighting
    search_words = search_term.split()
    
    # Create rows for each match
    for match in matches:
        row_frame = tk.Frame(results_scrollable_frame, bg=COLOR_BACKGROUND, relief=tk.RIDGE, bd=1)
        row_frame.pack(fill=tk.X, pady=2, padx=5)
        
        # Checkbox (checked by default) - custom larger checkbox
        checkbox_var = tk.BooleanVar(value=True)
        
        checkbox_btn = tk.Button(
            row_frame,
            text="☑",
            font=("Arial", 40),
            bg=COLOR_BACKGROUND,
            fg=COLOR_TEXT,
            relief=tk.FLAT,
            bd=0,
            padx=5,
            pady=5,
            cursor="hand2"
        )
        
        def update_checkbox_display(var=None, index=None, mode=None, btn=checkbox_btn, var_ref=checkbox_var):
            if var_ref.get():
                btn.config(text="☑", font=("Arial", 40))
            else:
                btn.config(text="☐", font=("Arial", 40))
        
        def toggle_checkbox(var_ref=checkbox_var):
            var_ref.set(not var_ref.get())
        
        checkbox_btn.config(command=toggle_checkbox)
        checkbox_btn.pack(side=tk.LEFT, padx=5, pady=5)
        match_checkboxes.append((match, checkbox_var))
        
        # Update display when variable changes
        checkbox_var.trace_add("write", update_checkbox_display)
        
        # Volume name
        volume_label = tk.Label(
            row_frame,
            text=match["folder"],
            font=("Arial", 30),
            bg=COLOR_BACKGROUND,
            fg=COLOR_TEXT,
            width=15,
            anchor="w",
            wraplength=150
        )
        volume_label.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Check if image exists (fallback: don't show page or button if image doesn't exist)
        image_exists = match.get("image_exists", True)  # Default to True for backward compatibility
        
        if image_exists:
            # Open image button (moved between Volume and Page)
            def make_open_button(match_data):
                btn = tk.Button(
                    row_frame,
                    text="📷 Deschide",
                    font=("Arial", 28),
                    bg=COLOR_BOOK,
                    fg=COLOR_TEXT,
                    relief=tk.RAISED,
                    padx=10,
                    pady=5,
                    cursor="hand2",
                    command=lambda: open_match_image(match_data)
                )
                return btn
            
            open_btn = make_open_button(match)
            open_btn.pack(side=tk.LEFT, padx=5, pady=5)
            
            # Page number
            page_label = tk.Label(
                row_frame,
                text=str(match["page_num"]),
                font=("Arial", 30),
                bg=COLOR_BACKGROUND,
                fg=COLOR_TEXT,
                width=10
            )
            page_label.pack(side=tk.LEFT, padx=5, pady=5)
        else:
            # Image doesn't exist - skip page number and button
            pass
        
        # Quote with bolded search words - use the exact matched fragment if available
        if "matched_fragment" in match and match["matched_fragment"]:
            # Use the exact matched fragment that was found during search
            matched_fragment = match["matched_fragment"]
            
            # Add 4 words before and 4 words after the matched fragment
            if "page_lines" in match and match["page_lines"]:
                # Get all words from page lines around the match area
                page_lines = match["page_lines"]
                match_start_line = match.get("match_start_line_idx", 0)
                match_end_line = match.get("match_end_line_idx", match_start_line + 1)
                
                # Get context lines (a few lines before and after the match)
                context_start = max(0, match_start_line - 2)
                context_end = min(len(page_lines), match_end_line + 2)
                context_lines = page_lines[context_start:context_end]
                
                # Convert context lines to words
                context_text = " ".join(context_lines)
                all_words = context_text.split()
                matched_words = matched_fragment.split()
                
                # Find where the matched fragment appears in all_words (using normalized comparison)
                matched_start_idx = None
                norm_matched_words = [normalize_text(w) for w in matched_words]
                for i in range(len(all_words) - len(matched_words) + 1):
                    # Compare normalized versions
                    norm_window = [normalize_text(w) for w in all_words[i:i+len(matched_words)]]
                    if norm_window == norm_matched_words:
                        matched_start_idx = i
                        break
                
                if matched_start_idx is not None:
                    # Extract 4 words before and 4 words after
                    start_idx = max(0, matched_start_idx - 4)
                    end_idx = min(len(all_words), matched_start_idx + len(matched_words) + 4)
                    quote_text = " ".join(all_words[start_idx:end_idx])
                else:
                    # Fallback: use matched fragment if we can't find it in context
                    quote_text = matched_fragment
            else:
                # Fallback: use matched fragment if no page_lines available
                quote_text = matched_fragment
        else:
            # Fallback: extract from snippet (for backward compatibility)
            full_quote_text = " ".join(match["snippet"])
            quote_words = full_quote_text.split()
            
            # Find where the search words appear in the quote
            norm_quote_words = [normalize_text(w) for w in quote_words]
            norm_search_words = [normalize_text(w) for w in search_words]
            
            # Find the first occurrence of any search word
            match_start_idx = None
            for i, norm_word in enumerate(norm_quote_words):
                for norm_search_word in norm_search_words:
                    if norm_search_word in norm_word or norm_word in norm_search_word:
                        match_start_idx = i
                        break
                if match_start_idx is not None:
                    break
            
            # If no match found, use the beginning
            if match_start_idx is None:
                match_start_idx = 0
            
            # Calculate window around the match
            # Try to center the match in the window
            half_span = word_span // 2
            start_idx = max(0, match_start_idx - half_span)
            end_idx = min(len(quote_words), start_idx + word_span)
            
            # Adjust start if we're near the end
            if end_idx - start_idx < word_span and start_idx > 0:
                start_idx = max(0, end_idx - word_span)
            
            # Extract the quote window
            quote_text = " ".join(quote_words[start_idx:end_idx])
        
        # Create a Text widget for formatted quote (supports bold)
        quote_text_widget = tk.Text(
            row_frame,
            font=("Arial", 30),
            bg=COLOR_BACKGROUND,
            fg=COLOR_TEXT,
            width=60,
            height=4,
            wrap=tk.WORD,
            relief=tk.FLAT,
            padx=5,
            pady=5
        )
        quote_text_widget.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.BOTH, expand=True)
        
        # Insert text first
        quote_text_widget.insert("1.0", quote_text)
        
        # Apply bold formatting to search words (case-insensitive)
        # Normalize the quote for matching
        norm_quote = normalize_text(quote_text)
        quote_words_list = quote_text.split()
        norm_quote_words = [normalize_text(w) for w in quote_words_list]
        
        # For each search word, find matching words in the quote and bold them
        for search_word in search_words:
            norm_search_word = normalize_text(search_word)
            if not norm_search_word:
                continue
            
            # Find all words in the quote that match the search word
            for i, norm_quote_word in enumerate(norm_quote_words):
                if words_match(norm_search_word, norm_quote_word):
                    # Calculate the start and end positions of this word in the text widget
                    # Count characters before this word
                    char_count = sum(len(w) + 1 for w in quote_words_list[:i])  # +1 for space
                    start_pos = f"1.0+{char_count}c"
                    end_pos = f"1.0+{char_count + len(quote_words_list[i])}c"
                    quote_text_widget.tag_add("bold", start_pos, end_pos)
        
        quote_text_widget.tag_config("bold", font=("Arial", 30, "bold"))
        quote_text_widget.config(state=tk.DISABLED)  # Make read-only after formatting
    
    # Update scroll region
    results_canvas.update_idletasks()
    results_canvas.configure(scrollregion=results_canvas.bbox("all"))
    
    # Add export button at the bottom
    global export_button_frame
    if export_button_frame:
        # Clear existing buttons
        for widget in export_button_frame.winfo_children():
            widget.destroy()
        
        # Add export button
        export_btn = tk.Button(
            export_button_frame,
            text="📄 Exportă la DOCX",
            command=lambda: export_selected_to_docx(search_term),
            bg=COLOR_SHELF,
            fg="white",
            font=("Arial", 36, "bold"),
            relief=tk.RAISED,
            padx=20,
            pady=10,
            cursor="hand2"
        )
        export_btn.pack()


def open_match_image(match: dict):
    """Open the image for a search match."""
    volume_path_str = match.get("volume_path", "")
    image_filename = match.get("image", "")
    
    if not volume_path_str or not image_filename:
        messagebox.showerror("Error", "Image information not available.")
        return
    
    volume_path = Path(volume_path_str)
    image_path = find_image_path(volume_path, image_filename)
    open_image(image_path)


def export_selected_to_docx(search_term: str):
    """Export selected matches to a Word document."""
    global match_checkboxes
    
    # Get selected matches
    selected_matches = [match for match, checkbox_var in match_checkboxes if checkbox_var.get()]
    
    if not selected_matches:
        messagebox.showwarning("No selection", "Please select at least one result to export.")
        return
    
    try:
        # Get output directory
        output_dir = get_output_path()
        
        # Find or create default.docx template
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
        search_run = title_para.add_run(search_term)
        search_run.italic = True
        title_para.add_run('"')
        doc.add_paragraph()
        
        # Add selected matches
        for match in selected_matches:
            # Get context: 5 lines above and 5 lines below the match
            if 'page_lines' in match and 'match_start_line_idx' in match:
                page_lines = match['page_lines']
                match_start = match.get('match_start_line_idx', 0)
                match_end = match.get('match_end_line_idx', match_start + 1)
                
                # Calculate context window: 5 lines above and 5 lines below
                context_start = max(0, match_start - 5)
                context_end = min(len(page_lines), match_end + 5)
                
                # Extract context lines
                context_lines = page_lines[context_start:context_end]
                
                # Add context lines to document
                for line in context_lines:
                    if line.strip():
                        para = doc.add_paragraph(line.strip())
                        try:
                            para.style = "2-Versuri-centru"
                        except KeyError:
                            pass
            else:
                # Fallback: use snippet if available, otherwise use matched_fragment
                if 'snippet' in match and match['snippet']:
                    for snippet_line in match['snippet']:
                        if snippet_line.strip():
                            para = doc.add_paragraph(snippet_line.strip())
                            try:
                                para.style = "2-Versuri-centru"
                            except KeyError:
                                pass
                elif 'matched_fragment' in match and match['matched_fragment']:
                    para = doc.add_paragraph(match['matched_fragment'].strip())
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
            return
        
        # Save document
        doc.save(save_path)
        
        # Show success message
        messagebox.showinfo(
            "Success",
            f"Report exported:\n{save_path}\n\n{len(selected_matches)} result(s) exported."
        )
    except Exception as e:
        # Show user-friendly error dialog
        messagebox.showerror(
            "Error",
            f"Failed to export report:\n{str(e)}"
        )


def on_generate():
    """Handle the Search button click."""
    search_term = entry.get().strip()
    words = search_term.split()

    # Validate search term length (1-10 words)
    if not (1 <= len(words) <= 10):
        messagebox.showerror(
            "Invalid input",
            "Please enter between 1 and 10 words."
        )
        return

    # Check if root folder is selected
    if not selected_root_folder:
        messagebox.showerror(
            "Missing folder",
            "Please select a Biblioteca folder first."
        )
        return
    
    # Check if any volumes are selected
    if not selected_volumes:
        messagebox.showerror(
            "No selection",
            "Please select at least one folder to search."
        )
        return

    try:
        # Disable button during processing
        button.config(state="disabled")
        root.update()  # Update UI to show disabled state
        
        # Get word span value
        word_span_value = word_span_var.get()
        if word_span_value and word_span_value.isdigit() and int(word_span_value) > 0:
            word_span = int(word_span_value)
        else:
            word_span = 10  # Default to 10 words if invalid
        
        # Get word order setting (exact or random)
        order_value = word_order_var.get()
        exact_order = (order_value == "Exact")
        
        # Run search only (no document generation)
        matches = run_search_only(
            query_text=search_term,
            ocr_root=Path(selected_root_folder),
            word_span=word_span,
            threshold=DEFAULT_THRESHOLD,
            selected_volumes=selected_volumes,
            exact_order=exact_order
        )
        
        # Display results in table with word_span limit
        display_results_in_table(matches, search_term, word_span=word_span)
        
    except Exception as e:
        # Show user-friendly error dialog
        messagebox.showerror(
            "Error",
            f"Failed to search:\n{str(e)}"
        )
    finally:
        # Re-enable button
        button.config(state="normal")


def _get_text_with_checkmark(text: str, add_checkmark: bool, tree_width: int = 600) -> str:
    """Get text with checkmark on the left before the title.
    tree_width: width of the tree column in pixels (default 600)"""
    # Remove existing checkmark if present (check both left and right)
    text = text.strip()
    if text.startswith("✓"):
        # Remove checkmark from the left
        text = text[1:].strip()
    elif "✓" in text:
        # Remove checkmark from the right (for backward compatibility)
        text = text.split("✓")[0].rstrip()
    
    if add_checkmark:
        # Put checkmark on the left before the text
        return "✓ " + text
    return text


def populate_library_tree(biblioteca_path: str):
    """Populate the library tree dynamically from folder structure.
    Marks folders containing ocr.txt as end folders (leaf nodes).
    All items are selected by default."""
    library_tree.delete(*library_tree.get_children())
    root_path = Path(biblioteca_path)
    
    # Add root folder - selected by default (with checkmark aligned to right)
    root_text = _get_text_with_checkmark(f"📚 {root_path.name}", True)
    root_id = library_tree.insert("", "end", text=root_text, 
                                  values=(str(root_path), "folder"),
                                  tags=("folder", "selected"))
    
    # Recursively build tree
    _build_tree_recursive(root_path, root_id)
    
    # Expand root level
    library_tree.item(root_id, open=True)
    
    # Update selected volumes after populating tree
    update_selected_volumes()


def _build_tree_recursive(folder: Path, parent_item):
    """Recursively build tree structure, marking folders with OCR as end folders."""
    # Check if this folder has OCR files
    has_ocr = folder_has_ocr(folder)
    
    # Get subfolders
    subfolders = sorted([p for p in folder.iterdir() if p.is_dir()])
    
    # If folder has OCR, mark it as end folder and don't show subfolders
    if has_ocr:
        # Mark parent as end folder (has OCR)
        current_tags = list(library_tree.item(parent_item, "tags"))
        if "end_folder" not in current_tags:
            library_tree.item(parent_item, tags=current_tags + ["end_folder"])
        return
    
    # If no OCR, show subfolders
    for subfolder in subfolders:
        # Determine icon and tags based on whether subfolder has OCR
        subfolder_has_ocr = folder_has_ocr(subfolder)
        
        # Get tree column width for proper alignment
        tree_width = library_tree.column("#0", "width") or 600
        
        if subfolder_has_ocr:
            # End folder (has OCR) - mark with special icon
            icon = "📗"
            tags = ("folder", "end_folder", "selected")
            text = _get_text_with_checkmark(f"{icon} {subfolder.name}", True, tree_width)  # Checkmark aligned to right
        else:
            # Regular folder (no OCR yet, may have subfolders)
            icon = "📁"
            tags = ("folder", "selected")
            text = _get_text_with_checkmark(f"{icon} {subfolder.name}", True, tree_width)  # Checkmark aligned to right
        
        subfolder_id = library_tree.insert(parent_item, "end", 
                                           text=text,
                                           values=(str(subfolder), "folder"),
                                           tags=tags)
        
        # Recursively process subfolders
        _build_tree_recursive(subfolder, subfolder_id)
    
    # Expand first level by default
    if parent_item:
        for child in library_tree.get_children(parent_item):
            library_tree.item(child, open=True)


def on_tree_check(event):
    """Handle checkbox clicks in the tree."""
    # Get the item that was clicked
    item = library_tree.selection()[0] if library_tree.selection() else None
    if not item:
        return
    
    # Toggle selection state
    current_tags = library_tree.item(item, "tags")
    if "selected" in current_tags:
        library_tree.item(item, tags=[tag for tag in current_tags if tag != "selected"])
    else:
        library_tree.item(item, tags=list(current_tags) + ["selected"])


def get_selected_paths():
    """Get all selected paths from the tree, handling parent-child relationships.
    Returns list of paths to search. If root folder is selected, returns just root.
    Otherwise returns selected folders (deduplicated)."""
    selected = []
    root_path = None
    
    # First pass: collect all selected items
    for item in library_tree.get_children():
        root_path, paths = _collect_selected(item, [])
        if root_path:
            # If root folder is selected, return it immediately (searches everything)
            return [str(root_path)]
        selected.extend(paths)
    
    # If no items selected, return empty list
    if not selected:
        return []
    
    # Deduplicate: if a parent is selected, remove its children
    # This handles cases like: Folder1 selected + SubFolder1 (child of Folder1) selected
    # We only need Folder1 since it includes SubFolder1
    selected_paths = []
    for path_str in selected:
        path = Path(path_str)
        # Check if this path is a child of any other selected path
        is_child = False
        for other_path_str in selected:
            if path_str != other_path_str:
                other_path = Path(other_path_str)
                try:
                    # Check if path is a subpath of other_path
                    path.relative_to(other_path)
                    is_child = True
                    break
                except ValueError:
                    pass
        
        if not is_child:
            selected_paths.append(path_str)
    
    return selected_paths


def _collect_selected(item, selected_list):
    """Recursively collect selected items. Returns (root_path, selected_paths).
    root_path is set if root folder itself is selected, otherwise None."""
    tags = library_tree.item(item, "tags")
    values = library_tree.item(item, "values")
    
    # Check if this item is selected
    if "selected" in tags and values:
        path_str = values[0]
        path = Path(path_str)
        
        # Check if this is the root item (first level child of tree root)
        parent = library_tree.parent(item)
        if parent == "":  # Root item
            return (Path(path_str), [])
        
        # Otherwise, add this path to the list
        selected_list.append(path_str)
    
    # Recursively check children
    for child in library_tree.get_children(item):
        root_path, child_paths = _collect_selected(child, [])
        if root_path:  # Root was selected in a child branch
            return (root_path, [])
        selected_list.extend(child_paths)
    
    return (None, selected_list)


def update_selected_volumes():
    """Update the global selected_volumes list with all selected end folders (with ocr.txt)."""
    global selected_volumes
    selected_volumes = []
    
    # Get all selected paths
    selected = get_selected_paths()
    
    # If root is selected, find all volumes recursively
    if selected_root_folder and len(selected) == 1 and Path(selected[0]) == Path(selected_root_folder):
        # Root selected - find all volumes
        _find_all_volumes(Path(selected_root_folder), selected_volumes)
    else:
        # Specific folders selected - extract only volumes (end folders with ocr.txt)
        for path_str in selected:
            path = Path(path_str)
            if folder_has_ocr(path):
                # This is a volume (end folder with OCR)
                selected_volumes.append(str(path))
            else:
                # Not a volume, search recursively for volumes within it
                _find_all_volumes(path, selected_volumes)


def _find_all_volumes(folder: Path, volumes_list: list):
    """Recursively find all volumes (folders with ocr.txt) within a folder."""
    if folder_has_ocr(folder):
        # This folder is a volume
        volumes_list.append(str(folder))
        return  # Don't search subfolders if this is an end folder
    
    # Search in subfolders
    for subfolder in sorted(p for p in folder.iterdir() if p.is_dir()):
        _find_all_volumes(subfolder, volumes_list)


def _deselect_parent_if_needed(item):
    """Deselect parent when any child is deselected.
    Recursively propagates up the tree."""
    parent = library_tree.parent(item)
    # If no parent (root level), stop
    if not parent or parent == "":
        return
    
    # Deselect the parent immediately (if it's selected)
    parent_tags = list(library_tree.item(parent, "tags"))
    if "selected" in parent_tags:
        parent_tags.remove("selected")
        tree_width = library_tree.column("#0", "width") or 600
        parent_text = library_tree.item(parent, "text")
        new_text = _get_text_with_checkmark(parent_text, False, tree_width)
        library_tree.item(parent, text=new_text, tags=parent_tags)
        # Recursively deselect parent's parent
        _deselect_parent_if_needed(parent)


def on_tree_click(event):
    """Handle clicks on tree items (toggle selection with cascading to children)."""
    region = library_tree.identify_region(event.x, event.y)
    if region == "cell" or region == "tree":
        item = library_tree.identify_row(event.y)
        if item:
            tags = list(library_tree.item(item, "tags"))
            # Determine new selection state (toggle)
            should_select = "selected" not in tags
            # Apply selection state to this item and all its children recursively
            _select_item_recursive(item, should_select)
            # If item was deselected, check if parent should also be deselected
            if not should_select:
                _deselect_parent_if_needed(item)
            # Update the selected volumes list
            update_selected_volumes()


# Create main window with library theme
root = tk.Tk()
root.title("Biblioteca - Căutare Documente")
root.geometry("800x600")
root.configure(bg="#F5E6D3")  # Warm beige background like old books

# Library-themed colors
COLOR_BACKGROUND = "#F5E6D3"
COLOR_SHELF = "#8B4513"  # Brown like wooden shelves
COLOR_BOOK = "#D4A574"  # Tan like book covers
COLOR_TEXT = "#2C1810"  # Dark brown text
COLOR_ACCENT = "#A0522D"  # Sienna accent

# Create scrollable frame for the entire window
main_canvas = tk.Canvas(root, bg=COLOR_BACKGROUND, highlightthickness=0)
main_scrollbar = ttk.Scrollbar(root, orient="vertical", command=main_canvas.yview)
scrollable_main_frame = tk.Frame(main_canvas, bg=COLOR_BACKGROUND)

scrollable_main_frame.bind(
    "<Configure>",
    lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
)

main_canvas.create_window((0, 0), window=scrollable_main_frame, anchor="nw")
main_canvas.configure(yscrollcommand=main_scrollbar.set)

# Pack canvas and scrollbar
main_canvas.pack(side="left", fill="both", expand=True)
main_scrollbar.pack(side="right", fill="y")

# Bind mousewheel to canvas
def _on_main_mousewheel(event):
    main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
main_canvas.bind_all("<MouseWheel>", _on_main_mousewheel)

# Update canvas width when scrollable frame changes
def configure_main_canvas_width(event):
    canvas_width = event.width
    main_canvas.itemconfig(main_canvas.find_all()[0], width=canvas_width)

scrollable_main_frame.bind("<Configure>", lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
main_canvas.bind("<Configure>", configure_main_canvas_width)

# Header frame with library title
header_frame = tk.Frame(scrollable_main_frame, bg=COLOR_BACKGROUND, pady=15)
header_frame.pack(fill=tk.X)

title_label = tk.Label(
    header_frame,
    text="📚 BIBLIOTECA 📚",
    font=("Times", 64, "bold"),
    bg=COLOR_BACKGROUND,
    fg=COLOR_TEXT
)
title_label.pack()

subtitle_label = tk.Label(
    header_frame,
    text="Selectează Biblioteca, Rafturi sau Volume pentru căutare",
    font=("Times", 32, "italic"),
    bg=COLOR_BACKGROUND,
    fg=COLOR_ACCENT
)
subtitle_label.pack(pady=(5, 0))

# Biblioteca selection section
biblioteca_frame = tk.Frame(scrollable_main_frame, bg=COLOR_BACKGROUND)
biblioteca_frame.pack(padx=15, pady=10, fill=tk.X)

select_folder_button = tk.Button(
    biblioteca_frame,
    text="📂 Selectează Biblioteca",
    command=on_select_folder,
    bg=COLOR_SHELF,
    fg="white",
    font=("Arial", 32, "bold"),
    relief=tk.RAISED,
    padx=15,
    pady=8,
    cursor="hand2"
)
select_folder_button.pack(side=tk.LEFT, padx=(0, 10))

folder_label = tk.Label(
    biblioteca_frame,
    text="Nicio bibliotecă selectată",
    fg=COLOR_ACCENT,
    bg=COLOR_BACKGROUND,
    font=("Arial", 30, "italic"),
    anchor="w"
)
folder_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

# Library tree frame with scrollbars
tree_frame = tk.Frame(scrollable_main_frame, bg=COLOR_BACKGROUND)
tree_frame.pack(padx=15, pady=10, fill=tk.X)

# Create scrollbars
v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

# Create tree with library styling
library_tree = ttk.Treeview(
    tree_frame,
    columns=("path", "type"),
    show="tree",
    yscrollcommand=v_scrollbar.set,
    xscrollcommand=h_scrollbar.set,
    selectmode="extended"
)
library_tree.column("#0", width=600, minwidth=200)
library_tree.column("path", width=0, stretch=False)  # Hidden column
library_tree.column("type", width=0, stretch=False)  # Hidden column

# Configure treeview font and row height
style = ttk.Style()
style.configure("Treeview", font=("Arial", 30), rowheight=60)
style.configure("Treeview.Heading", font=("Arial", 30, "bold"))

# Configure tags for styling
library_tree.tag_configure("folder", background="#F0E0C0", foreground=COLOR_TEXT)
library_tree.tag_configure("end_folder", background="#E8D5B7", foreground=COLOR_TEXT)  # Folders with OCR
library_tree.tag_configure("selected", background="#4A7C59", foreground="white")  # Green color for selected items

# Pack scrollbars and tree
v_scrollbar.config(command=library_tree.yview)
h_scrollbar.config(command=library_tree.xview)
v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
library_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Bind click events
library_tree.bind("<Button-1>", on_tree_click)
library_tree.bind("<Double-1>", lambda e: None)  # Prevent default double-click behavior

# Selection control buttons
selection_frame = tk.Frame(scrollable_main_frame, bg=COLOR_BACKGROUND)
selection_frame.pack(padx=15, pady=5, fill=tk.X)

def select_all_items():
    """Select all items in the tree."""
    for item in library_tree.get_children():
        _select_item_recursive(item, True)
    # Update the selected volumes list
    update_selected_volumes()

def deselect_all_items():
    """Deselect all items in the tree."""
    for item in library_tree.get_children():
        _select_item_recursive(item, False)
    # Update the selected volumes list
    update_selected_volumes()

def _select_item_recursive(item, select=True):
    """Recursively select or deselect an item and its children, updating text and colors."""
    tags = list(library_tree.item(item, "tags"))
    current_text = library_tree.item(item, "text")
    
    if select:
        if "selected" not in tags:
            tags.append("selected")
    else:
        if "selected" in tags:
            tags.remove("selected")
    
    # Get tree column width for proper alignment
    tree_width = library_tree.column("#0", "width") or 600
    
    # Update text with checkmark aligned to the right
    new_text = _get_text_with_checkmark(current_text, select, tree_width)
    library_tree.item(item, text=new_text, tags=tags)
    
    # Recursively process children
    for child in library_tree.get_children(item):
        _select_item_recursive(child, select)

select_all_btn = tk.Button(
    selection_frame,
    text="✓ Selectează Tot",
    command=select_all_items,
    bg=COLOR_BOOK,
    fg=COLOR_TEXT,
    font=("Arial", 28),
    relief=tk.RAISED,
    padx=10,
    pady=5,
    cursor="hand2"
)
select_all_btn.pack(side=tk.LEFT, padx=(0, 10))

deselect_all_btn = tk.Button(
    selection_frame,
    text="✗ Deselectează Tot",
    command=deselect_all_items,
    bg=COLOR_BOOK,
    fg=COLOR_TEXT,
    font=("Arial", 28),
    relief=tk.RAISED,
    padx=10,
    pady=5,
    cursor="hand2"
)
deselect_all_btn.pack(side=tk.LEFT)

# Instructions label
# instructions_label = tk.Label(
#     root,
#     text="💡 Click pe foldere pentru a le selecta/deselecta. Folderele cu 📗 conțin ocr.txt",
#     font=("Arial", 9),
#     bg=COLOR_BACKGROUND,
#     fg=COLOR_ACCENT,
#     pady=5
# )
# instructions_label.pack()

# Search section
search_frame = tk.Frame(scrollable_main_frame, bg=COLOR_BACKGROUND)
search_frame.pack(padx=15, pady=10, fill=tk.X)

tk.Label(
    search_frame,
    text="Căutare:",
    font=("Arial", 32, "bold"),
    bg=COLOR_BACKGROUND,
    fg=COLOR_TEXT
).pack(anchor="w", pady=(0, 5))

entry = tk.Entry(
    search_frame,
    width=60,
    font=("Arial", 32),
    relief=tk.SUNKEN,
    bd=2
)
entry.pack(fill=tk.X, pady=5)

# Word span section
word_span_frame = tk.Frame(scrollable_main_frame, bg=COLOR_BACKGROUND)
word_span_frame.pack(padx=15, pady=5, fill=tk.X)

tk.Label(
    word_span_frame,
    text="Span cuvinte:",
    font=("Arial", 30),
    bg=COLOR_BACKGROUND,
    fg=COLOR_TEXT
).pack(side=tk.LEFT, padx=(0, 10))

# Create dropdown for word span options
word_span_var = tk.StringVar(value="10")  # Default to 10 words
word_span_options = ["5", "10", "15", "20"]

word_span_dropdown = ttk.Combobox(
    word_span_frame,
    textvariable=word_span_var,
    values=word_span_options,
    state="readonly",
    width=15,
    font=("Arial", 30)
)
word_span_dropdown.pack(side=tk.LEFT)

# Word order section
word_order_frame = tk.Frame(scrollable_main_frame, bg=COLOR_BACKGROUND)
word_order_frame.pack(padx=15, pady=5, fill=tk.X)

tk.Label(
    word_order_frame,
    text="Ordinea cuvintelor:",
    font=("Arial", 30),
    bg=COLOR_BACKGROUND,
    fg=COLOR_TEXT
).pack(side=tk.LEFT, padx=(0, 10))

# Create dropdown for word order options
word_order_var = tk.StringVar(value="Exact")  # Default to exact order
word_order_options = ["Exact", "Aleatorie"]

word_order_dropdown = ttk.Combobox(
    word_order_frame,
    textvariable=word_order_var,
    values=word_order_options,
    state="readonly",
    width=15,
    font=("Arial", 30)
)
word_order_dropdown.pack(side=tk.LEFT)

# Generate button
button = tk.Button(
    scrollable_main_frame,
    text="🔍 Caută",
    command=on_generate,
    bg=COLOR_SHELF,
    fg="white",
    font=("Arial", 36, "bold"),
    relief=tk.RAISED,
    padx=20,
    pady=10,
    cursor="hand2"
)
button.pack(pady=15)

# Global variable to store results window
results_window = None
results_scrollable_frame = None
results_canvas = None
export_button_frame = None
results_title_label = None
match_checkboxes = []  # List of (match, checkbox_var) tuples for export

if __name__ == "__main__":
    root.mainloop()

