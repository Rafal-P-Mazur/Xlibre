import os
import concurrent.futures
import sys
import struct
import fitz  # PyMuPDF
from PIL import Image, ImageEnhance, ImageDraw, ImageFont, ImageOps, ImageFilter
import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup, NavigableString
import pyphen
import base64
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import threading
import re
import json
import glob
import io
from urllib.parse import unquote

# --- CONFIGURATION DEFAULTS ---
DEFAULT_SCREEN_WIDTH = 480
DEFAULT_SCREEN_HEIGHT = 800
DEFAULT_RENDER_SCALE = 3.0

# --- FACTORY DEFAULTS ---
FACTORY_DEFAULTS = {
    "font_size": 28,
    "font_weight": 400,
    "line_height": 1.4,
    "margin": 20,
    "top_padding": 15,
    "bottom_padding": 45,
    "orientation": "Portrait",
    "text_align": "justify",
    "font_name": "Default (System)",
    "preview_zoom": 300,
    "generate_toc": True,
    "show_footnotes": False,

    # --- ELEMENT VISIBILITY & POSITION ---
    "pos_title": "Footer",
    "pos_pagenum": "Footer",
    "pos_chap_page": "Hidden",
    "pos_percent": "Hidden",

    # --- ELEMENT ORDERING (Lower number = Left/First) ---
    "order_pagenum": 1,
    "order_title": 2,
    "order_chap_page": 3,
    "order_percent": 4,

    # --- PROGRESS BAR POSITION ---
    "pos_progress": "Footer (Below Text)",

    # --- PROGRESS BAR STYLING ---
    "bar_height": 4,
    "bar_show_marker": True,
    "bar_marker_color": "Black",
    "bar_marker_radius": 5,
    "bar_show_ticks": True,
    "bar_tick_height": 6,

    # --- HEADER STYLING ---
    "header_font_size": 16,
    "header_align": "Center",
    "header_margin": 10,

    # --- FOOTER STYLING ---
    "footer_font_size": 16,
    "footer_align": "Center",
    "footer_margin": 10,

    # --- RENDER MODES ---
    "render_mode": "Threshold",

    # Dither Settings (Also used for Images in Threshold mode)
    "white_clip": 220,
    "contrast": 1.2,

    # Threshold Settings (Text Only)
    "text_threshold": 130,
    "text_blur": 1.0,
}

SETTINGS_FILE = "default_settings.json"
PRESETS_DIR = "presets"


# --- UTILITY FUNCTIONS ---
def fix_css_font_paths(css_text, target_font_family="'CustomFont'"):
    if target_font_family is None:
        return css_text
    css_text = re.sub(r'font-family\s*:\s*[^;!]+', f'font-family: {target_font_family}', css_text)
    return css_text


def get_font_variants(font_path):
    if not font_path or not os.path.exists(font_path):
        return {}
    directory = os.path.dirname(font_path)
    try:
        all_files = [f for f in os.listdir(directory) if f.lower().endswith((".ttf", ".otf"))]
    except:
        return {}
    candidates = {"regular": [], "italic": [], "bold": [], "bold_italic": []}

    for f in all_files:
        full_path = os.path.join(directory, f).replace("\\", "/")
        name_lower = f.lower()
        has_bold = any(x in name_lower for x in ["bold", "bd"])
        has_italic = any(x in name_lower for x in ["italic", "oblique", "obl"])
        if has_bold and has_italic:
            candidates["bold_italic"].append(full_path)
            continue
        if has_bold and not has_italic:
            candidates["bold"].append(full_path)
            continue
        if has_italic and not has_bold:
            candidates["italic"].append(full_path)
            continue
        candidates["regular"].append(full_path)

    def pick_best(file_list):
        if not file_list: return None
        return sorted(file_list, key=lambda p: len(os.path.basename(p)))[0]

    return {
        "regular": font_path.replace("\\", "/"),
        "italic": pick_best(candidates["italic"]),
        "bold": pick_best(candidates["bold"]),
        "bold_italic": pick_best(candidates["bold_italic"])
    }


def get_pil_font(font_path, size):
    try:
        if font_path and os.path.exists(font_path):
            return ImageFont.truetype(font_path, size)
        return ImageFont.load_default()
    except:
        return ImageFont.load_default()


def extract_all_css(book):
    css_rules = []
    for item in book.get_items_of_type(ebooklib.ITEM_STYLE):
        try:
            css_rules.append(item.get_content().decode('utf-8', errors='ignore'))
        except:
            pass
    return "\n".join(css_rules)


def extract_images_to_base64(book):
    image_map = {}
    for item in book.get_items_of_type(ebooklib.ITEM_IMAGE):
        try:
            filename = os.path.basename(item.get_name())
            b64_data = base64.b64encode(item.get_content()).decode('utf-8')
            image_map[filename] = f"data:{item.media_type};base64,{b64_data}"
        except:
            pass
    return image_map


def get_official_toc_mapping(book):
    mapping = {}

    def add_entry(href, title):
        if '#' in href:
            href_clean, anchor = href.split('#', 1)
        else:
            href_clean, anchor = href, None
        filename = os.path.basename(href_clean)
        if filename not in mapping:
            mapping[filename] = []
        mapping[filename].append((anchor, title))

    def process_toc_item(item):
        if isinstance(item, tuple):
            if len(item) > 1 and isinstance(item[1], list):
                for sub in item[1]: process_toc_item(sub)
        elif isinstance(item, epub.Link):
            add_entry(item.href, item.title)

    for item in book.toc:
        process_toc_item(item)

    if not mapping:
        nav_item = next((item for item in book.get_items() if item.get_type() == ebooklib.ITEM_NAVIGATION), None)
        if nav_item:
            try:
                soup = BeautifulSoup(nav_item.get_content(), 'html.parser')
                nav_element = soup.find('nav', attrs={'epub:type': 'toc'}) or soup.find('nav')
                if nav_element:
                    for link in nav_element.find_all('a', href=True):
                        add_entry(link['href'], link.get_text().strip())
            except:
                pass
    return mapping


def hyphenate_html_text(soup, language_code):
    try:
        dic = pyphen.Pyphen(lang=language_code)
    except:
        try:
            dic = pyphen.Pyphen(lang='en')
        except:
            return soup
    word_pattern = re.compile(r'\w+', re.UNICODE)
    for text_node in soup.find_all(string=True):
        if text_node.parent.name in ['script', 'style', 'head', 'title', 'meta']: continue
        if not text_node.strip(): continue
        original_text = str(text_node)
        clean_text = original_text.replace('\u00A0', ' ')

        def replace_match(match):
            word = match.group(0)
            if len(word) < 6: return word
            return dic.inserted(word, hyphen='\u00AD')

        new_text = word_pattern.sub(replace_match, clean_text)
        if new_text != original_text:
            text_node.replace_with(NavigableString(new_text))
    return soup


def get_local_fonts():
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    fonts_dir = os.path.join(base_path, "fonts")
    if not os.path.exists(fonts_dir):
        try:
            os.makedirs(fonts_dir)
        except OSError:
            pass
    font_map = {}
    if os.path.exists(fonts_dir):
        for item in os.listdir(fonts_dir):
            item_path = os.path.join(fonts_dir, item)
            if os.path.isdir(item_path):
                files = [f for f in os.listdir(item_path) if f.lower().endswith((".ttf", ".otf"))]
                if not files: continue
                candidates = [f for f in files if "bold" not in f.lower() and "italic" not in f.lower()]
                main_file = candidates[0] if candidates else files[0]
                font_map[item] = os.path.abspath(os.path.join(item_path, main_file))
        for item in os.listdir(fonts_dir):
            item_path = os.path.join(fonts_dir, item)
            if os.path.isfile(item_path) and item.lower().endswith((".ttf", ".otf")):
                font_map[item] = os.path.abspath(item_path)
    return font_map


class CoverExportDialog(ctk.CTkToplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback = callback
        self.title("Export Cover Settings")
        self.geometry("400x350")
        self.transient(parent)
        self.grab_set()
        ctk.CTkLabel(self, text="Export Cover as BMP", font=("Arial", 16, "bold")).pack(pady=15)
        self.frm_dim = ctk.CTkFrame(self, fg_color="transparent")
        self.frm_dim.pack(pady=10)
        ctk.CTkLabel(self.frm_dim, text="Width:").grid(row=0, column=0, padx=5)
        self.entry_w = ctk.CTkEntry(self.frm_dim, width=60)
        self.entry_w.insert(0, "480")
        self.entry_w.grid(row=0, column=1, padx=5)
        ctk.CTkLabel(self.frm_dim, text="Height:").grid(row=0, column=2, padx=5)
        self.entry_h = ctk.CTkEntry(self.frm_dim, width=60)
        self.entry_h.insert(0, "800")
        self.entry_h.grid(row=0, column=3, padx=5)
        ctk.CTkLabel(self, text="Scaling Mode:").pack(pady=(15, 5))
        self.mode_var = ctk.StringVar(value="Crop to Fill (Best)")
        modes = ["Crop to Fill (Best)", "Fit (Add White Bars)", "Stretch (Distort)"]
        ctk.CTkOptionMenu(self, variable=self.mode_var, values=modes).pack(pady=5)
        ctk.CTkLabel(self,
                     text="Crop to Fill: Fills the screen, cuts off edges if ratio differs.\nFit: Keeps entire image, adds white bars.\nStretch: Forces image to size, may look squashed.",
                     font=("Arial", 11), text_color="gray").pack(pady=10)
        ctk.CTkButton(self, text="Export", command=self.confirm).pack(pady=20)

    def confirm(self):
        try:
            w = int(self.entry_w.get())
            h = int(self.entry_h.get())
        except ValueError:
            messagebox.showerror("Error", "Width and Height must be numbers.")
            return
        mode = self.mode_var.get()
        self.callback(w, h, mode)
        self.grab_release()
        self.destroy()


class ChapterSelectionDialog(ctk.CTkToplevel):
    def __init__(self, parent, chapters_list, callback):
        super().__init__(parent)
        self.callback = callback
        self.chapters_list = chapters_list
        self.title("Select Chapters for TOC")
        self.geometry("500x600")
        self.transient(parent)
        self.grab_set()
        self.lbl_info = ctk.CTkLabel(self,
                                     text="Uncheck chapters to hide them from the TOC and Progress Bar.\n(They will still be readable in the book)",
                                     wraplength=450)
        self.lbl_info.pack(pady=10)
        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=20)
        ctk.CTkButton(self.btn_frame, text="Select All", width=100, command=self.select_all).pack(side="left", padx=5)
        ctk.CTkButton(self.btn_frame, text="Select None", width=100, command=self.select_none).pack(side="left", padx=5)
        self.scroll_frame = ctk.CTkScrollableFrame(self)
        self.scroll_frame.pack(fill="both", expand=True, padx=20, pady=10)
        self.check_vars = []
        for i, chap in enumerate(chapters_list):
            is_generic = re.match(r"^Section \d+$", chap['title'])
            var = ctk.BooleanVar(value=not is_generic)
            self.check_vars.append(var)
            chk = ctk.CTkCheckBox(self.scroll_frame, text=f"{i + 1}. {chap['title']}", variable=var)
            chk.pack(anchor="w", pady=2, padx=5)
        self.btn_confirm = ctk.CTkButton(self, text="Confirm Selection", command=self.confirm)
        self.btn_confirm.pack(pady=20)

    def select_all(self):
        for var in self.check_vars: var.set(True)

    def select_none(self):
        for var in self.check_vars: var.set(False)

    def confirm(self):
        selected_indices = [i for i, var in enumerate(self.check_vars) if var.get()]
        if not selected_indices:
            if not messagebox.askyesno("Warning",
                                       "No chapters selected for TOC. The book will have no navigation. Continue?"): return
        self.grab_release()
        self.destroy()
        self.callback(selected_indices)


class EpubProcessor:
    def __init__(self):
        self.input_file = ""
        self.raw_chapters = []
        self.book_css = ""
        self.book_images = {}
        self.book_lang = 'en'
        self.cover_image_obj = None
        self.font_path = ""
        self.font_size = FACTORY_DEFAULTS["font_size"]
        self.margin = FACTORY_DEFAULTS["margin"]
        self.line_height = FACTORY_DEFAULTS["line_height"]
        self.font_weight = FACTORY_DEFAULTS["font_weight"]
        self.bottom_padding = FACTORY_DEFAULTS["bottom_padding"]
        self.top_padding = FACTORY_DEFAULTS["top_padding"]
        self.text_align = FACTORY_DEFAULTS["text_align"]
        self.screen_width = DEFAULT_SCREEN_WIDTH
        self.screen_height = DEFAULT_SCREEN_HEIGHT
        self.layout_settings = {}
        self.fitz_docs = []
        self.toc_data_final = []
        self.toc_pages_images = []
        self.page_map = []
        self.total_pages = 0
        self.toc_items_per_page = 18
        self.is_ready = False
        self.global_id_map = {}

    def _split_html_by_toc(self, soup, toc_entries):
        chunks = []
        if len(toc_entries) == 1 and not toc_entries[0][0]:
            return [{'title': toc_entries[0][1], 'soup': soup}]
        split_points = []
        for anchor, title in toc_entries:
            target = None
            if anchor: target = soup.find(id=anchor)
            if target or not anchor: split_points.append({'node': target, 'title': title})
        if not split_points: return [{'title': toc_entries[0][1], 'soup': soup}]

        current_idx = 0
        current_soup = BeautifulSoup("<body></body>", 'html.parser')
        body_children = list(soup.body.children) if soup.body else []
        for child in body_children:
            if isinstance(child, NavigableString) and not child.strip():
                if current_soup.body: current_soup.body.append(child.extract() if hasattr(child, 'extract') else child)
                continue
            if current_idx + 1 < len(split_points):
                next_node = split_points[current_idx + 1]['node']
                is_nested_target = False
                if hasattr(child, 'find_all'):
                    if next_node in child.find_all(): is_nested_target = True
                if next_node and (child == next_node or is_nested_target):
                    chunks.append({'title': split_points[current_idx]['title'], 'soup': current_soup})
                    current_idx += 1
                    current_soup = BeautifulSoup("<body></body>", 'html.parser')
            if current_soup.body: current_soup.body.append(child)
        chunks.append({'title': split_points[current_idx]['title'], 'soup': current_soup})
        return chunks

    def _smart_extract_content(self, elem):
        if elem.name == 'a':
            parent = elem.parent
            if parent and parent.name not in ['body', 'html', 'section']: return parent
            return elem
        if elem.name in ['aside', 'li', 'dd', 'div']: return elem
        text = elem.get_text(strip=True)
        if len(text) > 1: return elem
        parent = elem.parent
        if parent:
            if parent.name in ['body', 'html', 'section']: return elem
            return parent
        return elem

    def _build_global_id_map(self, book):
        id_map = {}
        for item in book.get_items_of_type(ebooklib.ITEM_DOCUMENT):
            try:
                soup = BeautifulSoup(item.get_content(), 'html.parser')
                filename = os.path.basename(item.get_name())
                for elem in soup.find_all(id=True):
                    target_node = self._smart_extract_content(elem)
                    import copy
                    content_node = copy.copy(target_node)
                    original_raw_html = content_node.decode_contents().strip()
                    for a in content_node.find_all('a'):
                        if a.get('role') in ['doc-backlink', 'doc-noteref']:
                            a.decompose()
                            continue
                        text = a.get_text(strip=True)
                        if any(x in text for x in ['↑', 'site', 'back', 'return', '↩']):
                            a.decompose()
                            continue
                        if len(text) < 5 and re.match(r'^[\s\[\(]*\d+[\.\)\]]*$', text):
                            a.decompose()
                            continue
                    final_html = content_node.decode_contents().strip()
                    if not final_html and original_raw_html: final_html = original_raw_html
                    if final_html: id_map[f"{filename}#{elem['id']}"] = final_html
            except Exception:
                pass
        return id_map

    def _inject_inline_footnotes(self, soup, current_filename):
        if not self.global_id_map: return soup
        links = soup.find_all('a', href=True)
        for link in reversed(list(links)):
            raw_href = link['href']
            href = unquote(raw_href)
            text = link.get_text(strip=True)
            if not text and not link.find('sup'): continue
            parent_classes = []
            for parent in link.parents:
                if parent.get('class'): parent_classes.extend(parent.get('class'))
            if any(x in [c.lower() for c in parent_classes] for x in
                   ['footnote', 'endnote', 'reflist', 'bibliography']): continue
            is_footnote = False
            if 'noteref' in link.get('epub:type', '') or link.get('role') == 'doc-noteref': is_footnote = True
            css = link.get('class', [])
            if isinstance(css, list): css = " ".join(css)
            if any(x in css.lower() for x in ['footnote', 'noteref', 'ref']): is_footnote = True
            if not is_footnote and text:
                clean_t = text.strip()
                if re.match(r'^[\(\[]?\d+[\)\]]?$', clean_t) or clean_t == '*':
                    is_footnote = True
                elif re.match(r'^[\(\[]?[ivx]+[\)\]]?$', clean_t.lower()):
                    is_footnote = True
            if not is_footnote: continue
            content = None
            if '#' in href:
                parts = href.rsplit('#', 1)
                href_path = parts[0]
                href_id = parts[1]
                f_name = os.path.basename(href_path) if href_path else current_filename
                key = f"{f_name}#{href_id}"
                content = self.global_id_map.get(key)
                if not content:
                    suffix = f"#{href_id}"
                    for k, v in self.global_id_map.items():
                        if k.endswith(suffix):
                            content = v
                            break
            if content:
                new_marker = soup.new_tag("sup")
                new_marker.string = text if text else "*"
                new_marker['class'] = "fn-marker"
                link.replace_with(new_marker)
                note_box = soup.new_tag("div")
                note_box['class'] = "inline-footnote-box"
                header = soup.new_tag("strong")
                header.string = f"{text}: "
                note_box.append(header)
                content_soup = BeautifulSoup(content, 'html.parser')
                note_box.append(content_soup)
                parent_block = new_marker.find_parent(['p', 'div', 'li', 'h1', 'h2', 'blockquote'])
                if parent_block:
                    parent_block.insert_after(note_box)
                else:
                    new_marker.insert_after(note_box)
        return soup

    def _find_cover_image(self, book):
        try:
            cover_data = book.get_metadata('OPF', 'cover')
            if cover_data:
                cover_id = cover_data[0][1]
                item = book.get_item_with_id(cover_id)
                if item: return Image.open(io.BytesIO(item.get_content()))
        except:
            pass
        for item in book.get_items_of_type(ebooklib.ITEM_IMAGE):
            name = item.get_name().lower()
            if 'cover' in name: return Image.open(io.BytesIO(item.get_content()))
        for item in book.get_items_of_type(ebooklib.ITEM_IMAGE):
            return Image.open(io.BytesIO(item.get_content()))
        return None

    def parse_book_structure(self, input_path):
        self.input_file = input_path
        self.raw_chapters = []
        self.cover_image_obj = None
        try:
            book = epub.read_epub(self.input_file)
        except Exception as e:
            print(f"Error reading EPUB: {e}")
            return False
        self.cover_image_obj = self._find_cover_image(book)
        self.global_id_map = self._build_global_id_map(book)
        try:
            self.book_lang = book.get_metadata('DC', 'language')[0][0]
        except:
            self.book_lang = 'en'
        self.book_images = extract_images_to_base64(book)
        self.book_css = extract_all_css(book)
        toc_mapping = get_official_toc_mapping(book)
        items = [book.get_item_with_id(item_ref[0]) for item_ref in book.spine if
                 isinstance(book.get_item_with_id(item_ref[0]), epub.EpubHtml)]
        for item in items:
            item_filename = os.path.basename(item.get_name())
            raw_html = item.get_content().decode('utf-8', errors='replace')
            soup = BeautifulSoup(raw_html, 'html.parser')
            has_image = bool(soup.find('img'))
            toc_entries = toc_mapping.get(item_filename)
            if toc_entries and len(toc_entries) > 1:
                split_chapters = self._split_html_by_toc(soup, toc_entries)
                for chunk in split_chapters:
                    self.raw_chapters.append(
                        {'title': chunk['title'], 'soup': chunk['soup'], 'has_image': bool(chunk['soup'].find('img')),
                         'filename': item_filename})
            else:
                chapter_title = toc_entries[0][1] if toc_entries else None
                if not chapter_title:
                    for tag in ['h1', 'h2', 'h3']:
                        header = soup.find(tag)
                        if header:
                            t = header.get_text().strip()
                            if t and len(t) < 150:
                                chapter_title = t
                                break
                    if not chapter_title: chapter_title = f"Section {len(self.raw_chapters) + 1}"
                self.raw_chapters.append(
                    {'title': chapter_title, 'soup': soup, 'has_image': has_image, 'filename': item_filename})
        return True

    def render_chapters(self, selected_indices, font_path, font_size, margin, line_height, font_weight, bottom_padding,
                        top_padding, text_align="justify", orientation="Portrait", add_toc=True, show_footnotes=True,
                        layout_settings=None, progress_callback=None):
        self.font_path = font_path if font_path != "DEFAULT" else ""
        self.font_size = font_size
        self.margin = margin
        self.line_height = line_height
        self.font_weight = font_weight
        self.bottom_padding = bottom_padding
        self.top_padding = top_padding
        self.text_align = text_align
        self.layout_settings = layout_settings if layout_settings else {}
        if orientation == "Landscape":
            self.screen_width = DEFAULT_SCREEN_HEIGHT
            self.screen_height = DEFAULT_SCREEN_WIDTH
        else:
            self.screen_width = DEFAULT_SCREEN_WIDTH
            self.screen_height = DEFAULT_SCREEN_HEIGHT
        for doc, _ in self.fitz_docs: doc.close()
        self.fitz_docs, self.page_map = [], []

        if self.font_path:
            variants = get_font_variants(self.font_path)
            font_rules = []
            font_rules.append(
                f"""@font-face {{ font-family: "CustomFont"; src: url("{variants['regular']}"); font-weight: normal; font-style: normal; }}""")
            if variants['italic']: font_rules.append(
                f"""@font-face {{ font-family: "CustomFont"; src: url("{variants['italic']}"); font-weight: normal; font-style: italic; }}""")
            if variants['bold']: font_rules.append(
                f"""@font-face {{ font-family: "CustomFont"; src: url("{variants['bold']}"); font-weight: bold; font-style: normal; }}""")
            if variants['bold_italic']: font_rules.append(
                f"""@font-face {{ font-family: "CustomFont"; src: url("{variants['bold_italic']}"); font-weight: bold; font-style: italic; }}""")
            font_face_rule = "\n".join(font_rules)
            font_family_val = '"CustomFont"'
        else:
            font_face_rule = ""
            font_family_val = "serif"
        patched_css = fix_css_font_paths(self.book_css, font_family_val)
        custom_css = f"""
                <style>
                    {font_face_rule}
                    @page {{ size: {self.screen_width}pt {self.screen_height}pt; margin: 0; }}
                    body {{
                        font-family: {font_family_val} !important;
                        font-size: {self.font_size}pt !important;
                        font-weight: {self.font_weight} !important;
                        line-height: {self.line_height} !important;
                        text-align: {self.text_align} !important;
                        color: black !important;
                        margin: 0 !important;
                        padding: {self.margin}px !important;
                        background-color: white !important;
                        width: 100% !important; 
                        height: 100% !important;
                        overflow-wrap: break-word;
                    }}
                    p, div, li, blockquote, dd, dt {{
                        font-family: inherit !important;
                        font-size: inherit !important;
                        font-weight: inherit !important;
                        line-height: inherit !important;
                        text-align: {self.text_align} !important;
                        color: inherit !important;
                    }}
                    span {{
                        font-family: {font_family_val} !important;
                        font-size: inherit !important;
                        line-height: inherit !important;
                        color: inherit !important;
                    }}
                    img {{ max-width: 95% !important; height: auto !important; display: block; margin: 20px auto !important; }}
                    h1, h2, h3 {{ text-align: center !important; margin-top: 1em; font-weight: {min(900, self.font_weight + 200)} !important; }}
                    .fn-marker {{ font-weight: bold; font-size: 0.7em !important; vertical-align: super; color: solid black !important; }}
                    .inline-footnote-box {{
                        display: block; margin: 15px 0px; padding: 0px 15px;
                        border-left: 4px solid solid black;
                        font-size: {int(self.font_size * 0.85)}pt !important;
                        line-height: {self.line_height} !important;
                    }}
                    .inline-footnote-box p {{ margin: 0 !important; padding: 0 !important; font-size: inherit !important; display: inline; }}
                </style>
                """
        temp_chapter_starts = []
        running_page_count = 0
        render_dir = os.path.dirname(self.input_file)
        temp_html_path = os.path.join(render_dir, "render_temp.html")
        final_toc_titles = []
        total_chaps = len(self.raw_chapters)
        selected_set = set(selected_indices)
        for idx, chapter in enumerate(self.raw_chapters):
            if progress_callback: progress_callback((idx / total_chaps) * 0.9)
            soup = chapter['soup']
            if show_footnotes: soup = self._inject_inline_footnotes(soup, chapter.get('filename', ''))
            for img_tag in soup.find_all('img'):
                src = os.path.basename(img_tag.get('src', ''))
                if src in self.book_images: img_tag['src'] = self.book_images[src]
            soup = hyphenate_html_text(soup, self.book_lang)
            if idx in selected_set:
                temp_chapter_starts.append(running_page_count)
                final_toc_titles.append(chapter['title'])
            body_content = "".join([str(x) for x in soup.body.contents]) if soup.body else str(soup)
            final_html = f"<html lang='{self.book_lang}'><head><style>{patched_css}</style>{custom_css}</head><body>{body_content}</body></html>"
            with open(temp_html_path, "w", encoding="utf-8") as f:
                f.write(final_html)
            doc = fitz.open(temp_html_path)
            rect = fitz.Rect(0, 0, self.screen_width, self.screen_height)
            doc.layout(rect=rect)
            self.fitz_docs.append((doc, chapter['has_image']))
            for i in range(len(doc)): self.page_map.append((len(self.fitz_docs) - 1, i))
            running_page_count += len(doc)
        if os.path.exists(temp_html_path): os.remove(temp_html_path)
        if add_toc and final_toc_titles:
            toc_main_size = self.font_size
            toc_header_size = int(self.font_size * 1.2)
            toc_row_height = int(self.font_size * self.line_height * 1.2)
            toc_header_space = 100 + self.top_padding
            reserved_top = toc_header_space
            reserved_bottom = self.bottom_padding
            available_h = self.screen_height - reserved_top - reserved_bottom
            self.toc_items_per_page = max(1, int(available_h // toc_row_height))
            num_toc_pages = (len(final_toc_titles) + self.toc_items_per_page - 1) // self.toc_items_per_page
            self.toc_data_final = [(t, temp_chapter_starts[i] + num_toc_pages + 1) for i, t in
                                   enumerate(final_toc_titles)]
            self.toc_pages_images = self._render_toc_pages(self.toc_data_final, toc_row_height, toc_main_size,
                                                           toc_header_size)
        else:
            self.toc_data_final = [(t, temp_chapter_starts[i] + 1) for i, t in enumerate(final_toc_titles)]
            self.toc_pages_images = []
        self.total_pages = len(self.toc_pages_images) + len(self.page_map)
        if progress_callback: progress_callback(1.0)
        self.is_ready = True
        return True

    def _get_ui_font(self, size):
        if self.font_path: return get_pil_font(self.font_path, size)
        try:
            return ImageFont.truetype("georgia.ttf", size)
        except:
            return ImageFont.load_default()

    def _render_toc_pages(self, toc_entries, row_height, font_size, header_size):
        pages = []

        def get_dynamic_toc_font(size):
            if self.font_path: return get_pil_font(self.font_path, size)
            try:
                return ImageFont.truetype("georgia.ttf", size)
            except:
                return ImageFont.load_default()

        font_main = get_dynamic_toc_font(font_size)
        font_header = get_dynamic_toc_font(header_size)
        left_margin = 40
        right_margin = 40
        column_gap = 20
        limit = self.toc_items_per_page
        for i in range(0, len(toc_entries), limit):
            chunk = toc_entries[i: i + limit]
            img = Image.new('1', (self.screen_width, self.screen_height), 1)
            draw = ImageDraw.Draw(img)
            header_text = "TABLE OF CONTENTS"
            header_w = font_header.getlength(header_text)
            header_y = 40 + self.top_padding
            draw.text(((self.screen_width - header_w) // 2, header_y), header_text, font=font_header, fill=0)
            line_y = header_y + int(header_size * 1.5)
            draw.line((left_margin, line_y, self.screen_width - right_margin, line_y), fill=0)
            y = line_y + int(font_size * 1.2)
            for title, pg_num in chunk:
                pg_str = str(pg_num)
                pg_w = font_main.getlength(pg_str)
                max_title_w = self.screen_width - left_margin - right_margin - pg_w - column_gap
                display_title = title
                if font_main.getlength(display_title) > max_title_w:
                    while font_main.getlength(display_title + "...") > max_title_w and len(
                            display_title) > 0: display_title = display_title[:-1]
                    display_title += "..."
                draw.text((left_margin, y), display_title, font=font_main, fill=0)
                title_end_x = left_margin + font_main.getlength(display_title) + 5
                dots_end_x = self.screen_width - right_margin - pg_w - 10
                if dots_end_x > title_end_x:
                    dots_text = "." * int((dots_end_x - title_end_x) / font_main.getlength("."))
                    draw.text((title_end_x, y), dots_text, font=font_main, fill=0)
                draw.text((self.screen_width - right_margin - pg_w, y), pg_str, font=font_main, fill=0)
                y += row_height
            pages.append(img)
        return pages

    def _draw_progress_bar(self, draw, y, height, global_page_index):
        if self.total_pages <= 0: return
        show_ticks = self.layout_settings.get("bar_show_ticks", True)
        tick_h = self.layout_settings.get("bar_tick_height", 6)
        show_marker = self.layout_settings.get("bar_show_marker", True)
        marker_r = self.layout_settings.get("bar_marker_radius", 5)
        marker_col_str = self.layout_settings.get("bar_marker_color", "Black")
        marker_fill = (255, 255, 255) if marker_col_str == "White" else (0, 0, 0)
        draw.rectangle([10, y, self.screen_width - 10, y + height], fill=(255, 255, 255), outline=(0, 0, 0))
        if show_ticks:
            bar_center_y = y + (height / 2)
            t_top = bar_center_y - (tick_h / 2)
            t_bot = bar_center_y + (tick_h / 2)
            chapter_pages = [item[1] for item in self.toc_data_final]
            for cp in chapter_pages:
                mx = int(((cp - 1) / self.total_pages) * (self.screen_width - 20)) + 10
                draw.line([mx, t_top, mx, t_bot], fill=(0, 0, 0), width=1)
        curr_page_disp = global_page_index + 1
        bar_width_px = self.screen_width - 20
        fill_width = int((curr_page_disp / self.total_pages) * bar_width_px)
        draw.rectangle([10, y, 10 + fill_width, y + height], fill=(0, 0, 0))
        if show_marker:
            cx = 10 + fill_width
            cy = y + (height / 2)
            draw.ellipse([cx - marker_r, cy - marker_r, cx + marker_r, cy + marker_r], fill=marker_fill,
                         outline=(0, 0, 0))

    def _get_page_text_elements(self, global_page_index):
        page_num_disp = global_page_index + 1
        percent = int((page_num_disp / self.total_pages) * 100)
        current_title = ""
        num_toc = len(self.toc_pages_images)
        if global_page_index < num_toc:
            current_title = "Table of Contents"
            chap_page_disp = f"{global_page_index + 1}/{num_toc}"
        else:
            for title, start_pg in reversed(self.toc_data_final):
                if page_num_disp >= start_pg:
                    current_title = title
                    break
            pm_idx = global_page_index - num_toc
            if 0 <= pm_idx < len(self.page_map):
                doc_idx, page_idx = self.page_map[pm_idx]
                doc_ref = self.fitz_docs[doc_idx][0]
                chap_total = len(doc_ref)
                chap_page_disp = f"{page_idx + 1}/{chap_total}"
            else:
                chap_page_disp = "1/1"
        return {'pagenum': f"{page_num_disp}/{self.total_pages}", 'title': current_title, 'chap_page': chap_page_disp,
                'percent': f"{percent}%"}

    def _draw_text_line(self, draw, y, font, elements_list, align):
        if not elements_list: return
        separator = "   |   "
        margin_x = 15
        if align == "Justify" and len(elements_list) > 1:
            left_item = elements_list[0]
            left_text = left_item[1]
            draw.text((margin_x, y), left_text, font=font, fill=(0, 0, 0))
            right_item = elements_list[-1]
            right_text = right_item[1]
            right_w = font.getlength(right_text)
            draw.text((self.screen_width - margin_x - right_w, y), right_text, font=font, fill=(0, 0, 0))
            if len(elements_list) > 2:
                mid_items = elements_list[1:-1]
                mid_text_parts = []
                title_in_middle = False
                title_original = ""
                for key, txt in mid_items:
                    if key == 'title':
                        title_in_middle = True
                        title_original = txt
                    mid_text_parts.append(txt)
                left_boundary = margin_x + font.getlength(left_text) + 20
                right_boundary = self.screen_width - margin_x - right_w - 20
                max_mid_w = right_boundary - left_boundary
                if max_mid_w < 50: return
                final_mid_text = separator.join(mid_text_parts)
                if title_in_middle and font.getlength(final_mid_text) > max_mid_w:
                    non_title_w = sum([font.getlength(p) for k, p in mid_items if k != 'title'])
                    sep_w_total = font.getlength(separator) * (len(mid_items) - 1) if len(mid_items) > 1 else 0
                    available_for_title = max_mid_w - non_title_w - sep_w_total
                    if available_for_title > 20:
                        trunc_title = title_original
                        while font.getlength(trunc_title + "...") > available_for_title and len(
                                trunc_title) > 0: trunc_title = trunc_title[:-1]
                        trunc_title += "..."
                        new_parts = []
                        for key, txt in mid_items:
                            if key == 'title':
                                new_parts.append(trunc_title)
                            else:
                                new_parts.append(txt)
                        final_mid_text = separator.join(new_parts)
                    else:
                        final_mid_text = separator.join([p for k, p in mid_items if k != 'title'])
                mid_w = font.getlength(final_mid_text)
                mid_x = (self.screen_width - mid_w) // 2
                if mid_x < left_boundary: mid_x = left_boundary
                draw.text((mid_x, y), final_mid_text, font=font, fill=(0, 0, 0))
        else:
            other_text = ""
            title_text = ""
            has_title = False
            non_title_parts = []
            for key, txt in elements_list:
                if key == 'title':
                    title_text = txt
                    has_title = True
                else:
                    non_title_parts.append(txt)
            sep_w = font.getlength(separator) * (len(elements_list) - 1) if len(elements_list) > 1 else 0
            others_w = sum([font.getlength(p) for p in non_title_parts])
            max_screen = self.screen_width - 30
            available_for_title = max_screen - others_w - sep_w
            if has_title:
                if font.getlength(title_text) > available_for_title:
                    while font.getlength(title_text + "...") > available_for_title and len(
                            title_text) > 0: title_text = title_text[:-1]
                    title_text += "..."
            final_parts = []
            for key, txt in elements_list:
                if key == 'title':
                    final_parts.append(title_text)
                else:
                    final_parts.append(txt)
            full_str = separator.join(final_parts)
            str_w = font.getlength(full_str)
            x_pos = 15
            if align == "Center":
                x_pos = (self.screen_width - str_w) // 2
            elif align == "Right":
                x_pos = self.screen_width - 15 - str_w
            draw.text((x_pos, y), full_str, font=font, fill=(0, 0, 0))

    def _get_active_elements(self, bar_role, text_data):
        s = self.layout_settings
        active = []
        for key in ['title', 'pagenum', 'chap_page', 'percent']:
            pos_val = s.get(f"pos_{key}", "Hidden")
            if pos_val == bar_role:
                order = int(s.get(f"order_{key}", 99))
                content = text_data.get(key, "")
                if content: active.append((order, key, content))
        active.sort(key=lambda x: x[0])
        return [(x[1], x[2]) for x in active]

    def _draw_header(self, draw, global_page_index):
        settings = self.layout_settings
        font_size = settings.get("header_font_size", 16)
        margin = settings.get("header_margin", 0)
        align = settings.get("header_align", "Center")
        bar_height = settings.get("bar_height", 4)
        pos_prog = settings.get("pos_progress", "Footer (Below Text)")
        text_data = self._get_page_text_elements(global_page_index)
        elements = self._get_active_elements("Header", text_data)
        has_text = len(elements) > 0
        has_bar = False
        bar_above_text = True
        if pos_prog == "Header (Above Text)":
            has_bar = True;
            bar_above_text = True
        elif pos_prog == "Header (Below Text)":
            has_bar = True;
            bar_above_text = False
        gap = 6
        font_ui = self._get_ui_font(font_size)
        current_y = margin
        if bar_above_text and has_bar:
            self._draw_progress_bar(draw, current_y, bar_height, global_page_index)
            current_y += bar_height + gap
        if has_text:
            self._draw_text_line(draw, current_y, font_ui, elements, align)
            current_y += font_size + gap
        if not bar_above_text and has_bar: self._draw_progress_bar(draw, current_y, bar_height, global_page_index)

    def _draw_footer(self, draw, global_page_index):
        settings = self.layout_settings
        font_size = settings.get("footer_font_size", 16)
        margin = settings.get("footer_margin", 0)
        align = settings.get("footer_align", "Center")
        bar_height = settings.get("bar_height", 4)
        pos_prog = settings.get("pos_progress", "Footer (Below Text)")
        text_data = self._get_page_text_elements(global_page_index)
        elements = self._get_active_elements("Footer", text_data)
        has_text = len(elements) > 0
        has_bar = False
        bar_above_text = True
        if pos_prog == "Footer (Above Text)":
            has_bar = True;
            bar_above_text = True
        elif pos_prog == "Footer (Below Text)":
            has_bar = True;
            bar_above_text = False
        gap = 6
        font_ui = self._get_ui_font(font_size)
        anchor_y = self.screen_height - margin
        text_h = font_size if has_text else 0
        bar_h = bar_height if has_bar else 0
        text_y = 0
        bar_y = 0
        if has_bar and has_text:
            if bar_above_text:
                text_y = anchor_y - text_h
                bar_y = text_y - gap - bar_h
            else:
                bar_y = anchor_y - bar_h
                text_y = bar_y - gap - text_h
        elif has_text:
            text_y = anchor_y - text_h
        elif has_bar:
            bar_y = anchor_y - bar_h
        if has_bar: self._draw_progress_bar(draw, bar_y, bar_height, global_page_index)
        if has_text: self._draw_text_line(draw, text_y, font_ui, elements, align)

    def render_page(self, global_page_index):
        if not self.is_ready: return None

        # --- GET SETTINGS ---
        sharpness_val = self.layout_settings.get("text_blur", 1.0)
        threshold_val = self.layout_settings.get("text_threshold", 130)
        mode = self.layout_settings.get("render_mode", "Threshold")
        white_clip = self.layout_settings.get("white_clip", 220)
        contrast = self.layout_settings.get("contrast", 1.2)

        num_toc = len(self.toc_pages_images)
        footer_padding = max(0, self.bottom_padding)
        header_padding = max(0, self.top_padding)

        # Calculate available height for the actual text content
        content_height = self.screen_height - footer_padding - header_padding
        if content_height < 1: content_height = 1  # Safety

        # --- STEP A: PREPARE CONTENT LAYER ---
        has_image_content = False

        if global_page_index < num_toc:
            # Table of Contents (Pre-rendered)
            img_content = self.toc_pages_images[global_page_index].copy().convert("L")
            is_toc = True
        else:
            is_toc = False
            doc_idx, page_idx = self.page_map[global_page_index - num_toc]
            doc, has_image_content = self.fitz_docs[doc_idx]
            page = doc[page_idx]

            # --- OPTIMIZATION START ---
            # Calculate the exact matrix to fit the target area
            # This replaces rendering at 3.0x and resizing with Lanczos
            src_w = page.rect.width
            src_h = page.rect.height

            # scaling factors
            sx = self.screen_width / src_w
            sy = content_height / src_h

            # Use PyMuPDF to render exactly at target size (Fast C++ rendering)
            mat = fitz.Matrix(sx, sy)
            pix = page.get_pixmap(matrix=mat, alpha=False)

            # Create PIL image from bytes (Zero-copy if possible, or fast copy)
            img_content = Image.frombytes("RGB", [pix.width, pix.height], pix.samples).convert("L")
            # --- OPTIMIZATION END ---

        # Create canvas
        full_page = Image.new("L", (self.screen_width, self.screen_height), 255)
        paste_y = 0 if is_toc else header_padding

        # Determine Paste coordinates (Center horizontally if slight mismatch)
        paste_x = (self.screen_width - img_content.width) // 2
        full_page.paste(img_content, (paste_x, paste_y))

        # --- STEP B: APPLY FILTERS ---
        if not is_toc:
            # 1. DITHER MODE
            if mode == "Dither":
                if contrast != 1.0:
                    full_page = ImageEnhance.Contrast(full_page).enhance(contrast)

                # Apply White Clip (Optimized lookup)
                if white_clip < 255:
                    full_page = full_page.point(lambda p: 255 if p > white_clip else p)

                # Floyd-Steinberg Dithering
                full_page = full_page.convert("1", dither=Image.Dither.FLOYDSTEINBERG).convert("L")

            # 2. THRESHOLD MODE
            else:
                # If page has images, we might want Dither even in Threshold mode
                # (Optional logic, sticking to your default logic here)
                if has_image_content:
                    # Simple contrast boost for mixed content
                    full_page = ImageEnhance.Contrast(full_page).enhance(1.2)
                    full_page = full_page.convert("1", dither=Image.Dither.FLOYDSTEINBERG).convert("L")
                else:
                    # Pure Text Optimization
                    if sharpness_val > 0:
                        enhancer = ImageEnhance.Sharpness(full_page)
                        full_page = enhancer.enhance(1.0 + (sharpness_val * 0.5))

                    # Fast Thresholding
                    full_page = full_page.point(lambda p: 255 if p > threshold_val else 0).convert("L")

        # --- STEP C: UI OVERLAY ---
        # Convert to RGB only for the colored UI drawing
        img_final = full_page.convert("RGB")
        draw = ImageDraw.Draw(img_final)

        if not is_toc:
            # Mask Header/Footer areas with White before drawing text
            if header_padding > 0:
                draw.rectangle([0, 0, self.screen_width, header_padding], fill=(255, 255, 255))
            if footer_padding > 0:
                draw.rectangle([0, self.screen_height - footer_padding, self.screen_width, self.screen_height],
                               fill=(255, 255, 255))

            self._draw_header(draw, global_page_index)
            self._draw_footer(draw, global_page_index)

        return img_final

    def save_xtc(self, out_name, progress_callback=None):
        if not self.is_ready: return

        # Prepare headers
        blob_accumulator = []
        idx_accumulator = []

        # Offset calculation
        # Header (56 bytes) + Index Entries (16 bytes * total_pages)
        data_off_start = 56 + (16 * self.total_pages)
        current_data_offset = data_off_start

        # --- HELPER WORKER FUNCTION ---
        def process_page_worker(page_index):
            # Render
            img_rgb = self.render_page(page_index)
            img_final = img_rgb.convert("1")
            w, h = img_final.size

            # Create XTC Page Block
            # 0x00475458 = 'XTG\0'
            img_bytes = img_final.tobytes()
            xtg_header = struct.pack("<IHHBBIQ",
                                     0x00475458,
                                     w, h,
                                     0, 0,
                                     ((w + 7) // 8) * h,  # Stride/Size calc
                                     0)

            full_page_blob = xtg_header + img_bytes
            return full_page_blob, w, h

        # --- PARALLEL EXECUTION ---
        # Using ThreadPoolExecutor because PyMuPDF fitz.Document isn't easily pickle-able
        # for ProcessPool, but fits handles threading reasonably well for reading.
        max_workers = os.cpu_count()
        results = []

        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all pages
            future_to_page = {executor.submit(process_page_worker, i): i for i in range(self.total_pages)}

            # Process results as they complete, but we need to store them in ORDER
            # So we use map, or just sort results later. Map is cleaner.
            pass

            # Map preserves order
            for i, result in enumerate(executor.map(process_page_worker, range(self.total_pages))):
                if progress_callback:
                    progress_callback((i + 1) / self.total_pages)

                page_blob, w, h = result
                blob_size = len(page_blob)

                # Build Index Entry
                # Offset (Q), Size (I), Width (H), Height (H)
                idx_entry = struct.pack("<QIHH", current_data_offset, blob_size, w, h)

                idx_accumulator.append(idx_entry)
                blob_accumulator.append(page_blob)

                current_data_offset += blob_size

        # --- WRITE FILE ---
        # Header: 'XTX\0' (0x00435458), Version 0x0100
        header = struct.pack("<IHHBBBBIQQQQQ",
                             0x00435458, 0x0100, self.total_pages,
                             0, 0, 0, 0, 0, 0,
                             56,  # Offset to Index
                             data_off_start,  # Offset to Data
                             0, 0)

        with open(out_name, "wb") as f:
            f.write(header)
            for idx_chunk in idx_accumulator:
                f.write(idx_chunk)
            for blob_chunk in blob_accumulator:
                f.write(blob_chunk)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.processor = EpubProcessor()
        self.current_page_index = 0
        self.debounce_timer = None
        self.is_processing = False
        self.selected_chapter_indices = []

        if not os.path.exists(PRESETS_DIR):
            os.makedirs(PRESETS_DIR)

        self.startup_settings = FACTORY_DEFAULTS.copy()
        self.load_startup_defaults()

        self.title("EPUB2XTC")
        self.geometry("1500x950")

        # --- MAIN GRID LAYOUT ---
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)
        self.grid_rowconfigure(0, weight=1)

        # ==============================================================================
        # [LEFT COLUMN] SIDEBAR: FILE, STRUCTURE, LAYOUT, TYPOGRAPHY
        # ==============================================================================
        self.left_sidebar = ctk.CTkFrame(self, width=320, corner_radius=0)
        self.left_sidebar.grid(row=0, column=0, sticky="nsew")

        # 1. FILE SELECTION
        self._add_section_divider(self.left_sidebar)
        ctk.CTkLabel(self.left_sidebar, text="FILE", font=("Arial", 14, "bold")).pack(pady=(15, 5))
        ctk.CTkButton(self.left_sidebar, text="Select EPUB", command=self.select_file, height=30).pack(padx=20, pady=5,
                                                                                                       fill="x")
        self.lbl_file = ctk.CTkLabel(self.left_sidebar, text="No file selected", text_color="gray", height=16)
        self.lbl_file.pack(padx=20, pady=(0, 5))

        self._add_section_divider(self.left_sidebar)

        # 2. STRUCTURE
        ctk.CTkLabel(self.left_sidebar, text="STRUCTURE", font=("Arial", 14, "bold")).pack(pady=5)
        self.var_toc = ctk.BooleanVar(value=self.startup_settings["generate_toc"])
        ctk.CTkCheckBox(self.left_sidebar, text="Generate TOC Pages", variable=self.var_toc,
                        command=self.schedule_update).pack(padx=20, pady=5, anchor="w")
        self.var_footnotes = ctk.BooleanVar(value=self.startup_settings.get("show_footnotes", True))
        ctk.CTkCheckBox(self.left_sidebar, text="Inline Footnotes", variable=self.var_footnotes,
                        command=self.schedule_update).pack(padx=20, pady=5, anchor="w")
        self.btn_chapters = ctk.CTkButton(self.left_sidebar, text="Edit Chapter Visibility", height=30,
                                          command=self.open_chapter_dialog, state="disabled", fg_color="gray")
        self.btn_chapters.pack(padx=20, pady=10, fill="x")

        self._add_section_divider(self.left_sidebar)

        # 3. PAGE LAYOUT
        ctk.CTkLabel(self.left_sidebar, text="PAGE LAYOUT", font=("Arial", 14, "bold")).pack(pady=5)

        row_orient = ctk.CTkFrame(self.left_sidebar, fg_color="transparent")
        row_orient.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(row_orient, text="Orientation:").pack(side="left", anchor="w")
        self.orientation_var = ctk.StringVar(value=self.startup_settings["orientation"])
        ctk.CTkOptionMenu(row_orient, values=["Portrait", "Landscape"], variable=self.orientation_var, width=120,
                          command=self.schedule_update).pack(side="right")

        self.create_slider_group(self.left_sidebar, "lbl_margin", f"Side Margin: {self.startup_settings['margin']}px",
                                 "slider_margin", 0, 100, self.update_margin_label, self.startup_settings['margin'])
        self.create_slider_group(self.left_sidebar, "lbl_top_padding",
                                 f"Top Padding: {self.startup_settings['top_padding']}px", "slider_top_padding", 0, 150,
                                 self.update_top_padding_label, self.startup_settings['top_padding'])
        self.create_slider_group(self.left_sidebar, "lbl_padding",
                                 f"Bottom Padding: {self.startup_settings['bottom_padding']}px", "slider_padding", 0,
                                 150, self.update_padding_label, self.startup_settings['bottom_padding'])

        self._add_section_divider(self.left_sidebar)

        # 4. TYPOGRAPHY
        ctk.CTkLabel(self.left_sidebar, text="TYPOGRAPHY", font=("Arial", 14, "bold")).pack(pady=5)

        row_font_tools = ctk.CTkFrame(self.left_sidebar, fg_color="transparent")
        row_font_tools.pack(fill="x", padx=20, pady=5)

        ctk.CTkLabel(row_font_tools, text="Font:").grid(row=0, column=0, sticky="w")
        self.available_fonts = get_local_fonts()
        self.font_options = ["Default (System)"] + sorted(list(self.available_fonts.keys()))
        self.font_map = self.available_fonts.copy()
        self.font_map["Default (System)"] = "DEFAULT"
        self.font_dropdown = ctk.CTkOptionMenu(row_font_tools, values=self.font_options, width=180,
                                               command=self.on_font_change)
        self.font_dropdown.set(self.startup_settings.get("font_name", "Default (System)"))
        self.font_dropdown.grid(row=0, column=1, sticky="e", padx=(10, 0))

        ctk.CTkLabel(row_font_tools, text="Align:").grid(row=1, column=0, sticky="w", pady=(5, 0))
        self.align_dropdown = ctk.CTkOptionMenu(row_font_tools, values=["justify", "left"], width=180,
                                                command=self.schedule_update)
        self.align_dropdown.set(self.startup_settings["text_align"])
        self.align_dropdown.grid(row=1, column=1, sticky="e", padx=(10, 0), pady=(5, 0))

        self.create_slider_group(self.left_sidebar, "lbl_size", f"Font Size: {self.startup_settings['font_size']}pt",
                                 "slider_size", 12, 48, self.update_size_label, self.startup_settings['font_size'])
        self.create_slider_group(self.left_sidebar, "lbl_weight",
                                 f"Font Weight: {self.startup_settings['font_weight']}", "slider_weight", 100, 900,
                                 self.update_weight_label, self.startup_settings['font_weight'])
        self.create_slider_group(self.left_sidebar, "lbl_line", f"Line Height: {self.startup_settings['line_height']}",
                                 "slider_line", 1.0, 2.5, self.update_line_label, self.startup_settings['line_height'])

        # ==============================================================================
        # [RIGHT COLUMN] SIDEBAR: HEADERS & FOOTERS
        # ==============================================================================
        self.right_sidebar = ctk.CTkScrollableFrame(self, width=320, corner_radius=0)
        self.right_sidebar.grid(row=0, column=2, sticky="nsew")
        self._create_header_footer_controls(self.right_sidebar)

        # ==============================================================================
        # [CENTER COLUMN] MAIN PREVIEW & BOTTOM CONTROL BAR
        # ==============================================================================
        self.center_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.center_frame.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)

        # 1. PREVIEW AREA
        self.preview_container = ctk.CTkFrame(self.center_frame, fg_color="transparent")
        self.preview_container.pack(side="top", fill="both", expand=True, padx=20, pady=20)

        # Create the scrollable frame
        self.preview_scroll = ctk.CTkScrollableFrame(self.preview_container, fg_color="transparent")
        self.preview_scroll.pack(expand=True, fill="both")

        # Hide scrollbar width
        self.preview_scroll._scrollbar.configure(width=0)

        # --- CENTER LOGIC ---
        # Configure the internal 'canvas' grid to center content
        self.preview_scroll.grid_columnconfigure(0, weight=1)
        self.preview_scroll.grid_rowconfigure(0, weight=1)

        self.img_label = ctk.CTkLabel(self.preview_scroll, text="Load EPUB to Preview")
        # Use grid instead of pack to respect the centering weights
        self.img_label.grid(row=0, column=0, sticky="nsew")

        # 2. NAVIGATION
        self.nav = ctk.CTkFrame(self.center_frame, fg_color="transparent")
        self.nav.pack(side="top", fill="x", pady=(0, 10))

        self.create_slider_group(self.center_frame, "lbl_preview_zoom",
                                 f"Preview Zoom: {self.startup_settings['preview_zoom']}", "slider_preview_zoom", 200,
                                 800, self.update_zoom_only, self.startup_settings['preview_zoom'])

        nav_inner = ctk.CTkFrame(self.nav, fg_color="transparent")
        nav_inner.pack()
        ctk.CTkButton(nav_inner, text="< Prev", width=80, height=30, command=self.prev_page).pack(side="left", padx=20)

        center_nav_info = ctk.CTkFrame(nav_inner, fg_color="transparent")
        center_nav_info.pack(side="left", padx=20)
        self.lbl_page = ctk.CTkLabel(center_nav_info, text="0 / 0", font=("Arial", 14))
        self.lbl_page.pack()
        goto_box = ctk.CTkFrame(center_nav_info, fg_color="transparent")
        goto_box.pack(pady=(2, 0))
        self.entry_page = ctk.CTkEntry(goto_box, width=50, height=22, placeholder_text="#")
        self.entry_page.pack(side="left", padx=(0, 5))
        self.entry_page.bind('<Return>', lambda event: self.go_to_page())
        ctk.CTkButton(goto_box, text="Go", width=40, height=22, command=self.go_to_page).pack(side="left")

        ctk.CTkButton(nav_inner, text="Next >", width=80, height=30, command=self.next_page).pack(side="right", padx=20)

        # 3. BOTTOM CONTROL BAR
        self.bottom_bar = ctk.CTkFrame(self.center_frame, height=180, fg_color="#2B2B2B", corner_radius=10)
        self.bottom_bar.pack(side="bottom", fill="x", padx=10, pady=10)
        self.bottom_bar.pack_propagate(False)

        self.bottom_bar.grid_columnconfigure(0, weight=1)  # Render (Left)
        self.bottom_bar.grid_columnconfigure(1, weight=1)  # Actions (Center)
        self.bottom_bar.grid_columnconfigure(2, weight=1)  # Presets (Right)
        self.bottom_bar.grid_rowconfigure(0, weight=1)

        # --- LEFT: RENDER OPTIONS ---
        self.frm_bot_left = ctk.CTkFrame(self.bottom_bar, fg_color="transparent")
        self.frm_bot_left.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        # CONTAINER FOR WIDTH SYNC:
        # Wraps both the Label/Dropdown AND the Sliders so they share width.
        self.render_container = ctk.CTkFrame(self.frm_bot_left, fg_color="transparent")
        self.render_container.pack(anchor="w")

        # Row 1: Label + Dropdown (fill="x")
        row_mode = ctk.CTkFrame(self.render_container, fg_color="transparent")
        row_mode.pack(fill="x", pady=(5, 5))
        ctk.CTkLabel(row_mode, text="Rendering mode:", font=("Arial", 12)).pack(side="left", padx=(0, 10))
        self.render_mode_var = ctk.StringVar(value=self.startup_settings.get("render_mode", "Threshold"))
        ctk.CTkOptionMenu(row_mode, values=["Dither", "Threshold"], variable=self.render_mode_var,
                          width=140, height=28, command=self.toggle_render_controls).pack(side="left")

        # Row 2: Dynamic Sliders Area (fill="x")
        self.frm_render_dynamic = ctk.CTkFrame(self.render_container, fg_color="transparent")
        self.frm_render_dynamic.pack(fill="x", expand=True)

        # Group: Dither
        self.frm_dither = ctk.CTkFrame(self.frm_render_dynamic, fg_color="transparent")
        # Note: fill="x" is important here to stretch the sliders
        self.frm_dither.pack(fill="x")
        self.sld_white = self.create_pixel_slider(self.frm_dither, "lbl_white_clip", "White Clip",
                                                  "slider_white_clip", 150, 255, 220)
        self.sld_contrast = self.create_pixel_slider(self.frm_dither, "lbl_contrast", "Contrast",
                                                     "slider_contrast", 0.5, 2.0, 1.2, is_float=True)

        # Group: Threshold
        self.frm_thresh = ctk.CTkFrame(self.frm_render_dynamic, fg_color="transparent")
        self.frm_thresh.pack(fill="x")
        self.sld_thresh = self.create_pixel_slider(self.frm_thresh, "lbl_threshold", "Threshold",
                                                   "slider_threshold", 50, 200, 130)
        self.sld_blur = self.create_pixel_slider(self.frm_thresh, "lbl_blur", "Definition", "slider_blur",
                                                 0.0, 3.0, 1.0, is_float=True)
        self.toggle_render_controls()

        # --- CENTER: ACTIONS & PROGRESS ---
        self.frm_bot_center = ctk.CTkFrame(self.bottom_bar, fg_color="transparent")
        self.frm_bot_center.grid(row=0, column=1, sticky="nsew", padx=5, pady=10)

        self.btn_row = ctk.CTkFrame(self.frm_bot_center, fg_color="transparent")
        self.btn_row.pack(expand=True)

        ctk.CTkButton(self.btn_row, text="Reset", command=self.reset_to_factory, fg_color="#555", width=70,
                      height=30).pack(side="left", padx=4)
        self.btn_run = ctk.CTkButton(self.btn_row, text="Force Refresh", fg_color="gray", width=110, height=30,
                                     command=self.run_processing)
        self.btn_run.pack(side="left", padx=4)
        self.btn_export = ctk.CTkButton(self.btn_row, text="Export XTC", state="disabled", width=110, height=30,
                                        command=self.export_file)
        self.btn_export.pack(side="left", padx=4)
        self.btn_export_cover = ctk.CTkButton(self.btn_row, text="Cover", state="disabled", width=70, height=30,
                                              command=self.open_cover_export, fg_color="#7B1FA2", hover_color="#4A148C")
        self.btn_export_cover.pack(side="left", padx=4)

        self.progress_bar = ctk.CTkProgressBar(self.frm_bot_center, height=6)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", pady=(10, 2), padx=20)
        self.progress_label = ctk.CTkLabel(self.frm_bot_center, text="Ready", font=("Arial", 11), text_color="gray")
        self.progress_label.pack()

        # --- RIGHT: PRESETS (Stick to Right Edge) ---
        self.frm_bot_right = ctk.CTkFrame(self.bottom_bar, fg_color="transparent")
        self.frm_bot_right.grid(row=0, column=2, sticky="nsew", padx=10, pady=10)

        # anchor="e" forces the content to the East (Right)
        self.preset_container = ctk.CTkFrame(self.frm_bot_right, fg_color="transparent")
        self.preset_container.pack(expand=True, anchor="e")

        # Row 1: Label + Dropdown
        row_preset_select = ctk.CTkFrame(self.preset_container, fg_color="transparent")
        row_preset_select.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(row_preset_select, text="Preset:", font=("Arial", 12)).pack(side="left", padx=(0, 10))
        self.preset_var = ctk.StringVar(value="Select Preset...")
        self.preset_dropdown = ctk.CTkOptionMenu(row_preset_select, variable=self.preset_var, values=[],
                                                 command=self.load_selected_preset, height=28, width=150)
        self.preset_dropdown.pack(side="left")
        self.refresh_presets_list()

        # Row 2 & 3: Buttons
        ctk.CTkButton(self.preset_container, text="Save Preset", command=self.save_new_preset,
                      height=28, fg_color="green").pack(fill="x", pady=3)

        ctk.CTkButton(self.preset_container, text="Set Default", command=self.save_current_as_default,
                      height=28, fg_color="#D35400").pack(fill="x", pady=3)

    def _add_section_divider(self, parent):
        ctk.CTkFrame(parent, height=2, fg_color="#444").pack(fill="x", padx=10, pady=(15, 5))

    # --- SLIDER FOR LAYOUT (Triggers Full Refresh) ---
    def create_slider_group(self, parent, label_attr, label_text, slider_attr, from_val, to_val, cmd, start_val):
        lbl = ctk.CTkLabel(parent, text=label_text)
        lbl.pack(pady=(5, 0), padx=20, anchor="w")
        setattr(self, label_attr, lbl)

        def wrapped_cmd(val):
            cmd(val)
            self.schedule_update()

        sld = ctk.CTkSlider(parent, from_=from_val, to=to_val, command=wrapped_cmd)
        sld.set(start_val)
        sld.pack(fill="x", padx=20, pady=5)
        setattr(self, slider_attr, sld)

    # --- SLIDER FOR PIXELS (Triggers Instant Refresh) ---
    def create_pixel_slider(self, parent, label_attr, label_text, slider_attr, from_val, to_val, default_val,
                            is_float=False):
        lbl = ctk.CTkLabel(parent, text=f"{label_text}: {default_val}")
        lbl.pack(pady=(2, 0), padx=20, anchor="w")
        setattr(self, label_attr, lbl)

        def update_lbl_and_refresh(val):
            v = float(val) if is_float else int(val)
            lbl.configure(text=f"{label_text}: {v:.1f}" if is_float else f"{label_text}: {v}")
            # Instant update for pixel filters
            self.update_render_settings_only()

        sld = ctk.CTkSlider(parent, from_=from_val, to=to_val, command=update_lbl_and_refresh)
        sld.set(self.startup_settings.get(slider_attr.replace("slider_", ""), default_val))
        sld.pack(fill="x", padx=20, pady=2)
        setattr(self, slider_attr, sld)
        return sld

    def toggle_render_controls(self, _=None):
        mode = self.render_mode_var.get()
        self.frm_dither.pack_forget()
        self.frm_thresh.pack_forget()

        if mode == "Dither":
            self.frm_dither.pack(fill="x")
        else:
            self.frm_thresh.pack(fill="x")

        # Trigger instant refresh instead of slow schedule_update
        self.update_render_settings_only()

    def update_render_settings_only(self):
        # Updates settings dict and refreshes ONLY the image without re-layout
        if not self.processor.is_ready: return

        # Update processor settings directly from pixel sliders
        s = self.processor.layout_settings
        s["render_mode"] = self.render_mode_var.get()
        s["white_clip"] = int(self.slider_white_clip.get())
        s["contrast"] = float(self.slider_contrast.get())
        s["text_threshold"] = int(self.slider_threshold.get())
        s["text_blur"] = float(self.slider_blur.get())

        # Refresh current page view immediately
        self.show_page(self.current_page_index)

    def _create_element_control(self, parent, label_text, key_pos, key_order):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", padx=20, pady=2)
        ctk.CTkLabel(f, text=label_text).pack(side="left")
        val_pos = self.startup_settings.get(key_pos, "Hidden")
        var_pos = ctk.StringVar(value=val_pos)
        setattr(self, f"var_{key_pos}", var_pos)
        val_order = str(self.startup_settings.get(key_order, 1))
        var_order = ctk.StringVar(value=val_order)
        setattr(self, f"var_{key_order}", var_order)
        entry_order = ctk.CTkEntry(f, textvariable=var_order, width=30)
        entry_order.pack(side="right", padx=(5, 0))
        ctk.CTkOptionMenu(f, variable=var_pos, values=["Header", "Footer", "Hidden"], width=100,
                          command=self.schedule_update).pack(side="right")
        ctk.CTkLabel(f, text="Ord:", font=("Arial", 10), text_color="gray").pack(side="right", padx=(10, 2))

    def _create_header_footer_controls(self, parent):
        self._add_section_divider(parent)
        ctk.CTkLabel(parent, text="HEADER & FOOTER", font=("Arial", 13, "bold")).pack(pady=(5, 10))
        self._create_element_control(parent, "Chapter Title:", "pos_title", "order_title")
        self._create_element_control(parent, "Page Number:", "pos_pagenum", "order_pagenum")
        self._create_element_control(parent, "Chapter Page (X/Y):", "pos_chap_page", "order_chap_page")
        self._create_element_control(parent, "Reading %:", "pos_percent", "order_percent")
        ctk.CTkLabel(parent, text="Progress Bar Settings", font=("Arial", 13, "bold")).pack(pady=(15, 5))
        f3 = ctk.CTkFrame(parent, fg_color="transparent")
        f3.pack(fill="x", padx=20, pady=2)
        ctk.CTkLabel(f3, text="Position:").pack(side="left")
        self.var_pos_progress = ctk.StringVar(value=self.startup_settings.get("pos_progress", "Footer (Below Text)"))
        ctk.CTkOptionMenu(f3, variable=self.var_pos_progress,
                          values=["Header (Above Text)", "Header (Below Text)", "Footer (Above Text)",
                                  "Footer (Below Text)", "Hidden"], width=160, command=self.schedule_update).pack(
            side="right")
        row_checks = ctk.CTkFrame(parent, fg_color="transparent")
        row_checks.pack(fill="x", padx=20, pady=(10, 10))
        self.var_bar_ticks = ctk.BooleanVar(value=self.startup_settings.get("bar_show_ticks", True))
        ctk.CTkCheckBox(row_checks, text="Show Ticks", variable=self.var_bar_ticks, command=self.schedule_update,
                        width=100).pack(side="left")
        self.var_bar_marker = ctk.BooleanVar(value=self.startup_settings.get("bar_show_marker", True))
        ctk.CTkCheckBox(row_checks, text="Show Marker", variable=self.var_bar_marker,
                        command=self.schedule_update).pack(side="right", padx=(10, 0), pady=10)
        f4 = ctk.CTkFrame(parent, fg_color="transparent")
        f4.pack(fill="x", padx=20, pady=2)
        ctk.CTkLabel(f4, text="Marker Color:").pack(side="left")
        self.var_marker_color = ctk.StringVar(value=self.startup_settings.get("bar_marker_color", "Black"))
        ctk.CTkOptionMenu(f4, variable=self.var_marker_color, values=["Black", "White"], width=100,
                          command=self.schedule_update).pack(side="right")
        self.create_slider_group(parent, "lbl_bar_thick", f"Bar Thickness: {self.startup_settings['bar_height']}px",
                                 "slider_bar_thick", 1, 10, self.update_bar_thick_label,
                                 self.startup_settings.get('bar_height', 4))
        self.create_slider_group(parent, "lbl_marker_size",
                                 f"Marker Radius: {self.startup_settings['bar_marker_radius']}px", "slider_marker_size",
                                 2, 10,
                                 lambda v: self.update_generic_label("lbl_marker_size", f"Marker Radius: {int(v)}px"),
                                 self.startup_settings.get('bar_marker_radius', 5))
        self.create_slider_group(parent, "lbl_tick_height",
                                 f"Tick Height: {self.startup_settings['bar_tick_height']}px", "slider_tick_height", 2,
                                 20, lambda v: self.update_generic_label("lbl_tick_height", f"Tick Height: {int(v)}px"),
                                 self.startup_settings.get('bar_tick_height', 6))
        ctk.CTkLabel(parent, text="Header Styles", font=("Arial", 13, "bold")).pack(pady=(15, 5))
        h_align = ctk.CTkFrame(parent, fg_color="transparent")
        h_align.pack(fill="x", padx=20, pady=2)
        ctk.CTkLabel(h_align, text="Alignment:").pack(side="left")
        self.var_header_align = ctk.StringVar(value=self.startup_settings.get("header_align", "Center"))
        ctk.CTkOptionMenu(h_align, variable=self.var_header_align, values=["Left", "Center", "Right", "Justify"],
                          width=100, command=self.schedule_update).pack(side="right")
        self.create_slider_group(parent, "lbl_header_size", f"Font Size: {self.startup_settings['header_font_size']}",
                                 "slider_header_size", 8, 30,
                                 lambda v: self.update_generic_label("lbl_header_size", f"Font Size: {int(v)}"),
                                 self.startup_settings.get('header_font_size', 16))
        self.create_slider_group(parent, "lbl_header_margin",
                                 f"Header Y-Offset: {self.startup_settings['header_margin']}px", "slider_header_margin",
                                 0, 80, lambda v: self.update_generic_label("lbl_header_margin",
                                                                            f"Header Y-Offset: {int(v)}px"),
                                 self.startup_settings.get('header_margin', 10))
        ctk.CTkLabel(parent, text="Footer Styles", font=("Arial", 13, "bold")).pack(pady=(15, 5))
        f_align = ctk.CTkFrame(parent, fg_color="transparent")
        f_align.pack(fill="x", padx=20, pady=2)
        ctk.CTkLabel(f_align, text="Alignment:").pack(side="left")
        self.var_footer_align = ctk.StringVar(value=self.startup_settings.get("footer_align", "Center"))
        ctk.CTkOptionMenu(f_align, variable=self.var_footer_align, values=["Left", "Center", "Right", "Justify"],
                          width=100, command=self.schedule_update).pack(side="right")
        self.create_slider_group(parent, "lbl_footer_size", f"Font Size: {self.startup_settings['footer_font_size']}",
                                 "slider_footer_size", 8, 30,
                                 lambda v: self.update_generic_label("lbl_footer_size", f"Font Size: {int(v)}"),
                                 self.startup_settings.get('footer_font_size', 16))
        self.create_slider_group(parent, "lbl_footer_margin",
                                 f"Footer Y-Offset: {self.startup_settings['footer_margin']}px", "slider_footer_margin",
                                 0, 80, lambda v: self.update_generic_label("lbl_footer_margin",
                                                                            f"Footer Y-Offset: {int(v)}px"),
                                 self.startup_settings.get('footer_margin', 10))

    def gather_current_ui_settings(self):
        def get_order(var_name):
            try:
                return int(getattr(self, var_name).get())
            except:
                return 99

        settings = {
            "font_size": int(self.slider_size.get()),
            "font_weight": int(self.slider_weight.get()),
            "line_height": float(self.slider_line.get()),
            "margin": int(self.slider_margin.get()),
            "top_padding": int(self.slider_top_padding.get()),
            "bottom_padding": int(self.slider_padding.get()),
            "orientation": self.orientation_var.get(),
            "text_align": self.align_dropdown.get(),
            "font_name": self.font_dropdown.get(),
            "preview_zoom": int(self.slider_preview_zoom.get()),
            "generate_toc": self.var_toc.get(),
            "show_footnotes": self.var_footnotes.get(),
            "bar_height": int(self.slider_bar_thick.get()),
            "pos_title": self.var_pos_title.get(),
            "pos_pagenum": self.var_pos_pagenum.get(),
            "pos_chap_page": self.var_pos_chap_page.get(),
            "pos_percent": self.var_pos_percent.get(),
            "pos_progress": self.var_pos_progress.get(),
            "order_title": get_order("var_order_title"),
            "order_pagenum": get_order("var_order_pagenum"),
            "order_chap_page": get_order("var_order_chap_page"),
            "order_percent": get_order("var_order_percent"),
            "bar_show_ticks": self.var_bar_ticks.get(),
            "bar_show_marker": self.var_bar_marker.get(),
            "bar_marker_color": self.var_marker_color.get(),
            "bar_marker_radius": int(self.slider_marker_size.get()),
            "bar_tick_height": int(self.slider_tick_height.get()),
            "header_align": self.var_header_align.get(),
            "header_font_size": int(self.slider_header_size.get()),
            "header_margin": int(self.slider_header_margin.get()),
            "footer_align": self.var_footer_align.get(),
            "footer_font_size": int(self.slider_footer_size.get()),
            "footer_margin": int(self.slider_footer_margin.get()),

            # --- NEW KEYS ---
            "render_mode": self.render_mode_var.get(),
            "white_clip": int(self.slider_white_clip.get()),
            "contrast": float(self.slider_contrast.get()),
            "text_threshold": int(self.slider_threshold.get()),
            "text_blur": float(self.slider_blur.get()),
        }
        return settings

    def apply_settings_dict(self, s):
        defaults = FACTORY_DEFAULTS.copy()
        defaults.update(s)
        s = defaults
        self.slider_size.set(s['font_size'])
        self.slider_weight.set(s['font_weight'])
        self.slider_line.set(s['line_height'])
        self.slider_margin.set(s['margin'])
        self.slider_top_padding.set(s['top_padding'])
        self.slider_padding.set(s['bottom_padding'])
        self.orientation_var.set(s['orientation'])
        self.align_dropdown.set(s['text_align'])
        self.slider_preview_zoom.set(s['preview_zoom'])
        self.var_toc.set(s['generate_toc'])
        self.var_footnotes.set(s.get('show_footnotes', True))
        self.slider_bar_thick.set(s['bar_height'])
        self.var_pos_title.set(s['pos_title'])
        self.var_pos_pagenum.set(s['pos_pagenum'])
        self.var_pos_chap_page.set(s['pos_chap_page'])
        self.var_pos_percent.set(s['pos_percent'])
        self.var_pos_progress.set(s['pos_progress'])
        self.var_order_title.set(str(s['order_title']))
        self.var_order_pagenum.set(str(s['order_pagenum']))
        self.var_order_chap_page.set(str(s['order_chap_page']))
        self.var_order_percent.set(str(s['order_percent']))
        self.var_bar_ticks.set(s['bar_show_ticks'])
        self.var_bar_marker.set(s['bar_show_marker'])
        self.var_marker_color.set(s['bar_marker_color'])
        self.slider_marker_size.set(s['bar_marker_radius'])
        self.slider_tick_height.set(s['bar_tick_height'])
        self.var_header_align.set(s['header_align'])
        self.slider_header_size.set(s['header_font_size'])
        self.slider_header_margin.set(s['header_margin'])
        self.var_footer_align.set(s['footer_align'])
        self.slider_footer_size.set(s['footer_font_size'])
        self.slider_footer_margin.set(s['footer_margin'])

        # --- NEW KEYS ---
        self.render_mode_var.set(s.get("render_mode", "Threshold"))
        self.slider_white_clip.set(s.get("white_clip", 220))
        self.slider_contrast.set(s.get("contrast", 1.2))
        self.slider_threshold.set(s.get("text_threshold", 130))
        self.slider_blur.set(s.get("text_blur", 1.0))

        # Update text labels
        self.lbl_white_clip.configure(text=f"White Clipping: {int(s.get('white_clip', 220))}")
        self.lbl_contrast.configure(text=f"Contrast Boost: {s.get('contrast', 1.2):.1f}")
        self.lbl_threshold.configure(text=f"Contrast Threshold: {int(s.get('text_threshold', 130))}")
        self.lbl_blur.configure(text=f"Definition: {s.get('text_blur', 1.0):.1f}")

        if s["font_name"] in self.font_options:
            self.font_dropdown.set(s["font_name"])
            self.processor.font_path = self.font_map[s["font_name"]]
        else:
            self.font_dropdown.set("Default (System)")
            self.processor.font_path = self.font_map["Default (System)"]
        self.update_size_label(s['font_size'])
        self.update_weight_label(s['font_weight'])
        self.update_line_label(s['line_height'])
        self.update_margin_label(s['margin'])
        self.update_top_padding_label(s['top_padding'])
        self.update_padding_label(s['bottom_padding'])
        self.update_zoom_only(s['preview_zoom'])
        self.update_bar_thick_label(s['bar_height'])
        self.update_generic_label("lbl_header_size", f"Font Size: {s['header_font_size']}")
        self.update_generic_label("lbl_header_margin", f"Header Y-Offset: {s['header_margin']}px")
        self.update_generic_label("lbl_footer_size", f"Font Size: {s['footer_font_size']}")
        self.update_generic_label("lbl_footer_margin", f"Footer Y-Offset: {s['footer_margin']}px")
        self.update_generic_label("lbl_marker_size", f"Marker Radius: {s['bar_marker_radius']}px")
        self.update_generic_label("lbl_tick_height", f"Tick Height: {s['bar_tick_height']}px")

        self.toggle_render_controls()
        self.schedule_update()

    def update_size_label(self, value):
        self.lbl_size.configure(text=f"Font Size: {int(value)}pt")

    def update_weight_label(self, value):
        self.lbl_weight.configure(text=f"Font Weight: {int(value)}")

    def update_line_label(self, value):
        self.lbl_line.configure(text=f"Line Height: {value:.1f}")

    def update_margin_label(self, value):
        self.lbl_margin.configure(text=f"Margin: {int(value)}px")

    def update_padding_label(self, value):
        self.lbl_padding.configure(text=f"Bottom Padding: {int(value)}px")

    def update_top_padding_label(self, value):
        self.lbl_top_padding.configure(text=f"Top Padding: {int(value)}px")

    def update_bar_thick_label(self, value):
        self.lbl_bar_thick.configure(text=f"Bar Thickness: {int(value)}px")

    def update_zoom_only(self, value):
        self.lbl_preview_zoom.configure(text=f"Preview Zoom: {int(value)}")

    def update_generic_label(self, attr_name, text):
        getattr(self, attr_name).configure(text=text)

    def on_font_change(self, choice):
        self.processor.font_path = self.font_map[choice]
        self.schedule_update()

    def load_startup_defaults(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r") as f:
                    saved = json.load(f)
                    self.startup_settings.update(saved)
            except:
                pass

    def save_current_as_default(self):
        data = self.gather_current_ui_settings()
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump(data, f, indent=4)
            messagebox.showinfo("Saved", "Current settings saved as startup default.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save settings: {e}")

    def refresh_presets_list(self):
        files = glob.glob(os.path.join(PRESETS_DIR, "*.json"))
        names = [os.path.basename(f).replace(".json", "") for f in files]
        names.sort()
        if not names:
            self.preset_dropdown.configure(values=["No Presets"])
            self.preset_var.set("No Presets")
        else:
            self.preset_dropdown.configure(values=names)
            self.preset_var.set("Select Preset...")

    def save_new_preset(self):
        name = simpledialog.askstring("New Preset", "Enter a name for this preset:")
        if not name: return
        safe_name = "".join(x for x in name if x.isalnum() or x in " -_")
        path = os.path.join(PRESETS_DIR, f"{safe_name}.json")
        data = self.gather_current_ui_settings()
        try:
            with open(path, "w") as f:
                json.dump(data, f, indent=4)
            self.refresh_presets_list()
            self.preset_var.set(safe_name)
            messagebox.showinfo("Success", f"Preset '{safe_name}' saved.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def load_selected_preset(self, choice):
        if choice == "No Presets" or choice == "Select Preset...": return
        path = os.path.join(PRESETS_DIR, f"{choice}.json")
        if os.path.exists(path):
            try:
                with open(path, "r") as f:
                    data = json.load(f)
                self.apply_settings_dict(data)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load preset: {e}")

    def reset_to_factory(self):
        if messagebox.askyesno("Confirm", "Reset everything to Factory Defaults?"): self.apply_settings_dict(
            FACTORY_DEFAULTS)

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("EPUB", "*.epub")])
        if path:
            self.processor.input_file = path
            self.current_page_index = 0
            self.lbl_file.configure(text=os.path.basename(path))
            self.progress_label.configure(text="Parsing structure...")
            threading.Thread(target=self._task_parse_structure).start()

    def _task_parse_structure(self):
        success = self.processor.parse_book_structure(self.processor.input_file)
        self.after(0, lambda: self._on_structure_parsed(success))

    def _on_structure_parsed(self, success):
        if not success:
            messagebox.showerror("Error", "Failed to parse EPUB.")
            return
        self.btn_chapters.configure(state="normal", fg_color="#3B8ED0")
        self.open_chapter_dialog()

    def open_chapter_dialog(self):
        if not self.selected_chapter_indices: self.selected_chapter_indices = list(
            range(len(self.processor.raw_chapters)))
        ChapterSelectionDialog(self, self.processor.raw_chapters, self._on_chapters_selected)

    def _on_chapters_selected(self, selected_indices):
        self.selected_chapter_indices = selected_indices
        self.run_processing()

    def schedule_update(self, _=None):
        if not self.processor.input_file: return

        # FIXED: Check if self.timer is valid before cancelling
        if self.debounce_timer is not None:
            try:
                self.after_cancel(self.debounce_timer)
            except ValueError:
                pass

        self.progress_label.configure(text="Waiting for changes...")
        self.debounce_timer = self.after(800, self.run_processing)

    def update_progress_ui(self, val, stage_text="Processing"):
        self.after(0, lambda: self.progress_bar.set(val))
        self.after(0, lambda: self.progress_label.configure(text=f"{stage_text}: {int(val * 100)}%"))

    def run_processing(self):
        if not self.processor.input_file: return
        if self.is_processing: return
        if self.selected_chapter_indices is None: self.selected_chapter_indices = list(
            range(len(self.processor.raw_chapters)))
        layout_settings = self.gather_current_ui_settings()
        self.is_processing = True
        self.btn_run.configure(state="disabled", text="Rendering...", fg_color="orange")
        self.progress_label.configure(text="Starting layout...")
        threading.Thread(target=lambda: self._task_render(layout_settings)).start()

    def _task_render(self, layout_settings):
        success = self.processor.render_chapters(
            self.selected_chapter_indices,
            self.processor.font_path,
            int(self.slider_size.get()),
            int(self.slider_margin.get()),
            float(self.slider_line.get()),
            int(self.slider_weight.get()),
            int(self.slider_padding.get()),
            int(self.slider_top_padding.get()),
            text_align=self.align_dropdown.get(),
            orientation=self.orientation_var.get(),
            add_toc=self.var_toc.get(),
            show_footnotes=self.var_footnotes.get(),
            layout_settings=layout_settings,
            progress_callback=lambda v: self.update_progress_ui(v, "Layout")
        )
        self.after(0, lambda: self._done(success))

    def _done(self, success):
        self.is_processing = False
        self.btn_run.configure(state="normal", text="Force Refresh", fg_color="green")
        if success:
            self.btn_export.configure(state="normal")
            self.btn_export_cover.configure(state="normal")
            old_idx = self.current_page_index
            total = self.processor.total_pages
            new_idx = min(old_idx, total - 1)
            self.show_page(new_idx)
        else:
            messagebox.showerror("Error", "Processing failed.")

    def show_page(self, idx):
        # Safety: check if img_label exists and processor is ready
        if not hasattr(self, 'img_label') or not self.processor.is_ready:
            return
        if not self.processor.is_ready: return
        self.current_page_index = idx
        img = self.processor.render_page(idx)
        base_size = int(self.slider_preview_zoom.get())
        if img.width > img.height:
            target_h = base_size
            aspect_ratio = img.width / img.height
            target_w = int(target_h * aspect_ratio)
        else:
            target_w = base_size
            aspect_ratio = img.height / img.width
            target_h = int(target_w * aspect_ratio)
        ctk_img = ctk.CTkImage(light_image=img, size=(target_w, target_h))
        self.img_label.configure(image=ctk_img, text="")
        # Ensure it stays centered
        self.img_label.grid(row=0, column=0, sticky="nsew")
        self.lbl_page.configure(text=f"Page {idx + 1} / {self.processor.total_pages}")

    def go_to_page(self):
        if not self.processor.is_ready: return
        txt = self.entry_page.get()
        if not txt.isdigit(): return
        target = int(txt)
        target = max(1, min(target, self.processor.total_pages))
        self.show_page(target - 1)
        self.entry_page.delete(0, 'end')

    def prev_page(self):
        self.show_page(max(0, self.current_page_index - 1))

    def next_page(self):
        self.show_page(min(self.processor.total_pages - 1, self.current_page_index + 1))

    def export_file(self):
        path = filedialog.asksaveasfilename(defaultextension=".xtc")
        if path: threading.Thread(target=lambda: self._run_export(path)).start()

    def _run_export(self, path):
        self.processor.save_xtc(path, progress_callback=lambda v: self.update_progress_ui(v, "Exporting"))
        self.after(0, lambda: messagebox.showinfo("Success", "XTC file saved."))

    def open_cover_export(self):
        if not self.processor.cover_image_obj:
            messagebox.showinfo("Info", "No cover image found in this book.")
            return
        CoverExportDialog(self, self._process_cover_export)

    def _process_cover_export(self, w, h, mode):
        path = filedialog.asksaveasfilename(defaultextension=".bmp", filetypes=[("Bitmap", "*.bmp")])
        if not path: return
        threading.Thread(target=lambda: self._run_cover_export(path, w, h, mode)).start()

    def _run_cover_export(self, path, w, h, mode):
        try:
            self.update_progress_ui(0.5, "Exporting Cover")
            img = self.processor.cover_image_obj.convert("RGB")
            if "Stretch" in mode:
                img = img.resize((w, h), Image.Resampling.LANCZOS)
            elif "Fit" in mode:
                img = ImageOps.pad(img, (w, h), color="white", centering=(0.5, 0.5))
            else:
                img = ImageOps.fit(img, (w, h), centering=(0.5, 0.5))
            img = img.convert("L")
            img = ImageEnhance.Contrast(img).enhance(1.3)
            img = ImageEnhance.Brightness(img).enhance(1.05)
            img = img.convert("1", dither=Image.Dither.FLOYDSTEINBERG)
            img.save(path, format="BMP")
            self.update_progress_ui(1.0, "Done")
            self.after(0, lambda: messagebox.showinfo("Success", f"Cover saved to {os.path.basename(path)}"))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error", f"Failed to export cover: {e}"))


if __name__ == "__main__":
    app = App()
    app.mainloop()
