import os
import sys
import struct
import fitz  # PyMuPDF
from PIL import Image, ImageEnhance, ImageDraw, ImageFont, ImageOps
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
}

SETTINGS_FILE = "default_settings.json"
PRESETS_DIR = "presets"


# --- UTILITY FUNCTIONS ---
def fix_css_font_paths(css_text, target_font_family="'CustomFont'"):
    if target_font_family is None:
        return css_text
    css_text = re.sub(r'font-family\s*:\s*[^;!]+', f'font-family: {target_font_family}', css_text)
    return css_text


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

    def process_toc_item(item):
        if isinstance(item, tuple):
            if len(item) > 1 and isinstance(item[1], list):
                for sub in item[1]: process_toc_item(sub)
        elif isinstance(item, epub.Link):
            href_clean = item.href.split('#')[0]
            filename = os.path.basename(href_clean)
            if filename not in mapping:
                mapping[filename] = item.title

    for item in book.toc:
        process_toc_item(item)

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
        if text_node.parent.name in ['script', 'style', 'head', 'title', 'meta']:
            continue
        if not text_node.strip():
            continue

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

    fonts = []
    if os.path.exists(fonts_dir):
        for f in os.listdir(fonts_dir):
            if f.lower().endswith((".ttf", ".otf")):
                fonts.append(os.path.abspath(os.path.join(fonts_dir, f)))
    return sorted(fonts)


# --- POPUP DIALOG FOR COVER EXPORT ---
class CoverExportDialog(ctk.CTkToplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback = callback
        self.title("Export Cover Settings")
        self.geometry("400x350")
        self.transient(parent)
        self.grab_set()

        ctk.CTkLabel(self, text="Export Cover as BMP", font=("Arial", 16, "bold")).pack(pady=15)

        # Dimensions
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

        # Scaling Mode
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


# --- POPUP DIALOG FOR CHAPTER SELECTION ---
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
                                       "No chapters selected for TOC. The book will have no navigation. Continue?"):
                return
        self.grab_release()
        self.destroy()
        self.callback(selected_indices)


# --- PROCESSING ENGINE ---

class EpubProcessor:
    def __init__(self):
        self.input_file = ""
        self.raw_chapters = []
        self.book_css = ""
        self.book_images = {}
        self.book_lang = 'en'
        self.cover_image_obj = None  # Store the PIL image of the cover

        # Default values
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

        # --- LAYOUT SETTINGS ---
        self.layout_settings = {}

        self.fitz_docs = []
        self.toc_data_final = []
        self.toc_pages_images = []
        self.page_map = []
        self.total_pages = 0
        self.toc_items_per_page = 18
        self.is_ready = False
        self.global_id_map = {}

    def _smart_extract_content(self, elem):
        if elem.name == 'a':
            parent = elem.parent
            if parent and parent.name not in ['body', 'html', 'section']:
                return parent
            return elem
        if elem.name in ['aside', 'li', 'dd', 'div']:
            return elem
        text = elem.get_text(strip=True)
        if len(text) > 1:
            return elem
        parent = elem.parent
        if parent:
            if parent.name in ['body', 'html', 'section']:
                return elem
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
                    if not final_html and original_raw_html:
                        final_html = original_raw_html
                    if final_html:
                        id_map[f"{filename}#{elem['id']}"] = final_html
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
                   ['footnote', 'endnote', 'reflist', 'bibliography']):
                continue

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
        # 1. Try Metadata
        try:
            cover_data = book.get_metadata('OPF', 'cover')
            if cover_data:
                cover_id = cover_data[0][1]
                item = book.get_item_with_id(cover_id)
                if item:
                    return Image.open(io.BytesIO(item.get_content()))
        except:
            pass

        # 2. Try Item Names
        for item in book.get_items_of_type(ebooklib.ITEM_IMAGE):
            name = item.get_name().lower()
            if 'cover' in name:
                return Image.open(io.BytesIO(item.get_content()))

        # 3. Fallback: First image
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

        # Extract Cover immediately
        self.cover_image_obj = self._find_cover_image(book)

        self.global_id_map = self._build_global_id_map(book)
        try:
            self.book_lang = book.get_metadata('DC', 'language')[0][0]
        except:
            self.book_lang = 'en'

        self.book_images = extract_images_to_base64(book)
        self.book_css = extract_all_css(book)
        toc_mapping = get_official_toc_mapping(book)

        items = [book.get_item_with_id(item_ref[0]) for item_ref in book.spine
                 if isinstance(book.get_item_with_id(item_ref[0]), epub.EpubHtml)]

        for idx, item in enumerate(items):
            item_name = item.get_name()
            item_filename = os.path.basename(item_name)
            raw_html = item.get_content().decode('utf-8', errors='replace')
            soup = BeautifulSoup(raw_html, 'html.parser')
            text_content = soup.get_text().strip()
            has_image = bool(soup.find('img'))

            chapter_title = toc_mapping.get(item_filename)
            if not chapter_title:
                for tag in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                    header = soup.find(tag)
                    if header:
                        t = header.get_text().strip()
                        if t and len(t) < 150:
                            chapter_title = t
                            break
                if not chapter_title:
                    class_titles = soup.find_all(['p', 'div', 'span'], class_=re.compile(r'title|header|chapter', re.I))
                    for ct in class_titles:
                        t = ct.get_text().strip()
                        if t and len(t) < 150:
                            chapter_title = t
                            break
                if not chapter_title and soup.title and soup.title.string:
                    t = soup.title.string.strip()
                    if t: chapter_title = t
                if not chapter_title:
                    chapter_title = f"Section {len(self.raw_chapters) + 1}"

            self.raw_chapters.append({
                'title': chapter_title,
                'soup': soup,
                'has_image': has_image,
                'filename': item_filename
            })

        return True

    def render_chapters(self, selected_indices, font_path, font_size, margin, line_height, font_weight,
                        bottom_padding, top_padding, text_align="justify", orientation="Portrait", add_toc=True,
                        show_footnotes=True, layout_settings=None, progress_callback=None):

        self.font_path = font_path if font_path != "DEFAULT" else ""
        self.font_size = font_size
        self.margin = margin
        self.line_height = line_height
        self.font_weight = font_weight
        self.bottom_padding = bottom_padding
        self.top_padding = top_padding
        self.text_align = text_align

        # Store Layout Settings for page rendering
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
            css_font_path = self.font_path.replace("\\", "/")
            font_face_rule = f'@font-face {{ font-family: "CustomFont"; src: url("{css_font_path}"); }}'
            font_family_val = '"CustomFont"'
        else:
            font_face_rule = ""
            font_family_val = "serif"

        patched_css = fix_css_font_paths(self.book_css, font_family_val)

        custom_css = f"""
                <style>
                    {font_face_rule}
                    @page {{ size: {self.screen_width}pt {self.screen_height}pt; margin: 0; }}

                    body, p, div, span, li, blockquote, dd, dt {{
                        font-family: {font_family_val} !important;
                        font-size: {self.font_size}pt !important;
                        font-weight: {self.font_weight} !important;
                        line-height: {self.line_height} !important;
                        text-align: {self.text_align} !important;
                        color: black !important;
                        overflow-wrap: break-word;
                    }}

                    body {{
                        margin: 0 !important;
                        padding: {self.margin}px !important;
                        background-color: white !important;
                        width: 100% !important; 
                        height: 100% !important;
                    }}

                    img {{ max-width: 95% !important; height: auto !important; display: block; margin: 20px auto !important; }}
                    h1, h2, h3 {{ text-align: center !important; margin-top: 1em; font-weight: {min(900, self.font_weight + 200)} !important; }}

                    .fn-marker {{
                        font-weight: bold;
                        font-size: 0.7em !important;
                        vertical-align: super;
                        color: solid black !important;
                    }}

                    .inline-footnote-box {{
                        display: block;
                        margin: 15px 0px;
                        padding: 0px 15px;
                        border-left: 4px solid solid black;
                        font-size: {int(self.font_size * 0.85)}pt !important;
                        line-height: {self.line_height} !important;
                    }}

                    .inline-footnote-box p {{
                         margin: 0 !important;
                         padding: 0 !important;
                         font-size: inherit !important;
                         display: inline;
                    }}
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
            if show_footnotes:
                soup = self._inject_inline_footnotes(soup, chapter.get('filename', ''))

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
        if self.font_path:
            return get_pil_font(self.font_path, size)
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
                    while font_main.getlength(display_title + "...") > max_title_w and len(display_title) > 0:
                        display_title = display_title[:-1]
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
            draw.ellipse([cx - marker_r, cy - marker_r, cx + marker_r, cy + marker_r],
                         fill=marker_fill, outline=(0, 0, 0))

    def _get_page_text_elements(self, global_page_index):
        # 1. Page Number & Percentage
        page_num_disp = global_page_index + 1
        percent = int((page_num_disp / self.total_pages) * 100)

        # 2. Title & Chapter Stats
        current_title = ""

        # Calculation for Chapter Page X of Y
        num_toc = len(self.toc_pages_images)
        if global_page_index < num_toc:
            # We are in TOC
            current_title = "Table of Contents"
            chap_page_disp = f"{global_page_index + 1}/{num_toc}"
        else:
            # We are in a real chapter
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
                chap_page_disp = "1/1"  # Fallback

        return {
            'pagenum': f"{page_num_disp}/{self.total_pages}",
            'title': current_title,
            'chap_page': chap_page_disp,
            'percent': f"{percent}%"
        }

    def _draw_text_line(self, draw, y, font, elements_list, align):
        """
        Draws a list of text elements on the same line.
        elements_list: list of (key, text) tuples.
        """
        if not elements_list: return

        # Separator definition
        separator = "   |   "
        margin_x = 15

        if align == "Justify" and len(elements_list) > 1:
            # JUSTIFY MODE Logic
            # 1. First element goes Left
            left_item = elements_list[0]
            left_text = left_item[1]
            draw.text((margin_x, y), left_text, font=font, fill=(0, 0, 0))

            # 2. Last element goes Right
            right_item = elements_list[-1]
            right_text = right_item[1]
            right_w = font.getlength(right_text)
            draw.text((self.screen_width - margin_x - right_w, y), right_text, font=font, fill=(0, 0, 0))

            # 3. Middle Elements (if any)
            if len(elements_list) > 2:
                mid_items = elements_list[1:-1]

                # Check if Title is in middle to apply truncation
                mid_text_parts = []
                title_in_middle = False
                title_original = ""

                for key, txt in mid_items:
                    if key == 'title':
                        title_in_middle = True
                        title_original = txt
                    mid_text_parts.append(txt)

                # Calculate available width for middle
                # Left Width + Gap
                left_boundary = margin_x + font.getlength(left_text) + 20
                # Right Width + Gap
                right_boundary = self.screen_width - margin_x - right_w - 20

                max_mid_w = right_boundary - left_boundary
                if max_mid_w < 50: return  # No space for middle

                # --- CHANGE 1: Use separator for joining ---
                final_mid_text = separator.join(mid_text_parts)

                if title_in_middle and font.getlength(final_mid_text) > max_mid_w:
                    # Truncation needed on Title
                    # Calculate how much space non-title parts take
                    non_title_w = sum([font.getlength(p) for k, p in mid_items if k != 'title'])

                    # --- CHANGE 2: Calculate separator width instead of space width ---
                    # Calculate total width of separators
                    sep_w_total = font.getlength(separator) * (len(mid_items) - 1) if len(mid_items) > 1 else 0

                    available_for_title = max_mid_w - non_title_w - sep_w_total

                    if available_for_title > 20:
                        trunc_title = title_original
                        while font.getlength(trunc_title + "...") > available_for_title and len(trunc_title) > 0:
                            trunc_title = trunc_title[:-1]
                        trunc_title += "..."

                        # Rebuild string with truncated title
                        new_parts = []
                        for key, txt in mid_items:
                            if key == 'title':
                                new_parts.append(trunc_title)
                            else:
                                new_parts.append(txt)
                        # --- CHANGE 3: Join with separator ---
                        final_mid_text = separator.join(new_parts)
                    else:
                        # Not enough space for title at all
                        # --- CHANGE 4: Join with separator ---
                        final_mid_text = separator.join([p for k, p in mid_items if k != 'title'])

                # Draw Center Anchored
                mid_w = font.getlength(final_mid_text)
                mid_x = (self.screen_width - mid_w) // 2

                # Ensure it doesn't overlap left
                if mid_x < left_boundary: mid_x = left_boundary

                draw.text((mid_x, y), final_mid_text, font=font, fill=(0, 0, 0))

        else:
            # Standard Join Mode (Left/Center/Right) for non-Justify
            # ... (Existing logic remains exactly the same) ...
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
                    while font.getlength(title_text + "...") > available_for_title and len(title_text) > 0:
                        title_text = title_text[:-1]
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
                if content:
                    active.append((order, key, content))
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

        if not bar_above_text and has_bar:
            self._draw_progress_bar(draw, current_y, bar_height, global_page_index)

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

        if has_bar:
            self._draw_progress_bar(draw, bar_y, bar_height, global_page_index)

        if has_text:
            self._draw_text_line(draw, text_y, font_ui, elements, align)

    def render_page(self, global_page_index):
        if not self.is_ready: return None
        num_toc = len(self.toc_pages_images)

        footer_padding = max(0, self.bottom_padding)
        header_padding = max(0, self.top_padding)
        content_height = self.screen_height - footer_padding - header_padding

        if global_page_index < num_toc:
            img = self.toc_pages_images[global_page_index].copy().convert("RGB")
        else:
            doc_idx, page_idx = self.page_map[global_page_index - num_toc]
            doc, has_image = self.fitz_docs[doc_idx]
            page = doc[page_idx]
            mat = fitz.Matrix(DEFAULT_RENDER_SCALE, DEFAULT_RENDER_SCALE)
            pix = page.get_pixmap(matrix=mat, alpha=False)

            img_content = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_content = img_content.resize((self.screen_width, content_height), Image.Resampling.LANCZOS).convert("L")

            img = Image.new("RGB", (self.screen_width, self.screen_height), (255, 255, 255))
            img.paste(img_content, (0, header_padding))

            if has_image:
                img = img.convert("L")
                img = ImageEnhance.Contrast(ImageEnhance.Brightness(img).enhance(1.15)).enhance(1.4)
                img = img.convert("1", dither=Image.Dither.FLOYDSTEINBERG)
            else:
                img = img.convert("L")
                img = ImageEnhance.Contrast(img).enhance(2.0).point(lambda p: 255 if p > 140 else 0, mode='1')
            img = img.convert("RGB")

        draw = ImageDraw.Draw(img)
        if header_padding > 0:
            draw.rectangle([0, 0, self.screen_width, header_padding], fill=(255, 255, 255))
        if footer_padding > 0:
            draw.rectangle([0, self.screen_height - footer_padding, self.screen_width, self.screen_height],
                           fill=(255, 255, 255))

        self._draw_header(draw, global_page_index)
        self._draw_footer(draw, global_page_index)

        return img

    def save_xtc(self, out_name, progress_callback=None):
        if not self.is_ready: return
        blob, idx = bytearray(), bytearray()
        data_off = 56 + (16 * self.total_pages)

        for i in range(self.total_pages):
            if progress_callback: progress_callback((i + 1) / self.total_pages)

            # --- OLD LINE (CAUSES BAD IMAGES) ---
            # img = self.render_page(i).convert("L").point(lambda p: 255 if p > 128 else 0, mode='1')

            # --- NEW LINE (FIXES IMAGES) ---
            # 1. Get the page
            # 2. Convert to Grayscale ("L")
            # 3. Convert to 1-bit ("1") using Floyd-Steinberg Dithering
            img = self.render_page(i).convert("L").convert("1", dither=Image.Dither.FLOYDSTEINBERG)

            w, h = img.size
            xtg = struct.pack("<IHHBBIQ", 0x00475458, w, h, 0, 0, ((w + 7) // 8) * h, 0) + img.tobytes()
            idx.extend(struct.pack("<QIHH", data_off + len(blob), len(xtg), w, h))
            blob.extend(xtg)

        header = struct.pack("<IHHBBBBIQQQQQ", 0x00435458, 0x0100, self.total_pages, 0, 0, 0, 0, 0, 0, 56, data_off, 0,
                             0)
        with open(out_name, "wb") as f:
            f.write(header + idx + blob)


# --- GUI APPLICATION ---

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.processor = EpubProcessor()
        self.current_page_index = 0
        self.debounce_timer = None
        self.is_processing = False
        self.selected_chapter_indices = []

        # --- PRESETS SETUP ---
        if not os.path.exists(PRESETS_DIR):
            os.makedirs(PRESETS_DIR)

        self.startup_settings = FACTORY_DEFAULTS.copy()
        self.load_startup_defaults()

        self.title("EPUB2XTC")
        self.geometry("1400x1000")

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)
        self.grid_rowconfigure(0, weight=1)

        # ==========================================
        # LEFT SIDEBAR
        # ==========================================
        self.left_sidebar = ctk.CTkScrollableFrame(self, width=320, corner_radius=0)
        self.left_sidebar.grid(row=0, column=0, sticky="nsew")

        # --- SECTION: FILE & PRESETS ---
        self._add_section_divider(self.left_sidebar)  # Divider first
        ctk.CTkLabel(self.left_sidebar, text="FILE & PRESETS", font=("Arial", 14, "bold")).pack(pady=2)

        ctk.CTkButton(self.left_sidebar, text="Select EPUB", command=self.select_file, height=30).pack(padx=10, pady=5,
                                                                                                       fill="x")
        self.lbl_file = ctk.CTkLabel(self.left_sidebar, text="No file selected", text_color="gray", height=16)
        self.lbl_file.pack(padx=10, pady=(0, 10))

        self.frm_presets = ctk.CTkFrame(self.left_sidebar, fg_color="#333")
        self.frm_presets.pack(fill="x", padx=10, pady=5)

        self.preset_var = ctk.StringVar(value="Select Preset...")
        self.preset_dropdown = ctk.CTkOptionMenu(self.frm_presets, variable=self.preset_var, values=[],
                                                 command=self.load_selected_preset, height=28)
        self.preset_dropdown.pack(fill="x", padx=10, pady=10)
        self.refresh_presets_list()

        btn_preset_row = ctk.CTkFrame(self.frm_presets, fg_color="transparent")
        btn_preset_row.pack(fill="x", padx=5, pady=(0, 10))
        ctk.CTkButton(btn_preset_row, text="Save", command=self.save_new_preset, width=60, height=28,
                      fg_color="green").pack(side="left", padx=(5, 5), expand=True, fill="x")
        ctk.CTkButton(btn_preset_row, text="Default", command=self.save_current_as_default, width=60, height=28,
                      fg_color="#D35400").pack(side="left", padx=(0, 5), expand=True, fill="x")

        # --- SECTION: STRUCTURE ---
        self._add_section_divider(self.left_sidebar)
        ctk.CTkLabel(self.left_sidebar, text="STRUCTURE", font=("Arial", 14, "bold")).pack(pady=2)

        self.var_toc = ctk.BooleanVar(value=self.startup_settings["generate_toc"])
        ctk.CTkCheckBox(self.left_sidebar, text="Generate TOC Pages", variable=self.var_toc,
                        command=self.schedule_update).pack(padx=20, pady=5, anchor="w")
        self.var_footnotes = ctk.BooleanVar(value=self.startup_settings.get("show_footnotes", True))
        ctk.CTkCheckBox(self.left_sidebar, text="Inline Footnotes", variable=self.var_footnotes,
                        command=self.schedule_update).pack(padx=20, pady=5, anchor="w")
        self.btn_chapters = ctk.CTkButton(self.left_sidebar, text="Edit Chapter Visibility", height=30,
                                          command=self.open_chapter_dialog, state="disabled", fg_color="gray")
        self.btn_chapters.pack(padx=20, pady=10, fill="x")

        # --- SECTION: PAGE LAYOUT ---
        self._add_section_divider(self.left_sidebar)
        ctk.CTkLabel(self.left_sidebar, text="PAGE LAYOUT", font=("Arial", 14, "bold")).pack(pady=2)

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

        # --- SECTION: TYPOGRAPHY ---
        self._add_section_divider(self.left_sidebar)
        ctk.CTkLabel(self.left_sidebar, text="TYPOGRAPHY", font=("Arial", 14, "bold")).pack(pady=2)

        # --- Font Family Row ---
        row_font = ctk.CTkFrame(self.left_sidebar, fg_color="transparent")
        row_font.pack(fill="x", padx=20, pady=2)  # Reduced pady to match marker row style

        ctk.CTkLabel(row_font, text="Font Family:").pack(side="left")

        self.available_fonts = get_local_fonts()
        self.font_options = ["Default (System)"] + [os.path.basename(f) for f in self.available_fonts]
        self.font_map = {os.path.basename(f): f for f in self.available_fonts}
        self.font_map["Default (System)"] = "DEFAULT"

        self.font_dropdown = ctk.CTkOptionMenu(row_font, values=self.font_options, width=150,
                                               command=self.on_font_change)
        self.font_dropdown.set(self.startup_settings.get("font_name", "Default (System)"))
        self.font_dropdown.pack(side="right")

        # --- Alignment Row ---
        row_align = ctk.CTkFrame(self.left_sidebar, fg_color="transparent")
        row_align.pack(fill="x", padx=20, pady=2)

        ctk.CTkLabel(row_align, text="Alignment:").pack(side="left")

        self.align_dropdown = ctk.CTkOptionMenu(row_align, values=["justify", "left"], width=150,
                                                command=self.schedule_update)
        self.align_dropdown.set(self.startup_settings["text_align"])
        self.align_dropdown.pack(side="right")

        self.create_slider_group(self.left_sidebar, "lbl_size", f"Font Size: {self.startup_settings['font_size']}pt",
                                 "slider_size", 12, 48, self.update_size_label, self.startup_settings['font_size'])
        self.create_slider_group(self.left_sidebar, "lbl_weight",
                                 f"Weight: {self.startup_settings['font_weight']}", "slider_weight", 100, 900,
                                 self.update_weight_label, self.startup_settings['font_weight'])
        self.create_slider_group(self.left_sidebar, "lbl_line", f"Line Height: {self.startup_settings['line_height']}",
                                 "slider_line", 1.0, 2.5, self.update_line_label, self.startup_settings['line_height'])

        # ==========================================
        # RIGHT SIDEBAR
        # ==========================================
        self.right_sidebar = ctk.CTkScrollableFrame(self, width=350, corner_radius=0)
        self.right_sidebar.grid(row=0, column=2, sticky="nsew")

        # --- NEW HEADER & FOOTER ELEMENTS ---
        self._create_header_footer_controls(self.right_sidebar)

        # ==========================================
        # CENTER: PREVIEW & ACTIONS
        # ==========================================
        self.preview_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.preview_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

        # 1. Image
        self.img_label = ctk.CTkLabel(self.preview_frame, text="Load EPUB to Preview")
        self.img_label.pack(expand=True, fill="both", pady=(0, 10))

        # 2. Page Navigation
        self.nav = ctk.CTkFrame(self.preview_frame, fg_color="transparent")
        self.nav.pack(side="top", fill="x", pady=10)

        self.create_slider_group(self.preview_frame, "lbl_preview_zoom",
                                 f"Preview Zoom: {self.startup_settings['preview_zoom']}", "slider_preview_zoom", 200,
                                 800, self.update_zoom_only, self.startup_settings['preview_zoom'])

        ctk.CTkButton(self.nav, text="< Prev", width=80, height=30, command=self.prev_page).pack(side="left", padx=20)

        self.center_nav = ctk.CTkFrame(self.nav, fg_color="transparent")
        self.center_nav.pack(side="left", expand=True, fill="both")
        self.lbl_page = ctk.CTkLabel(self.center_nav, text="0 / 0", font=("Arial", 14))
        self.lbl_page.pack(side="top", pady=(0, 2))

        self.goto_frame = ctk.CTkFrame(self.center_nav, fg_color="transparent")
        self.goto_frame.pack(side="top")
        self.entry_page = ctk.CTkEntry(self.goto_frame, width=50, height=24, placeholder_text="#")
        self.entry_page.pack(side="left", padx=(0, 5))
        self.entry_page.bind('<Return>', lambda event: self.go_to_page())
        self.btn_go = ctk.CTkButton(self.goto_frame, text="Go", width=40, height=24, command=self.go_to_page)
        self.btn_go.pack(side="left")

        ctk.CTkButton(self.nav, text="Next >", width=80, height=30, command=self.next_page).pack(side="right", padx=20)

        # 3. Bottom Action Bar
        self.bottom_panel = ctk.CTkFrame(self.preview_frame, fg_color="#2B2B2B", height=150)
        self.bottom_panel.pack(side="bottom", fill="x", pady=(20, 0))
        self.bottom_panel.pack_propagate(False)

        # Container for centering
        self.action_container = ctk.CTkFrame(self.bottom_panel, fg_color="transparent")
        self.action_container.place(relx=0.5, rely=0.5, anchor="center")

        # Buttons
        self.btn_row = ctk.CTkFrame(self.action_container, fg_color="transparent")
        self.btn_row.pack(pady=(0, 15))

        ctk.CTkButton(self.btn_row, text="Reset", command=self.reset_to_factory,
                      fg_color="#555", width=90, height=34).pack(side="left", padx=10)

        self.btn_run = ctk.CTkButton(self.btn_row, text="Force Refresh", fg_color="gray", width=150, height=34,
                                     command=self.run_processing)
        self.btn_run.pack(side="left", padx=10)

        self.btn_export = ctk.CTkButton(self.btn_row, text="Export XTC", state="disabled", width=150, height=34,
                                        command=self.export_file)
        self.btn_export.pack(side="left", padx=10)

        # ADDED EXPORT COVER BUTTON
        self.btn_export_cover = ctk.CTkButton(self.btn_row, text="Export Cover", state="disabled", width=150, height=34,
                                              command=self.open_cover_export, fg_color="#7B1FA2", hover_color="#4A148C")
        self.btn_export_cover.pack(side="left", padx=10)

        # Progress
        self.progress_bar = ctk.CTkProgressBar(self.action_container, width=450, height=12)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=(5, 5))

        self.progress_label = ctk.CTkLabel(self.action_container, text="Ready", font=("Arial", 12))
        self.progress_label.pack()

    def _add_section_divider(self, parent):
        ctk.CTkFrame(parent, height=2, fg_color="#444").pack(fill="x", padx=10, pady=(15, 5))

    def _create_element_control(self, parent, label_text, key_pos, key_order):
        # Frame row
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", padx=20, pady=2)

        # Label (Left)
        ctk.CTkLabel(f, text=label_text).pack(side="left")

        # Dropdown (Right, then Spinbox)
        val_pos = self.startup_settings.get(key_pos, "Hidden")
        var_pos = ctk.StringVar(value=val_pos)
        setattr(self, f"var_{key_pos}", var_pos)

        val_order = str(self.startup_settings.get(key_order, 1))
        var_order = ctk.StringVar(value=val_order)
        setattr(self, f"var_{key_order}", var_order)

        # Order Input
        entry_order = ctk.CTkEntry(f, textvariable=var_order, width=30)
        entry_order.pack(side="right", padx=(5, 0))

        # Dropdown
        ctk.CTkOptionMenu(f, variable=var_pos, values=["Header", "Footer", "Hidden"], width=100,
                          command=self.schedule_update).pack(side="right")

        # Small Label "Ord:"
        ctk.CTkLabel(f, text="Ord:", font=("Arial", 10), text_color="gray").pack(side="right", padx=(10, 2))

    def _create_header_footer_controls(self, parent):
        # Divider Line
        self._add_section_divider(parent)

        # SECTION TITLE
        ctk.CTkLabel(parent, text="HEADER & FOOTER", font=("Arial", 13, "bold")).pack(pady=(5, 10))

        # --- ELEMENT VISIBILITY & ORDERING ---
        self._create_element_control(parent, "Chapter Title:", "pos_title", "order_title")
        self._create_element_control(parent, "Page Number:", "pos_pagenum", "order_pagenum")
        self._create_element_control(parent, "Chapter Page (X/Y):", "pos_chap_page", "order_chap_page")
        self._create_element_control(parent, "Reading %:", "pos_percent", "order_percent")

        # --- PROGRESS BAR SETTINGS ---
        ctk.CTkLabel(parent, text="Progress Bar Settings", font=("Arial", 13, "bold")).pack(pady=(15, 5))

        # Position
        f3 = ctk.CTkFrame(parent, fg_color="transparent")
        f3.pack(fill="x", padx=20, pady=2)
        ctk.CTkLabel(f3, text="Position:").pack(side="left")
        self.var_pos_progress = ctk.StringVar(value=self.startup_settings.get("pos_progress", "Footer (Below Text)"))
        ctk.CTkOptionMenu(f3, variable=self.var_pos_progress,
                          values=["Header (Above Text)", "Header (Below Text)", "Footer (Above Text)",
                                  "Footer (Below Text)", "Hidden"],
                          width=160,
                          command=self.schedule_update).pack(side="right")

        # Checkboxes
        row_checks = ctk.CTkFrame(parent, fg_color="transparent")
        row_checks.pack(fill="x", padx=20, pady=(10, 10))

        self.var_bar_ticks = ctk.BooleanVar(value=self.startup_settings.get("bar_show_ticks", True))
        ctk.CTkCheckBox(row_checks, text="Show Ticks", variable=self.var_bar_ticks,
                        command=self.schedule_update, width=100).pack(side="left")

        self.var_bar_marker = ctk.BooleanVar(value=self.startup_settings.get("bar_show_marker", True))
        ctk.CTkCheckBox(row_checks, text="Show Marker", variable=self.var_bar_marker,
                        command=self.schedule_update).pack(side="right", padx=(10, 0), pady=10)

        # Marker Color
        f4 = ctk.CTkFrame(parent, fg_color="transparent")
        f4.pack(fill="x", padx=20, pady=2)
        ctk.CTkLabel(f4, text="Marker Color:").pack(side="left")
        self.var_marker_color = ctk.StringVar(value=self.startup_settings.get("bar_marker_color", "Black"))
        ctk.CTkOptionMenu(f4, variable=self.var_marker_color, values=["Black", "White"], width=100,
                          command=self.schedule_update).pack(side="right")

        # Sliders
        self.create_slider_group(parent, "lbl_bar_thick",
                                 f"Bar Thickness: {self.startup_settings['bar_height']}px", "slider_bar_thick", 1,
                                 10, self.update_bar_thick_label, self.startup_settings.get('bar_height', 4))

        self.create_slider_group(parent, "lbl_marker_size",
                                 f"Marker Radius: {self.startup_settings['bar_marker_radius']}px", "slider_marker_size",
                                 2,
                                 10,
                                 lambda v: self.update_generic_label("lbl_marker_size", f"Marker Radius: {int(v)}px"),
                                 self.startup_settings.get('bar_marker_radius', 5))

        self.create_slider_group(parent, "lbl_tick_height",
                                 f"Tick Height: {self.startup_settings['bar_tick_height']}px", "slider_tick_height", 2,
                                 20, lambda v: self.update_generic_label("lbl_tick_height", f"Tick Height: {int(v)}px"),
                                 self.startup_settings.get('bar_tick_height', 6))

        # --- HEADER STYLING ---
        ctk.CTkLabel(parent, text="Header Styles", font=("Arial", 13, "bold")).pack(pady=(15, 5))

        # Align
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
                                 0, 80,
                                 lambda v: self.update_generic_label("lbl_header_margin",
                                                                     f"Header Y-Offset: {int(v)}px"),
                                 self.startup_settings.get('header_margin', 10))

        # --- FOOTER STYLING ---
        ctk.CTkLabel(parent, text="Footer Styles", font=("Arial", 13, "bold")).pack(pady=(15, 5))

        # Align
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
                                 0, 80,
                                 lambda v: self.update_generic_label("lbl_footer_margin",
                                                                     f"Footer Y-Offset: {int(v)}px"),
                                 self.startup_settings.get('footer_margin', 10))

    def create_slider_group(self, parent, label_attr, label_text, slider_attr, from_val, to_val, cmd, start_val):
        lbl = ctk.CTkLabel(parent, text=label_text)
        lbl.pack(pady=(5, 0), padx=20, anchor="w")
        setattr(self, label_attr, lbl)

        # Wrapper to update label AND schedule update
        def wrapped_cmd(val):
            cmd(val)
            self.schedule_update()

        sld = ctk.CTkSlider(parent, from_=from_val, to=to_val, command=wrapped_cmd)
        sld.set(start_val)
        sld.pack(fill="x", padx=20, pady=5)
        setattr(self, slider_attr, sld)

    def gather_current_ui_settings(self):
        # Safe int conversion for order
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

            # Visibility
            "pos_title": self.var_pos_title.get(),
            "pos_pagenum": self.var_pos_pagenum.get(),
            "pos_chap_page": self.var_pos_chap_page.get(),
            "pos_percent": self.var_pos_percent.get(),
            "pos_progress": self.var_pos_progress.get(),

            # Ordering
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

        if s["font_name"] in self.font_options:
            self.font_dropdown.set(s["font_name"])
            self.processor.font_path = self.font_map[s["font_name"]]
        else:
            self.font_dropdown.set("Default (System)")
            self.processor.font_path = self.font_map["Default (System)"]

        # Update text labels
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

        self.schedule_update()

    # --- UI UPDATERS ---
    def update_size_label(self, value):
        self.lbl_size.configure(text=f"Font Size: {int(value)}pt");

    def update_weight_label(self, value):
        self.lbl_weight.configure(text=f"Font Weight: {int(value)}");

    def update_line_label(self, value):
        self.lbl_line.configure(text=f"Line Height: {value:.1f}");

    def update_margin_label(self, value):
        self.lbl_margin.configure(text=f"Margin: {int(value)}px");

    def update_padding_label(self, value):
        self.lbl_padding.configure(text=f"Bottom Padding: {int(value)}px");

    def update_top_padding_label(self, value):
        self.lbl_top_padding.configure(text=f"Top Padding: {int(value)}px");

    def update_bar_thick_label(self, value):
        self.lbl_bar_thick.configure(text=f"Bar Thickness: {int(value)}px");

    def update_zoom_only(self, value):
        self.lbl_preview_zoom.configure(text=f"Preview Zoom: {int(value)}");

    def update_generic_label(self, attr_name, text):
        getattr(self, attr_name).configure(text=text)

    def on_font_change(self, choice):
        self.processor.font_path = self.font_map[choice];
        self.schedule_update()

    # --- PRESET/FILE LOGIC ---
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
        if messagebox.askyesno("Confirm", "Reset everything to Factory Defaults?"):
            self.apply_settings_dict(FACTORY_DEFAULTS)

    # --- MAIN FUNCTIONALITY ---
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
        if not self.selected_chapter_indices:
            self.selected_chapter_indices = list(range(len(self.processor.raw_chapters)))
        ChapterSelectionDialog(self, self.processor.raw_chapters, self._on_chapters_selected)

    def _on_chapters_selected(self, selected_indices):
        self.selected_chapter_indices = selected_indices
        self.run_processing()

    def schedule_update(self, _=None):
        if not self.processor.input_file: return
        if self.debounce_timer is not None: self.after_cancel(self.debounce_timer)
        self.progress_label.configure(text="Waiting for changes...")
        self.debounce_timer = self.after(800, self.run_processing)

    def update_progress_ui(self, val, stage_text="Processing"):
        self.after(0, lambda: self.progress_bar.set(val))
        self.after(0, lambda: self.progress_label.configure(text=f"{stage_text}: {int(val * 100)}%"))

    def run_processing(self):
        if not self.processor.input_file: return
        if self.is_processing: return
        if self.selected_chapter_indices is None:
            self.selected_chapter_indices = list(range(len(self.processor.raw_chapters)))

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
            self.btn_export_cover.configure(state="normal")  # Enable Cover Button
            old_idx = self.current_page_index
            total = self.processor.total_pages
            new_idx = min(old_idx, total - 1)
            self.show_page(new_idx)
        else:
            messagebox.showerror("Error", "Processing failed.")

    def show_page(self, idx):
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

    # --- COVER EXPORT LOGIC ---
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

            # 1. Resize / Scale Logic
            if "Stretch" in mode:
                img = img.resize((w, h), Image.Resampling.LANCZOS)
            elif "Fit" in mode:
                img = ImageOps.pad(img, (w, h), color="white", centering=(0.5, 0.5))
            else:  # Crop to Fill (Best)
                img = ImageOps.fit(img, (w, h), centering=(0.5, 0.5))

            # 2. Prepare for E-Ink (Grayscale + Contrast)
            img = img.convert("L")

            # Enhancing contrast slightly usually makes dithering look much sharper on E-Ink
            img = ImageEnhance.Contrast(img).enhance(1.3)
            img = ImageEnhance.Brightness(img).enhance(1.05)

            # 3. Convert to 1-bit B/W using Floyd-Steinberg Dithering
            # PIL's default .convert('1') applies Floyd-Steinberg dithering automatically
            img = img.convert("1", dither=Image.Dither.FLOYDSTEINBERG)

            img.save(path, format="BMP")
            self.update_progress_ui(1.0, "Done")
            self.after(0, lambda: messagebox.showinfo("Success", f"Cover saved to {os.path.basename(path)}"))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error", f"Failed to export cover: {e}"))


if __name__ == "__main__":
    app = App()

    app.mainloop()
