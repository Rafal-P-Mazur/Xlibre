import os
import sys
import struct
import fitz  # PyMuPDF
from PIL import Image, ImageEnhance, ImageDraw, ImageFont
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

# --- CONFIGURATION DEFAULTS ---
DEFAULT_SCREEN_WIDTH = 480
DEFAULT_SCREEN_HEIGHT = 800
DEFAULT_RENDER_SCALE = 3.0

# --- FACTORY DEFAULTS ---
FACTORY_DEFAULTS = {
    "font_size": 22,
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

    # --- FOOTER DEFAULTS ---
    "footer_visible": True,
    "footer_font_size": 16,
    "footer_show_progress": True,
    "footer_show_pagenum": True,
    "footer_show_title": True,
    "footer_text_pos": "Text Above Bar",
    "footer_bar_height": 4,
    "footer_bottom_margin": 10
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
            clean_href = item.href.split('#')[0]
            mapping[clean_href] = item.title

    for item in book.toc: process_toc_item(item)
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
            var = ctk.BooleanVar(value=True)
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

        # Footer Settings
        self.footer_visible = FACTORY_DEFAULTS["footer_visible"]
        self.footer_font_size = FACTORY_DEFAULTS["footer_font_size"]
        self.footer_show_progress = FACTORY_DEFAULTS["footer_show_progress"]
        self.footer_show_pagenum = FACTORY_DEFAULTS["footer_show_pagenum"]
        self.footer_show_title = FACTORY_DEFAULTS["footer_show_title"]
        self.footer_text_pos = FACTORY_DEFAULTS["footer_text_pos"]
        self.footer_bar_height = FACTORY_DEFAULTS["footer_bar_height"]
        self.footer_bottom_margin = FACTORY_DEFAULTS["footer_bottom_margin"]

        self.fitz_docs = []
        self.toc_data_final = []
        self.toc_pages_images = []
        self.page_map = []
        self.total_pages = 0
        self.toc_items_per_page = 18
        self.is_ready = False

    def parse_book_structure(self, input_path):
        self.input_file = input_path
        self.raw_chapters = []

        try:
            book = epub.read_epub(self.input_file)
        except Exception as e:
            print(f"Error reading EPUB: {e}")
            return False

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
            raw_html = item.get_content().decode('utf-8', errors='replace')
            soup = BeautifulSoup(raw_html, 'html.parser')
            text_content = soup.get_text().strip()
            has_image = bool(soup.find('img'))

            if item_name not in toc_mapping and len(text_content) < 50 and not has_image:
                continue

            chapter_title = toc_mapping.get(item_name) or (soup.find(['h1', 'h2']).get_text().strip() if soup.find(
                ['h1', 'h2']) else f"Section {len(self.raw_chapters) + 1}")

            self.raw_chapters.append({
                'title': chapter_title,
                'soup': soup,
                'has_image': has_image
            })

        return True

    def render_chapters(self, selected_indices, font_path, font_size, margin, line_height, font_weight,
                        bottom_padding, top_padding, text_align="justify", orientation="Portrait", add_toc=True,
                        footer_settings=None, progress_callback=None):

        self.font_path = font_path if font_path != "DEFAULT" else ""
        self.font_size = font_size
        self.margin = margin
        self.line_height = line_height
        self.font_weight = font_weight
        self.bottom_padding = bottom_padding
        self.top_padding = top_padding
        self.text_align = text_align

        # Apply Footer Settings
        if footer_settings:
            self.footer_visible = footer_settings.get("footer_visible", True)
            self.footer_font_size = footer_settings.get("footer_font_size", 16)
            self.footer_show_progress = footer_settings.get("footer_show_progress", True)
            self.footer_show_pagenum = footer_settings.get("footer_show_pagenum", True)
            self.footer_show_title = footer_settings.get("footer_show_title", True)
            self.footer_text_pos = footer_settings.get("footer_text_pos", "Text Above Bar")
            self.footer_bar_height = footer_settings.get("footer_bar_height", 4)
            self.footer_bottom_margin = footer_settings.get("footer_bottom_margin", 10)

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

        # --- MATCHING WEB MARGINS ---
        # Web uses hardcoded margins for TOC pages
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

            # Web Logic: 40 + top_padding
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

    def render_page(self, global_page_index):
        if not self.is_ready: return None
        num_toc = len(self.toc_pages_images)

        # Footer Rendering Logic
        footer_height = max(0, self.bottom_padding)
        header_height = max(0, self.top_padding)
        content_height = self.screen_height - footer_height - header_height

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
            img.paste(img_content, (0, header_height))

            if has_image:
                img = img.convert("L")
                img = ImageEnhance.Contrast(ImageEnhance.Brightness(img).enhance(1.15)).enhance(1.4)
                img = img.convert("1", dither=Image.Dither.FLOYDSTEINBERG)
            else:
                img = img.convert("L")
                img = ImageEnhance.Contrast(img).enhance(2.0).point(lambda p: 255 if p > 140 else 0, mode='1')
            img = img.convert("RGB")

        # --- DYNAMIC FOOTER RENDERING ---
        if self.footer_visible:
            draw = ImageDraw.Draw(img)

            # 1. Clean background for footer area based on user margin + buffer
            clean_start_y = self.screen_height - self.bottom_padding
            draw.rectangle([0, clean_start_y, self.screen_width, self.screen_height], fill=(255, 255, 255))

            # 2. Calculate Layout
            font_ui = self._get_ui_font(self.footer_font_size)
            text_height_px = int(self.footer_font_size * 1.3)
            bar_thickness = self.footer_bar_height if self.footer_show_progress else 0

            element_gap = 5

            # Position Reference: footer_bottom_margin is distance from BOTTOM of screen
            base_y = self.screen_height - self.footer_bottom_margin

            bar_y = 0
            text_y = 0

            if self.footer_text_pos == "Text Above Bar":
                # Layout from bottom up: Margin -> Bar -> Gap -> Text
                if self.footer_show_progress:
                    bar_y = base_y - bar_thickness
                    text_ref = bar_y - element_gap
                else:
                    text_ref = base_y

                text_y = text_ref - text_height_px + 3  # +3 adjustment for font baseline

            else:  # "Text Below Bar"
                # Layout from bottom up: Margin -> Text -> Gap -> Bar
                has_text = self.footer_show_pagenum or self.footer_show_title
                if has_text:
                    text_y = base_y - text_height_px
                    bar_ref = text_y - element_gap
                else:
                    bar_ref = base_y

                bar_y = bar_ref - bar_thickness

            # 3. Draw Progress Bar
            if self.footer_show_progress:
                draw.rectangle([10, bar_y, self.screen_width - 10, bar_y + bar_thickness], fill=(255, 255, 255),
                               outline=(0, 0, 0))

                # Chapters ticks
                chapter_pages = [item[1] for item in self.toc_data_final]
                for cp in chapter_pages:
                    if self.total_pages > 0:
                        mx = int(((cp - 1) / self.total_pages) * (self.screen_width - 20)) + 10
                        draw.line([mx, bar_y - 1, mx, bar_y + bar_thickness + 1], fill=(0, 0, 0), width=1)

                # Fill progress
                page_num_disp = global_page_index + 1
                if self.total_pages > 0:
                    bw = int((page_num_disp / self.total_pages) * (self.screen_width - 20))
                    draw.rectangle([10, bar_y, 10 + bw, bar_y + bar_thickness], fill=(0, 0, 0))

            # 4. Draw Text Elements
            has_text = self.footer_show_pagenum or self.footer_show_title
            if has_text:
                page_num_disp = global_page_index + 1
                current_title = ""
                for title, start_pg in reversed(self.toc_data_final):
                    if page_num_disp >= start_pg:
                        current_title = title
                        break

                cursor_x = 15

                # Draw Page Num
                if self.footer_show_pagenum:
                    pg_text = f"{page_num_disp}/{self.total_pages}"
                    draw.text((cursor_x, text_y), pg_text, font=font_ui, fill=(0, 0, 0))
                    cursor_x += font_ui.getlength(pg_text) + 15

                    if self.footer_show_title and current_title:
                        draw.text((cursor_x - 10, text_y), "|", font=font_ui, fill=(0, 0, 0))

                # Draw Title
                if self.footer_show_title and current_title:
                    available_width = self.screen_width - cursor_x - 10
                    display_title = current_title
                    if font_ui.getlength(display_title) > available_width:
                        while font_ui.getlength(display_title + "...") > available_width and len(display_title) > 0:
                            display_title = display_title[:-1]
                        display_title += "..."
                    draw.text((cursor_x, text_y), display_title, font=font_ui, fill=(0, 0, 0))

        return img

    def save_xtc(self, out_name, progress_callback=None):
        if not self.is_ready: return
        blob, idx = bytearray(), bytearray()
        data_off = 56 + (16 * self.total_pages)
        for i in range(self.total_pages):
            if progress_callback: progress_callback((i + 1) / self.total_pages)
            img = self.render_page(i).convert("L").point(lambda p: 255 if p > 128 else 0, mode='1')
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

        # LOAD STARTUP DEFAULTS
        self.startup_settings = FACTORY_DEFAULTS.copy()
        self.load_startup_defaults()

        self.title("EPUB2XTC")
        self.geometry("1280x950")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # WIDER SIDEBAR
        self.sidebar = ctk.CTkScrollableFrame(self, width=480, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")

        # --- SECTION: FILE & PRESETS ---
        ctk.CTkButton(self.sidebar, text="Select EPUB", command=self.select_file).pack(padx=10, pady=(20, 5), fill="x")
        self.lbl_file = ctk.CTkLabel(self.sidebar, text="No file", text_color="gray")
        self.lbl_file.pack()

        self.frm_presets = ctk.CTkFrame(self.sidebar, fg_color="#333")
        self.frm_presets.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(self.frm_presets, text="Manage Presets").pack(pady=(5, 0))

        self.preset_var = ctk.StringVar(value="Select Preset...")
        self.preset_dropdown = ctk.CTkOptionMenu(self.frm_presets, variable=self.preset_var, values=[],
                                                 command=self.load_selected_preset)
        self.preset_dropdown.pack(fill="x", padx=10, pady=5)
        self.refresh_presets_list()

        self.btn_save_preset = ctk.CTkButton(self.frm_presets, text="Save New Preset", command=self.save_new_preset,
                                             height=24, fg_color="green")
        self.btn_save_preset.pack(fill="x", padx=10, pady=5)

        self.btn_save_default = ctk.CTkButton(self.frm_presets, text="Save as Startup Default",
                                              command=self.save_current_as_default, height=24, fg_color="#D35400")
        self.btn_save_default.pack(fill="x", padx=10, pady=(5, 10))

        # --- SECTION: GLOBAL SETTINGS ---
        self.var_toc = ctk.BooleanVar(value=self.startup_settings["generate_toc"])
        self.check_toc = ctk.CTkCheckBox(self.sidebar, text="Generate TOC Pages", variable=self.var_toc,
                                         command=self.schedule_update)
        self.check_toc.pack(padx=20, pady=5, anchor="w")

        self.btn_chapters = ctk.CTkButton(self.sidebar, text="Edit Chapter Visibility",
                                          command=self.open_chapter_dialog, state="disabled", fg_color="gray")
        self.btn_chapters.pack(padx=20, pady=5, fill="x")

        ctk.CTkLabel(self.sidebar, text="--- Layout & Typography ---").pack(pady=(15, 5))

        # --- LAYOUT GRID ---
        self.cols_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.cols_frame.pack(fill="x", padx=5, pady=5)

        self.col_left = ctk.CTkFrame(self.cols_frame, fg_color="transparent")
        self.col_left.pack(side="left", fill="both", expand=True, padx=5)

        self.col_right = ctk.CTkFrame(self.cols_frame, fg_color="transparent")
        self.col_right.pack(side="right", fill="both", expand=True, padx=5)

        # === LEFT COLUMN ===
        ctk.CTkLabel(self.col_left, text="Orientation:").pack(pady=(5, 0))
        self.orientation_var = ctk.StringVar(value=self.startup_settings["orientation"])
        self.orientation_dropdown = ctk.CTkOptionMenu(self.col_left, values=["Portrait", "Landscape"],
                                                      variable=self.orientation_var, command=self.schedule_update)
        self.orientation_dropdown.pack(fill="x", pady=5)

        ctk.CTkLabel(self.col_left, text="Font Family:").pack(pady=(5, 0))
        self.available_fonts = get_local_fonts()
        self.font_options = ["Default (System)"] + [os.path.basename(f) for f in self.available_fonts]
        self.font_map = {os.path.basename(f): f for f in self.available_fonts}
        self.font_map["Default (System)"] = "DEFAULT"
        self.font_dropdown = ctk.CTkOptionMenu(self.col_left, values=self.font_options, command=self.on_font_change)

        if self.startup_settings["font_name"] in self.font_options:
            self.font_dropdown.set(self.startup_settings["font_name"])
        else:
            self.font_dropdown.set("Default (System)")
        self.font_dropdown.pack(fill="x", pady=5)
        self.processor.font_path = self.font_map[self.font_dropdown.get()]

        self.lbl_size = ctk.CTkLabel(self.col_left, text=f"Font Size: {self.startup_settings['font_size']}pt")
        self.lbl_size.pack(pady=(5, 0))
        self.slider_size = ctk.CTkSlider(self.col_left, from_=12, to=48, command=self.update_size_label)
        self.slider_size.set(self.startup_settings['font_size'])
        self.slider_size.pack(fill="x", pady=5)

        self.lbl_margin = ctk.CTkLabel(self.col_left, text=f"Margin: {self.startup_settings['margin']}px")
        self.lbl_margin.pack(pady=(5, 0))
        self.slider_margin = ctk.CTkSlider(self.col_left, from_=0, to=100, command=self.update_margin_label)
        self.slider_margin.set(self.startup_settings['margin'])
        self.slider_margin.pack(fill="x", pady=5)

        self.lbl_top_padding = ctk.CTkLabel(self.col_left,
                                            text=f"Top Padding: {self.startup_settings['top_padding']}px")
        self.lbl_top_padding.pack(pady=(5, 0))
        self.slider_top_padding = ctk.CTkSlider(self.col_left, from_=0, to=100, command=self.update_top_padding_label)
        self.slider_top_padding.set(self.startup_settings['top_padding'])
        self.slider_top_padding.pack(fill="x", pady=5)

        # === RIGHT COLUMN ===
        ctk.CTkLabel(self.col_right, text="Text Alignment:").pack(pady=(5, 0))
        self.align_dropdown = ctk.CTkOptionMenu(self.col_right, values=["justify", "left"],
                                                command=self.schedule_update)
        self.align_dropdown.set(self.startup_settings["text_align"])
        self.align_dropdown.pack(fill="x", pady=5)

        self.lbl_weight = ctk.CTkLabel(self.col_right, text=f"Font Weight: {self.startup_settings['font_weight']}")
        self.lbl_weight.pack(pady=(5, 0))
        self.slider_weight = ctk.CTkSlider(self.col_right, from_=100, to=900, number_of_steps=8,
                                           command=self.update_weight_label)
        self.slider_weight.set(self.startup_settings['font_weight'])
        self.slider_weight.pack(fill="x", pady=5)

        self.lbl_line = ctk.CTkLabel(self.col_right, text=f"Line Height: {self.startup_settings['line_height']}")
        self.lbl_line.pack(pady=(5, 0))
        self.slider_line = ctk.CTkSlider(self.col_right, from_=1.0, to=2.5, command=self.update_line_label)
        self.slider_line.set(self.startup_settings['line_height'])
        self.slider_line.pack(fill="x", pady=5)

        self.lbl_preview_zoom = ctk.CTkLabel(self.col_right,
                                             text=f"Preview Zoom: {self.startup_settings['preview_zoom']}")
        self.lbl_preview_zoom.pack(pady=(5, 0))
        self.slider_preview_zoom = ctk.CTkSlider(self.col_right, from_=200, to=800, command=self.update_zoom_only)
        self.slider_preview_zoom.set(self.startup_settings['preview_zoom'])
        self.slider_preview_zoom.pack(fill="x", pady=5)

        self.lbl_padding = ctk.CTkLabel(self.col_right,
                                        text=f"Bottom Padding: {self.startup_settings['bottom_padding']}px")
        self.lbl_padding.pack(pady=(5, 0))
        self.slider_padding = ctk.CTkSlider(self.col_right, from_=0, to=150, command=self.update_padding_label)
        self.slider_padding.set(self.startup_settings['bottom_padding'])
        self.slider_padding.pack(fill="x", pady=5)

        # --- SECTION: FOOTER CUSTOMIZATION (MATCHING STYLE) ---
        ctk.CTkLabel(self.sidebar, text="--- Footer Settings ---").pack(pady=(15, 5))

        self.var_footer_visible = ctk.BooleanVar(value=self.startup_settings.get("footer_visible", True))
        self.check_footer_vis = ctk.CTkCheckBox(self.sidebar, text="Show Footer Area", variable=self.var_footer_visible,
                                                command=self.schedule_update)
        self.check_footer_vis.pack(padx=20, pady=5, anchor="center")

        # Two-Column Layout (Matching above)
        self.footer_cols = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.footer_cols.pack(fill="x", padx=5, pady=5)

        self.f_col1 = ctk.CTkFrame(self.footer_cols, fg_color="transparent")
        self.f_col1.pack(side="left", fill="both", expand=True, padx=5)

        self.f_col2 = ctk.CTkFrame(self.footer_cols, fg_color="transparent")
        self.f_col2.pack(side="right", fill="both", expand=True, padx=5)

        # COL 1: CONTENT & POS
        self.var_footer_prog = ctk.BooleanVar(value=self.startup_settings.get("footer_show_progress", True))
        ctk.CTkCheckBox(self.f_col1, text="Progress Bar", variable=self.var_footer_prog,
                        command=self.schedule_update).pack(pady=2, anchor="w")

        self.var_footer_pg = ctk.BooleanVar(value=self.startup_settings.get("footer_show_pagenum", True))
        ctk.CTkCheckBox(self.f_col1, text="Page Numbers", variable=self.var_footer_pg,
                        command=self.schedule_update).pack(pady=2, anchor="w")

        self.var_footer_title = ctk.BooleanVar(value=self.startup_settings.get("footer_show_title", True))
        ctk.CTkCheckBox(self.f_col1, text="Chapter Title", variable=self.var_footer_title,
                        command=self.schedule_update).pack(pady=2, anchor="w")

        ctk.CTkLabel(self.f_col1, text="Text Position:").pack(pady=(10, 0))
        self.footer_pos_var = ctk.StringVar(value=self.startup_settings.get("footer_text_pos", "Text Above Bar"))
        ctk.CTkOptionMenu(self.f_col1, variable=self.footer_pos_var, values=["Text Above Bar", "Text Below Bar"],
                          command=self.schedule_update).pack(fill="x", pady=2)

        # COL 2: SLIDERS
        self.lbl_footer_size = ctk.CTkLabel(self.f_col2,
                                            text=f"Text Size: {self.startup_settings.get('footer_font_size', 16)}pt")
        self.lbl_footer_size.pack(anchor="w", pady=(0, 0))
        self.slider_footer_size = ctk.CTkSlider(self.f_col2, from_=10, to=24, height=16,
                                                command=self.update_footer_size_label)
        self.slider_footer_size.set(self.startup_settings.get('footer_font_size', 16))
        self.slider_footer_size.pack(fill="x", pady=5)

        self.lbl_bar_thick = ctk.CTkLabel(self.f_col2,
                                          text=f"Bar Thick: {self.startup_settings.get('footer_bar_height', 4)}px")
        self.lbl_bar_thick.pack(anchor="w", pady=(5, 0))
        self.slider_bar_thick = ctk.CTkSlider(self.f_col2, from_=1, to=10, height=16,
                                              command=self.update_bar_thick_label)
        self.slider_bar_thick.set(self.startup_settings.get('footer_bar_height', 4))
        self.slider_bar_thick.pack(fill="x", pady=5)

        self.lbl_footer_margin = ctk.CTkLabel(self.f_col2,
                                              text=f"Footer position: {self.startup_settings.get('footer_bottom_margin', 10)}px")
        self.lbl_footer_margin.pack(anchor="w", pady=(5, 0))
        self.slider_footer_margin = ctk.CTkSlider(self.f_col2, from_=0, to=50, height=16,
                                                  command=self.update_footer_margin_label)
        self.slider_footer_margin.set(self.startup_settings.get('footer_bottom_margin', 10))
        self.slider_footer_margin.pack(fill="x", pady=5)

        # --- SECTION: FOOTER ACTIONS ---
        ctk.CTkButton(self.sidebar, text="Reset to Factory Defaults", command=self.reset_to_factory,
                      fg_color="#555").pack(padx=20, pady=20, fill="x")

        self.btn_run = ctk.CTkButton(self.sidebar, text="Force Refresh", fg_color="gray", command=self.run_processing)
        self.btn_run.pack(padx=20, pady=5, fill="x")

        self.btn_export = ctk.CTkButton(self.sidebar, text="Export XTC", state="disabled", command=self.export_file)
        self.btn_export.pack(padx=20, pady=5, fill="x")

        self.progress_bar = ctk.CTkProgressBar(self.sidebar)
        self.progress_bar.set(0)
        self.progress_bar.pack(padx=20, pady=10, fill="x")
        self.progress_label = ctk.CTkLabel(self.sidebar, text="Progress: Ready")
        self.progress_label.pack()

        # --- PREVIEW AREA ---
        self.preview_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.preview_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.img_label = ctk.CTkLabel(self.preview_frame, text="Load EPUB to Preview")
        self.img_label.pack(expand=True, fill="both")

        self.nav = ctk.CTkFrame(self.preview_frame, fg_color="transparent")
        self.nav.pack(side="bottom", fill="x", pady=10)

        ctk.CTkButton(self.nav, text="< Previous", width=90, command=self.prev_page).pack(side="left", padx=20)
        ctk.CTkButton(self.nav, text="Next >", width=90, command=self.next_page).pack(side="right", padx=20)

        self.center_nav = ctk.CTkFrame(self.nav, fg_color="transparent")
        self.center_nav.pack(side="left", expand=True, fill="both")

        self.lbl_page = ctk.CTkLabel(self.center_nav, text="Page 0/0", font=("Arial", 16))
        self.lbl_page.pack(side="top", pady=(0, 2))

        self.goto_frame = ctk.CTkFrame(self.center_nav, fg_color="transparent")
        self.goto_frame.pack(side="top")

        self.entry_page = ctk.CTkEntry(self.goto_frame, width=60, height=24, placeholder_text="#")
        self.entry_page.pack(side="left", padx=(0, 5))
        self.entry_page.bind('<Return>', lambda event: self.go_to_page())

        self.btn_go = ctk.CTkButton(self.goto_frame, text="Go", width=40, height=24, command=self.go_to_page)
        self.btn_go.pack(side="left")

    # --- PRESET & SETTINGS LOGIC ---

    def gather_current_ui_settings(self):
        return {
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

            # Footer
            "footer_visible": self.var_footer_visible.get(),
            "footer_font_size": int(self.slider_footer_size.get()),
            "footer_show_progress": self.var_footer_prog.get(),
            "footer_show_pagenum": self.var_footer_pg.get(),
            "footer_show_title": self.var_footer_title.get(),
            "footer_text_pos": self.footer_pos_var.get(),
            "footer_bar_height": int(self.slider_bar_thick.get()),
            "footer_bottom_margin": int(self.slider_footer_margin.get())
        }

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

        # Footer
        self.var_footer_visible.set(s['footer_visible'])
        self.slider_footer_size.set(s['footer_font_size'])
        self.var_footer_prog.set(s['footer_show_progress'])
        self.var_footer_pg.set(s['footer_show_pagenum'])
        self.var_footer_title.set(s['footer_show_title'])
        self.footer_pos_var.set(s['footer_text_pos'])
        self.slider_bar_thick.set(s['footer_bar_height'])
        self.slider_footer_margin.set(s['footer_bottom_margin'])

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
        self.update_footer_size_label(s['footer_font_size'])
        self.update_bar_thick_label(s['footer_bar_height'])
        self.update_footer_margin_label(s['footer_bottom_margin'])

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

        # Sanitize filename
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

    def update_zoom_only(self, value):
        self.lbl_preview_zoom.configure(text=f"Preview Zoom: {int(value)}")
        if self.processor.is_ready: self.show_page(self.current_page_index)

    def on_font_change(self, choice):
        self.processor.font_path = self.font_map[choice]
        self.schedule_update()

    def update_size_label(self, value):
        self.lbl_size.configure(text=f"Font Size: {int(value)}pt")
        self.schedule_update()

    def update_weight_label(self, value):
        self.lbl_weight.configure(text=f"Font Weight: {int(value)}")
        self.schedule_update()

    def update_line_label(self, value):
        self.lbl_line.configure(text=f"Line Height: {value:.1f}")
        self.schedule_update()

    def update_margin_label(self, value):
        self.lbl_margin.configure(text=f"Margin: {int(value)}px")
        self.schedule_update()

    def update_padding_label(self, value):
        self.lbl_padding.configure(text=f"Bottom Padding: {int(value)}px")
        self.schedule_update()

    def update_top_padding_label(self, value):
        self.lbl_top_padding.configure(text=f"Top Padding: {int(value)}px")
        self.schedule_update()

    def update_footer_size_label(self, value):
        self.lbl_footer_size.configure(text=f"Text Size: {int(value)}")
        self.schedule_update()

    def update_bar_thick_label(self, value):
        self.lbl_bar_thick.configure(text=f"Bar Thick: {int(value)}px")
        self.schedule_update()

    def update_footer_margin_label(self, value):
        self.lbl_footer_margin.configure(text=f"Footer position: {int(value)}px")
        self.schedule_update()

    def update_progress_ui(self, val, stage_text="Processing"):
        self.after(0, lambda: self.progress_bar.set(val))
        self.after(0, lambda: self.progress_label.configure(text=f"{stage_text}: {int(val * 100)}%"))

    def run_processing(self):
        if not self.processor.input_file: return
        if self.is_processing: return

        if self.selected_chapter_indices is None:
            self.selected_chapter_indices = list(range(len(self.processor.raw_chapters)))

        # Capture current footer config
        footer_settings = {
            "footer_visible": self.var_footer_visible.get(),
            "footer_font_size": int(self.slider_footer_size.get()),
            "footer_show_progress": self.var_footer_prog.get(),
            "footer_show_pagenum": self.var_footer_pg.get(),
            "footer_show_title": self.var_footer_title.get(),
            "footer_text_pos": self.footer_pos_var.get(),
            "footer_bar_height": int(self.slider_bar_thick.get()),
            "footer_bottom_margin": int(self.slider_footer_margin.get())
        }

        self.is_processing = True
        self.btn_run.configure(state="disabled", text="Rendering...", fg_color="orange")
        self.progress_label.configure(text="Starting layout...")

        threading.Thread(target=lambda: self._task_render(footer_settings)).start()

    def _task_render(self, footer_settings):
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
            footer_settings=footer_settings,
            progress_callback=lambda v: self.update_progress_ui(v, "Layout")
        )
        self.after(0, lambda: self._done(success))

    def _done(self, success):
        self.is_processing = False
        self.btn_run.configure(state="normal", text="Force Refresh", fg_color="green")

        if success:
            self.btn_export.configure(state="normal")
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

        # --- SMART SCALING LOGIC ---
        base_size = int(self.slider_preview_zoom.get())

        if img.width > img.height:
            # Landscape: Slider controls Height
            target_h = base_size
            aspect_ratio = img.width / img.height
            target_w = int(target_h * aspect_ratio)
        else:
            # Portrait: Slider controls Width
            target_w = base_size
            aspect_ratio = img.height / img.width
            target_h = int(target_w * aspect_ratio)

        ctk_img = ctk.CTkImage(light_image=img, size=(target_w, target_h))
        self.img_label.configure(image=ctk_img, text="")
        self.lbl_page.configure(text=f"Page {idx + 1} / {self.processor.total_pages}")

    def go_to_page(self):
        if not self.processor.is_ready: return
        txt = self.entry_page.get()
        if not txt.isdigit():
            return

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


if __name__ == "__main__":
    app = App()
    app.mainloop()