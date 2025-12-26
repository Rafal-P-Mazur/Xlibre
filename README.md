# EPUB to XTC Converter for Xteink X4
Epub to XTC converter

A GUI-based tool designed to convert standard `.epub` files into the `.xtc` binary format required by the **Xteink X4** e-reader. It renders HTML content into paginated, bitmapped images optimized for e-ink displays.

[![Web Version](https://img.shields.io/badge/Web_Version-Live-green)](https://epub2xtc.streamlit.app/)

## Main Features

* **Smart Table of Contents:** Automatically generates visual TOC pages that dynamically adapt to your selected font size and margins. Now features **robust file mapping** to accurately detect chapter titles even in complex EPUB folder structures.
* **Chapter Visibility Control:** Non-destructively hide specific chapters from the **Table of Contents** and **Progress Bar** while keeping the text readable in the book.
* **Advanced Header & Footer Engine:** Fully customizable top and bottom areas. Assign elements (**Title**, **Page Number**, **Reading %**, **Chapter Page**) to either the **Header** or **Footer**. Supports custom **ordering** and **alignment** (Left/Center/Right/Justify).
* **Enhanced Progress Bar:** Visual bar featuring **Chapter Ticks** and a **Current Position Marker**. Can be positioned in the Header or Footer (Above/Below text).
* **Inline Footnotes:** Automatically detects and renders internal links/footnotes at the bottom of the relevant text block, preventing the need to jump pages to read references.
* **Cover Export Tool:** Built-in utility to resize, crop, and dither the book cover into a high-quality 1-bit BMP image.
* **Image Optimization:** Automatically extracts, scales, contrast-enhances, and applies **Floyd-Steinberg dithering** to images for optimal display on e-ink screens.
* **Preset Management:** Save and load your favorite layout configurations via the `presets/` folder. Presets are JSON-based and cross-compatible with the Web version.
* **Custom Typography:** Drop external `.ttf` or `.otf` files into the `fonts/` directory to use them instantly. Includes sliders for font weight, size, and line height.
* **Smart Hyphenation:** Uses `pyphen` to inject soft hyphens into text nodes, ensuring proper line breaks and justified text flow.

![App Screenshot 1](images/xtc.png)
![App Screenshot 2](images/xtc2.png)


## 📥 Installation

### Option 1: Run from Source
1. **Install the dependencies:**
    ```bash
    pip install pymupdf Pillow EbookLib beautifulsoup4 pyphen customtkinter
    ```
2.  **Clone the repository:**
    ```bash
    git clone https://github.com/Rafal-P-Mazur/EPUB2XTC.git
    cd EPUB2XTC
    ```
3.  **Run the App:**
    ```bash
    python EPUB2XTC.py
    ```

### Option 2: Standalone Executable (.exe)
If you have downloaded the [Release version](https://github.com/Rafal-P-Mazur/EPUB2XTC/releases), simply unzip the file and run `EPUB2XTC.exe`. No Python installation is required.

## 📖 User Manual

1.  **Load an EPUB:**
    * Click **Select EPUB** in the **Left Sidebar**. The application will instantly parse the book structure and attempt to detect the cover.

2.  **Select Chapters:**
    * A dialog will automatically appear displaying all detected chapters.
    * **Uncheck** any chapters you wish to hide from the **Table of Contents** and **Progress Bar** (e.g., Copyright pages or generic "Section" headers).
    * *Note:* These chapters are **not deleted**; they remain in the book for reading but will not clutter your navigation.

3.  **Manage Presets (Left Sidebar):**
    * Use the dropdown menu to load saved layouts from the `presets` folder.
    * Click **Save** to store your current configuration as a new JSON file.
    * Click **Default** to save the current settings as your startup configuration.

4.  **Configure Layout (Left Sidebar):**
    * **Structure:** Toggle **"Generate TOC Pages"** or enable **"Inline Footnotes"** (renders footnotes at the bottom of the text block instead of jumping pages).
    * **Typography:** Select fonts (place `.ttf` or `.otf` files in the `fonts` directory to add more), and adjust Size, Weight, and Line Height.
    * **Margins:** Fine-tune Side Margins, Top Padding (Header area), and Bottom Padding (Footer area).

5.  **Header & Footer (Right Sidebar):**
    * **Positioning:** Assign elements (Chapter Title, Page Number, Reading %, Chapter Page) to the **Header**, **Footer**, or **Hidden**.
    * **Ordering:** Change the numeric order (1–4) to rearrange elements on the line.
    * **Progress Bar:** Choose to display the bar in the Header or Footer. Toggle **"Show Ticks"** (chapter markers) and **"Show Marker"** (current location dot).
    * **Alignment:** independently align the Header and Footer text (Left, Center, Right, Justify).

6.  **Navigate & Preview (Center):**
    * Use the **< Prev** and **Next >** buttons or the **"Go"** input box to navigate pages.
    * Use the **Preview Zoom** slider to resize the image (Smart Scaling automatically optimizes this based on orientation).
    * *Note:* The preview **automatically updates** after a short delay when settings are changed.

7.  **Export:**
    * **Export XTC:** Generates the final binary file for the device.
    * **Export Cover:** Opens a tool to resize, crop, and dither the book cover into a 1-bit BMP image.
   
## 📦 Dependencies

* `customtkinter` (GUI)
* `PyMuPDF` (Rendering)
* `Pillow` (Image Processing)
* `EbookLib` & `BeautifulSoup4` (Parsing)
* `Pyphen` (Hyphenation)

---

## 📄 License
MIT License

