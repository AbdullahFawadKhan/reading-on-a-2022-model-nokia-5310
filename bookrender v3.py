import os
import re
import sys
import shutil
import numpy as np
import fitz  # PyMuPDF
from PIL import Image, ImageOps

class PDFProcessor:
    def __init__(self):
        self.test_mode = False
        self.test_page_limit = 10
        self.script_dir = os.path.dirname(os.path.abspath(__file__))

    def show_menu(self):
        while True:
            print("\n=== Nokia PDF Processor ===")
            print(f"1. Toggle Test Mode (Currently: {'ON' if self.test_mode else 'OFF'})")
            print("2. Process All PDFs")
            print("3. Exit")
            
            try:
                choice = input("Select option (1-3): ").strip()
                if choice == "1":
                    self.test_mode = not self.test_mode
                elif choice == "2":
                    self.process_pdfs()
                elif choice == "3":
                    sys.exit(0)
            except (KeyboardInterrupt, EOFError):
                print("\nExiting...")
                sys.exit(0)

    def process_pdfs(self):
        if not os.path.exists("processed_books"):
            os.makedirs("processed_books")

        for pdf in [f for f in os.listdir(self.script_dir) if f.lower().endswith('.pdf')]:
            print(f"\nProcessing: {pdf}")
            try:
                self._process_pdf(pdf)
            except Exception as e:
                print(f"✗ Failed: {str(e)}")

    def _process_pdf(self, filename):
        doc = fitz.open(os.path.join(self.script_dir, filename))
        book_title = self._sanitize(os.path.splitext(filename)[0])
        base_dir = os.path.join(self.script_dir, "processed_books", book_title)
        
        if os.path.exists(base_dir):
            shutil.rmtree(base_dir)
        os.makedirs(base_dir)

        # Chapter detection from TOC
        toc = doc.get_toc()
        chapter_ranges = self._get_chapter_ranges(doc, toc)
        
        # Process all pages with chapter validation
        max_pages = self.test_page_limit if self.test_mode else len(doc)
        for page_num in range(min(len(doc), max_pages)):
            try:
                chapter_dir = self._get_chapter_folder(page_num, chapter_ranges, base_dir)
                self._process_page(doc.load_page(page_num), page_num, chapter_dir)
            except Exception as e:
                print(f"Page {page_num} error: {str(e)}")

        doc.close()

    def _get_chapter_ranges(self, doc, toc):
        """Returns {chapter_name: (start_page, end_page)} with 2-digit numbering"""
        chapters = {}
        current_chapter = "00_cover"
        start_page = 0
        
        for entry in toc:
            if entry[0] == 1:  # Level-1 entries are chapters
                chapter_num = len(chapters) + 1
                chapter_name = f"{chapter_num:02d}_{self._sanitize(entry[1])}"
                chapters[current_chapter] = (start_page, entry[2] - 2)#Changed from -1 to -2
                current_chapter = chapter_name
                start_page = entry[2] - 1  # Changed from entry[2] to entry[2]-1
        
        chapters[current_chapter] = (start_page, len(doc) - 1)
        chapters["99_back"] = (len(doc), len(doc))  # Empty back matter folder
        return chapters

    def _get_chapter_folder(self, page_num, chapter_ranges, base_dir):
        """Returns the correct chapter folder for a page"""
        for chapter, (start, end) in chapter_ranges.items():
            if start <= page_num <= end:
                folder = os.path.join(base_dir, chapter)
                os.makedirs(folder, exist_ok=True)
                return folder
        return os.path.join(base_dir, "00_cover")

    def _crop_whitespace(self, img):
        """Precision four-edge cropping with content detection"""
        arr = np.array(img.convert("L"))
        threshold = 220  # Text detection threshold
        margin = int(min(arr.shape) * 0.03)  # 3% protection margin

        # Projections to detect content edges
        def find_edge(projection, margin, threshold):
            for i in range(margin, len(projection) - margin):
                if projection[i] < threshold:
                    return max(0, i - margin)
            return margin

        # Horizontal and vertical projections
        h_proj = np.min(arr, axis=1)  # Horizontal projection (rows)
        v_proj = np.min(arr, axis=0)  # Vertical projection (columns)

        # Detect all four edges independently
        top = find_edge(h_proj, margin, threshold)
        bottom = arr.shape[0] - find_edge(h_proj[::-1], margin, threshold)
        left = find_edge(v_proj, margin, threshold)
        right = arr.shape[1] - find_edge(v_proj[::-1], margin, threshold)

        # Validate crop area
        if (right - left) < 50 or (bottom - top) < 50:  # Minimum 50px size
            return img  # Return original if crop is too aggressive
        
        return img.crop((left, top, right, bottom))

    def _process_page(self, page, page_num, output_dir):
        # Render page at Nokia-friendly resolution
        pix = page.get_pixmap(dpi=144)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        gray = ImageOps.autocontrast(img.convert("L"), cutoff=5)
        
        # 1. Crop whitespace (four-edge detection)
        cropped = self._crop_whitespace(gray)
        if not cropped:
            cropped = gray  # Fallback to original if crop failed

        # 2. Get dimensions of cropped image
        width, height = cropped.size

        # 3. Create splits with 45%-55% overlap (as requested)
        top_split = cropped.crop((
            0,            # Left edge
            0,            # Top edge (0%)
            width,        # Right edge
            int(height * 0.55)  # Bottom edge (55%)
        ))
        
        bottom_split = cropped.crop((
            0,            # Left edge
            int(height * 0.45),  # Top edge (45%)
            width,        # Right edge
            height        # Bottom edge (100%)
        ))

        # 4. Save as Nokia-compatible 1-bit BMP
        def save_split(img, suffix):
            output_path = os.path.join(output_dir, f"{page_num+1:04d}{suffix}.bmp")
            img.convert("1", dither=Image.FLOYDSTEINBERG) \
               .transpose(Image.ROTATE_270) \
               .save(output_path)
            return output_path

        # Verify splits contain content before saving
        if np.array(top_split.convert("L")).min() < 240:  # Has content
            save_split(top_split, "a")
        if np.array(bottom_split.convert("L")).min() < 240:  # Has content
            save_split(bottom_split, "b")

    def _sanitize(self, text):
        return re.sub(r'[\\/*?:"<>|]', "", text)[:30].replace(" ", "_")

if __name__ == "__main__":
    os.system('cls' if os.name == 'nt' else 'clear')
    PDFProcessor().show_menu()