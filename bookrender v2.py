import os
import re
import sys
import shutil
import numpy as np
import fitz  # PyMuPDF
from PIL import Image, ImageFilter, ImageOps

class PDFProcessor:
    def __init__(self):
        self.test_mode = False
        self.script_dir = os.path.dirname(os.path.abspath(__file__))

    def show_menu(self):
        while True:
            print("\n=== Nokia PDF Processor ===")
            print("1. Toggle Test Mode (Currently: " + ("ON" if self.test_mode else "OFF") + ")")
            print("2. Process All PDFs")
            print("3. Exit")
            
            try:
                choice = input("Select option (1-3): ").strip()
                
                if choice == "1":
                    self.test_mode = not self.test_mode
                elif choice == "2":
                    self.process_pdfs()
                elif choice == "3":
                    print("Exiting program...")
                    sys.exit(0)
                else:
                    print("Invalid choice. Please enter 1-3.")
            except KeyboardInterrupt:
                print("\nExiting program...")
                sys.exit(0)
            except Exception as e:
                print(f"Error: {str(e)}")

    def process_pdfs(self):
        if not os.path.exists("processed_books"):
            os.makedirs("processed_books")

        pdf_files = [f for f in os.listdir(self.script_dir) 
                   if f.lower().endswith('.pdf') and not f.startswith('~$')]

        if not pdf_files:
            print("\nNo PDF files found in current directory!")
            return

        for pdf in pdf_files:
            print(f"\nProcessing: {pdf}")
            try:
                self._process_pdf(pdf)
                print("✓ Success")
            except Exception as e:
                print(f"✗ Failed: {str(e)}")

    def _process_pdf(self, filename):
        doc = fitz.open(os.path.join(self.script_dir, filename))
        book_title = os.path.splitext(filename)[0]
        output_dir = os.path.join(self.script_dir, "processed_books", self._sanitize(book_title))

        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir)

        max_pages = 10 if self.test_mode else len(doc)
        for page_num in range(min(len(doc), max_pages)):
            try:
                page = doc.load_page(page_num)
                self._process_page(page, page_num, output_dir)
            except Exception as e:
                print(f"Error processing page {page_num}: {str(e)}")

        doc.close()

    def _process_page(self, page, page_num, output_dir):
        # Render at Nokia-friendly resolution
        pix = page.get_pixmap(dpi=144)  # 144 DPI = 6px/mm (Nokia native)
        
        # Convert to Nokia-compatible format
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Full-page image detection
        if len(page.get_images()) > 3:
            img = img.convert("L")  # Grayscale
            img = ImageOps.autocontrast(img, cutoff=5)
            img.save(os.path.join(output_dir, f"{page_num + 1:04d}c.bmp"), format="BMP")
            return

        # Convert to grayscale for processing
        gray = img.convert("L")
        gray = ImageOps.autocontrast(gray, cutoff=5)
        
        # Crop whitespace
        cropped = self._crop_whitespace(gray)
        if cropped:
            width, height = cropped.size
            #   temporarily changed for testing        overlap = int(height * 0.10)  # 10% overlap
            overlap = max(int(height * 0.03), 10)
            
            # Split with overlap
            top = cropped.crop((0, 0, width, height//2 + overlap))
            bottom = cropped.crop((0, height//2 - overlap, width, height))
            
            # Apply Nokia-specific optimizations
            for half, suffix in [(top, "a"), (bottom, "b")]:
                # Convert to 1-bit with dithering for best Nokia display
                bw = half.convert("1", dither=Image.FLOYDSTEINBERG)
                rotated = bw.transpose(Image.ROTATE_270)
                
                # Save as uncompressed BMP
                rotated.save(
                    os.path.join(output_dir, f"{page_num + 1:04d}{suffix}.bmp"),
                    format="BMP"
                )

    def _crop_whitespace(self, img):
        """Advanced cropping ignoring decorative margins."""
        arr = np.array(img)
        margin = int(min(arr.shape) * 0.05)
        
        # Detect content area (more tolerant of light decorations)
        content_mask = arr < 220  # 220 threshold catches near-white
        
        vertical_proj = np.max(content_mask, axis=0)
        horizontal_proj = np.max(content_mask, axis=1)
        
        left = max(np.argmax(vertical_proj[margin:]) + margin - 5, 0)
        right = min(len(vertical_proj) - np.argmax(vertical_proj[::-1][margin:]) - margin + 5, arr.shape[1])
        top = max(np.argmax(horizontal_proj[margin:]) + margin - 5, 0)
        bottom = min(len(horizontal_proj) - np.argmax(horizontal_proj[::-1][margin:]) - margin + 5, arr.shape[0])
        
        return img.crop((left, top, right, bottom))

    def _sanitize(self, text):
        """Clean filenames for Nokia compatibility."""
        return re.sub(r'[\\/*?:"<>|]', "", text)[:30].replace(" ", "_")

if __name__ == "__main__":
    os.system('cls' if os.name == 'nt' else 'clear')
    print("Nokia PDF Processor v2.2")
    processor = PDFProcessor()
    processor.show_menu()