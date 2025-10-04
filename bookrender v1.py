import os
import re
import sys
import shutil
import fitz  # PyMuPDF

class PDFBookProcessor:
    def __init__(self):
        self.include_images = True
        self.script_dir = os.path.dirname(os.path.abspath(__file__))

    def main_menu(self):
        while True:
            print("\n=== PDF Book Processor ===")
            print(f"1. Toggle Images (currently: {'INCLUDED' if self.include_images else 'EXCLUDED'})")
            print("2. Process All PDFs")
            print("3. Exit")
            
            choice = input("Select option (1-3): ").strip()
            
            if choice == "1":
                self.include_images = not self.include_images
                print(f"Images will be {'included' if self.include_images else 'excluded'}")
            elif choice == "2":
                self.process_pdfs_in_directory()
            elif choice == "3":
                print("Exiting program...")
                sys.exit()
            else:
                print("Invalid choice. Please enter 1-3.")

    def process_pdfs_in_directory(self):
        """Process all PDFs in the script directory."""
        if not os.path.exists("processed_books"):
            os.makedirs("processed_books")

        processed_count = 0
        for filename in os.listdir(self.script_dir):
            full_path = os.path.join(self.script_dir, filename)
            
            if (os.path.isdir(full_path) or filename.startswith('~$') or not filename.lower().endswith('.pdf')):
                continue
                
            try:
                print(f"Processing {filename}...", end='', flush=True)
                self.process_pdf_file(full_path)
                processed_count += 1
                print("✓")
            except Exception as e:
                print(f"✗ Failed: {self.format_error(e)}")

        if processed_count == 0:
            print("\nNo PDFs found in the current directory.")
        else:
            print(f"\nProcessed {processed_count} PDF(s).")

    def format_error(self, error):
        """Convert errors to user-friendly messages."""
        error_str = str(error)
        if "no objects found" in error_str:
            return "Not a valid PDF file"
        elif "password" in error_str:
            return "Password-protected PDF"
        return error_str

    def process_pdf_file(self, filepath):
        """Process a single PDF file with Nokia-compatible output."""
        doc = None
        try:
            doc = fitz.open(filepath)
            if not doc.is_pdf:
                raise ValueError("Not a valid PDF file")
                
            book_title = os.path.splitext(os.path.basename(filepath))[0]
            output_dir = os.path.join(self.script_dir, "processed_books", self.sanitize_filename(book_title))
            
            if os.path.exists(output_dir):
                shutil.rmtree(output_dir)
            os.makedirs(output_dir)

            toc = doc.get_toc()
            chapter_ranges = self._get_chapter_ranges(doc, toc)
            self._process_all_pages(doc, chapter_ranges, output_dir)
            
        finally:
            if doc:
                doc.close()

    def _get_chapter_ranges(self, doc, toc):
        """Identify chapter ranges and front/back matter."""
        chapters = {}
        first_chapter_start = len(doc)  # Initialize with max value
        last_chapter_end = 0
        
        # Find all level-1 chapters (main chapters only)
        chapter_num = 1
        for entry in toc:
            if len(entry) >= 3 and entry[0] == 1:  # Level 1 entries only
                page_num = max(0, entry[2] - 1)  # Convert to 0-based
                chapter_name = f"{chapter_num:02d}_{self.sanitize_title(entry[1])}"
                chapters[chapter_name] = page_num
                first_chapter_start = min(first_chapter_start, page_num)
                last_chapter_end = max(last_chapter_end, page_num)
                chapter_num += 1

        # Build final structure
        ranges = {}
        if chapters:
            # Front matter (everything before first chapter)
            ranges["0000_cover"] = (0, first_chapter_start - 1)
            
            # Chapters (sorted by their number)
            sorted_chapters = sorted(chapters.items(), key=lambda x: int(x[0][:2]))
            for i, (name, start) in enumerate(sorted_chapters):
                end = (sorted_chapters[i+1][1] - 1) if i+1 < len(sorted_chapters) else last_chapter_end
                ranges[name] = (start, min(end, len(doc) - 1))
            
            # Back matter (everything after last chapter)
            ranges["9999_back"] = (last_chapter_end + 1, len(doc) - 1)
        else:
            # No chapters found - treat entire document as "cover"
            ranges["0000_cover"] = (0, len(doc) - 1)
            
        return ranges

    def _process_all_pages(self, doc, ranges, output_dir):
        """Process all pages according to the ranges."""
        for folder_name, (start, end) in ranges.items():
            if start > end:  # Skip invalid ranges
                continue
                
            folder_path = os.path.join(output_dir, folder_name)
            os.makedirs(folder_path, exist_ok=True)
            
            print(f"  Processing {folder_name} (pages {start+1}-{end+1})")
            for page_num in range(start, end + 1):
                self._process_page(doc.load_page(page_num), page_num, folder_path)

    def _process_page(self, page, page_num, folder_path):
        # Nokia-compatible filename (8.3 format, uppercase)
        file_prefix = f"P{page_num+1:04d}"  # P0001.TXT, P0002.TXT, etc.
        text_path = os.path.join(folder_path, f"{file_prefix}.TXT")

        # Extract text and convert to Nokia-friendly format
        text = page.get_text("text")
        
        # Step 1: Force ANSI encoding (drop unsupported chars)
        text = text.encode('cp1252', errors='ignore').decode('cp1252')
        
        # Step 2: Normalize line endings to Windows CR+LF
        text = text.replace('\r\n', '\n').replace('\n', '\r\n')
        
        # Step 3: Remove trailing spaces/blank lines
        text = '\r\n'.join(line.strip() for line in text.splitlines() if line.strip())
        
        # Write the file
        with open(text_path, 'w', encoding='cp1252') as f:
            f.write(text)

        # Images (unchanged)
        if self.include_images and page.get_images():
            pix = page.get_pixmap()
            pix.save(os.path.join(folder_path, f"{file_prefix}.PNG"))

    def sanitize_filename(self, name):
        """Make safe filenames."""
        return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

    def sanitize_title(self, title):
        """Clean and shorten titles."""
        clean = re.sub(r'[\\/*?:"<>|]', "", title)
        return clean[:30].replace(" ", "_")  # Shorter for Nokia compatibility

if __name__ == "__main__":
    print("PDF Book Processor - Nokia Compatible")
    try:
        processor = PDFBookProcessor()
        processor.main_menu()
    except Exception as e:
        print(f"Fatal error: {str(e)}")
        sys.exit(1)