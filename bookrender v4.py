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
        self.max_path_length = 200  # Conservative limit for Windows



    #NEW FUNCTION TO FIX CHAPTER FOLDER TO PAGES MISMATCH IN CROPPED PDF

    def _fix_trimmed_pdf_toc(self, toc, doc, filename):
        """Adjust TOC for manually-trimmed PDFs where early chapters were deleted."""
        if not toc:
            return toc
        
        # Check if this is your specific problematic PDF
        if "SS ch 601" not in filename:
            return toc
        
        # Find the first chapter that actually exists in the physical PDF
        first_visible_page = 0  # Your trimmed PDF starts at physical page 0
        
        for entry in toc:
            if entry[0] == 1:  # Only look at chapter-level entries
                # Convert TOC's 1-based page to 0-based
                toc_page_0based = entry[2] - 1
                
                # If this chapter starts within the physical PDF
                if toc_page_0based >= first_visible_page:
                    # Rebuild the TOC from this point onward
                    new_toc = [entry] + [
                        e for e in toc 
                        if e[0] == 1 and e[2] > entry[2]
                    ]
                    print(f"Fixed TOC: First chapter is now '{entry[1]}' (page {entry[2]})")
                    return new_toc
        
        return toc  # Fallback if no adjustment possible





    def show_menu(self):
        while True:
            print("\n=== Nokia PDF Processor ===")
            print(f"1. Toggle Test Mode (Currently: {'ON' if self.test_mode else 'OFF'})")
            print("2. Process All PDFs")
            print("3. Process Specific Chapter")
            print("4. Exit")
            
            try:
                choice = input("Select option (1-4): ").strip()
                if choice == "1":
                    self.test_mode = not self.test_mode
                elif choice == "2":
                    self.process_pdfs()
                elif choice == "3":
                    self.process_specific_chapter()
                elif choice == "4":
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

    def process_specific_chapter(self):
        pdfs = [f for f in os.listdir(self.script_dir) if f.lower().endswith('.pdf')]
        if not pdfs:
            print("No PDFs found in directory!")
            return

        print("\nAvailable PDFs:")
        for i, pdf in enumerate(pdfs, 1):
            print(f"{i}. {pdf}")
        
        try:
            pdf_choice = int(input("Select PDF (number): ")) - 1
            selected_pdf = pdfs[pdf_choice]
        except (ValueError, IndexError):
            print("Invalid selection!")
            return

        chapter_num = input("Enter chapter number to process: ").strip()
        if not chapter_num.isdigit():
            print("Chapter number must be digits only!")
            return

        print(f"\nProcessing chapter {chapter_num} from {selected_pdf}")
        try:
            self._process_pdf(selected_pdf, specific_chapter=chapter_num)
        except Exception as e:
            print(f"✗ Failed: {str(e)}")

    def _process_pdf(self, filename, specific_chapter=None):
        doc = fitz.open(os.path.join(self.script_dir, filename))
        book_title = self._sanitize(os.path.splitext(filename)[0])
        base_dir = os.path.join(self.script_dir, "processed_books", book_title)
        
        if not os.path.exists(base_dir):
            os.makedirs(base_dir)

        # Chapter detection from TOC
        toc = doc.get_toc()
        chapter_ranges = self._get_chapter_ranges(doc, toc, filename)
        
        # If processing specific chapter, find its range
        if specific_chapter:
            chapter_found = False
            for chapter_name, (start, end) in self._flatten_chapter_ranges(chapter_ranges).items():
                # Extract all numbers from chapter name
                chap_numbers = re.findall(r'\d+', chapter_name)
                # Check if any of them match the requested chapter (with leading zeros stripped)
                if any(num.lstrip('0') == specific_chapter.lstrip('0') for num in chap_numbers):
                    print(f"Found chapter: {chapter_name} (pages {start}-{end})")
                    self._process_chapter(doc, start, end, chapter_name, base_dir, chapter_ranges)
                    chapter_found = True
                    break
            
            if not chapter_found:
                print(f"Chapter {specific_chapter} not found in table of contents!")
            doc.close()
            return

        # Process all pages with chapter validation
        max_pages = self.test_page_limit if self.test_mode else len(doc)
        for page_num in range(min(len(doc), max_pages)):
            try:
                chapter_dir = self._get_chapter_folder(page_num, chapter_ranges, base_dir)
                self._process_page(doc.load_page(page_num), page_num, chapter_dir)
            except Exception as e:
                print(f"Page {page_num} error: {str(e)}")

        doc.close()

    def _flatten_chapter_ranges(self, chapter_ranges):
        """Flatten the chapter ranges structure for easier searching"""
        flat_ranges = {}
        for key, value in chapter_ranges.items():
            if isinstance(value, dict):  # Part folder
                flat_ranges.update(value)
            else:  # Regular chapter
                flat_ranges[key] = value
        return flat_ranges

    def _process_chapter(self, doc, start_page, end_page, chapter_name, base_dir, chapter_ranges):
        """Process only a specific chapter range"""
        chapter_dir = self._get_chapter_folder(start_page, chapter_ranges, base_dir)
        
        print(f"Processing pages {start_page}-{end_page} to {chapter_dir}")
        for page_num in range(start_page, min(end_page + 1, len(doc))):
            try:
                self._process_page(doc.load_page(page_num), page_num, chapter_dir)
            except Exception as e:
                print(f"Page {page_num} error: {str(e)}")

    def _get_chapter_ranges(self, doc, toc, filename):  # Added filename parameter
        """Returns {chapter_name: (start_page, end_page)} with 4-digit numbering for numbered chapters"""
        

        #NEW CODEFIX FOR CHAPTER FOLDER TO PAGES MISMATCH IN CROPPED PDF
        # Apply TOC fix for trimmed PDFs
        toc = self._fix_trimmed_pdf_toc(doc.get_toc(), doc, filename)




        chapters = {}
        current_chapter = "0000-cover"
        start_page = 0
        numbered_chapters = []
        
        for entry in toc:
            if entry[0] == 1:  # Level-1 entries are chapters
                chapter_title = self._sanitize(entry[1])  # Sanitize early!
                # Check if chapter title contains a number
                numbers = re.findall(r'\d+', chapter_title)
                
                if numbers:
                    # Format numbers to 4 digits but keep original text
                    numbered_chapters.append((start_page, entry[2] - 2, chapter_title))             #OFFSET ERROR FIX HERE
                else:
                    # Unnumbered section - goes to cover or back matter
                    if not chapters:  # First section goes to cover
                        chapters[current_chapter] = (start_page, entry[2] - 2)
                    else:  # Later unnumbered sections go to back matter
                        if "9999-back" not in chapters:
                            chapters["9999-back"] = (start_page, len(doc) - 1)
                        else:
                            # Extend back matter range
                            old_start, old_end = chapters["9999-back"]
                            chapters["9999-back"] = (old_start, entry[2] - 2)
                
                start_page = entry[2] - 1                                                        #OFFSET ERROR FIX HERE

        # Handle numbered chapters (group into parts if more than 100)
        if numbered_chapters:
            if len(numbered_chapters) > 100:
                # Split into parts of max 100 chapters each
                part_num = 1
                for i in range(0, len(numbered_chapters), 100):
                    part_name = f"part-{part_num:03d}"
                    part_chapters = numbered_chapters[i:i+100]
                    
                    # Initialize part dictionary
                    chapters[part_name] = {}
                    
                    # First part includes cover if it exists
                    if part_num == 1 and "0000-cover" in chapters:
                        chapters[part_name]["0000-cover"] = chapters["0000-cover"]
                        del chapters["0000-cover"]
                    
                    # Add numbered chapters to part
                    for chap_start, chap_end, chap_title in part_chapters:
                        safe_title = self._sanitize(chap_title)
                        chapters[part_name][safe_title] = (chap_start, chap_end)
                    
                    # Last part includes back matter if it exists
                    if i + 100 >= len(numbered_chapters) and "9999-back" in chapters:
                        chapters[part_name]["9999-back"] = chapters["9999-back"]
                        del chapters["9999-back"]
                    
                    part_num += 1
            else:
                # Less than 100 chapters - add directly to root
                for chap_start, chap_end, chap_title in numbered_chapters:
                    safe_title = self._sanitize(chap_title)
                    chapters[safe_title] = (chap_start, chap_end)
                
                # Include cover and back if they exist
                if "0000-cover" in chapters:
                    chapters["0000-cover"] = chapters["0000-cover"]
                if "9999-back" in chapters:
                    chapters["9999-back"] = chapters["9999-back"]

        # Final fallback if no chapters were detected
        if not chapters:
            chapters["0000-cover"] = (0, len(doc) - 1)
        
        return chapters

    def _get_chapter_folder(self, page_num, chapter_ranges, base_dir):
        """Returns the correct chapter folder for a page, handling part subfolders"""
        # First check parts (if any)
        for part_name, part_data in chapter_ranges.items():
            if isinstance(part_data, dict):  # This is a part folder
                part_path = os.path.join(base_dir, self._sanitize(part_name))
                os.makedirs(part_path, exist_ok=True)
                
                for chapter, (start, end) in part_data.items():
                    if start <= page_num <= end:
                        folder = os.path.join(part_path, self._sanitize(chapter))
                        os.makedirs(folder, exist_ok=True)
                        return folder
            else:  # Regular chapter (cover or back)
                if part_name in ["0000-cover", "9999-back"]:
                    start, end = part_data
                    if start <= page_num <= end:
                        folder = os.path.join(base_dir, part_name)
                        os.makedirs(folder, exist_ok=True)
                        return folder
        
        # Fallback for any unclassified pages
        folder = os.path.join(base_dir, "0000-cover")
        os.makedirs(folder, exist_ok=True)
        return folder

    def _crop_whitespace(self, img):
        """Precision four-edge cropping with content detection"""
        arr = np.array(img.convert("L"))
        threshold = 220  # Text detection threshold
        margin = int(min(arr.shape) * 0.03)  # 3% protection margin

        def find_edge(projection, margin, threshold):
            for i in range(margin, len(projection) - margin):
                if projection[i] < threshold:
                    return max(0, i - margin)
            return margin

        h_proj = np.min(arr, axis=1)
        v_proj = np.min(arr, axis=0)

        top = find_edge(h_proj, margin, threshold)
        bottom = arr.shape[0] - find_edge(h_proj[::-1], margin, threshold)
        left = find_edge(v_proj, margin, threshold)
        right = arr.shape[1] - find_edge(v_proj[::-1], margin, threshold)

        if (right - left) < 50 or (bottom - top) < 50:
            return img
        
        return img.crop((left, top, right, bottom))

    def _process_page(self, page, page_num, output_dir):
        pix = page.get_pixmap(dpi=144)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        gray = ImageOps.autocontrast(img.convert("L"), cutoff=5)
        
        cropped = self._crop_whitespace(gray)
        if not cropped:
            cropped = gray

        width, height = cropped.size

        top_split = cropped.crop((
            0,
            0,
            width,
            int(height * 0.55)
        ))
        
        bottom_split = cropped.crop((
            0,
            int(height * 0.45),
            width,
            height
        ))

        def save_split(img, suffix):
            output_path = os.path.join(output_dir, f"{page_num+1:04d}{suffix}.bmp")
            # Ensure path isn't too long
            if len(output_path) > self.max_path_length:
                output_path = output_path[:self.max_path_length-4] + ".bmp"
            img.convert("1", dither=Image.FLOYDSTEINBERG) \
               .transpose(Image.ROTATE_270) \
               .save(output_path)
            return output_path

        if np.array(top_split.convert("L")).min() < 240:
            save_split(top_split, "a")
        if np.array(bottom_split.convert("L")).min() < 240:
            save_split(bottom_split, "b")

    def _sanitize(self, text):
        """Enhanced sanitization for Windows paths with length limits"""
        # Remove all problematic characters
        sanitized = re.sub(r'[\\/*?:"<>|()…]', "", text)
        # Replace spaces with underscores
        sanitized = sanitized.replace(" ", "_")
        # Trim to reasonable length (leave room for path components)
        max_length = 50  # Conservative limit for individual folder names
        return sanitized[:max_length]

if __name__ == "__main__":
    os.system('cls' if os.name == 'nt' else 'clear')
    PDFProcessor().show_menu()