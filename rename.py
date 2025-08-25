import os
import re
import logging
from datetime import datetime
import PyPDF2
import pdfplumber

# -----------------------
# Configuration
# -----------------------
INPUT_PDF = "hackthon.pdf"           # Your input multi-page PDF
OUTPUT_FOLDER = "renamed_certificates"  # Output folder
SKIPPED_LOG = "skipped_files.txt"       # Log for failed pages

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("certificate_renamer.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# -----------------------
# Helper Functions
# -----------------------

def sanitize_filename(name):
    """Keep only safe characters for filenames"""
    name = re.sub(r'[<>:"/\\|?*]', '_', name)  # Replace invalid chars
    name = re.sub(r'[^\w\s\-\_]', '', name)   # Remove non-ASCII
    name = re.sub(r'\s+', ' ', name).strip()  # Normalize spaces
    name = name.strip(" ._-")
    if not name:
        raise ValueError("Filename became empty after sanitization")
    return name[:50]  # Limit length

def should_skip_line(line):
    """
    Check if a line is likely not a person's name.
    Less aggressive to avoid skipping actual names.
    """
    line_lower = line.lower().strip()
    
    # Skip only obvious header/footer/template lines
    skip_keywords = [
        'certificate', 'certifies', 'congratulations', 'participant',
        'diploma', 'completion', 'attendance', 'workshop', 'seminar',
        'course', 'training', 'program', 'event', 'this', 'is awarded to',
        'hackathon', 'team darion', 'darion', 'date', 'signed', 'winner',
        'certificate of', 'issued to', 'presented to', 'name', 'student',
        'the', 'and', 'for', 'this is to', 'we present', 'in recognition of',
        'please find', 'below', 'attached', 'email', 'gmail', 'outlook'
    ]
    
    # If it looks like a real name (2+ words, starts with letter), don't skip
    words = line.split()
    if len(words) >= 2 and all(w.isalpha() for w in words):
        first = words[0].lower()
        if first not in ['certificate', 'this', 'issued', 'presented', 'participant', 'congratulations']:
            return False  # Could be a name

    return any(kw in line_lower for kw in skip_keywords)

def extract_participant_name_from_page(page):
    """
    Extract the participant's name using:
    1. Layout-based text extraction
    2. Largest font size detection (likely the name)
    3. Smart filtering
    """
    # Extract full text first
    text = page.extract_text()
    if not text:
        raise ValueError("No text found on page")

    lines = [line.strip() for line in text.split('\n') if line.strip()]
    if not lines:
        raise ValueError("No readable text on page")

    # Try 1: Use largest font size (most likely the name)
    try:
        words = page.extract_words(extra_attrs=["size"])
        if words:
            # Group by line (y-position)
            lines_with_size = {}
            for word in words:
                y = round(word['top'])
                lines_with_size.setdefault(y, []).append(word)

            # Reconstruct lines and get their average font size
            line_data = []
            for y, word_list in lines_with_size.items():
                line_text = ' '.join(w['text'] for w in word_list)
                avg_size = sum(w['size'] for w in word_list) / len(word_list)
                line_data.append((avg_size, line_text.strip()))

            # Sort by font size (descending)
            line_data.sort(reverse=True, key=lambda x: x[0])

            # Check top 3 largest lines
            for size, line in line_data[:3]:
                if not should_skip_line(line):
                    clean_name = re.sub(r'[^a-zA-Z\s]', ' ', line)
                    clean_name = re.sub(r'\s+', ' ', clean_name).strip()
                    if len(clean_name) >= 2 and clean_name.lower() != 'name':
                        return clean_name
    except Exception as e:
        logger.debug(f"Font size analysis failed: {e}")

    # Try 2: Use normal line extraction
    for line in lines:
        if should_skip_line(line):
            continue
        # Clean the line
        clean_name = re.sub(r'[^a-zA-Z\s]', ' ', line)
        clean_name = re.sub(r'\s+', ' ', clean_name).strip()
        if len(clean_name) >= 2 and clean_name.lower() != 'name':
            return clean_name

    # Final fallback
    raise ValueError(f"Could not extract name. Found: {lines[:5]}")

# -----------------------
# Main Processing
# -----------------------

def split_and_rename_certificates():
    """Split PDF and save each page as Name.pdf"""
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # Reset skipped log
    with open(SKIPPED_LOG, "w", encoding='utf-8') as f:
        f.write(f"Skipped Files Log - {datetime.now()}\n")
        f.write("=" * 60 + "\n\n")

    if not os.path.exists(INPUT_PDF):
        logger.error(f"❌ Input PDF not found: {INPUT_PDF}")
        return

    try:
        with pdfplumber.open(INPUT_PDF) as pdf:
            total_pages = len(pdf.pages)
            logger.info(f"📄 Loaded '{INPUT_PDF}' with {total_pages} pages")
    except Exception as e:
        logger.error(f"❌ Failed to open PDF with pdfplumber: {e}")
        return

    try:
        pdf_reader = PyPDF2.PdfReader(INPUT_PDF)
    except Exception as e:
        logger.error(f"❌ Failed to read PDF with PyPDF2: {e}")
        return

    processed = 0
    skipped = 0
    used_names = {}  # Track duplicates: Harsha → Harsha.pdf, Harsha_1.pdf

    for page_idx in range(total_pages):
        page_num = page_idx + 1
        try:
            logger.info(f"🔍 Processing page {page_num}/{total_pages}")

            # Extract text and name
            with pdfplumber.open(INPUT_PDF) as pdf:
                page = pdf.pages[page_idx]
                raw_name = extract_participant_name_from_page(page)
                name = sanitize_filename(raw_name)

            # Handle duplicates
            final_name = name
            if name in used_names:
                used_names[name] += 1
                final_name = f"{name}_{used_names[name]}"
            else:
                used_names[name] = 0

            # Create filename
            output_filename = f"{final_name}.pdf"
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)

            # Avoid overwriting
            counter = 1
            while os.path.exists(output_path):
                output_filename = f"{final_name}_{counter}.pdf"
                output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                counter += 1

            # Save single page
            writer = PyPDF2.PdfWriter()
            writer.add_page(pdf_reader.pages[page_idx])
            with open(output_path, 'wb') as out_file:
                writer.write(out_file)

            logger.info(f"✅ Saved: {output_filename}")
            processed += 1

        except Exception as e:
            error_msg = f"Page {page_num} - ❌ Skipped | Reason: {str(e)}"
            logger.error(error_msg)
            with open(SKIPPED_LOG, "a", encoding='utf-8') as log:
                log.write(error_msg + "\n")
            skipped += 1

    # Final summary
    logger.info("\n" + "=" * 60)
    logger.info("🎉 Certificate Splitting & Renaming Complete!")
    logger.info(f"📄 Total Pages: {total_pages}")
    logger.info(f"✅ Successfully Created: {processed}")
    logger.info(f"❌ Skipped: {skipped}")
    logger.info(f"📁 Output Folder: {os.path.abspath(OUTPUT_FOLDER)}")
    logger.info(f"📝 Skipped Details: {SKIPPED_LOG}")
    logger.info("📋 Generated Files:")
    for f in sorted(os.listdir(OUTPUT_FOLDER)):
        if f.endswith(".pdf"):
            logger.info(f"   • {f}")

# -----------------------
# Run Script
# -----------------------

if __name__ == "__main__":
    try:
        split_and_rename_certificates()
    except Exception as e:
        logger.critical(f"💥 Fatal error: {e}")
        raise