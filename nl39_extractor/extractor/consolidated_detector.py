"""
consolidated_detector.py

Finds the page range of the NL-39 form within a consolidated PDF.

A consolidated PDF contains multiple IRDAI forms merged into one file.
This module scans page text to find where the NL-39 Ageing of Claims
section starts and ends.

Detection strategy:
  START: First page where >= min_matches NL-39 keywords appear
  END:   Page before the next form header appears, or last page of PDF

NL-39 is a 2-page form (page 0 = For the Quarter, page 1 = Upto the Quarter).
"""

import re
import logging
import tempfile
import os
from typing import Optional, Tuple, List

logger = logging.getLogger(__name__)

DEFAULT_KEYWORDS = [
    "NL-39",
    "AGEING OF CLAIMS",
    "NO. OF CLAIMS PAID",       # column header present only on the data table
    "AMOUNT OF CLAIMS PAID",    # column header present only on the data table
]

# Minimum matches to distinguish the actual NL-39 data page from TOC entries.
# TOC pages match at most 2 of these (NL-39 + AGEING OF CLAIMS);
# the real form page matches all 4.
DEFAULT_MIN_MATCHES = 3

# Regex to detect any IRDAI form header on a page
FORM_HEADER_PATTERN = re.compile(
    r"^\s*(?:FORM\s+)?NL[-\s]?(\d+)|\bFORM\s+NL[-\s]?(\d+)", 
    re.IGNORECASE | re.MULTILINE
)
def is_toc_page(text: str) -> bool:
    if re.search(r"TABLE\s+OF\s+CONTENTS|FORM\s+INDEX|INDEX\s+OF\s+FORMS", text, re.IGNORECASE):
        return True
    matches = re.findall(r"\bNL[-\s]?(\d+)\b", text, re.IGNORECASE)
    valid_forms = set(m for m in matches if 1 <= int(m) <= 45)
    return len(valid_forms) >= 4

def _page_keyword_count(text: str, keywords: List[str]) -> int:
    """Count how many keywords appear in the page text (case-insensitive)."""
    text_upper = text.upper()
    return sum(1 for kw in keywords if kw.upper() in text_upper)


def find_nl39_pages(
    pdf_path: str,
    keywords: Optional[List[str]] = None,
    min_matches: int = DEFAULT_MIN_MATCHES,
) -> Optional[Tuple[int, int]]:
    """
    Scan the consolidated PDF and return (start_page, end_page) 0-indexed
    for the NL-39 section. Returns None if not found.

    start_page and end_page are both inclusive.
    """
    try:
        import pdfplumber
    except ImportError:
        logger.error("pdfplumber not available")
        return None

    if keywords is None:
        keywords = DEFAULT_KEYWORDS

    try:
        with pdfplumber.open(pdf_path) as pdf:
            n_pages = len(pdf.pages)
            page_texts = []

            for page in pdf.pages:
                try:
                    text = page.extract_text() or ""
                except Exception:
                    text = ""
                page_texts.append(text)

        # --- Find start page ---
        start_page = None
        for i, text in enumerate(page_texts):
            if is_toc_page(text):
                logger.debug(f"  page {i + 1}: TOC page, skipping")
                continue
            if _page_keyword_count(text, keywords) >= min_matches:
                start_page = i
                break

        if start_page is None:
            logger.warning(f"NL-39 section not found in: {pdf_path}")
            return None

        # --- Find end page ---
        # Stop when a DIFFERENT form number appears after the start page.
        end_page = n_pages - 1  # default: end of document

        for i in range(start_page + 1, n_pages):
            text = page_texts[i]
            matches = FORM_HEADER_PATTERN.findall(text)
            flat_matches = []
            for m in matches:
                flat_matches.extend(g for g in m if g)
            non_nl39 = [m for m in flat_matches if m != "39"]
            if non_nl39:
                end_page = i - 1
                logger.debug(
                    f"NL-39 ends at page {end_page} "
                    f"(NL-{non_nl39[0]} starts at page {i})"
                )
                break

        logger.info(
            f"NL-39 found at pages {start_page}-{end_page} "
            f"(0-indexed) in {os.path.basename(pdf_path)}"
        )
        return (start_page, end_page)

    except Exception as e:
        logger.error(f"Error scanning consolidated PDF {pdf_path}: {e}")
        return None


def extract_nl39_to_temp(
    pdf_path: str,
    start_page: int,
    end_page: int,
) -> Optional[str]:
    """
    Extract pages start_page..end_page from pdf_path into a temporary PDF file.
    Returns the path to the temp file, or None on failure.
    Caller is responsible for deleting the temp file after use.
    """
    try:
        import pypdf
    except ImportError:
        try:
            import PyPDF2 as pypdf
        except ImportError:
            logger.error("pypdf or PyPDF2 not available — cannot extract pages")
            return None

    try:
        reader = pypdf.PdfReader(pdf_path)
        writer = pypdf.PdfWriter()

        for page_num in range(start_page, end_page + 1):
            if page_num < len(reader.pages):
                writer.add_page(reader.pages[page_num])

        tmp = tempfile.NamedTemporaryFile(
            suffix=".pdf", delete=False, prefix="nl39_extract_"
        )
        with open(tmp.name, "wb") as f:
            writer.write(f)

        logger.debug(f"Extracted pages {start_page}-{end_page} to {tmp.name}")
        return tmp.name

    except Exception as e:
        logger.error(f"Error extracting pages from {pdf_path}: {e}")
        return None
