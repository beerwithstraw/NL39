"""
Dedicated parser for ICICI Lombard NL-39.
ICICI publishes a 3-page PDF:
  - Page 0: Total (Leader + Follower)
  - Page 1: Leader (The one the user wants)
  - Page 2: Follower

Each page contains stacked QTR and YTD tables.
"""

import logging
from extractor.models import CompanyExtract, PeriodData
from extractor.normaliser import clean_number
from extractor.parser import _detect_lob_col, _split_stacked_table, _extract_page
import pdfplumber
from pathlib import Path
from config.company_registry import COMPANY_DISPLAY_NAMES

logger = logging.getLogger(__name__)

def parse_icici_lombard(pdf_path: str, company_key: str, quarter: str = "", year: str = "") -> CompanyExtract:
    logger.info(f"Using dedicated ICICI parser for: {pdf_path}")
    
    company_name = COMPANY_DISPLAY_NAMES.get(company_key, "ICICI Lombard")
    
    extract = CompanyExtract(
        source_file=Path(pdf_path).name,
        company_key=company_key,
        company_name=company_name,
        form_type="NL39",
        quarter=quarter,
        year=year,
    )
    
    period_data = PeriodData(period_label="current")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # We specifically want Page 1 (Leader data)
            if len(pdf.pages) < 2:
                logger.error(f"ICICI PDF {pdf_path} has fewer than 2 pages. Cannot extract Leader data.")
                extract.extraction_errors.append("PDF too short for Leader data (need page 2)")
                return extract
            
            page = pdf.pages[1]
            table = page.extract_table()
            if not table:
                logger.error(f"No table found on Page 1 of {pdf_path}")
                extract.extraction_errors.append("No table found on Page 2")
                return extract
            
            lob_col = _detect_lob_col(table)
            qtr_rows, ytd_rows = _split_stacked_table(table)
            
            n_qtr = _extract_page(qtr_rows, company_key, "qtr", period_data, lob_col)
            logger.info(f"Extracted {n_qtr} QTR LOBs from Page 1 (Leader)")
            
            if ytd_rows:
                n_ytd = _extract_page(ytd_rows, company_key, "ytd", period_data, lob_col)
                logger.info(f"Extracted {n_ytd} YTD LOBs from Page 1 (Leader)")
            else:
                logger.warning("No YTD section found on Page 1 (Leader)")
                extract.extraction_warnings.append("YTD section not found in stacked table")
                
    except Exception as e:
        logger.error(f"Dedicated ICICI parser failed: {e}", exc_info=True)
        extract.extraction_errors.append(str(e))
        return extract

    extract.current_year = period_data
    return extract
