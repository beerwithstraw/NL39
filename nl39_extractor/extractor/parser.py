"""
NL-39 Generic Parser.

NL-39 (Ageing of Claims) has a transposed layout vs NL-4/NL-5/NL-7:
  - PDF rows   = LOBs  (Fire, Marine Cargo, Motor OD, ...)
  - PDF cols   = Metrics (ageing buckets: upto 1m, >1m-3m, ..., totals)
  - Page 1 (index 0) = "For the Quarter"   → qtr values
  - Page 2 (index 1) = "Upto the Quarter"  → ytd values

No CY/PY split — NL-39 is a single-period form.
Only extract.current_year is populated; extract.prior_year is always None.

Output model (identical to other extractors):
  data[lob_key][metric_key] = {"qtr": <float|None>, "ytd": <float|None>}
"""

import logging
from pathlib import Path

import pdfplumber

from config.company_registry import COMPANY_DISPLAY_NAMES, DEDICATED_PARSER
from config.lob_registry import LOB_ALIASES, COMPANY_SPECIFIC_ALIASES
from config.row_registry import COLUMN_SCHEMA
from extractor.models import CompanyExtract, PeriodData
from extractor.normaliser import clean_number, normalise_text

logger = logging.getLogger(__name__)


def _resolve_lob(raw_label: str, company_key: str):
    """
    Normalise a PDF row label → canonical LOB key.
    Returns None if the row is a header / footer / blank / unrecognised.
    """
    normalised = normalise_text(str(raw_label or ""))
    if not normalised:
        return None

    # Company-specific override first
    company_aliases = COMPANY_SPECIFIC_ALIASES.get(company_key, {})
    if normalised in company_aliases:
        return company_aliases[normalised]

    return LOB_ALIASES.get(normalised)


def _extract_page(table, company_key: str, period_key: str, period_data: PeriodData) -> int:
    """
    Parse one pdfplumber table (one page) and store values into period_data.

    Args:
        table:       list[list[str|None]] from pdfplumber.extract_table()
        company_key: e.g. "bajaj_allianz"
        period_key:  "qtr" (page 0) or "ytd" (page 1)
        period_data: PeriodData mutated in-place

    Returns number of LOB rows extracted.
    """
    lobs_found = 0

    for row in table:
        if not row or len(row) < 3:
            continue

        # Col 1 = LOB label; col 0 = Sl.No. (ignored)
        raw_label = row[1] if len(row) > 1 else (row[0] or "")
        lob_key = _resolve_lob(raw_label, company_key)
        if lob_key is None:
            continue

        if lob_key not in period_data.data:
            period_data.data[lob_key] = {}

        for col_idx, metric_key in COLUMN_SCHEMA.items():
            if col_idx >= len(row):
                continue
            val = clean_number(row[col_idx])
            if metric_key not in period_data.data[lob_key]:
                period_data.data[lob_key][metric_key] = {"qtr": None, "ytd": None}
            if val is not None:
                period_data.data[lob_key][metric_key][period_key] = val

        lobs_found += 1

    return lobs_found


def parse_pdf(pdf_path: str, company_key: str, quarter: str = "", year: str = "") -> CompanyExtract:
    """
    Main entry point — parses one NL-39 PDF.

    Page 0 → qtr values (For the Quarter)
    Page 1 → ytd values (Upto the Quarter)
    Both merged into a single PeriodData stored in extract.current_year.
    """
    logger.info(f"Parsing NL-39 PDF: {pdf_path} for company: {company_key}")

    company_name = COMPANY_DISPLAY_NAMES.get(company_key, str(company_key).title())

    # Route to dedicated parser if registered
    dedicated_func_name = DEDICATED_PARSER.get(company_key)
    if dedicated_func_name:
        from extractor.companies import PARSER_REGISTRY
        dedicated_func = PARSER_REGISTRY.get(dedicated_func_name)
        if dedicated_func:
            logger.info(f"Routing to dedicated parser: {dedicated_func_name}")
            return dedicated_func(pdf_path, company_key, quarter, year)

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
            n_pages = len(pdf.pages)

            if n_pages == 1:
                # Single-page layout: both QTR and YTD tables stacked on page 0.
                # Use extract_tables() to get both; first = qtr, second = ytd.
                tables = pdf.pages[0].extract_tables()
                if len(tables) >= 2:
                    n = _extract_page(tables[0], company_key, "qtr", period_data)
                    logger.debug(f"Page 0 table[0] (qtr): {n} LOBs extracted")
                    n = _extract_page(tables[1], company_key, "ytd", period_data)
                    logger.debug(f"Page 0 table[1] (ytd): {n} LOBs extracted")
                elif len(tables) == 1:
                    n = _extract_page(tables[0], company_key, "qtr", period_data)
                    logger.debug(f"Page 0 (qtr only, 1 table): {n} LOBs extracted")
                else:
                    logger.warning(f"Page 0: no table found in {pdf_path}")
            else:
                # Two-page layout: page 0 = For the Quarter, page 1 = Upto the Quarter
                if n_pages >= 1:
                    table = pdf.pages[0].extract_table()
                    if table:
                        n = _extract_page(table, company_key, "qtr", period_data)
                        logger.debug(f"Page 0 (qtr): {n} LOBs extracted")
                    else:
                        logger.warning(f"Page 0: no table found in {pdf_path}")

                if n_pages >= 2:
                    table = pdf.pages[1].extract_table()
                    if table:
                        n = _extract_page(table, company_key, "ytd", period_data)
                        logger.debug(f"Page 1 (ytd): {n} LOBs extracted")
                    else:
                        logger.warning(f"Page 1: no table found in {pdf_path}")

    except Exception as e:
        logger.error(f"Failed to parse {pdf_path}: {e}", exc_info=True)
        extract.extraction_errors.append(str(e))
        return extract

    lobs_with_data = len(period_data.data)
    if lobs_with_data == 0:
        logger.warning(f"No LOBs extracted from {pdf_path}")
        extract.extraction_warnings.append("No LOBs extracted")
    else:
        logger.info(f"Extraction complete: {lobs_with_data} LOBs")

    extract.current_year = period_data
    return extract
