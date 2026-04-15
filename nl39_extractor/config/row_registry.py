"""
Metric (Ageing Bucket) Registry for NL-39.

NL-39 has a transposed layout vs NL-4/NL-5/NL-7:
  - Rows in the PDF  = LOBs  (Fire, Marine Cargo, ...)
  - Columns in the PDF = Metrics (ageing buckets + totals)

These 16 canonical metric keys correspond to the fixed column positions
in every NL-39 PDF table.
"""

# Canonical metric keys — ordered to match PDF column layout (left→right).
ROW_ORDER = [
    # --- No. of claims paid ---
    "count_upto_1m",        # col 2: upto 1 month
    "count_1m_3m",          # col 3: >1 month and <=3 months
    "count_3m_6m",          # col 4: >3 months and <=6 months
    "count_6m_1y",          # col 5: >6 months and <=1 year
    "count_1y_3y",          # col 6: >1 year and <=3 years
    "count_3y_5y",          # col 7: >3 years and <=5 years
    "count_gt5y",           # col 8: >5 years
    # --- Amount of claims paid ---
    "amount_upto_1m",       # col 9: upto 1 month
    "amount_1m_3m",         # col 10: >1 month and <=3 months
    "amount_3m_6m",         # col 11: >3 months and <=6 months
    "amount_6m_1y",         # col 12: >6 months and <=1 year
    "amount_1y_3y",         # col 13: >1 year and <=3 years
    "amount_3y_5y",         # col 14: >3 years and <=5 years
    "amount_gt5y",          # col 15: >5 years
    # --- Totals ---
    "total_count",          # col 16: Total No. of claims paid
    "total_amount",         # col 17: Total amount of claims paid
]

# Display-friendly names for verification sheets and Excel headers.
ROW_DISPLAY_NAMES = {
    "count_upto_1m":    "No. of Claims — Upto 1 Month",
    "count_1m_3m":      "No. of Claims — >1M to 3M",
    "count_3m_6m":      "No. of Claims — >3M to 6M",
    "count_6m_1y":      "No. of Claims — >6M to 1Y",
    "count_1y_3y":      "No. of Claims — >1Y to 3Y",
    "count_3y_5y":      "No. of Claims — >3Y to 5Y",
    "count_gt5y":       "No. of Claims — >5Y",
    "amount_upto_1m":   "Amount — Upto 1 Month",
    "amount_1m_3m":     "Amount — >1M to 3M",
    "amount_3m_6m":     "Amount — >3M to 6M",
    "amount_6m_1y":     "Amount — >6M to 1Y",
    "amount_1y_3y":     "Amount — >1Y to 3Y",
    "amount_3y_5y":     "Amount — >3Y to 5Y",
    "amount_gt5y":      "Amount — >5Y",
    "total_count":      "Total No. of Claims Paid",
    "total_amount":     "Total Amount of Claims Paid",
}

# Fixed column index (0-based in pdfplumber extracted table) → metric key.
# Columns 0 and 1 are Sl.No. and Line of Business (LOB label) — skipped.
COLUMN_SCHEMA = {
    2:  "count_upto_1m",
    3:  "count_1m_3m",
    4:  "count_3m_6m",
    5:  "count_6m_1y",
    6:  "count_1y_3y",
    7:  "count_3y_5y",
    8:  "count_gt5y",
    9:  "amount_upto_1m",
    10: "amount_1m_3m",
    11: "amount_3m_6m",
    12: "amount_6m_1y",
    13: "amount_1y_3y",
    14: "amount_3y_5y",
    15: "amount_gt5y",
    16: "total_count",
    17: "total_amount",
}

# Count bucket keys (for validation sum checks)
COUNT_BUCKETS = ["count_upto_1m", "count_1m_3m", "count_3m_6m", "count_6m_1y",
                 "count_1y_3y", "count_3y_5y", "count_gt5y"]

# Amount bucket keys (for validation sum checks)
AMOUNT_BUCKETS = ["amount_upto_1m", "amount_1m_3m", "amount_3m_6m", "amount_6m_1y",
                  "amount_1y_3y", "amount_3y_5y", "amount_gt5y"]
