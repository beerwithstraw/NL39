"""
Validation Checks for NL-39 (Ageing of Claims).

Three check families:
  1. COUNT_BUCKET_SUM  — sum of 7 count buckets ≈ total_count per LOB
  2. AMOUNT_BUCKET_SUM — sum of 7 amount buckets ≈ total_amount per LOB
  3. QTR_LE_YTD        — qtr total_count ≤ ytd total_count per LOB (cumulative)
  4. COMPLETENESS      — mandatory LOBs present (fire, health, motor_od, motor_tp)
"""

import csv
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional
from dataclasses import dataclass, asdict

from extractor.models import CompanyExtract, PeriodData
from config.lob_registry import LOB_ORDER
from config.row_registry import COUNT_BUCKETS, AMOUNT_BUCKETS
from config.company_registry import COMPLETENESS_IGNORE, BUCKET_SUM_IGNORE

logger = logging.getLogger(__name__)

TOLERANCE = 2.0   # rounding tolerance for bucket-sum checks (PDFs round independently)

MANDATORY_LOBS = {"fire", "health", "motor_od", "motor_tp"}


@dataclass
class ValidationResult:
    company: str
    quarter: str
    year: str
    lob: str
    period: str   # "current_qtr", "current_ytd"
    check_name: str
    status: str   # PASS, WARN, FAIL, SKIP
    expected: Optional[float]
    actual: Optional[float]
    delta: Optional[float]
    note: str


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def run_validations(extractions: List[CompanyExtract]) -> List[ValidationResult]:
    """Run all NL-39 validation checks on the provided extractions."""
    results: List[ValidationResult] = []

    for exc in extractions:
        period_data = exc.current_year
        if not period_data:
            continue

        for timewise in ["qtr", "ytd"]:
            p_id = f"current_{timewise}"
            for lob, lob_data in period_data.data.items():
                r = _check_count_bucket_sum(exc, lob, p_id, lob_data, timewise)
                if r:
                    results.append(r)
                r = _check_amount_bucket_sum(exc, lob, p_id, lob_data, timewise)
                if r:
                    results.append(r)

        # QTR ≤ YTD check (once per LOB, not per timewise)
        for lob, lob_data in period_data.data.items():
            r = _check_qtr_le_ytd(exc, lob, lob_data)
            if r:
                results.append(r)

        # Completeness
        results.extend(_check_completeness(exc, period_data))

    return results


# ---------------------------------------------------------------------------
# Individual checks
# ---------------------------------------------------------------------------

def _get_val(lob_data: Dict, metric: str, p_key: str) -> Optional[float]:
    v = lob_data.get(metric, {}).get(p_key)
    return float(v) if v is not None else None


def _check_count_bucket_sum(
    exc: CompanyExtract,
    lob: str,
    pid: str,
    lob_data: Dict,
    p_key: str,
) -> Optional[ValidationResult]:
    """COUNT_BUCKET_SUM: sum of the 7 count buckets should equal total_count."""
    total = _get_val(lob_data, "total_count", p_key)
    if total is None:
        return None

    bucket_vals = [_get_val(lob_data, b, p_key) for b in COUNT_BUCKETS]
    if all(v is None for v in bucket_vals):
        return None

    bucket_sum = sum(v for v in bucket_vals if v is not None)
    delta = abs(total - bucket_sum)
    status = "PASS" if delta <= TOLERANCE else "FAIL"
    return ValidationResult(
        exc.company_name, exc.quarter, exc.year, lob, pid,
        "COUNT_BUCKET_SUM", status,
        expected=bucket_sum, actual=total, delta=delta, note="",
    )


def _check_amount_bucket_sum(
    exc: CompanyExtract,
    lob: str,
    pid: str,
    lob_data: Dict,
    p_key: str,
) -> Optional[ValidationResult]:
    """AMOUNT_BUCKET_SUM: sum of the 7 amount buckets should equal total_amount."""
    if lob in BUCKET_SUM_IGNORE.get(exc.company_key, set()):
        return None
    total = _get_val(lob_data, "total_amount", p_key)
    if total is None:
        return None

    bucket_vals = [_get_val(lob_data, b, p_key) for b in AMOUNT_BUCKETS]
    if all(v is None for v in bucket_vals):
        return None

    bucket_sum = sum(v for v in bucket_vals if v is not None)
    delta = abs(total - bucket_sum)
    status = "PASS" if delta <= TOLERANCE else "FAIL"
    return ValidationResult(
        exc.company_name, exc.quarter, exc.year, lob, pid,
        "AMOUNT_BUCKET_SUM", status,
        expected=bucket_sum, actual=total, delta=delta, note="",
    )


def _check_qtr_le_ytd(
    exc: CompanyExtract,
    lob: str,
    lob_data: Dict,
) -> Optional[ValidationResult]:
    """QTR_LE_YTD: qtr total_count should be ≤ ytd total_count (claims are cumulative)."""
    qtr_val = _get_val(lob_data, "total_count", "qtr")
    ytd_val = _get_val(lob_data, "total_count", "ytd")
    if qtr_val is None or ytd_val is None:
        return None

    status = "PASS" if qtr_val <= ytd_val + TOLERANCE else "WARN"
    return ValidationResult(
        exc.company_name, exc.quarter, exc.year, lob, "current",
        "QTR_LE_YTD", status,
        expected=ytd_val, actual=qtr_val,
        delta=qtr_val - ytd_val if qtr_val > ytd_val else 0.0,
        note="" if status == "PASS" else "QTR total_count exceeds YTD",
    )


def _check_completeness(
    exc: CompanyExtract,
    period_data: PeriodData,
) -> List[ValidationResult]:
    """COMPLETENESS: mandatory LOBs must appear with at least one non-None value."""
    results = []
    ignore = COMPLETENESS_IGNORE.get(exc.company_key, set())
    for lob in LOB_ORDER:
        if lob in ignore:
            continue
        lob_data = period_data.data.get(lob, {})
        has_data = any(
            any(v is not None for v in cell.values())
            for cell in lob_data.values()
        )
        if not has_data:
            status = "FAIL" if lob in MANDATORY_LOBS else "WARN"
            results.append(ValidationResult(
                exc.company_name, exc.quarter, exc.year, lob, "current",
                "COMPLETENESS", status,
                expected=None, actual=None, delta=None,
                note=f"LOB '{lob}' is missing",
            ))
    return results


# ---------------------------------------------------------------------------
# Reporting helpers
# ---------------------------------------------------------------------------

def build_validation_summary_table(results: List[ValidationResult]):
    """Build a Rich Table summarising PASS/SKIP/WARN/FAIL counts."""
    from rich.table import Table
    counts: Dict[str, int] = {"PASS": 0, "SKIP": 0, "WARN": 0, "FAIL": 0}
    for r in results:
        counts[r.status] = counts.get(r.status, 0) + 1
    t = Table(title="Validation Summary")
    t.add_column("Status", style="bold")
    t.add_column("Count", justify="right")
    t.add_row("[green]PASS[/green]", str(counts["PASS"]))
    t.add_row("[blue]SKIP[/blue]", str(counts["SKIP"]))
    t.add_row("[yellow]WARN[/yellow]", str(counts["WARN"]))
    t.add_row("[red]FAIL[/red]", str(counts["FAIL"]))
    return t


def write_validation_report(results: List[ValidationResult], output_path: str):
    """Write validation results to a CSV file."""
    with open(output_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "company", "quarter", "year", "lob", "period",
            "check_name", "status", "expected", "actual", "delta", "note",
        ])
        writer.writeheader()
        for r in results:
            writer.writerow(asdict(r))
    logger.info(f"Validation report saved to {output_path}")
