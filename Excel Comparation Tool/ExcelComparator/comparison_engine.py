\
from __future__ import annotations

import math
from dataclasses import dataclass, asdict
from typing import Dict, List, Any, Tuple

import numpy as np
import pandas as pd


@dataclass
class StructureIssue:
    sheet: str
    issue: str
    detail: str


@dataclass
class DtypeIssue:
    sheet: str
    column: str
    ssrs_dtype: str
    powerbi_dtype: str


@dataclass
class ValueMismatch:
    sheet: str
    row: int
    column: str
    ssrs_value: Any
    powerbi_value: Any


@dataclass
class ComparisonResult:
    structure_issues: List[StructureIssue]
    dtype_issues: List[DtypeIssue]
    value_mismatches: List[ValueMismatch]
    summary: Dict[str, Any]

    def as_dict(self) -> Dict[str, Any]:
        return {
            "structure_issues": [asdict(x) for x in self.structure_issues],
            "dtype_issues": [asdict(x) for x in self.dtype_issues],
            "value_mismatches": [asdict(x) for x in self.value_mismatches],
            "summary": self.summary,
        }


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Keep original order; just ensure columns are strings
    df = df.copy()
    df.columns = [str(c) for c in df.columns]
    return df


def _safe_read_excel(path: str) -> Dict[str, pd.DataFrame]:
    """Read all sheets to DataFrames, preserving headers as-is."""
    sheets = pd.read_excel(path, sheet_name=None, dtype=None)  # allow dtype inference
    return {name: _normalize_columns(df) for name, df in sheets.items()}


def _dtype_name(dtype: Any) -> str:
    # Produce friendly dtype names
    d = str(dtype)
    if d.startswith("Int") or d.startswith("Float") or d.startswith("bool") or d.startswith("datetime64"):
        return d
    return d


def _is_number(x) -> bool:
    try:
        return np.isfinite(float(x))
    except Exception:
        return False


def _approx_equal(a, b, tol: float) -> bool:
    # Treat NaN/None/empty as equal if both missing
    if (a is None or (isinstance(a, float) and math.isnan(a)) or (isinstance(a, str) and a.strip() == "")) and \
       (b is None or (isinstance(b, float) and math.isnan(b)) or (isinstance(b, str) and b.strip() == "")):
        return True

    # Try numeric comparison with tolerance
    if _is_number(a) and _is_number(b):
        try:
            return abs(float(a) - float(b)) <= tol
        except Exception:
            pass

    # Fallback to string compare (trim)
    return str(a).strip() == str(b).strip()


def compare_workbooks(
    ssrs_path: str,
    powerbi_path: str,
    numeric_tolerance: float = 1e-9
) -> ComparisonResult:
    ssrs = _safe_read_excel(ssrs_path)
    pbi = _safe_read_excel(powerbi_path)

    structure_issues: List[StructureIssue] = []
    dtype_issues: List[DtypeIssue] = []
    value_mismatches: List[ValueMismatch] = []

    ssrs_sheets = set(ssrs.keys())
    pbi_sheets = set(pbi.keys())

    # Sheet presence
    missing_in_pbi = ssrs_sheets - pbi_sheets
    missing_in_ssrs = pbi_sheets - ssrs_sheets

    for s in sorted(missing_in_pbi):
        structure_issues.append(StructureIssue(sheet=s, issue="Missing in Power BI", detail="Sheet not found in Power BI workbook"))

    for s in sorted(missing_in_ssrs):
        structure_issues.append(StructureIssue(sheet=s, issue="Missing in SSRS", detail="Sheet not found in SSRS workbook"))

    # Compare common sheets
    common = sorted(ssrs_sheets & pbi_sheets)
    for sheet in common:
        df_a = ssrs[sheet].copy()
        df_b = pbi[sheet].copy()

        # Row/column counts
        if len(df_a) != len(df_b):
            structure_issues.append(StructureIssue(sheet=sheet, issue="Row count mismatch", detail=f"SSRS={len(df_a)}, PowerBI={len(df_b)}"))

        if df_a.shape[1] != df_b.shape[1]:
            structure_issues.append(StructureIssue(sheet=sheet, issue="Column count mismatch", detail=f"SSRS={df_a.shape[1]}, PowerBI={df_b.shape[1]}"))

        # Column names and order
        cols_a = list(df_a.columns)
        cols_b = list(df_b.columns)
        if cols_a != cols_b:
            structure_issues.append(StructureIssue(sheet=sheet, issue="Column order/name mismatch", detail=f"SSRS={cols_a}, PowerBI={cols_b}"))

        # Dtype comparison for intersection columns
        for col in set(cols_a).intersection(cols_b):
            da = _dtype_name(df_a[col].dtype)
            db = _dtype_name(df_b[col].dtype)
            if da != db:
                dtype_issues.append(DtypeIssue(sheet=sheet, column=col, ssrs_dtype=da, powerbi_dtype=db))

        # Align and compare values (cell-by-cell) on shared columns
        shared_cols = [c for c in cols_a if c in cols_b]
        # Reindex to max length to avoid dropping rows; compare row by row
        max_len = max(len(df_a), len(df_b))
        df_a2 = df_a.reindex(range(max_len))
        df_b2 = df_b.reindex(range(max_len))

        for r in range(max_len):
            for c in shared_cols:
                va = df_a2.at[r, c] if r < len(df_a2) else np.nan
                vb = df_b2.at[r, c] if r < len(df_b2) else np.nan
                if not _approx_equal(va, vb, numeric_tolerance):
                    value_mismatches.append(ValueMismatch(sheet=sheet, row=r+1, column=c, ssrs_value=va, powerbi_value=vb))

    total_mismatch_cells = len(value_mismatches)
    summary = {
        "sheets_ssrs": sorted(ssrs_sheets),
        "sheets_powerbi": sorted(pbi_sheets),
        "common_sheets": common,
        "structure_issue_count": len(structure_issues),
        "dtype_issue_count": len(dtype_issues),
        "value_mismatch_count": total_mismatch_cells,
        "all_matched": (len(structure_issues) == 0 and len(dtype_issues) == 0 and total_mismatch_cells == 0),
    }
    return ComparisonResult(structure_issues, dtype_issues, value_mismatches, summary)
