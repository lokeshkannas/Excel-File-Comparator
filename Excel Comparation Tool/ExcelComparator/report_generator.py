\
from __future__ import annotations

from typing import Dict, Any, List
import os
import pandas as pd

from comparison_engine import ComparisonResult


def _ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def generate_report(output_path: str, result: ComparisonResult) -> str:
    """Create a multi-sheet Excel report and return the path."""
    out_dir = os.path.dirname(output_path)
    if out_dir:
        _ensure_dir(out_dir)

    # Build DataFrames
    summary_df = pd.DataFrame([result.summary])

    struct_df = pd.DataFrame([
        {"sheet": s.sheet, "issue": s.issue, "detail": s.detail}
        for s in result.structure_issues
    ])

    dtype_df = pd.DataFrame([
        {"sheet": d.sheet, "column": d.column, "ssrs_dtype": d.ssrs_dtype, "powerbi_dtype": d.powerbi_dtype}
        for d in result.dtype_issues
    ])

    mismatch_df = pd.DataFrame([
        {"sheet": v.sheet, "row": v.row, "column": v.column, "ssrs_value": v.ssrs_value, "powerbi_value": v.powerbi_value}
        for v in result.value_mismatches
    ])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        (struct_df if not struct_df.empty else pd.DataFrame(columns=["sheet","issue","detail"])).to_excel(writer, sheet_name="Structure_Issues", index=False)
        (dtype_df if not dtype_df.empty else pd.DataFrame(columns=["sheet","column","ssrs_dtype","powerbi_dtype"])).to_excel(writer, sheet_name="Dtype_Issues", index=False)
        (mismatch_df if not mismatch_df.empty else pd.DataFrame(columns=["sheet","row","column","ssrs_value","powerbi_value"])).to_excel(writer, sheet_name="Value_Mismatches", index=False)

    return output_path
