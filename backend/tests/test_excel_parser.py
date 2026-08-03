"""Excel model parser regression coverage.

Run with: python -m pytest tests/ -q     (from the backend/ directory)
"""
from pathlib import Path

import openpyxl

from app.parsing.excel_parser import extract_facts


def test_memo_charts_sheet_takes_priority_over_the_whole_workbook_scan(tmp_path):
    """A "Memo Charts" tab is meant to hold the deck's own numbers -- when one
    exists, the exhibit should read from it exclusively rather than the
    whole-workbook scan, which can pick up an unrelated figure sitting next
    to the same label on an underwriting tab."""
    wb = openpyxl.Workbook()
    underwriting = wb.active
    underwriting.title = "Underwriting"
    underwriting["A1"] = "Total Sources"
    underwriting["B1"] = 200000  # a decoy figure from an unrelated adjacent cell

    memo = wb.create_sheet("Memo Charts")
    memo["A1"] = "Total Sources"
    memo["B1"] = 42065000

    path = tmp_path / "model.xlsx"
    wb.save(str(path))

    base_facts, _ = extract_facts(str(path))
    assert base_facts["total_sources"] == "$42,065,000"


def test_no_memo_charts_sheet_falls_back_to_the_whole_workbook_scan(tmp_path):
    """Without a Memo Charts tab, behavior is unchanged: scan every sheet."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Underwriting"
    ws["A1"] = "Total Sources"
    ws["B1"] = 42065000

    path = tmp_path / "model.xlsx"
    wb.save(str(path))

    base_facts, _ = extract_facts(str(path))
    assert base_facts["total_sources"] == "$42,065,000"
