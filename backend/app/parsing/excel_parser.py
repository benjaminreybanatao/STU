"""
Underwriting Excel model parser.

Every shop lays its model out differently, so instead of hard-coding
absolute cell references (A1-style) this uses label-adjacent scanning:
for each metric we care about, search every cell in the workbook for text
that matches one of that metric's label patterns, then read the value out
of the cell immediately to the right (same row) or, failing that, directly
below it. That mirrors how a human skims a model -- find the row label,
read the number next to it -- and is far more robust to a firm's specific
layout than fixed coordinates.

If a specific firm's model layout is known, pass a `cell_map` (dict of
key -> "SheetName!A1") to skip the heuristic and read exact cells instead --
that is the "map the model's specific cell layout" hook the spec calls for.

Scenario sheets: any sheet whose name matches the downside/bear pattern is
scanned into a second result set, so a model with a Base Case tab and a
Downside tab produces two comparable KPI sets for slide 2 / slide 2b.
"""
import re
from datetime import datetime

import openpyxl

LABEL_PATTERNS: dict[str, list[str]] = {
    "purchase_price": [r"purchase price", r"acquisition price", r"total purchase price"],
    "price_per_unit": [r"price\s*/\s*unit", r"price\s*per\s*unit", r"price\s*/\s*sf", r"price\s*per\s*sf"],
    "going_in_cap": [r"going[- ]in cap", r"entry cap rate", r"year 1 cap rate", r"purchase cap rate"],
    "exit_cap": [r"exit cap", r"terminal cap", r"reversion cap"],
    "leverage": [r"leverage", r"ltv", r"loan[- ]to[- ]value"],
    "debt_rate": [r"interest rate", r"debt rate", r"spread", r"coupon"],
    "initial_equity": [r"initial equity", r"equity required", r"total equity invested"],
    "peak_equity": [r"peak equity", r"max(?:imum)? equity"],
    "hold_period": [r"hold period", r"holding period", r"investment period"],
    "unlevered_irr": [r"unlevered irr", r"unleveraged\.? irr"],
    "levered_irr": [r"(?<!un)levered irr", r"(?<!un)leveraged\.? irr"],
    "unlevered_em": [r"unlevered equity multiple", r"unlevered em\b"],
    "levered_em": [r"(?<!un)levered equity multiple", r"(?<!un)levered em\b"],
    "cash_on_cash": [r"cash[- ]on[- ]cash", r"\bcoc\b"],
    "sf_or_units": [r"total (?:rentable )?(?:square feet|sf|rsf)", r"unit count", r"number of units"],
    "occupancy": [r"occupancy", r"leased %", r"% leased"],
    "walt": [r"walt", r"weighted average lease term"],
    "lease_term_assumption": [r"lease term assumption", r"average lease term"],
    "downtime_assumption": [r"downtime assumption", r"months? vacant", r"downtime \(months\)"],
    "exit_assumption": [r"exit assumption", r"sale assumption", r"disposition assumption"],
}

DOWNSIDE_SHEET_RE = re.compile(r"downside|bear|stress|sensitivity", re.IGNORECASE)
BASE_SHEET_RE = re.compile(r"base\s*case|upside|primary", re.IGNORECASE)

PERCENT_KEYS = {"going_in_cap", "exit_cap", "leverage", "debt_rate", "unlevered_irr", "levered_irr", "cash_on_cash", "occupancy"}
CURRENCY_KEYS = {"purchase_price", "price_per_unit", "initial_equity", "peak_equity"}
MULTIPLE_KEYS = {"unlevered_em", "levered_em"}


def _normalize(key: str, raw) -> str | None:
    if raw is None or raw == "":
        return None
    if isinstance(raw, datetime):
        return raw.strftime("%Y-%m-%d")
    if isinstance(raw, (int, float)):
        if key in PERCENT_KEYS:
            # Excel stores 8% as 0.08; also tolerate models that store "8" literally.
            value = raw * 100 if abs(raw) <= 1 else raw
            return f"{value:.1f}%"
        if key in CURRENCY_KEYS:
            return f"${raw:,.0f}"
        if key in MULTIPLE_KEYS:
            return f"{raw:.2f}x"
        return str(raw)
    text = str(raw).strip()
    return text or None


def _compiled_patterns() -> dict[str, list[re.Pattern]]:
    return {key: [re.compile(p, re.IGNORECASE) for p in pats] for key, pats in LABEL_PATTERNS.items()}


def _scan_sheet(ws, patterns: dict[str, list[re.Pattern]]) -> dict:
    found: dict = {}
    max_row = min(ws.max_row, 500)
    max_col = min(ws.max_column, 60)
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            if not isinstance(cell.value, str):
                continue
            label = cell.value.strip()
            if not label:
                continue
            for key, regs in patterns.items():
                if key in found:
                    continue
                if any(r.search(label) for r in regs):
                    value = ws.cell(row=row, column=col + 1).value
                    if value in (None, ""):
                        value = ws.cell(row=row + 1, column=col).value
                    normalized = _normalize(key, value)
                    if normalized is not None:
                        found[key] = normalized
    return found


def _scan_workbook(wb, sheet_filter=None) -> dict:
    patterns = _compiled_patterns()
    results: dict = {}
    for ws in wb.worksheets:
        if sheet_filter is not None and not sheet_filter(ws.title):
            continue
        sheet_results = _scan_sheet(ws, patterns)
        for key, value in sheet_results.items():
            results.setdefault(key, value)
    return results


def extract_facts(xlsx_path: str, cell_map: dict | None = None) -> dict:
    """Returns (base_case_facts, downside_case_facts)."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    if cell_map:
        base_facts = {}
        for key, ref in cell_map.items():
            sheet_name, coord = ref.split("!")
            ws = wb[sheet_name]
            base_facts[key] = _normalize(key, ws[coord].value)
        wb.close()
        return base_facts, {}

    has_named_scenarios = any(DOWNSIDE_SHEET_RE.search(ws.title) for ws in wb.worksheets)

    if has_named_scenarios:
        base_facts = _scan_workbook(wb, sheet_filter=lambda name: not DOWNSIDE_SHEET_RE.search(name))
        downside_facts = _scan_workbook(wb, sheet_filter=lambda name: bool(DOWNSIDE_SHEET_RE.search(name)))
    else:
        base_facts = _scan_workbook(wb)
        downside_facts = {}

    wb.close()
    return base_facts, downside_facts
