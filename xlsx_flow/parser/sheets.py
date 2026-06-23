"""Worksheet-level analysis: type classification and (Task 4) header recovery."""
from __future__ import annotations


def classify_sheet(ws, pq_targets: set[str]) -> tuple[str, dict]:
    if ws.title in pq_targets:
        return "pq_output", {"formula_ratio": 0.0, "cell_count": 0}

    cell_count = 0
    formula_cells = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            cell_count += 1
            if cell.data_type == "f":
                formula_cells += 1

    ratio = (formula_cells / cell_count) if cell_count else 0.0
    detail = {"formula_ratio": ratio, "cell_count": cell_count}

    if formula_cells == 0:
        return "raw", detail
    if ratio > 0.5:
        return "formula", detail
    return "mixed", detail
