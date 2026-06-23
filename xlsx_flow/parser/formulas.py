"""Extract cross-sheet references from worksheet formulas."""
from __future__ import annotations

import re

from openpyxl.utils import column_index_from_string

# Matches  Sheet!A1  or  'Sheet Name'!A1:B2 .  Captures sheet name and ref.
_SHEET_REF = re.compile(
    r"(?:'(?P<q>[^']+)'|(?P<u>[A-Za-z0-9_　-鿿＀-￯]+))"
    r"!\$?(?P<ref>[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)"
)


def cell_to_column_index(cell_ref: str) -> int:
    m = re.match(r"\$?([A-Z]+)", cell_ref)
    if m is None:
        raise ValueError(f"invalid cell reference: {cell_ref!r}")
    return column_index_from_string(m.group(1))


def extract_references(ws) -> list[dict]:
    refs: list[dict] = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.data_type != "f" or cell.value is None:
                continue
            formula = str(cell.value)
            for m in _SHEET_REF.finditer(formula):
                sheet = m.group("q") or m.group("u")
                refs.append(
                    {
                        "src_cell": cell.coordinate,
                        "target_sheet": sheet,
                        "target_ref": m.group("ref"),
                        "formula": formula,
                    }
                )
    return refs
