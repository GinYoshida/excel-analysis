"""Extract cross-sheet references from worksheet formulas."""
from __future__ import annotations

import re

from openpyxl.utils import column_index_from_string

# Matches  Sheet!A1  or  'Sheet Name'!A1:B2 .  Captures sheet name and ref.
_SHEET_REF = re.compile(
    r"(?:'(?P<q>[^']+)'|(?P<u>[A-Za-z0-9_　-鿿＀-￯]+))"
    r"!\$?(?P<ref>[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)"
)


# Any cell/range token, optionally sheet-qualified. The lookbehind/lookahead
# keep it from matching inside identifiers or function names (e.g. LOG10().
_CELL_REF = re.compile(
    r"(?<![A-Za-z0-9_!])"
    r"(?:(?:'(?P<q>[^']+)'|(?P<u>[A-Za-z0-9_　-鿿＀-￯]+))!)?"
    r"(?P<c1>\$?[A-Z]{1,3}\$?\d+)"
    r"(?::(?P<c2>\$?[A-Z]{1,3}\$?\d+))?"
    r"(?![A-Za-z0-9_(])"
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


def extract_cell_dependencies(ws) -> list[dict]:
    """For every formula cell, list the cells/ranges it directly depends on
    (same-sheet refs unqualified, cross-sheet refs carry their sheet name).

    Returns dicts: {src_cell, formula, precedents: [{sheet|None, ref}]}.
    """
    out: list[dict] = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.data_type != "f" or cell.value is None:
                continue
            formula = str(cell.value)
            precedents: list[dict] = []
            seen: set[tuple] = set()
            for m in _CELL_REF.finditer(formula):
                sheet = m.group("q") or m.group("u")
                ref = m.group("c1").replace("$", "")
                if m.group("c2"):
                    ref += ":" + m.group("c2").replace("$", "")
                # A bare same-sheet self-reference to this very cell is noise.
                if sheet is None and ref == cell.coordinate:
                    continue
                key = (sheet, ref)
                if key in seen:
                    continue
                seen.add(key)
                precedents.append({"sheet": sheet, "ref": ref})
            if precedents:
                out.append({"src_cell": cell.coordinate, "formula": formula,
                            "precedents": precedents})
    return out
