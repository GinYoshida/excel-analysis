"""Collect free-text notes (long text cells + cell comments) for the AI agent."""
from __future__ import annotations

NOTES_BUDGET = 60


def extract_notes(ws, min_len: int = 20) -> list[dict]:
    """Return natural-language memos on this sheet: long non-formula text cells
    and cell comments. Numbers, formulas, and short labels are skipped."""
    out: list[dict] = []
    for row in ws.iter_rows():
        for cell in row:
            comment = getattr(cell, "comment", None)
            if comment is not None and comment.text:
                out.append({"sheet": ws.title, "addr": cell.coordinate,
                            "text": str(comment.text).strip(), "kind": "comment"})
            if cell.data_type == "s" and isinstance(cell.value, str):
                text = cell.value.strip()
                if len(text) >= min_len:
                    out.append({"sheet": ws.title, "addr": cell.coordinate,
                                "text": text, "kind": "text"})
    return out
