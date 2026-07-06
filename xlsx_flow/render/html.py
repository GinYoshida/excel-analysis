"""Render a Model into a single self-contained interactive HTML file."""
from __future__ import annotations

import html as _html
import json
from pathlib import Path

from xlsx_flow.model import Model

_TEMPLATE_DIR = Path(__file__).parent / "template"


def _read(name: str) -> str:
    return (_TEMPLATE_DIR / name).read_text(encoding="utf-8")


def _warnings_html(warnings: list[str]) -> str:
    if not warnings:
        return ""
    items = "".join("<div>⚠ " + _html.escape(w) + "</div>" for w in warnings)
    return items


def render_html(model: Model) -> str:
    template = _read("index.html")
    model_json = json.dumps(model.to_dict(), ensure_ascii=False)
    # Escape "</" so a "</script>" inside any string value (e.g. m_code)
    # cannot terminate the inline <script> block that embeds MODEL.
    model_json = model_json.replace("</", "<\\/")
    replacements = {
        "__STYLE_CSS__": _read("style.css"),
        "__CYTOSCAPE_JS__": _read("cytoscape.min.js"),
        "__APP_JS__": _read("app.js"),
        "__CHAT_JS__": _read("chat.js"),
        "__MODEL_JSON__": model_json,
        "__WARNINGS_HTML__": _warnings_html(model.warnings),
    }
    out = template
    for key, val in replacements.items():
        out = out.replace(key, val)
    return out
