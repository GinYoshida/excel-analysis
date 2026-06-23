"""Extract PowerQuery (M) definitions from the xlsx DataMashup part."""
from __future__ import annotations

import base64
import io
import re
import struct
import zipfile
from dataclasses import dataclass, field

_SOURCE_FUNCS = [
    "Csv.Document",
    "Excel.Workbook",
    "Web.Contents",
    "Sql.Database",
    "File.Contents",
    "Json.Document",
]

# shared NAME =   where NAME is  #"quoted"  or a bare identifier
_QUERY_HEAD = re.compile(r'shared\s+(#"(?P<q>[^"]+)"|(?P<u>[A-Za-z_][A-Za-z0-9_.]*))\s*=')


@dataclass
class Query:
    name: str
    m_code: str
    sources: list[str] = field(default_factory=list)
    refs: list[str] = field(default_factory=list)


def _decode_datamashup(raw_b64: str) -> bytes:
    raw = base64.b64decode(raw_b64)
    # version(int32 LE), package length(int32 LE), then package zip
    _version, length = struct.unpack_from("<ii", raw, 0)
    return raw[8:8 + length]


def _read_section_m(xlsx_path: str, warnings: list[str]) -> str | None:
    try:
        with zipfile.ZipFile(xlsx_path) as z:
            items = [n for n in z.namelist()
                     if re.search(r"(^|/)customXml/item\d+\.xml$", n)]
            for name in items:
                xml = z.read(name).decode("utf-8", errors="replace")
                m = re.search(r"<DataMashup[^>]*>(.*?)</DataMashup>", xml, re.DOTALL)
                if not m:
                    continue
                b64 = m.group(1).strip()
                try:
                    package = _decode_datamashup(b64)
                    with zipfile.ZipFile(io.BytesIO(package)) as pz:
                        return pz.read("Formulas/Section1.m").decode("utf-8")
                except Exception as exc:  # noqa: BLE001 - report, never crash
                    warnings.append(f"DataMashup の復号に失敗: {exc}")
                    continue
    except Exception as exc:  # noqa: BLE001
        warnings.append(f"xlsx を zip として開けませんでした: {exc}")
        return None
    return None


def _split_queries(section_m: str) -> list[tuple[str, str]]:
    matches = list(_QUERY_HEAD.finditer(section_m))
    pairs: list[tuple[str, str]] = []
    for i, m in enumerate(matches):
        name = m.group("q") or m.group("u")
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(section_m)
        body = section_m[start:end].rstrip().rstrip(";")
        pairs.append((name, body))
    return pairs


def extract_powerquery(xlsx_path: str) -> tuple[list[Query], list[str]]:
    warnings: list[str] = []
    section = _read_section_m(xlsx_path, warnings)
    if section is None:
        if not warnings:
            warnings.append("PowerQuery(DataMashup)が見つかりませんでした")
        return [], warnings

    pairs = _split_queries(section)
    all_names = {name for name, _ in pairs}
    queries: list[Query] = []
    for name, body in pairs:
        sources = [fn for fn in _SOURCE_FUNCS if fn in body]
        refs = []
        for other in all_names:
            if other == name:
                continue
            # match #"other" or bare-word boundary occurrence
            pattern = r'#"' + re.escape(other) + r'"|(?<![\w.])' + re.escape(other) + r'(?![\w.])'
            if re.search(pattern, body):
                refs.append(other)
        queries.append(Query(name=name, m_code=body, sources=sources, refs=refs))
    return queries, warnings
