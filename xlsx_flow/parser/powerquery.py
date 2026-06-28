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

# Excel.CurrentWorkbook(){[Name="X"]} — an in-workbook table or named range.
_WB_ENTITY = re.compile(
    r'Excel\.CurrentWorkbook\(\)\s*\{\s*\[\s*Name\s*=\s*"([^"]+)"\s*\]\s*\}')

# Literal file path / URL arguments to external source functions.
_URI_ARG = re.compile(r'(?:File\.Contents|Web\.Contents)\(\s*"([^"]+)"')


def _dedup(seq: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for x in seq:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


@dataclass
class Step:
    name: str
    expr: str
    refs: list[str] = field(default_factory=list)


def _refs_in(expr: str, candidates: list[str], exclude: str) -> list[str]:
    out: list[str] = []
    for other in candidates:
        if other == exclude:
            continue
        patt = (r'#"' + re.escape(other) + r'"|(?<![\w.])'
                + re.escape(other) + r'(?![\w.])')
        if re.search(patt, expr):
            out.append(other)
    return out


def _split_let_bindings(s: str) -> list[str]:
    """Split the top-level bindings of a `let ... in ...` body, honouring
    brackets, strings (incl. #"quoted" names), and nested let/in."""
    m = re.match(r"\s*let\b", s)
    if not m:
        return []
    i, n = m.end(), len(s)
    depth = 0          # () [] {} nesting
    let_depth = 0      # nested let..in nesting
    buf: list[str] = []
    bindings: list[str] = []
    while i < n:
        c = s[i]
        if c == '"':  # string literal or #"quoted name" span
            buf.append(c)
            i += 1
            while i < n:
                buf.append(s[i])
                if s[i] == '"':
                    if i + 1 < n and s[i + 1] == '"':  # "" escape
                        buf.append('"')
                        i += 2
                        continue
                    i += 1
                    break
                i += 1
            continue
        if c in "([{":
            depth += 1
            buf.append(c)
            i += 1
            continue
        if c in ")]}":
            depth -= 1
            buf.append(c)
            i += 1
            continue
        if depth == 0 and (c.isalpha() or c == "_"):
            j = i
            while j < n and (s[j].isalnum() or s[j] in "_."):
                j += 1
            word = s[i:j]
            if word == "let":
                let_depth += 1
            elif word == "in" and let_depth == 0:
                bindings.append("".join(buf))
                return bindings
            elif word == "in":
                let_depth -= 1
            buf.append(word)
            i = j
            continue
        if c == "," and depth == 0 and let_depth == 0:
            bindings.append("".join(buf))
            buf = []
            i += 1
            continue
        buf.append(c)
        i += 1
    bindings.append("".join(buf))
    return bindings


_BINDING_HEAD = re.compile(
    r'\s*(#"(?P<q>[^"]+)"|(?P<u>[A-Za-z_][A-Za-z0-9_.]*))\s*=(?P<expr>.*)',
    re.DOTALL)


def decompose_steps(m_code: str) -> list[Step]:
    """Break a query's `let` body into ordered M steps with inter-step refs.
    Returns [] for single-expression (non-let) queries."""
    parsed: list[tuple[str, str]] = []
    for b in _split_let_bindings(m_code):
        if not b.strip():
            continue
        mm = _BINDING_HEAD.match(b)
        if not mm:
            continue
        name = mm.group("q") or mm.group("u")
        parsed.append((name, mm.group("expr").strip()))
    names = [n for n, _ in parsed]
    return [Step(name=name, expr=expr, refs=_refs_in(expr, names, name))
            for name, expr in parsed]


@dataclass
class Query:
    name: str
    m_code: str
    sources: list[str] = field(default_factory=list)
    refs: list[str] = field(default_factory=list)
    wb_entities: list[str] = field(default_factory=list)
    uris: list[str] = field(default_factory=list)


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
        wb_entities = _dedup(_WB_ENTITY.findall(body))
        uris = _dedup(_URI_ARG.findall(body))
        queries.append(Query(name=name, m_code=body, sources=sources, refs=refs,
                             wb_entities=wb_entities, uris=uris))
    return queries, warnings
