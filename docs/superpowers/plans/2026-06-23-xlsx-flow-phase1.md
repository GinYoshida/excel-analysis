# xlsx-flow Phase 1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Excel ワークブック（結合セル・複数シートまたぎ・PowerQuery）を解析し、データフローを自己完結のインタラクティブ HTML として可視化する CLI ツールを作る。

**Architecture:** 2 ステージ・パイプライン。`parser/*`（Python）が xlsx を読んで中間モデル（`model.py` の dataclass）を生成し、`render/*` がそのモデルを Cytoscape.js 埋め込みの 1 ファイル HTML に描画する。パーサは HTML を知らず、レンダラは xlsx を知らない。両者の唯一の契約が `model.py`。

**Tech Stack:** Python 3.11+ / openpyxl / 標準ライブラリ（zipfile, struct, xml.etree, re, json, base64, dataclasses）/ pytest。HTML は CDN 非依存で Cytoscape.js を vendoring してインライン。

## Global Constraints

- Python 3.11+（dataclass / `list[...]` 記法を使用）
- パッケージ名: `xlsx_flow`、配布名: `xlsx-flow`
- 依存は `openpyxl` のみ（実行時）。テストは `pytest`。それ以外の実行時依存を増やさない。
- 出力 HTML はネットワーク不要で単体で開ける（Cytoscape.js をインライン）。
- パーサは例外でクラッシュさせず、抽出失敗は `Model.warnings` に積んで部分結果を必ず返す。
- すべての文字列は UTF-8。日本語シート名・ヘッダを壊さない。
- 中間モデルのノード ID 規約: `sheet:<名>` / `col:<シート>/<親>/<子>` / `query:<名>` / `source:<名>`。

---

### Task 0: プロジェクト雛形とツールチェーン

**Files:**
- Create: `pyproject.toml`
- Create: `xlsx_flow/__init__.py`
- Create: `tests/__init__.py`
- Create: `tests/test_smoke.py`
- Create: `.gitignore`

**Interfaces:**
- Consumes: なし
- Produces: `import xlsx_flow` が成功し `xlsx_flow.__version__ == "0.1.0"`。`pytest` が走る環境。

- [ ] **Step 1: `.gitignore` を作成**

```
__pycache__/
*.pyc
*.egg-info/
.pytest_cache/
build/
dist/
*.xlsx
!tests/fixtures/*.xlsx
out.html
out.json
```

- [ ] **Step 2: `pyproject.toml` を作成**

```toml
[build-system]
requires = ["setuptools>=68"]
build-backend = "setuptools.build_meta"

[project]
name = "xlsx-flow"
version = "0.1.0"
description = "Visualize Excel workbook data flow (merged cells, cross-sheet refs, PowerQuery) as interactive HTML"
requires-python = ">=3.11"
dependencies = ["openpyxl>=3.1"]

[project.optional-dependencies]
dev = ["pytest>=7.4"]

[project.scripts]
xlsx-flow = "xlsx_flow.cli:main"

[tool.setuptools.packages.find]
include = ["xlsx_flow*"]

[tool.setuptools.package-data]
xlsx_flow = ["render/template/*"]
```

- [ ] **Step 3: パッケージ初期化ファイルを作成**

`xlsx_flow/__init__.py`:
```python
__version__ = "0.1.0"
```

`tests/__init__.py`: （空ファイル）

- [ ] **Step 4: スモークテストを書く**

`tests/test_smoke.py`:
```python
import xlsx_flow


def test_version():
    assert xlsx_flow.__version__ == "0.1.0"
```

- [ ] **Step 5: 依存をインストールして実行**

```bash
pip3 install -e ".[dev]"
pytest tests/test_smoke.py -v
```
Expected: PASS（1 passed）

- [ ] **Step 6: Commit**

```bash
git add pyproject.toml xlsx_flow/__init__.py tests/__init__.py tests/test_smoke.py .gitignore
git commit -m "chore: scaffold xlsx-flow package and tooling"
```

---

### Task 1: 中間データモデル（`model.py`）

**Files:**
- Create: `xlsx_flow/model.py`
- Test: `tests/test_model.py`

**Interfaces:**
- Consumes: なし
- Produces:
  - `Node(id: str, type: str, attrs: dict)` — `type` は `"source"|"query"|"sheet"|"column"`。
  - `Edge(source: str, target: str, edge_type: str, granularity: str = "L1", detail: str = "")` — `edge_type` は `"dataflow"|"reference"`、`granularity` は `"L1"|"L2"|"L3"|"pq"`。
  - `Model(file: str, nodes: list[Node], edges: list[Edge], warnings: list[str])`。
  - `Model.add_node(node)`, `Model.add_edge(edge)`, `Model.warn(msg)`。
  - `Model.to_dict() -> dict`, `Model.to_json(indent=2) -> str`。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_model.py`:
```python
import json
from xlsx_flow.model import Node, Edge, Model


def test_node_to_dict_flattens_attrs():
    n = Node(id="sheet:Sales", type="sheet", attrs={"sheet_type": "formula"})
    assert n.to_dict() == {"id": "sheet:Sales", "type": "sheet", "sheet_type": "formula"}


def test_model_collects_and_serializes():
    m = Model(file="sample.xlsx")
    m.add_node(Node(id="sheet:Sales", type="sheet", attrs={"sheet_type": "raw"}))
    m.add_edge(Edge(source="query:Q", target="sheet:Sales", edge_type="dataflow"))
    m.warn("could not parse query Z")

    data = json.loads(m.to_json())
    assert data["meta"]["file"] == "sample.xlsx"
    assert data["meta"]["sheet_count"] == 1
    assert data["nodes"][0]["id"] == "sheet:Sales"
    assert data["edges"][0]["edge_type"] == "dataflow"
    assert data["edges"][0]["granularity"] == "L1"
    assert data["warnings"] == ["could not parse query Z"]
```

- [ ] **Step 2: 実行して失敗を確認**

Run: `pytest tests/test_model.py -v`
Expected: FAIL（`ModuleNotFoundError: No module named 'xlsx_flow.model'`）

- [ ] **Step 3: 最小実装を書く**

`xlsx_flow/model.py`:
```python
from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime, timezone


@dataclass
class Node:
    id: str
    type: str  # "source" | "query" | "sheet" | "column"
    attrs: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {"id": self.id, "type": self.type, **self.attrs}


@dataclass
class Edge:
    source: str
    target: str
    edge_type: str  # "dataflow" | "reference"
    granularity: str = "L1"  # "L1" | "L2" | "L3" | "pq"
    detail: str = ""

    def to_dict(self) -> dict:
        return {
            "source": self.source,
            "target": self.target,
            "edge_type": self.edge_type,
            "granularity": self.granularity,
            "detail": self.detail,
        }


@dataclass
class Model:
    file: str
    nodes: list[Node] = field(default_factory=list)
    edges: list[Edge] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def add_node(self, node: Node) -> None:
        self.nodes.append(node)

    def add_edge(self, edge: Edge) -> None:
        self.edges.append(edge)

    def warn(self, msg: str) -> None:
        self.warnings.append(msg)

    def to_dict(self) -> dict:
        sheet_count = sum(1 for n in self.nodes if n.type == "sheet")
        return {
            "meta": {
                "file": self.file,
                "generated_at": datetime.now(timezone.utc).isoformat(),
                "sheet_count": sheet_count,
            },
            "nodes": [n.to_dict() for n in self.nodes],
            "edges": [e.to_dict() for e in self.edges],
            "warnings": list(self.warnings),
        }

    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)
```

- [ ] **Step 4: 実行して成功を確認**

Run: `pytest tests/test_model.py -v`
Expected: PASS（2 passed）

- [ ] **Step 5: Commit**

```bash
git add xlsx_flow/model.py tests/test_model.py
git commit -m "feat: add intermediate graph model"
```

---

### Task 2: 合成サンプル xlsx ジェネレータ（`samples/gen_sample.py`）

検証の土台。openpyxl でシートを作り、zip を後処理して DataMashup（PowerQuery）を注入する。

**Files:**
- Create: `samples/__init__.py`（空）
- Create: `samples/gen_sample.py`
- Test: `tests/test_gen_sample.py`

**Interfaces:**
- Consumes: なし
- Produces:
  - `build_datamashup(section_m: str) -> bytes` — `section Section1; ...` の M テキストから DataMashup の base64 デコード前バイト列（version+length+package zip）を作る。
  - `datamashup_base64(section_m: str) -> str` — 上記を base64 文字列化。
  - `generate(path: str) -> None` — 既知構造の 5 シート xlsx を `path` に書き出し、DataMashup customXml を注入する。
  - 生成されるシート（固定・テストが依存）:
    - `生データ`（raw、数式なし、縦持ち）
    - `売上集計`（formula、結合ヘッダ「売上→当月/前月/前月差」、`前月差`列に `=当月-前月`、`生データ` を参照）
    - `値貼り付け`（pasted、数式なし）
    - `罫線表頭`（罫線で 3 列を囲う擬似ヘッダ、結合なし）
    - `PQ出力`（pq_output、Q_売上 のロード先）
  - DataMashup の M（固定）:
    - `Q_元データ = Csv.Document(File.Contents("sales.csv"))`
    - `Q_売上 = ... Q_元データ を参照`

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_gen_sample.py`:
```python
import zipfile
import base64
import struct

import openpyxl

from samples.gen_sample import build_datamashup, generate


def test_build_datamashup_roundtrips_section_m():
    section = 'section Section1;\nshared Q = 1;'
    raw = build_datamashup(section)
    # header: version(int32 LE), package length(int32 LE), then package zip
    version, length = struct.unpack_from("<ii", raw, 0)
    assert version == 0
    pkg = raw[8:8 + length]
    with zipfile.ZipFile(__import__("io").BytesIO(pkg)) as z:
        assert z.read("Formulas/Section1.m").decode("utf-8") == section


def test_generate_creates_five_sheets_with_expected_names(tmp_path):
    out = tmp_path / "sample.xlsx"
    generate(str(out))
    wb = openpyxl.load_workbook(out)
    assert set(wb.sheetnames) == {"生データ", "売上集計", "値貼り付け", "罫線表頭", "PQ出力"}


def test_generate_injects_datamashup(tmp_path):
    out = tmp_path / "sample.xlsx"
    generate(str(out))
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        item = [n for n in names if n.startswith("customXml/item") and n.endswith(".xml")]
        assert item, f"no customXml item part found in {names}"
        xml = z.read(item[0]).decode("utf-8")
        assert "DataMashup" in xml


def test_generate_uri_has_merged_sales_header(tmp_path):
    out = tmp_path / "sample.xlsx"
    generate(str(out))
    wb = openpyxl.load_workbook(out)
    ws = wb["売上集計"]
    merged = [str(r) for r in ws.merged_cells.ranges]
    # "売上" spans 3 columns on the header row
    assert any(":" in r for r in merged), f"expected a merged range, got {merged}"
```

- [ ] **Step 2: 実行して失敗を確認**

Run: `pytest tests/test_gen_sample.py -v`
Expected: FAIL（`ModuleNotFoundError: No module named 'samples.gen_sample'`）

- [ ] **Step 3: 最小実装を書く**

`samples/__init__.py`: （空ファイル）

`samples/gen_sample.py`:
```python
"""Generate a synthetic .xlsx that mirrors the user's hard-to-read workbooks.

Sheets: raw vertical data, a formula sheet with a merged 3-column "売上" header,
a value-pasted sheet, a border-drawn pseudo-header sheet, and a PowerQuery load
target. A DataMashup (PowerQuery M) part is injected post-hoc because openpyxl
cannot write PowerQuery.
"""
from __future__ import annotations

import base64
import io
import re
import struct
import zipfile

import openpyxl
from openpyxl.styles import Border, Side

SECTION_M = """section Section1;

shared #"Q_元データ" = let
    Source = Csv.Document(File.Contents("sales.csv"))
in
    Source;

shared #"Q_売上" = let
    Source = #"Q_元データ",
    Added = Table.AddColumn(Source, "前月差", each [当月] - [前月])
in
    Added;
"""


def build_datamashup(section_m: str) -> bytes:
    """Build the pre-base64 DataMashup bytes: version + package-length + package zip.

    The package is a zip whose only entry is Formulas/Section1.m. Real Excel
    appends Permissions/Metadata sections after the package; our parser only
    reads the package, so a trailing-empty structure is sufficient here.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("Formulas/Section1.m", section_m)
        z.writestr("[Content_Types].xml", _PACKAGE_CONTENT_TYPES)
    package = buf.getvalue()
    header = struct.pack("<ii", 0, len(package))
    return header + package


def datamashup_base64(section_m: str) -> str:
    return base64.b64encode(build_datamashup(section_m)).decode("ascii")


_PACKAGE_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="utf-8"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="m" ContentType="application/x-ms-m"/>'
    '<Default Extension="xml" ContentType="text/xml"/>'
    "</Types>"
)


def _item_xml(section_m: str) -> str:
    b64 = datamashup_base64(section_m)
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">'
        f"{b64}</DataMashup>"
    )


def _build_workbook() -> openpyxl.Workbook:
    wb = openpyxl.Workbook()

    # Sheet 1: 生データ (raw, vertical/long format, no formulas)
    ws = wb.active
    ws.title = "生データ"
    ws.append(["月", "店舗", "売上"])
    for month, shop, val in [
        ("1月", "A店", 100), ("1月", "B店", 120),
        ("2月", "A店", 110), ("2月", "B店", 130),
    ]:
        ws.append([month, shop, val])

    # Sheet 2: 売上集計 (formula, merged 3-col header, references 生データ)
    ws = wb.create_sheet("売上集計")
    ws["A1"] = "店舗"
    ws["B1"] = "売上"            # merged across B:D
    ws.merge_cells("B1:D1")
    ws["B2"] = "当月"
    ws["C2"] = "前月"
    ws["D2"] = "前月差"
    ws["A3"] = "A店"
    ws["B3"] = 110
    ws["C3"] = 100
    ws["D3"] = "=B3-C3"          # 前月差 = 当月 - 前月
    ws["A4"] = "B店"
    ws["B4"] = "=生データ!C3"    # cross-sheet reference
    ws["C4"] = 120
    ws["D4"] = "=B4-C4"

    # Sheet 3: 値貼り付け (pasted, no formulas, dense values)
    ws = wb.create_sheet("値貼り付け")
    ws.append(["店舗", "当月", "前月", "前月差"])
    ws.append(["A店", 110, 100, 10])
    ws.append(["B店", 120, 130, -10])

    # Sheet 4: 罫線表頭 (border-drawn pseudo header, no merge)
    ws = wb.create_sheet("罫線表頭")
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws["A1"] = "売上"
    for col in ("A1", "B1", "C1"):
        ws[col].border = box
    ws["A2"] = "当月"
    ws["B2"] = "前月"
    ws["C2"] = "前月差"
    ws.append([])  # row 3 spacer handled by explicit rows below
    ws["A3"] = 110
    ws["B3"] = 100
    ws["C3"] = 10

    # Sheet 5: PQ出力 (PowerQuery load target placeholder)
    ws = wb.create_sheet("PQ出力")
    ws.append(["店舗", "当月", "前月", "前月差"])
    ws.append(["A店", 110, 100, 10])

    return wb


def _inject_datamashup(path: str, section_m: str) -> None:
    """Append a customXml DataMashup part into the already-written xlsx zip."""
    with zipfile.ZipFile(path) as zin:
        items = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in items}

    item_name = "customXml/item1.xml"
    data[item_name] = _item_xml(section_m).encode("utf-8")

    # Register content type for the customXml part.
    ct = data["[Content_Types].xml"].decode("utf-8")
    override = '<Override PartName="/customXml/item1.xml" ContentType="application/xml"/>'
    if override not in ct:
        ct = ct.replace("</Types>", override + "</Types>")
        data["[Content_Types].xml"] = ct.encode("utf-8")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, payload in data.items():
            zout.writestr(name, payload)


def generate(path: str) -> None:
    wb = _build_workbook()
    wb.save(path)
    _inject_datamashup(path, SECTION_M)


if __name__ == "__main__":
    import sys

    target = sys.argv[1] if len(sys.argv) > 1 else "sample.xlsx"
    generate(target)
    print(f"wrote {target}")
```

- [ ] **Step 4: 実行して成功を確認**

Run: `pytest tests/test_gen_sample.py -v`
Expected: PASS（4 passed）

- [ ] **Step 5: Commit**

```bash
git add samples/__init__.py samples/gen_sample.py tests/test_gen_sample.py
git commit -m "feat: add synthetic sample xlsx generator with injected PowerQuery"
```

---

### Task 3: シート種別判定（`parser/sheets.py`）

**Files:**
- Create: `xlsx_flow/parser/__init__.py`（空）
- Create: `xlsx_flow/parser/sheets.py`
- Test: `tests/test_sheets.py`

**Interfaces:**
- Consumes: openpyxl の `Worksheet`。
- Produces:
  - `classify_sheet(ws, pq_targets: set[str]) -> tuple[str, dict]` — 戻り値は `(sheet_type, detail)`。`sheet_type` は `"raw"|"pasted"|"formula"|"pq_output"|"mixed"`。`detail` は `{"formula_ratio": float, "cell_count": int}`。
  - 判定規則:
    - `ws.title in pq_targets` → `"pq_output"`。
    - 数式比率 > 0 かつ数式セルが 1 個以上 → `"formula"`（数式比率 < 1.0 でデータも多い場合は `"mixed"`）。具体的には: formula_cells == 0 → `"raw"`、0 < formula_ratio <= 0.5 → `"mixed"`、formula_ratio > 0.5 → `"formula"`。
    - 値のみ（formula_cells == 0）は既定 `"raw"`。呼び出し側が外部リンク痕跡で `"pasted"` に上書きする余地は将来対応（本タスクでは raw のまま）。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_sheets.py`:
```python
import openpyxl

from xlsx_flow.parser.sheets import classify_sheet
from samples.gen_sample import generate


def _load(tmp_path):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    return openpyxl.load_workbook(out)


def test_raw_sheet_has_no_formulas(tmp_path):
    wb = _load(tmp_path)
    stype, detail = classify_sheet(wb["生データ"], pq_targets=set())
    assert stype == "raw"
    assert detail["formula_ratio"] == 0.0


def test_formula_sheet_detected(tmp_path):
    wb = _load(tmp_path)
    stype, _ = classify_sheet(wb["売上集計"], pq_targets=set())
    assert stype in ("formula", "mixed")


def test_pq_target_overrides(tmp_path):
    wb = _load(tmp_path)
    stype, _ = classify_sheet(wb["PQ出力"], pq_targets={"PQ出力"})
    assert stype == "pq_output"
```

- [ ] **Step 2: 実行して失敗を確認**

Run: `pytest tests/test_sheets.py -v`
Expected: FAIL（`ModuleNotFoundError: No module named 'xlsx_flow.parser'`）

- [ ] **Step 3: 最小実装を書く**

`xlsx_flow/parser/__init__.py`: （空ファイル）

`xlsx_flow/parser/sheets.py`:
```python
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
```

- [ ] **Step 4: 実行して成功を確認**

Run: `pytest tests/test_sheets.py -v`
Expected: PASS（3 passed）

- [ ] **Step 5: Commit**

```bash
git add xlsx_flow/parser/__init__.py xlsx_flow/parser/sheets.py tests/test_sheets.py
git commit -m "feat: classify sheet type by formula ratio"
```

---

### Task 4: 結合セル → 論理ヘッダ復元（`parser/sheets.py` に追記）

**Files:**
- Modify: `xlsx_flow/parser/sheets.py`
- Test: `tests/test_headers.py`

**Interfaces:**
- Consumes: openpyxl の `Worksheet`。
- Produces:
  - `LogicalColumn` dataclass: `header_path: list[str]`, `col_letter: str`, `col_index: int`, `is_formula: bool`, `confidence: str`（`"high"|"low"`）。
  - `recover_headers(ws, max_header_rows: int = 3) -> list[LogicalColumn]`。
    - 横結合の親ラベルを配下列へ前方フィルし、子ラベルと結合して `header_path` を作る。
    - 結合がない罫線擬似ヘッダは、ヘッダ帯の各セル値を `header_path` にしつつ `confidence="low"`。
    - 結合由来は `confidence="high"`。
    - `is_formula` は当該列のデータ行に数式セルが 1 つでもあれば True。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_headers.py`:
```python
import openpyxl

from xlsx_flow.parser.sheets import recover_headers
from samples.gen_sample import generate


def _load(tmp_path, sheet):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    return openpyxl.load_workbook(out)[sheet]


def test_merged_header_builds_parent_child_path(tmp_path):
    ws = _load(tmp_path, "売上集計")
    cols = recover_headers(ws)
    paths = [c.header_path for c in cols]
    assert ["店舗"] in paths or ["店舗", ""] in [p for p in paths]
    assert ["売上", "当月"] in paths
    assert ["売上", "前月"] in paths
    assert ["売上", "前月差"] in paths


def test_merged_columns_are_high_confidence(tmp_path):
    ws = _load(tmp_path, "売上集計")
    cols = {tuple(c.header_path): c for c in recover_headers(ws)}
    assert cols[("売上", "前月差")].confidence == "high"
    assert cols[("売上", "前月差")].is_formula is True


def test_border_pseudo_header_is_low_confidence(tmp_path):
    ws = _load(tmp_path, "罫線表頭")
    cols = recover_headers(ws)
    assert cols, "expected some columns"
    assert all(c.confidence == "low" for c in cols)
```

- [ ] **Step 2: 実行して失敗を確認**

Run: `pytest tests/test_headers.py -v`
Expected: FAIL（`ImportError: cannot import name 'recover_headers'`）

- [ ] **Step 3: 最小実装を追記**

`xlsx_flow/parser/sheets.py` の先頭付近に追記:
```python
from dataclasses import dataclass, field
from openpyxl.utils import get_column_letter


@dataclass
class LogicalColumn:
    header_path: list[str]
    col_letter: str
    col_index: int
    is_formula: bool = False
    confidence: str = "high"
    sheet: str = ""


def _horizontal_merges(ws) -> dict[int, tuple[int, int, str]]:
    """Map each covered column index (1-based) on the merge's top row to
    (row, top_left_col, label) for horizontal merges spanning >1 column."""
    out: dict[int, tuple[int, int, str]] = {}
    for rng in ws.merged_cells.ranges:
        if rng.max_col > rng.min_col:  # horizontal span
            label = ws.cell(row=rng.min_row, column=rng.min_col).value
            label = "" if label is None else str(label)
            for col in range(rng.min_col, rng.max_col + 1):
                out[col] = (rng.min_row, rng.min_col, label)
    return out


def _column_has_formula(ws, col_index: int, start_row: int) -> bool:
    for row in range(start_row, ws.max_row + 1):
        cell = ws.cell(row=row, column=col_index)
        if cell.data_type == "f":
            return True
    return False


def recover_headers(ws, max_header_rows: int = 3) -> list[LogicalColumn]:
    merges = _horizontal_merges(ws)
    has_merge = bool(merges)
    header_rows = 2 if has_merge else 1
    data_start = header_rows + 1

    cols: list[LogicalColumn] = []
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        parts: list[str] = []
        confidence = "low"
        if col in merges:
            _, _, parent = merges[col]
            if parent:
                parts.append(parent)
            child = ws.cell(row=header_rows, column=col).value
            if child is not None and str(child) != "":
                parts.append(str(child))
            confidence = "high"
        else:
            top = ws.cell(row=1, column=col).value
            if top is not None and str(top) != "":
                parts.append(str(top))
            if has_merge:
                child = ws.cell(row=header_rows, column=col).value
                if child is not None and str(child) != "" and str(child) != parts[-1:] :
                    parts.append(str(child))

        if not parts:
            continue
        is_formula = _column_has_formula(ws, col, data_start)
        cols.append(
            LogicalColumn(
                header_path=parts,
                col_letter=get_column_letter(col),
                col_index=col,
                is_formula=is_formula,
                confidence=confidence,
                sheet=ws.title,
            )
        )
    return cols
```

> 注意: 上の `_horizontal_merges` は親ラベルが空の結合（罫線表頭シートのように単独セル）には反応しない。罫線表頭シートは結合を持たないため `has_merge=False` となり、ヘッダ 1 行・全列 `confidence="low"` で復元される（テスト期待どおり）。

- [ ] **Step 4: 実行して成功を確認**

Run: `pytest tests/test_headers.py -v`
Expected: PASS（3 passed）。失敗する場合は `売上集計` の `A1="店舗"` が `header_path=["店舗"]` になるか、子行 `A2` の空セル処理を確認。

- [ ] **Step 5: 全テスト実行**

Run: `pytest -v`
Expected: ここまでの全テスト PASS

- [ ] **Step 6: Commit**

```bash
git add xlsx_flow/parser/sheets.py tests/test_headers.py
git commit -m "feat: recover logical headers from merged cells with confidence flag"
```

---

### Task 5: シート間参照の抽出（`parser/formulas.py`）

**Files:**
- Create: `xlsx_flow/parser/formulas.py`
- Test: `tests/test_formulas.py`

**Interfaces:**
- Consumes: openpyxl の `Worksheet`、Task 4 の `LogicalColumn`。
- Produces:
  - `extract_references(ws) -> list[dict]` — 各要素 `{"src_cell": "B4", "target_sheet": "生データ", "target_ref": "C3", "formula": "=生データ!C3"}`。
    - シート修飾参照 `Sheet!A1` / `'シート 名'!A1:B2` を正規表現で抽出。
    - シート修飾のない参照（同一シート内 `=B3-C3` 等）は対象外（シート間フローのみ）。
  - `cell_to_column_index(cell_ref: str) -> int` — `"C3"` → 3（列インデックス）。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_formulas.py`:
```python
import openpyxl

from xlsx_flow.parser.formulas import extract_references, cell_to_column_index
from samples.gen_sample import generate


def test_cell_to_column_index():
    assert cell_to_column_index("A1") == 1
    assert cell_to_column_index("C3") == 3
    assert cell_to_column_index("AA10") == 27


def test_extracts_cross_sheet_reference(tmp_path):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    ws = openpyxl.load_workbook(out)["売上集計"]
    refs = extract_references(ws)
    targets = {(r["target_sheet"], r["target_ref"]) for r in refs}
    assert ("生データ", "C3") in targets


def test_same_sheet_formula_ignored(tmp_path):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    ws = openpyxl.load_workbook(out)["売上集計"]
    refs = extract_references(ws)
    # =B3-C3 has no sheet qualifier, must not appear
    assert all(r["target_sheet"] for r in refs)
```

- [ ] **Step 2: 実行して失敗を確認**

Run: `pytest tests/test_formulas.py -v`
Expected: FAIL（`ModuleNotFoundError: No module named 'xlsx_flow.parser.formulas'`）

- [ ] **Step 3: 最小実装を書く**

`xlsx_flow/parser/formulas.py`:
```python
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
    letters = re.match(r"\$?([A-Z]+)", cell_ref).group(1)
    return column_index_from_string(letters)


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
```

- [ ] **Step 4: 実行して成功を確認**

Run: `pytest tests/test_formulas.py -v`
Expected: PASS（3 passed）

- [ ] **Step 5: Commit**

```bash
git add xlsx_flow/parser/formulas.py tests/test_formulas.py
git commit -m "feat: extract cross-sheet references from formulas"
```

---

### Task 6: PowerQuery（DataMashup）抽出（`parser/powerquery.py`）

**Files:**
- Create: `xlsx_flow/parser/powerquery.py`
- Test: `tests/test_powerquery.py`

**Interfaces:**
- Consumes: xlsx ファイルパス（zip として開く）。
- Produces:
  - `Query` dataclass: `name: str`, `m_code: str`, `sources: list[str]`, `refs: list[str]`（参照する他クエリ名）。
  - `extract_powerquery(xlsx_path: str) -> tuple[list[Query], list[str]]` — `(queries, warnings)`。
    - `customXml/item*.xml` から DataMashup base64 を取り出し復号、package zip 内 `Formulas/Section1.m` を読む。
    - `shared <name> = ... ;` でクエリ分割（`#"..."` 形式の名前に対応）。
    - 各本文から `Csv.Document` / `Excel.Workbook` / `Web.Contents` / `Sql.Database` / `File.Contents` を検出して `sources` に。
    - 本文中に出現する他クエリ名（`#"名前"` または素の識別子）を `refs` に。
    - DataMashup 不在や復号失敗は warnings に積み、`([], warnings)` を返す（例外を投げない）。
  - `_decode_datamashup(raw_b64: str) -> bytes` — base64 → package zip バイト列（version+length ヘッダを外す）。
  - `_split_queries(section_m: str) -> list[tuple[str, str]]` — `[(name, body), ...]`。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_powerquery.py`:
```python
from xlsx_flow.parser.powerquery import extract_powerquery, _split_queries
from samples.gen_sample import generate, SECTION_M


def test_split_queries_finds_both_names():
    pairs = _split_queries(SECTION_M)
    names = {n for n, _ in pairs}
    assert "Q_元データ" in names
    assert "Q_売上" in names


def test_extract_finds_source_and_dependency(tmp_path):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    queries, warnings = extract_powerquery(str(out))
    by_name = {q.name: q for q in queries}
    assert "Q_元データ" in by_name
    assert "Q_売上" in by_name
    # Q_元データ reads a Csv source
    assert any("Csv.Document" in s or "File.Contents" in s for s in by_name["Q_元データ"].sources)
    # Q_売上 references Q_元データ
    assert "Q_元データ" in by_name["Q_売上"].refs


def test_missing_datamashup_returns_warning(tmp_path):
    import openpyxl
    out = tmp_path / "plain.xlsx"
    wb = openpyxl.Workbook()
    wb.save(out)
    queries, warnings = extract_powerquery(str(out))
    assert queries == []
    assert warnings  # at least one warning, no exception
```

- [ ] **Step 2: 実行して失敗を確認**

Run: `pytest tests/test_powerquery.py -v`
Expected: FAIL（`ModuleNotFoundError`）

- [ ] **Step 3: 最小実装を書く**

`xlsx_flow/parser/powerquery.py`:
```python
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
                     if n.startswith("customXml/item") and n.endswith(".xml")]
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
                    return None
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
```

- [ ] **Step 4: 実行して成功を確認**

Run: `pytest tests/test_powerquery.py -v`
Expected: PASS（3 passed）

- [ ] **Step 5: Commit**

```bash
git add xlsx_flow/parser/powerquery.py tests/test_powerquery.py
git commit -m "feat: extract PowerQuery M queries, sources, and dependencies"
```

---

### Task 7: ワークブック統括 → モデル生成（`parser/workbook.py`）

**Files:**
- Create: `xlsx_flow/parser/workbook.py`
- Test: `tests/test_workbook.py`

**Interfaces:**
- Consumes: Task 1 `Model/Node/Edge`、Task 3 `classify_sheet`、Task 4 `recover_headers`、Task 5 `extract_references`、Task 6 `extract_powerquery`。
- Produces:
  - `analyze(xlsx_path: str) -> Model` — 全パーサを束ね、ノード/エッジ/警告を集約した `Model` を返す。
    - 各 query → `Node(type="query")`、各 source → `Node(type="source")`、`source→query`/`query→query` の `dataflow` エッジ（granularity `"pq"`）。
    - 各 sheet → `Node(type="sheet", attrs={sheet_type, has_merged})`。
    - 各論理列 → `Node(type="column", attrs={header_path, is_formula, confidence})`（granularity の概念はノードには持たせず、列ノードは L2 でのみ表示する旨を render が判断）。
    - シート間参照 → `reference` エッジ。参照先列にマップできれば `col:...`→`col:...`（granularity `"L2"`）、できなければ `sheet:...`→`sheet:...`（granularity `"L1"`）。
    - PQ ロード先シートの推定は best-effort（`pq_targets` は connections 解析が未実装のため空集合で渡す。将来拡張）。クエリ名と一致するシートがあれば `query→sheet` の dataflow を張り、そのシートを pq_output として扱う。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_workbook.py`:
```python
from xlsx_flow.parser.workbook import analyze
from samples.gen_sample import generate


def test_analyze_produces_sheet_nodes(tmp_path):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    model = analyze(str(out))
    sheet_ids = {n.id for n in model.nodes if n.type == "sheet"}
    assert "sheet:売上集計" in sheet_ids
    assert "sheet:生データ" in sheet_ids


def test_analyze_produces_query_and_source_nodes(tmp_path):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    model = analyze(str(out))
    query_ids = {n.id for n in model.nodes if n.type == "query"}
    source_ids = {n.id for n in model.nodes if n.type == "source"}
    assert "query:Q_元データ" in query_ids
    assert "query:Q_売上" in query_ids
    assert source_ids  # at least one external source


def test_analyze_produces_cross_sheet_reference_edge(tmp_path):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    model = analyze(str(out))
    ref_edges = [e for e in model.edges if e.edge_type == "reference"]
    assert any(e.target.startswith("sheet:生データ") or "生データ" in e.target
               for e in ref_edges)


def test_analyze_produces_query_dependency_edge(tmp_path):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    model = analyze(str(out))
    df = [(e.source, e.target) for e in model.edges if e.edge_type == "dataflow"]
    assert ("query:Q_元データ", "query:Q_売上") in df
```

- [ ] **Step 2: 実行して失敗を確認**

Run: `pytest tests/test_workbook.py -v`
Expected: FAIL（`ModuleNotFoundError`）

- [ ] **Step 3: 最小実装を書く**

`xlsx_flow/parser/workbook.py`:
```python
"""Orchestrate all parsers and assemble the intermediate Model."""
from __future__ import annotations

import openpyxl

from xlsx_flow.model import Model, Node, Edge
from xlsx_flow.parser.sheets import classify_sheet, recover_headers
from xlsx_flow.parser.formulas import extract_references, cell_to_column_index
from xlsx_flow.parser.powerquery import extract_powerquery


def _col_id(sheet: str, header_path: list[str]) -> str:
    return "col:" + sheet + "/" + "/".join(header_path)


def analyze(xlsx_path: str) -> Model:
    model = Model(file=xlsx_path)

    # --- PowerQuery layer ---
    queries, pq_warnings = extract_powerquery(xlsx_path)
    for w in pq_warnings:
        model.warn(w)
    query_names = {q.name for q in queries}
    source_ids: set[str] = set()
    for q in queries:
        model.add_node(Node(id=f"query:{q.name}", type="query",
                            attrs={"m_code": q.m_code}))
        for src in q.sources:
            sid = f"source:{src}"
            if sid not in source_ids:
                model.add_node(Node(id=sid, type="source", attrs={"kind": "function"}))
                source_ids.add(sid)
            model.add_edge(Edge(source=sid, target=f"query:{q.name}",
                                edge_type="dataflow", granularity="pq"))
        for ref in q.refs:
            model.add_edge(Edge(source=f"query:{ref}", target=f"query:{q.name}",
                                edge_type="dataflow", granularity="pq"))

    # --- Workbook / sheets layer ---
    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    pq_targets = {name for name in wb.sheetnames if name in query_names}

    # Map sheet -> {col_index: LogicalColumn} for reference resolution.
    sheet_cols: dict[str, dict[int, list[str]]] = {}

    for ws in wb.worksheets:
        stype, detail = classify_sheet(ws, pq_targets)
        has_merged = bool(ws.merged_cells.ranges)
        model.add_node(Node(id=f"sheet:{ws.title}", type="sheet",
                            attrs={"sheet_type": stype, "has_merged": has_merged,
                                   "formula_ratio": round(detail["formula_ratio"], 3)}))
        cols = recover_headers(ws)
        sheet_cols[ws.title] = {c.col_index: c.header_path for c in cols}
        for c in cols:
            model.add_node(Node(
                id=_col_id(ws.title, c.header_path), type="column",
                attrs={"sheet": ws.title, "header_path": c.header_path,
                       "is_formula": c.is_formula, "confidence": c.confidence}))

    # query -> sheet dataflow when a sheet name matches a query name
    for q in queries:
        if q.name in wb.sheetnames:
            model.add_edge(Edge(source=f"query:{q.name}", target=f"sheet:{q.name}",
                                edge_type="dataflow", granularity="pq"))

    # --- Cross-sheet references ---
    for ws in wb.worksheets:
        for ref in extract_references(ws):
            tgt_sheet = ref["target_sheet"]
            if tgt_sheet not in {s.title for s in wb.worksheets}:
                continue
            # try to resolve source cell's column -> logical column on this sheet
            src_col_idx = cell_to_column_index(ref["src_cell"])
            src_path = sheet_cols.get(ws.title, {}).get(src_col_idx)
            tgt_col_idx = cell_to_column_index(ref["target_ref"].split(":")[0])
            tgt_path = sheet_cols.get(tgt_sheet, {}).get(tgt_col_idx)

            if src_path and tgt_path:
                model.add_edge(Edge(
                    source=_col_id(ws.title, src_path),
                    target=_col_id(tgt_sheet, tgt_path),
                    edge_type="reference", granularity="L2",
                    detail=ref["formula"]))
            else:
                model.add_edge(Edge(
                    source=f"sheet:{ws.title}", target=f"sheet:{tgt_sheet}",
                    edge_type="reference", granularity="L1",
                    detail=ref["formula"]))

    return model
```

- [ ] **Step 4: 実行して成功を確認**

Run: `pytest tests/test_workbook.py -v`
Expected: PASS（4 passed）

- [ ] **Step 5: 全テスト実行**

Run: `pytest -v`
Expected: 全 PASS

- [ ] **Step 6: Commit**

```bash
git add xlsx_flow/parser/workbook.py tests/test_workbook.py
git commit -m "feat: assemble full model from all parsers"
```

---

### Task 8: HTML レンダラ（`render/html.py` + テンプレート + Cytoscape vendoring）

**Files:**
- Create: `xlsx_flow/render/__init__.py`（空）
- Create: `xlsx_flow/render/html.py`
- Create: `xlsx_flow/render/template/cytoscape.min.js`（vendoring）
- Create: `xlsx_flow/render/template/app.js`
- Create: `xlsx_flow/render/template/style.css`
- Create: `xlsx_flow/render/template/index.html`
- Test: `tests/test_render.py`

**Interfaces:**
- Consumes: Task 1 `Model`。
- Produces:
  - `render_html(model: Model) -> str` — テンプレート（index.html）に `__MODEL_JSON__` / `__CYTOSCAPE_JS__` / `__APP_JS__` / `__STYLE_CSS__` を差し込んだ自己完結 HTML 文字列を返す。
  - 生成 HTML は: 粒度トグル（L1/L2/L3/PQ）ボタン、警告パネル、Cytoscape グラフ、クリック時の詳細サイドパネルを含む。

- [ ] **Step 1: Cytoscape.js を vendoring**

```bash
mkdir -p xlsx_flow/render/template
pip3 download cytoscape 2>/dev/null || true   # not a python pkg; use the JS UMD build below
```
ネットワークがある場合:
```bash
curl -fsSL https://unpkg.com/cytoscape@3.30.2/dist/cytoscape.min.js \
  -o xlsx_flow/render/template/cytoscape.min.js
```
ネットワークが無い／取得できない場合は、最小フォールバックとして次の内容を `cytoscape.min.js` に置く（描画ライブラリ不在でもページが壊れないためのスタブ。実機では必ず本物に差し替えること。差し替え必須を README に明記）:
```javascript
// FALLBACK STUB — replace with the real cytoscape UMD build before release.
window.cytoscape = window.cytoscape || function(opts){
  var el = opts && opts.container;
  if (el) { el.innerHTML = '<pre style="padding:1em">cytoscape.min.js is a stub. '
    + 'Replace xlsx_flow/render/template/cytoscape.min.js with the real build.</pre>'; }
  return { on:function(){}, layout:function(){return {run:function(){}}},
           nodes:function(){return [];}, elements:function(){return [];},
           add:function(){}, remove:function(){}, fit:function(){}, json:function(){return {};} };
};
```

- [ ] **Step 2: 失敗するテストを書く**

`tests/test_render.py`:
```python
from xlsx_flow.model import Model, Node, Edge
from xlsx_flow.render.html import render_html


def _model():
    m = Model(file="s.xlsx")
    m.add_node(Node(id="sheet:売上集計", type="sheet", attrs={"sheet_type": "formula"}))
    m.add_node(Node(id="query:Q_売上", type="query", attrs={"m_code": "let ... in ..."}))
    m.add_edge(Edge(source="query:Q_売上", target="sheet:売上集計",
                    edge_type="dataflow", granularity="pq"))
    m.warn("test warning")
    return m


def test_render_is_self_contained_html():
    html = render_html(_model())
    assert html.lstrip().startswith("<!DOCTYPE html>")
    # placeholders must be gone
    assert "__MODEL_JSON__" not in html
    assert "__CYTOSCAPE_JS__" not in html
    assert "__APP_JS__" not in html
    assert "__STYLE_CSS__" not in html


def test_render_embeds_model_and_warnings():
    html = render_html(_model())
    assert "sheet:売上集計" in html
    assert "Q_売上" in html
    assert "test warning" in html


def test_render_has_granularity_controls():
    html = render_html(_model())
    for level in ("L1", "L2", "L3", "PQ"):
        assert level in html
```

- [ ] **Step 3: 実行して失敗を確認**

Run: `pytest tests/test_render.py -v`
Expected: FAIL（`ModuleNotFoundError: No module named 'xlsx_flow.render.html'`）

- [ ] **Step 4: テンプレート CSS を書く**

`xlsx_flow/render/template/style.css`:
```css
* { box-sizing: border-box; }
body { margin: 0; font-family: "Segoe UI", "Hiragino Sans", sans-serif; }
#toolbar { padding: 8px 12px; background: #1f2937; color: #fff; display: flex;
  gap: 8px; align-items: center; }
#toolbar button { background: #374151; color: #fff; border: 1px solid #4b5563;
  padding: 6px 12px; border-radius: 4px; cursor: pointer; }
#toolbar button.active { background: #2563eb; border-color: #3b82f6; }
#warnings { background: #fef3c7; color: #92400e; padding: 6px 12px;
  font-size: 13px; border-bottom: 1px solid #f59e0b; }
#warnings:empty { display: none; }
#main { display: flex; height: calc(100vh - 84px); }
#cy { flex: 1; height: 100%; background: #f9fafb; }
#detail { width: 320px; border-left: 1px solid #e5e7eb; padding: 12px;
  overflow: auto; font-size: 13px; }
#detail h3 { margin-top: 0; }
#detail pre { white-space: pre-wrap; background: #f3f4f6; padding: 8px;
  border-radius: 4px; }
.legend span { display: inline-block; margin-right: 10px; font-size: 12px; }
.dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%;
  margin-right: 4px; vertical-align: middle; }
```

- [ ] **Step 5: テンプレート app.js を書く**

`xlsx_flow/render/template/app.js`:
```javascript
// MODEL is injected as a global by index.html before this script runs.
(function () {
  var TYPE_COLOR = {
    source: "#a855f7", query: "#0ea5e9", sheet: "#22c55e", column: "#94a3b8",
  };
  var SHEET_TYPE_COLOR = {
    raw: "#22c55e", pasted: "#eab308", formula: "#f97316",
    pq_output: "#0ea5e9", mixed: "#ef4444",
  };

  function nodeColor(n) {
    if (n.type === "sheet" && n.sheet_type) {
      return SHEET_TYPE_COLOR[n.sheet_type] || TYPE_COLOR.sheet;
    }
    return TYPE_COLOR[n.type] || "#9ca3af";
  }

  function label(n) {
    if (n.type === "column" && n.header_path) return n.header_path.join(" / ");
    return n.id.split(":").slice(1).join(":");
  }

  // Which node types / edge granularities are visible at each level.
  var LEVELS = {
    L1: { nodes: ["source", "query", "sheet"], edges: ["L1", "pq"] },
    L2: { nodes: ["source", "query", "sheet", "column"], edges: ["L1", "L2", "pq"] },
    L3: { nodes: ["source", "query", "sheet", "column"], edges: ["L1", "L2", "L3", "pq"] },
    PQ: { nodes: ["source", "query"], edges: ["pq"] },
  };

  function buildElements(level) {
    var spec = LEVELS[level];
    var nodeOk = {};
    var els = [];
    MODEL.nodes.forEach(function (n) {
      if (spec.nodes.indexOf(n.type) === -1) return;
      nodeOk[n.id] = true;
      els.push({ data: { id: n.id, label: label(n), color: nodeColor(n),
        raw: n }, classes: n.type });
    });
    MODEL.edges.forEach(function (e, i) {
      if (spec.edges.indexOf(e.granularity) === -1) return;
      if (!nodeOk[e.source] || !nodeOk[e.target]) return;
      els.push({ data: { id: "e" + i, source: e.source, target: e.target,
        etype: e.edge_type, detail: e.detail || "" } });
    });
    return els;
  }

  var cy = cytoscape({
    container: document.getElementById("cy"),
    style: [
      { selector: "node", style: {
        "background-color": "data(color)", "label": "data(label)",
        "font-size": "11px", "text-valign": "center", "color": "#111",
        "text-outline-color": "#fff", "text-outline-width": 2,
        "width": "label", "padding": "8px", "shape": "round-rectangle" } },
      { selector: "node.query", style: { "shape": "round-tag" } },
      { selector: "node.source", style: { "shape": "barrel" } },
      { selector: "edge", style: {
        "width": 1.5, "line-color": "#9ca3af", "target-arrow-color": "#9ca3af",
        "target-arrow-shape": "triangle", "curve-style": "bezier" } },
      { selector: 'edge[etype="reference"]', style: {
        "line-style": "dashed", "line-color": "#3b82f6",
        "target-arrow-color": "#3b82f6" } },
    ],
    elements: [],
  });

  function setLevel(level) {
    cy.elements().remove();
    cy.add(buildElements(level));
    cy.layout({ name: "breadthfirst", directed: true, padding: 20,
      spacingFactor: 1.1 }).run();
    cy.fit(undefined, 30);
    var btns = document.querySelectorAll("#toolbar button[data-level]");
    btns.forEach(function (b) {
      b.classList.toggle("active", b.getAttribute("data-level") === level);
    });
  }

  cy.on("tap", "node", function (evt) {
    var n = evt.target.data("raw");
    var d = document.getElementById("detail");
    var rows = Object.keys(n).filter(function (k) { return k !== "id"; })
      .map(function (k) {
        var v = n[k];
        if (typeof v === "object") v = JSON.stringify(v);
        return "<div><b>" + k + ":</b> " + String(v) + "</div>";
      }).join("");
    d.innerHTML = "<h3>" + label(n) + "</h3>" + rows;
  });

  document.querySelectorAll("#toolbar button[data-level]").forEach(function (b) {
    b.addEventListener("click", function () { setLevel(b.getAttribute("data-level")); });
  });

  setLevel("L1");
})();
```

- [ ] **Step 6: テンプレート index.html を書く**

`xlsx_flow/render/template/index.html`:
```html
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>xlsx-flow</title>
<style>__STYLE_CSS__</style>
</head>
<body>
<div id="toolbar">
  <strong>xlsx-flow</strong>
  <button data-level="L1">L1 フロー</button>
  <button data-level="L2">L2 列</button>
  <button data-level="L3">L3 セル</button>
  <button data-level="PQ">PQ のみ</button>
  <span class="legend" style="margin-left:auto">
    <span><span class="dot" style="background:#22c55e"></span>生/raw</span>
    <span><span class="dot" style="background:#f97316"></span>関数</span>
    <span><span class="dot" style="background:#0ea5e9"></span>PQ</span>
    <span><span class="dot" style="background:#a855f7"></span>外部源</span>
  </span>
</div>
<div id="warnings">__WARNINGS_HTML__</div>
<div id="main">
  <div id="cy"></div>
  <div id="detail"><h3>詳細</h3><div>ノードをクリックすると表示されます。</div></div>
</div>
<script>__CYTOSCAPE_JS__</script>
<script>var MODEL = __MODEL_JSON__;</script>
<script>__APP_JS__</script>
</body>
</html>
```

- [ ] **Step 7: レンダラ実装を書く**

`xlsx_flow/render/__init__.py`: （空ファイル）

`xlsx_flow/render/html.py`:
```python
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
    replacements = {
        "__STYLE_CSS__": _read("style.css"),
        "__CYTOSCAPE_JS__": _read("cytoscape.min.js"),
        "__APP_JS__": _read("app.js"),
        "__MODEL_JSON__": model_json,
        "__WARNINGS_HTML__": _warnings_html(model.warnings),
    }
    out = template
    for key, val in replacements.items():
        out = out.replace(key, val)
    return out
```

- [ ] **Step 8: 実行して成功を確認**

Run: `pytest tests/test_render.py -v`
Expected: PASS（3 passed）

- [ ] **Step 9: Commit**

```bash
git add xlsx_flow/render tests/test_render.py
git commit -m "feat: render model to self-contained interactive HTML"
```

---

### Task 9: CLI 配線とエンドツーエンド（`cli.py`）

**Files:**
- Create: `xlsx_flow/cli.py`
- Create: `README.md`
- Test: `tests/test_cli.py`

**Interfaces:**
- Consumes: Task 7 `analyze`、Task 8 `render_html`。
- Produces:
  - `main(argv: list[str] | None = None) -> int` — `xlsx-flow analyze <input.xlsx> -o out.html [--json out.json]`。HTML を書き出し、`--json` 指定時はモデル JSON も書く。戻り値は終了コード（0=成功）。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_cli.py`:
```python
from xlsx_flow.cli import main
from samples.gen_sample import generate


def test_cli_writes_html_and_json(tmp_path):
    src = tmp_path / "in.xlsx"
    generate(str(src))
    out_html = tmp_path / "out.html"
    out_json = tmp_path / "out.json"
    rc = main(["analyze", str(src), "-o", str(out_html), "--json", str(out_json)])
    assert rc == 0
    text = out_html.read_text(encoding="utf-8")
    assert text.lstrip().startswith("<!DOCTYPE html>")
    assert "売上集計" in text
    assert out_json.exists()
    assert "Q_売上" in out_json.read_text(encoding="utf-8")
```

- [ ] **Step 2: 実行して失敗を確認**

Run: `pytest tests/test_cli.py -v`
Expected: FAIL（`ModuleNotFoundError: No module named 'xlsx_flow.cli'`）

- [ ] **Step 3: 最小実装を書く**

`xlsx_flow/cli.py`:
```python
"""Command-line entry point for xlsx-flow."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from xlsx_flow.parser.workbook import analyze
from xlsx_flow.render.html import render_html


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="xlsx-flow")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("analyze", help="Analyze an .xlsx and emit interactive HTML")
    p.add_argument("input", help="Path to the input .xlsx")
    p.add_argument("-o", "--output", default="out.html", help="Output HTML path")
    p.add_argument("--json", dest="json_out", default=None,
                   help="Also write the intermediate model JSON to this path")

    args = parser.parse_args(argv)

    if args.cmd == "analyze":
        model = analyze(args.input)
        html = render_html(model)
        Path(args.output).write_text(html, encoding="utf-8")
        if args.json_out:
            Path(args.json_out).write_text(model.to_json(), encoding="utf-8")
        print(f"wrote {args.output}"
              + (f" and {args.json_out}" if args.json_out else ""))
        if model.warnings:
            print(f"  ({len(model.warnings)} warning(s) — see the HTML panel)")
        return 0
    return 1


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: 実行して成功を確認**

Run: `pytest tests/test_cli.py -v`
Expected: PASS（1 passed）

- [ ] **Step 5: README を書く**

`README.md`:
```markdown
# xlsx-flow

結合セル・複数シートまたぎ・PowerQuery を含む Excel ワークブックの
データフローを、自己完結のインタラクティブ HTML として可視化する CLI。

## インストール

    pip install -e ".[dev]"

## 使い方

    xlsx-flow analyze path/to/book.xlsx -o out.html --json out.json

`out.html` をブラウザで開くと、粒度トグル（L1 フロー / L2 列 / L3 セル / PQ のみ）、
ズーム・パン、ノードクリックで詳細パネルが使える。

## 検証用サンプルの生成

    python -m samples.gen_sample sample.xlsx
    xlsx-flow analyze sample.xlsx -o sample.html

## 粒度レベル

- **L1**: 外部ソース → クエリ → シートのデータフロー（シート種別を色分け）
- **L2**: シートを論理列（結合ヘッダ復元）まで展開、列レベルの参照
- **L3**: セル/範囲レベルの数式依存（詳細）
- **PQ のみ**: PowerQuery の M クエリ依存グラフ

## 注意

`xlsx_flow/render/template/cytoscape.min.js` はオフライン描画のため vendoring
している。スタブが入っている場合は実物の Cytoscape.js UMD ビルドに差し替えること。

## スコープ

本ツールはフェーズ1（可視化）。CSV / Python コードへの変換はフェーズ2で対象外。
```

- [ ] **Step 6: 全テスト実行 + 手動エンドツーエンド確認**

```bash
pytest -v
python -m samples.gen_sample /tmp/sample.xlsx
xlsx-flow analyze /tmp/sample.xlsx -o /tmp/out.html --json /tmp/out.json
```
Expected: 全テスト PASS、`/tmp/out.html` が生成され `<!DOCTYPE html>` で始まる。

- [ ] **Step 7: Commit**

```bash
git add xlsx_flow/cli.py README.md tests/test_cli.py
git commit -m "feat: wire CLI end-to-end and add README"
```

---

## 完了の定義

- `pytest -v` が全 PASS。
- `xlsx-flow analyze <xlsx> -o out.html` が自己完結 HTML を生成し、ブラウザで粒度トグル・ズーム・クリック詳細が動く。
- 合成サンプルで、結合ヘッダ復元・シート種別・シート間参照・PowerQuery 依存が図に出る。
- 抽出失敗は warnings として HTML に表示され、クラッシュしない。

## 既知の制約（フェーズ1で許容、将来拡張）

- PQ ロード先シートの推定はクエリ名＝シート名の一致に限定（`connections.xml` の完全解析は未実装）。
- L3（セル/範囲）は参照エッジの `detail`（数式文字列）表示まで。セル単位ノード展開はしない。
- DataMashup は package（`Section1.m`）のみ解析。Permissions/Metadata は読まない。
- `cytoscape.min.js` は実物の vendoring が前提（スタブ時は描画されない）。
- 罫線擬似ヘッダは `confidence="low"` のヒューリスティック。誤読の可能性を UI 上で明示。
