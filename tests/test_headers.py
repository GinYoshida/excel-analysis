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
