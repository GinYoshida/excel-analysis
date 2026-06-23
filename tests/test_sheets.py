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
