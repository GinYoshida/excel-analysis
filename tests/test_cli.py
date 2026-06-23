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
