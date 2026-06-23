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
