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


def test_analyze_survives_sheet_parse_failure(tmp_path, monkeypatch):
    import xlsx_flow.parser.workbook as wbmod
    out = tmp_path / "s.xlsx"
    generate(str(out))

    def boom(ws, pq_targets):
        raise RuntimeError("boom")

    monkeypatch.setattr(wbmod, "classify_sheet", boom)
    model = wbmod.analyze(str(out))
    # No exception escaped; a Model came back with warnings recorded.
    assert any("解析に失敗" in w for w in model.warnings)


def test_analyze_survives_unreadable_file(tmp_path):
    from xlsx_flow.parser.workbook import analyze
    bad = tmp_path / "corrupt.xlsx"
    bad.write_bytes(b"this is not a zip / xlsx")
    model = analyze(str(bad))
    # No exception escaped; a Model came back with a warning recorded.
    assert isinstance(model.warnings, list)
    assert model.warnings


def test_analyze_maps_pq_entity_to_real_sheet(tmp_path):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    model = analyze(str(out))
    rng = next((n for n in model.nodes if n.id == "range:T_値貼付"), None)
    assert rng is not None
    assert rng.type == "range"
    assert rng.attrs["sheet"] == "値貼り付け"
    assert rng.attrs["ref"] == "A1:D3"
    df = {(e.source, e.target) for e in model.edges if e.edge_type == "dataflow"}
    # entity feeds the query, and the real sheet feeds the entity
    assert ("range:T_値貼付", "query:Q_テーブル取込") in df
    assert ("sheet:値貼り付け", "range:T_値貼付") in df


def test_analyze_creates_pq_step_nodes_and_edges(tmp_path):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    model = analyze(str(out))
    ids = {n.id for n in model.nodes}
    assert "step:Q_売上/Source" in ids
    assert "step:Q_売上/Added" in ids
    added = next(n for n in model.nodes if n.id == "step:Q_売上/Added")
    assert added.type == "step"
    assert added.attrs["query"] == "Q_売上"
    df = {(e.source, e.target) for e in model.edges if e.edge_type == "dataflow"}
    assert ("step:Q_売上/Source", "step:Q_売上/Added") in df


def test_analyze_produces_l3_cell_edges(tmp_path):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    model = analyze(str(out))
    l3 = {(e.source, e.target) for e in model.edges if e.granularity == "L3"}
    # same-sheet: B3 -> D3 ; cross-sheet: 生データ!C3 -> 売上集計!B4
    assert ("cell:売上集計!B3", "cell:売上集計!D3") in l3
    assert ("cell:生データ!C3", "cell:売上集計!B4") in l3
    cell_nodes = {n.id for n in model.nodes if n.type == "cell"}
    assert "cell:売上集計!D3" in cell_nodes
