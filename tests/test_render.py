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
