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


def test_render_escapes_script_close_in_model():
    from xlsx_flow.model import Model, Node
    m = Model(file="s.xlsx")
    m.add_node(Node(id="query:Q", type="query",
                    attrs={"m_code": "let x = </script><img src=x onerror=alert(1)> in x"}))
    html = render_html(m)
    # The raw, dangerous sequence must not survive into the output.
    assert "</script><img" not in html
    # The escaped form is present instead.
    assert "<\\/script><img" in html


def test_render_includes_chat_panel_and_grounding(tmp_path):
    from xlsx_flow.parser.workbook import analyze
    from xlsx_flow.render.html import render_html
    from samples.gen_sample import generate
    out = tmp_path / "s.xlsx"
    generate(str(out))
    html = render_html(analyze(str(out)))
    assert 'id="chat"' in html                     # チャットパネル
    assert 'id="chat-key"' in html                 # APIキー入力
    assert "anthropic-dangerous-direct-browser-access" in html  # 直呼出ヘッダ
    assert "claude-opus-4-8" in html               # 既定モデル
    assert '"notes"' in html                        # grounding に notes 同梱
