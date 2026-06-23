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
