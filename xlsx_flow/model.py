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
    notes: list[dict] = field(default_factory=list)

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
            "notes": list(self.notes),
        }

    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)
