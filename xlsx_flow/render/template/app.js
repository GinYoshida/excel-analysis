// MODEL is injected as a global by index.html before this script runs.
(function () {
  var TYPE_COLOR = {
    source: "#a855f7", query: "#0ea5e9", sheet: "#22c55e", column: "#94a3b8",
  };
  var SHEET_TYPE_COLOR = {
    raw: "#22c55e", pasted: "#eab308", formula: "#f97316",
    pq_output: "#0ea5e9", mixed: "#ef4444",
  };

  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

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
        return "<div><b>" + esc(k) + ":</b> " + esc(v) + "</div>";
      }).join("");
    d.innerHTML = "<h3>" + esc(label(n)) + "</h3>" + rows;
  });

  document.querySelectorAll("#toolbar button[data-level]").forEach(function (b) {
    b.addEventListener("click", function () { setLevel(b.getAttribute("data-level")); });
  });

  setLevel("L1");
})();
