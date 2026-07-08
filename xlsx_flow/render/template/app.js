// MODEL is injected as a global by index.html before this script runs.
(function () {
  var TYPE_COLOR = {
    source: "#a855f7", query: "#0ea5e9", sheet: "#22c55e", column: "#94a3b8",
    range: "#14b8a6", step: "#7dd3fc", cell: "#cbd5e1",
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
    if (n.type === "cell" && n.is_formula) return "#fcd34d";  // formula cells stand out
    return TYPE_COLOR[n.type] || "#9ca3af";
  }

  function label(n) {
    if (n.type === "column" && n.header_path) return n.header_path.join(" / ");
    if (n.type === "step" && n.name) return n.name;
    if (n.type === "cell" && n.addr) return n.addr;
    return n.id.split(":").slice(1).join(":");
  }

  // Which node types / edge granularities are visible at each level.
  var LEVELS = {
    L1: { nodes: ["source", "query", "sheet", "range"], edges: ["L1", "pq"] },
    L2: { nodes: ["source", "query", "sheet", "column", "range"], edges: ["L1", "L2", "pq"] },
    L3: { nodes: ["source", "query", "sheet", "range", "cell"], edges: ["L3", "pq"] },
    PQ: { nodes: ["source", "query", "range", "step"], edges: ["pq", "step"] },
  };

  function buildElements(level) {
    var spec = LEVELS[level];
    var nodeOk = {};
    var els = [];
    // First pass: decide which nodes are visible so parent links can be checked.
    MODEL.nodes.forEach(function (n) {
      if (spec.nodes.indexOf(n.type) !== -1) nodeOk[n.id] = true;
    });
    // Second pass: build node elements; columns nest under their sheet (compound).
    MODEL.nodes.forEach(function (n) {
      if (!nodeOk[n.id]) return;
      var data = { id: n.id, label: label(n), color: nodeColor(n), raw: n };
      if (n.type === "column" && n.sheet) {
        var parentId = "sheet:" + n.sheet;
        if (nodeOk[parentId]) data.parent = parentId;
      }
      if (n.type === "step" && n.query) {
        var qid = "query:" + n.query;
        if (nodeOk[qid]) data.parent = qid;
      }
      if (n.type === "cell" && n.sheet) {
        var csid = "sheet:" + n.sheet;
        if (nodeOk[csid]) data.parent = csid;
      }
      els.push({ data: data, classes: n.type });
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
        "font-size": "12px", "text-valign": "center", "color": "#111",
        "text-outline-color": "#fff", "text-outline-width": 2,
        "width": "label", "height": "label", "padding": "10px",
        "shape": "round-rectangle" } },
      { selector: "node.query", style: { "shape": "round-tag" } },
      { selector: "node.source", style: { "shape": "barrel" } },
      { selector: "node.range", style: { "shape": "cut-rectangle" } },
      { selector: "node.step", style: { "shape": "ellipse" } },
      { selector: "node.cell", style: { "shape": "round-rectangle", "font-size": "11px" } },
      // Sheet acts as a container box when it holds columns (L2/L3).
      { selector: ":parent", style: {
        "background-color": "data(color)", "background-opacity": 0.12,
        "border-color": "data(color)", "border-width": 2,
        "label": "data(label)", "text-valign": "top", "text-halign": "center",
        "font-size": "13px", "font-weight": "bold", "color": "#111",
        "text-outline-color": "#fff", "text-outline-width": 2,
        "padding": "16px", "shape": "round-rectangle" } },
      { selector: "node:selected", style: {
        "border-width": 3, "border-color": "#2563eb" } },
      { selector: "edge", style: {
        "width": 1.5, "line-color": "#9ca3af", "target-arrow-color": "#9ca3af",
        "target-arrow-shape": "triangle", "curve-style": "bezier" } },
      { selector: 'edge[etype="reference"]', style: {
        "line-style": "dashed", "line-color": "#3b82f6",
        "target-arrow-color": "#3b82f6" } },
      { selector: "edge:selected", style: {
        "width": 3, "line-color": "#2563eb", "target-arrow-color": "#2563eb" } },
    ],
    elements: [],
  });

  function runLayout() {
    // cose respects compound (sheet) containers; breadthfirst gives a clean
    // directed flow when there are no nested columns (L1 / PQ).
    var hasCompound = cy.nodes().some(function (n) { return n.isParent(); });
    var opts = hasCompound
      ? { name: "cose", padding: 24, animate: false, nodeDimensionsIncludeLabels: true,
          idealEdgeLength: 70, nodeRepulsion: 9000, nestingFactor: 1.1 }
      : { name: "breadthfirst", directed: true, padding: 24, spacingFactor: 1.15 };
    cy.layout(opts).run();
  }

  function setLevel(level) {
    cy.elements().remove();
    cy.add(buildElements(level));
    runLayout();
    cy.resize();
    cy.fit(undefined, 30);
    var btns = document.querySelectorAll("#toolbar button[data-level]");
    btns.forEach(function (b) {
      b.classList.toggle("active", b.getAttribute("data-level") === level);
    });
  }

  function showDetail(html) {
    var d = document.getElementById("detail");
    d.innerHTML = html;
    // On stacked (mobile) layout the panel sits below the graph; make sure a
    // fresh selection is brought into view.
    d.scrollTop = 0;
    d.scrollIntoView({ block: "nearest" });
  }

  // Render a sheet preview grid (values + formulas) as a compact HTML table.
  function sheetPreviewHtml(p) {
    if (!p || !p.cells || !p.cells.length) return "";
    var body = p.cells.map(function (row) {
      return "<tr>" + row.map(function (c) {
        var isF = typeof c === "string" && c.charAt(0) === "=";
        return "<td" + (isF ? ' class="f"' : "") + ">" + esc(c) + "</td>";
      }).join("") + "</tr>";
    }).join("");
    var note = "<div class=\"muted\">" + p.dims.rows + "行 × " + p.dims.cols + "列"
      + (p.truncated ? "（先頭のみ表示）" : "") + "</div>";
    return note + "<table class=\"preview\">" + body + "</table>";
  }

  // Render a column preview: range, sample values, and a formula example.
  function columnPreviewHtml(p) {
    if (!p) return "";
    var out = "<div class=\"muted\">範囲 " + esc(p.range) + "</div>";
    if (p.samples && p.samples.length) {
      out += "<table class=\"preview\">" + p.samples.map(function (v) {
        var isF = typeof v === "string" && v.charAt(0) === "=";
        return "<tr><td" + (isF ? ' class="f"' : "") + ">" + esc(v) + "</td></tr>";
      }).join("") + "</table>";
    }
    if (p.formula_example) {
      out += "<div class=\"muted\">数式例</div><pre>" + esc(p.formula_example) + "</pre>";
    }
    return out;
  }

  cy.on("tap", "node", function (evt) {
    var n = evt.target.data("raw");
    window.__selectedNodeId = evt.target.id();
    if (!n) return;  // compound parents created implicitly still carry raw
    var previewHtml = "";
    if (n.preview) {
      previewHtml = n.type === "sheet"
        ? sheetPreviewHtml(n.preview)
        : columnPreviewHtml(n.preview);
    } else if (n.type === "step" && n.expr) {
      previewHtml = "<div class=\"muted\">Mステップ式</div><pre>" + esc(n.expr) + "</pre>";
    } else if (n.type === "cell" && n.formula) {
      previewHtml = "<div class=\"muted\">数式</div><pre>" + esc(n.formula) + "</pre>";
    }
    var rows = Object.keys(n).filter(function (k) {
      return k !== "id" && k !== "preview" && k !== "expr" && k !== "formula";
    }).map(function (k) {
        var v = n[k];
        if (typeof v === "object") v = JSON.stringify(v);
        return "<div><b>" + esc(k) + ":</b> " + esc(v) + "</div>";
      }).join("");
    showDetail("<h3>" + esc(label(n)) + "</h3>" + rows + previewHtml);
  });

  cy.on("tap", "edge", function (evt) {
    var e = evt.target.data();
    showDetail("<h3>" + esc(e.etype) + "</h3>"
      + "<div><b>" + esc(e.source) + " → " + esc(e.target) + "</b></div>"
      + (e.detail ? "<pre>" + esc(e.detail) + "</pre>" : "<div>(詳細なし)</div>"));
  });

  document.querySelectorAll("#toolbar button[data-level]").forEach(function (b) {
    b.addEventListener("click", function () { setLevel(b.getAttribute("data-level")); });
  });

  // Re-fit when the container settles or the device rotates (mobile webviews
  // sometimes report a zero/late size at first paint).
  window.addEventListener("resize", function () {
    cy.resize();
    cy.fit(undefined, 30);
  });

  setLevel("L1");
})();
