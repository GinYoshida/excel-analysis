// Pure helpers for the workbook chat agent. Node-requireable for tests;
// browser wiring is added in a later task guarded by `typeof document`.
(function (root, factory) {
  var api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  if (typeof window !== "undefined") window.XlsxChat = api;
})(this, function () {
  var DEFAULT_MODEL = "claude-opus-4-8";
  var MAXLEN = 400;   // m_code / expr / detail / formula
  var CELLLEN = 60;   // preview grid cell

  function clip(s, n) {
    s = String(s);
    return s.length > n ? s.slice(0, n - 1) + "…" : s;
  }

  function compactModel(model) {
    var nodes = (model.nodes || []).map(function (n) {
      var c = {};
      for (var k in n) {
        if (k === "preview" && n.preview && n.preview.cells) {
          c.preview = { dims: n.preview.dims, truncated: n.preview.truncated,
            cells: n.preview.cells.map(function (row) {
              return row.map(function (v) { return clip(v, CELLLEN); });
            }) };
        } else if (k === "m_code" || k === "expr" || k === "formula") {
          c[k] = clip(n[k], MAXLEN);
        } else {
          c[k] = n[k];
        }
      }
      return c;
    });
    var edges = (model.edges || []).map(function (e) {
      var c = {};
      for (var k in e) c[k] = (k === "detail") ? clip(e[k], MAXLEN) : e[k];
      return c;
    });
    return { meta: model.meta, nodes: nodes, edges: edges,
      warnings: model.warnings || [], notes: model.notes || [] };
  }

  var SYSTEM_PROMPT = [
    "あなたはExcelワークブックの解析アシスタントです。",
    "以下のJSONは、あるワークブックの構造・シート種別・列・数式のセル依存(L3)・",
    "PowerQuery(Mコード/ステップ)・セルの実値プレビュー・自然言語メモ(notes)を表します。",
    "このJSONだけを根拠に、日本語で簡潔に答えてください。JSONにない事実は推測せず",
    "「情報がありません」と述べること。用途概要・シート要約・数式やPQの平易な説明・",
    "notesの整形に対応します。",
  ].join("\n");

  function buildRequest(model, history, userText, opts) {
    opts = opts || {};
    var grounding = JSON.stringify(compactModel(model));
    return {
      model: opts.model || DEFAULT_MODEL,
      max_tokens: 4096,
      system: SYSTEM_PROMPT + "\n\n# WORKBOOK MODEL (JSON)\n" + grounding,
      messages: (history || []).concat([{ role: "user", content: userText }]),
    };
  }

  function extractText(resp) {
    if (!resp) return "(応答がありません)";
    if (resp.stop_reason === "refusal") {
      return "(安全上の理由で回答できませんでした)";
    }
    var parts = (resp.content || []).filter(function (b) { return b.type === "text"; })
      .map(function (b) { return b.text; });
    return parts.length ? parts.join("") : "(テキスト応答がありません)";
  }

  return { DEFAULT_MODEL: DEFAULT_MODEL, compactModel: compactModel,
    buildRequest: buildRequest, extractText: extractText,
    SYSTEM_PROMPT: SYSTEM_PROMPT };
});
