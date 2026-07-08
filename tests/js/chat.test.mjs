import assert from "node:assert";
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const chat = require("../../xlsx_flow/render/template/chat.js");

const MODEL = {
  meta: { file: "s.xlsx", sheet_count: 2 },
  nodes: [
    { id: "sheet:売上集計", type: "sheet", preview: { cells: [["x".repeat(200)]] } },
    { id: "query:Q", type: "query", m_code: "let\n" + "a".repeat(1000) + "\nin a" },
  ],
  edges: [{ source: "a", target: "b", detail: "d".repeat(1000) }],
  warnings: [],
  notes: [{ sheet: "手順", addr: "A2", text: "月次売上の作業表", kind: "text" }],
};

// compactModel は長文を切詰め、notes は保持する
const c = chat.compactModel(MODEL);
assert.ok(c.nodes[1].m_code.length <= 400, "m_code truncated");
assert.ok(c.nodes[0].preview.cells[0][0].length <= 60, "preview cell truncated");
assert.ok(c.edges[0].detail.length <= 400, "detail truncated");
assert.deepStrictEqual(c.notes, MODEL.notes, "notes preserved");

// buildRequest は既定モデルと system/messages を組み立てる
const req = chat.buildRequest(MODEL, [{ role: "user", content: "前" }], "このブックの概要は？");
assert.strictEqual(req.model, "claude-opus-4-8");
assert.ok(req.system.includes("売上集計"), "system carries grounding");
assert.strictEqual(req.messages.at(-1).content, "このブックの概要は？");
assert.strictEqual(req.messages.at(-1).role, "user");
assert.strictEqual(req.messages.length, 2, "history + new user turn");

// opts.model で上書きできる
assert.strictEqual(
  chat.buildRequest(MODEL, [], "q", { model: "claude-haiku-4-5" }).model,
  "claude-haiku-4-5",
);

// extractText は text ブロックを連結し、refusal を扱う
assert.strictEqual(
  chat.extractText({ stop_reason: "end_turn", content: [
    { type: "text", text: "答え1" }, { type: "text", text: "答え2" }] }),
  "答え1答え2",
);
assert.ok(chat.extractText({ stop_reason: "refusal", content: [] }).length > 0);

console.log("chat.test.mjs OK");
