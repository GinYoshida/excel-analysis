# ワークブック対話エージェント Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 生成HTMLに、埋め込み済み `MODEL` を文脈にブラウザから直接 Anthropic API を呼ぶ、任意の対話エージェント（全体概要 / シート要約 / PQ・数式説明 / メモ整形）を追加する。

**Architecture:** パーサに自然言語ノート抽出を足して `Model.notes` に集約し、`model.to_dict()` 経由で既存の `MODEL` JSON に同梱する。UIは純関数（`compactModel`/`buildRequest`/`extractText`）を `chat.js` に切り出して node で単体テストし、DOM/ネットワーク配線はブラウザ実行時のみ動く。既存の `analyze` 静的HTMLと可視化は不変で、キー未入力時はチャットが不活性。

**Tech Stack:** Python 3 / openpyxl / pytest（既存）、バニラJS（テンプレート `app.js`/新規 `chat.js`）、Anthropic Messages API（ブラウザ `fetch` 直呼出）。

## Global Constraints

- 出力HTMLは自己完結を維持する。外部通信は**ユーザーが起こすAnthropic fetch のみ**。CDN/外部アセット追加禁止。
- APIキーは localStorage キー `xlsxflow_anthropic_key` のみに保存。**生成HTMLへ書き込まない**。
- 既定モデルは `claude-opus-4-8`（任意で `claude-haiku-4-5` 切替）。ハードコードのモデルID文字列に日付サフィックスを付けない。
- API呼出ヘッダ: `x-api-key`, `anthropic-version: 2023-06-01`, `anthropic-dangerous-direct-browser-access: true`, `content-type: application/json`。
- 送信は既定オフ（オプトイン）。キー未入力なら一切送信しない。
- ノート抽出のしきい値: 長文セル `>=20` 文字。ワークブック全体の上限 `NOTES_BUDGET = 60` 件、超過は `model.warn`。
- 既存テスト（49件）を壊さない。UI文言は日本語。

---

### Task 1: 自然言語ノート抽出 → `Model.notes`

**Files:**
- Create: `xlsx_flow/parser/notes.py`
- Modify: `xlsx_flow/model.py`（`Model` に `notes` フィールド、`to_dict()` に同梱）
- Modify: `xlsx_flow/parser/workbook.py`（各シートで `extract_notes` を呼び集約・上限）
- Modify: `samples/gen_sample.py`（メモセルとセルコメントを追加）
- Test: `tests/test_notes.py`

**Interfaces:**
- Consumes: `openpyxl` の worksheet、既存 `Model`（`xlsx_flow/model.py`）、`analyze`（`xlsx_flow/parser/workbook.py`）。
- Produces:
  - `extract_notes(ws, min_len: int = 20) -> list[dict]` — 各要素 `{"sheet": str, "addr": str, "text": str, "kind": "text"|"comment"}`。
  - `Model.notes: list[dict]`（`to_dict()` のトップレベル `"notes"` に出力）。
  - `xlsx_flow.parser.notes.NOTES_BUDGET = 60`。

- [ ] **Step 1: `Model.notes` の失敗テストを書く**

`tests/test_notes.py` を新規作成:

```python
from xlsx_flow.model import Model, Node


def test_model_to_dict_includes_notes():
    m = Model(file="x.xlsx")
    m.notes.append({"sheet": "S", "addr": "A1", "text": "説明メモ", "kind": "text"})
    d = m.to_dict()
    assert d["notes"] == [{"sheet": "S", "addr": "A1", "text": "説明メモ", "kind": "text"}]


def test_model_notes_defaults_empty():
    assert Model(file="x.xlsx").to_dict()["notes"] == []
```

- [ ] **Step 2: 実行して失敗を確認**

Run: `python -m pytest tests/test_notes.py -q`
Expected: FAIL（`Model` に `notes` が無く `KeyError`/`AttributeError`）

- [ ] **Step 3: `Model` に `notes` を追加**

`xlsx_flow/model.py` の `Model` データクラスを修正:

```python
@dataclass
class Model:
    file: str
    nodes: list[Node] = field(default_factory=list)
    edges: list[Edge] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    notes: list[dict] = field(default_factory=list)
```

`to_dict()` の戻り dict に `notes` を追加:

```python
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
```

- [ ] **Step 4: 実行してパスを確認**

Run: `python -m pytest tests/test_notes.py -q`
Expected: PASS（2件）

- [ ] **Step 5: `extract_notes` の失敗テストを追加**

`tests/test_notes.py` に追記:

```python
import openpyxl
from openpyxl.comments import Comment

from xlsx_flow.parser.notes import extract_notes


def test_extract_notes_picks_long_text_cell():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "手順"
    ws["A1"] = "短い"  # 20文字未満: 拾わない
    ws["A2"] = "このシートは毎月の売上を店舗別に集計するための作業表です"  # 長文: 拾う
    ws["A3"] = 12345  # 数値: 拾わない
    notes = extract_notes(ws, min_len=20)
    picked = {(n["addr"], n["kind"]) for n in notes}
    assert ("A2", "text") in picked
    assert all(n["addr"] != "A1" for n in notes)
    assert all(n["addr"] != "A3" for n in notes)
    assert all(n["sheet"] == "手順" for n in notes)


def test_extract_notes_picks_cell_comment():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "集計"
    ws["B2"] = 10
    ws["B2"].comment = Comment("前月比の計算に注意すること", "作成者")
    notes = extract_notes(ws, min_len=20)
    comments = [n for n in notes if n["kind"] == "comment"]
    assert comments and comments[0]["addr"] == "B2"
    assert "前月比" in comments[0]["text"]
```

- [ ] **Step 6: 実行して失敗を確認**

Run: `python -m pytest tests/test_notes.py -q`
Expected: FAIL（`xlsx_flow.parser.notes` が無い）

- [ ] **Step 7: `notes.py` を実装**

`xlsx_flow/parser/notes.py` を新規作成:

```python
"""Collect free-text notes (long text cells + cell comments) for the AI agent."""
from __future__ import annotations

NOTES_BUDGET = 60


def extract_notes(ws, min_len: int = 20) -> list[dict]:
    """Return natural-language memos on this sheet: long non-formula text cells
    and cell comments. Numbers, formulas, and short labels are skipped."""
    out: list[dict] = []
    for row in ws.iter_rows():
        for cell in row:
            comment = getattr(cell, "comment", None)
            if comment is not None and comment.text:
                out.append({"sheet": ws.title, "addr": cell.coordinate,
                            "text": str(comment.text).strip(), "kind": "comment"})
            if cell.data_type == "s" and isinstance(cell.value, str):
                text = cell.value.strip()
                if len(text) >= min_len:
                    out.append({"sheet": ws.title, "addr": cell.coordinate,
                                "text": text, "kind": "text"})
    return out
```

- [ ] **Step 8: 実行してパスを確認**

Run: `python -m pytest tests/test_notes.py -q`
Expected: PASS（4件）

- [ ] **Step 9: サンプルにメモとコメントを追加**

`samples/gen_sample.py` の `_build_workbook()` の「Sheet 1: 生データ」ブロック直後（`ws.append([...])` ループの後）に追記:

```python
    # 自然言語メモ（AIエージェント検証用）: 長文セルとセルコメント
    from openpyxl.comments import Comment
    ws["E1"] = "このシートは各店舗の月次売上を手入力で記録する元データ表です"
    ws["C1"].comment = Comment("売上は税抜・単位は千円で入力すること", "運用担当")
```

- [ ] **Step 10: `analyze` が notes を集約する失敗テストを追加**

`tests/test_notes.py` に追記:

```python
from xlsx_flow.parser.workbook import analyze
from samples.gen_sample import generate


def test_analyze_collects_notes(tmp_path):
    out = tmp_path / "s.xlsx"
    generate(str(out))
    model = analyze(str(out))
    texts = " ".join(n["text"] for n in model.notes)
    assert "月次売上" in texts        # 長文セル E1
    assert "税抜" in texts            # セルコメント C1
    assert all({"sheet", "addr", "text", "kind"} <= set(n) for n in model.notes)


def test_analyze_notes_respects_budget(tmp_path, monkeypatch):
    import xlsx_flow.parser.workbook as wbmod
    out = tmp_path / "s.xlsx"
    generate(str(out))
    monkeypatch.setattr(wbmod, "NOTES_BUDGET", 1)
    model = wbmod.analyze(str(out))
    assert len(model.notes) <= 1
    assert any("メモ" in w or "notes" in w.lower() for w in model.warnings)
```

- [ ] **Step 11: 実行して失敗を確認**

Run: `python -m pytest tests/test_notes.py -q`
Expected: FAIL（`analyze` が notes を埋めない / `NOTES_BUDGET` 未 import）

- [ ] **Step 12: `workbook.analyze` に配線**

`xlsx_flow/parser/workbook.py` の import 群に追記:

```python
from xlsx_flow.parser.notes import extract_notes, NOTES_BUDGET
```

`analyze` 内、シートループ（`for ws in wb.worksheets:` でノード生成している箇所）の中の
`try:` ブロック末尾、`recover_headers` 由来のノード追加が終わった後に追記:

```python
            for note in extract_notes(ws):
                if len(model.notes) >= NOTES_BUDGET:
                    break
                model.notes.append(note)
```

同 `analyze` の `return model` の直前に上限警告を追記:

```python
    if len(model.notes) >= NOTES_BUDGET:
        model.warn(f"自然言語メモは上限({NOTES_BUDGET})に達したため一部のみ収集しました")
```

- [ ] **Step 13: 実行してパスを確認**

Run: `python -m pytest tests/test_notes.py -q`
Expected: PASS（6件）

- [ ] **Step 14: 全テストで回帰確認**

Run: `python -m pytest -q`
Expected: PASS（既存49 + 新規6 = 55件）

- [ ] **Step 15: コミット**

```bash
git add xlsx_flow/model.py xlsx_flow/parser/notes.py xlsx_flow/parser/workbook.py samples/gen_sample.py tests/test_notes.py
git commit -m "⑤a ノート抽出: 長文セル/コメントをModel.notesに集約"
```

---

### Task 2: グラウンディング純関数 `chat.js`（compactModel / buildRequest / extractText）

**Files:**
- Create: `xlsx_flow/render/template/chat.js`
- Test: `tests/js/chat.test.mjs`

**Interfaces:**
- Consumes: `MODEL` 形状（`model.to_dict()` の出力: `meta/nodes/edges/warnings/notes`）。
- Produces（`chat.js` が node では `module.exports`、ブラウザでは `window.XlsxChat` に公開）:
  - `compactModel(model) -> object` — `m_code`/`expr`/`detail`/`formula` を 400 字、`preview.cells` の各文字列を 60 字で切詰めた縮約コピー。`notes` はそのまま。
  - `buildRequest(model, history, userText, opts?) -> {model, max_tokens, system, messages}` — `opts.model`（既定 `"claude-opus-4-8"`）。`system` は指示文＋`compactModel` の JSON。`messages` は `history` に `{role:"user", content:userText}` を追加。
  - `extractText(apiResponse) -> string` — `content` 内 `type==="text"` を連結。`stop_reason==="refusal"` は定型メッセージ。
  - 定数 `DEFAULT_MODEL = "claude-opus-4-8"`。

- [ ] **Step 1: JSテスト用ディレクトリと失敗テストを書く**

`tests/js/chat.test.mjs` を新規作成:

```javascript
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
```

- [ ] **Step 2: 実行して失敗を確認**

Run: `node tests/js/chat.test.mjs`
Expected: FAIL（`chat.js` が無く `Cannot find module`）

- [ ] **Step 3: `chat.js` の純関数部を実装**

`xlsx_flow/render/template/chat.js` を新規作成（この Step ではブラウザ配線は入れない）:

```javascript
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
    return s.length > n ? s.slice(0, n) + "…" : s;
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
```

- [ ] **Step 4: 実行してパスを確認**

Run: `node tests/js/chat.test.mjs`
Expected: 出力 `chat.test.mjs OK`（例外なし）

- [ ] **Step 5: 構文チェック**

Run: `node --check xlsx_flow/render/template/chat.js`
Expected: エラー無し（何も出力されない）

- [ ] **Step 6: コミット**

```bash
git add xlsx_flow/render/template/chat.js tests/js/chat.test.mjs
git commit -m "⑤b グラウンディング純関数: compactModel/buildRequest/extractText"
```

---

### Task 3: チャットUIとレンダ統合（fetch配線・キー保存・サジェスト）

**Files:**
- Modify: `xlsx_flow/render/html.py`（`__CHAT_JS__` 置換を追加）
- Modify: `xlsx_flow/render/template/index.html`（チャットパネルの markup と `chat.js` の読込）
- Modify: `xlsx_flow/render/template/style.css`（チャットのスタイル）
- Modify: `xlsx_flow/render/template/chat.js`（`typeof document` ガード下のDOM/fetch配線を追記）
- Test: `tests/test_render.py`（レンダにチャット要素と grounding が含まれること）

**Interfaces:**
- Consumes: Task 2 の `window.XlsxChat`（`buildRequest`/`extractText`）、埋め込み `window.MODEL`、`app.js` が選択中ノードを公開する `window.__selectedNodeId`（本タスクで `app.js` に追加）。
- Produces: 生成HTMLに `id="chat"` パネル、キー入力 `id="chat-key"`、送信 `id="chat-send"`、サジェストボタン群 `.chat-suggest[data-q]`。

- [ ] **Step 1: レンダ統合の失敗テストを書く**

`tests/test_render.py` を確認し、末尾に追記:

```python
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
```

- [ ] **Step 2: 実行して失敗を確認**

Run: `python -m pytest tests/test_render.py -q`
Expected: FAIL（チャット markup / chat.js 未組込）

- [ ] **Step 3: `html.py` に `__CHAT_JS__` を追加**

`xlsx_flow/render/html.py` の `replacements` dict に1行追加:

```python
    replacements = {
        "__STYLE_CSS__": _read("style.css"),
        "__CYTOSCAPE_JS__": _read("cytoscape.min.js"),
        "__APP_JS__": _read("app.js"),
        "__CHAT_JS__": _read("chat.js"),
        "__MODEL_JSON__": model_json,
        "__WARNINGS_HTML__": _warnings_html(model.warnings),
    }
```

- [ ] **Step 4: `index.html` にチャット markup と読込を追加**

`xlsx_flow/render/template/index.html` の `#detail` div の直後（`</div>` で `#main` を閉じる直前）にチャットパネルを追加し、`app.js` の後に `chat.js` を読込む。`#main` ブロックを次に置換:

```html
<div id="main">
  <div id="cy"></div>
  <div id="side">
    <div id="detail"><h3>詳細</h3><div>ノードをクリックすると表示されます。</div></div>
    <div id="chat">
      <div id="chat-head">AIに質問 <span class="chat-note">※内容がAnthropic APIへ送信されます</span></div>
      <div id="chat-key-row">
        <input id="chat-key" type="password" placeholder="Anthropic APIキー (sk-ant-...)">
        <button id="chat-key-save">保存</button>
        <button id="chat-key-clear" title="キー削除">×</button>
      </div>
      <div id="chat-log"></div>
      <div id="chat-suggest">
        <button class="chat-suggest" data-q="このワークブック全体の用途とデータフローの概要を教えて">全体概要</button>
        <button class="chat-suggest" data-q="選択中（または主要な）シートは何をするシートか説明して">このシートは？</button>
        <button class="chat-suggest" data-q="PowerQueryと主要な数式が何をしているか平易に説明して">PQ・数式の説明</button>
        <button class="chat-suggest" data-q="notesに含まれる自然言語メモを整形して要点をまとめて">メモを整形</button>
      </div>
      <div id="chat-input-row">
        <input id="chat-input" type="text" placeholder="質問を入力…">
        <button id="chat-send">送信</button>
      </div>
    </div>
  </div>
</div>
```

同ファイル末尾のスクリプト群を、`app.js` の後に `chat.js` を足す形へ置換:

```html
<script>__CYTOSCAPE_JS__</script>
<script>var MODEL = __MODEL_JSON__;</script>
<script>__APP_JS__</script>
<script>__CHAT_JS__</script>
```

- [ ] **Step 5: 実行して render テストのパスを確認**

Run: `python -m pytest tests/test_render.py -q`
Expected: PASS（`anthropic-dangerous-direct-browser-access` は Step 7 で chat.js に入るため、この時点では失敗する場合は Step 7 まで進めてから再実行）

> 注: `anthropic-dangerous-direct-browser-access` と `claude-opus-4-8` の文字列は Step 7 の `chat.js` 配線に含まれる。Step 4→Step 7 を続けて実装し、Step 8 でまとめて確認する。

- [ ] **Step 6: `app.js` に選択ノードIDの公開を追加**

`xlsx_flow/render/template/app.js` の `cy.on("tap", "node", function (evt) {` ハンドラ冒頭（`var n = evt.target.data("raw");` の直後）に追記:

```javascript
    window.__selectedNodeId = evt.target.id();
```

- [ ] **Step 7: `chat.js` にDOM/fetch配線を追記**

`xlsx_flow/render/template/chat.js` の `return { ... };` の**前**（factory 関数内）に、ブラウザ限定の初期化を追加:

```javascript
  var KEY_STORAGE = "xlsxflow_anthropic_key";

  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function initChat() {
    if (typeof document === "undefined") return;
    var el = document.getElementById("chat");
    if (!el) return;
    var keyInput = document.getElementById("chat-key");
    var log = document.getElementById("chat-log");
    var input = document.getElementById("chat-input");
    var history = [];

    var saved = null;
    try { saved = window.localStorage.getItem(KEY_STORAGE); } catch (e) {}
    if (saved) keyInput.value = saved;

    function append(role, text) {
      var d = document.createElement("div");
      d.className = "chat-msg chat-" + role;
      d.innerHTML = "<b>" + (role === "user" ? "あなた" : role === "error" ? "エラー" : "AI")
        + ":</b> " + esc(text);
      log.appendChild(d);
      log.scrollTop = log.scrollHeight;
      return d;
    }

    function saveKey() {
      try { window.localStorage.setItem(KEY_STORAGE, keyInput.value.trim()); } catch (e) {}
    }
    document.getElementById("chat-key-save").addEventListener("click", saveKey);
    document.getElementById("chat-key-clear").addEventListener("click", function () {
      keyInput.value = "";
      try { window.localStorage.removeItem(KEY_STORAGE); } catch (e) {}
    });

    function send(text) {
      var key = keyInput.value.trim();
      if (!key) { append("error", "APIキーを入力してください。"); return; }
      if (!text) return;
      // 選択中ノードがあれば文脈を付与
      var sel = window.__selectedNodeId;
      var userText = sel ? (text + "\n\n（選択中の要素: " + sel + "）") : text;
      append("user", text);
      var pending = append("assistant", "…");
      var req = buildRequest(window.MODEL, history, userText);
      fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-api-key": key,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify(req),
      }).then(function (r) {
        return r.json().then(function (j) { return { ok: r.ok, status: r.status, j: j }; });
      }).then(function (res) {
        if (!res.ok) {
          var msg = (res.j && res.j.error && res.j.error.message) || ("HTTP " + res.status);
          pending.remove(); append("error", msg); return;
        }
        var answer = extractText(res.j);
        pending.remove();
        append("assistant", answer);
        history.push({ role: "user", content: userText });
        history.push({ role: "assistant", content: answer });
      }).catch(function (err) {
        pending.remove(); append("error", String(err));
      });
    }

    document.getElementById("chat-send").addEventListener("click", function () {
      var t = input.value.trim(); input.value = ""; send(t);
    });
    input.addEventListener("keydown", function (e) {
      if (e.key === "Enter") { var t = input.value.trim(); input.value = ""; send(t); }
    });
    var sugg = document.querySelectorAll(".chat-suggest");
    for (var i = 0; i < sugg.length; i++) {
      sugg[i].addEventListener("click", function () { send(this.getAttribute("data-q")); });
    }
  }

  if (typeof document !== "undefined") {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initChat);
    } else {
      initChat();
    }
  }
```

- [ ] **Step 8: 純関数テストとレンダテストを再実行**

Run: `node tests/js/chat.test.mjs && node --check xlsx_flow/render/template/chat.js && python -m pytest tests/test_render.py -q`
Expected: `chat.test.mjs OK` ＋ 構文OK ＋ render テスト PASS

- [ ] **Step 9: チャットのスタイルを追加**

`xlsx_flow/render/template/style.css` の末尾に追記（既存 `#detail` は `#side` 内に入るため幅指定を調整）:

```css
#side { width: 340px; border-left: 1px solid #e5e7eb; display: flex;
  flex-direction: column; min-height: 0; }
#side #detail { width: auto; border-left: none; flex: 1 1 auto; }
#chat { border-top: 2px solid #e5e7eb; padding: 8px; display: flex;
  flex-direction: column; gap: 6px; max-height: 46%; }
#chat-head { font-weight: bold; font-size: 13px; }
.chat-note { font-weight: normal; color: #b45309; font-size: 11px; }
#chat-key-row, #chat-input-row { display: flex; gap: 4px; }
#chat-key, #chat-input { flex: 1; min-width: 0; padding: 4px 6px;
  border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px; }
#chat button { background: #2563eb; color: #fff; border: none; border-radius: 4px;
  padding: 4px 8px; cursor: pointer; font-size: 12px; }
#chat-key-clear { background: #6b7280; }
#chat-suggest { display: flex; flex-wrap: wrap; gap: 4px; }
#chat-suggest .chat-suggest { background: #eef2ff; color: #3730a3;
  border: 1px solid #c7d2fe; }
#chat-log { flex: 1 1 auto; overflow: auto; font-size: 12px; min-height: 60px; }
.chat-msg { margin: 4px 0; white-space: pre-wrap; word-break: break-word; }
.chat-error { color: #b91c1c; }
@media (max-width: 700px) {
  #side { width: 100%; }
  #chat { max-height: 40vh; }
}
```

既存の `#detail` 幅指定（`#detail { width: 320px; ... }`）は `#side` に内包されるため、`width: 320px;` の行は残置で問題ない（`#side #detail { width: auto; }` が上書きする）。

- [ ] **Step 10: 全テストで回帰確認**

Run: `python -m pytest -q && node tests/js/chat.test.mjs`
Expected: pytest PASS（56件: 55 + render1）＋ `chat.test.mjs OK`

- [ ] **Step 11: サンプル生成と目視用HTMLの確認**

Run:
```bash
python -m samples.gen_sample /tmp/s.xlsx >/dev/null
python -m xlsx_flow.cli analyze /tmp/s.xlsx -o /tmp/s.html
node --check xlsx_flow/render/template/app.js
python -c "h=open('/tmp/s.html').read(); print('chat', 'id=\"chat\"' in h); print('notes', '\"notes\"' in h); print('nokey', 'sk-ant-' not in h)"
```
Expected: `chat True` / `notes True` / `nokey True`（キーはHTMLに埋め込まれない）。構文エラー無し。

- [ ] **Step 12: コミット**

```bash
git add xlsx_flow/render/html.py xlsx_flow/render/template/index.html xlsx_flow/render/template/style.css xlsx_flow/render/template/chat.js xlsx_flow/render/template/app.js tests/test_render.py
git commit -m "⑤c チャットUI: ブラウザ直呼出・キー保存・サジェスト・メモ整形"
```

---

### Task 4: ドキュメント更新

**Files:**
- Modify: `README.md`（AIチャットの使い方・キーの扱い・プライバシー）

**Interfaces:**
- Consumes: Task 1–3 の成果（`notes`、チャットパネル）。
- Produces: なし（ドキュメントのみ）。

- [ ] **Step 1: README にAIチャット節を追記**

`README.md` の「## 注意」節の直前に追記:

```markdown
## AIチャット（任意）

生成HTMLの右下パネルに Anthropic APIキー（`sk-ant-...`）を入力すると、ワークブックに
ついて日本語で質問できる。「全体概要 / このシートは？ / PQ・数式の説明 / メモを整形」の
サジェスト、または自由入力で回答が得られる（既定モデル `claude-opus-4-8`）。

- キーは**ブラウザのlocalStorageのみ**に保存され、生成HTMLには書き込まれない。
- 質問すると**ワークブックの構造・内容（プレビュー値やメモを含む）が Anthropic API に送信**される。
- キー未入力ならチャットは不活性で、グラフ可視化は従来通りオフラインで動作する。
```

- [ ] **Step 2: 表示確認**

Run: `python -c "print('AIチャット' in open('README.md').read())"`
Expected: `True`

- [ ] **Step 3: コミット**

```bash
git add README.md
git commit -m "docs: AIチャット（対話エージェント）の使い方とプライバシーを追記"
```

---

## 完了時

- Run: `python -m pytest -q`（56件 PASS）と `node tests/js/chat.test.mjs`（OK）で最終確認。
- ブランチ `claude/superpower-brainstorm-6q800x` に `git push -u origin`（ネットワーク失敗は指数バックオフで最大4回）。
- finishing-a-development-branch スキルで統合方法（PR等）を判断。
