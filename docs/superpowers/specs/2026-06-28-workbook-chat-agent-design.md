# xlsx-flow フェーズ2 ⑤: ワークブック対話エージェント 設計

- 日付: 2026-06-28
- ステータス: 承認済み（実装計画へ）
- 関連: フェーズ1（可視化）、フェーズ2 W1（①〜④グラフ深掘り、実装済み）

## 1. 目的

生成済みの自己完結HTMLに、ワークブックについて自然言語で質問でき、自然言語で
回答が返る**任意の対話エージェント**を追加する。次の4種の出力を会話で得られる:

1. ワークブック全体の概要（用途・データフロー・入出力）
2. シートごとの要約（種別・入力/出力・注意点）
3. PowerQuery / 数式の平易な説明
4. セル・シートに散在する自然言語メモの整形・構造化

## 2. アーキテクチャ方針

- 既存の `xlsx-flow analyze`（静的HTML出力）は**変更しない**。チャットは生成HTMLに
  追加される**任意機能**で、APIキー未入力なら不活性。グラフ可視化は従来通りフル機能。
- HTMLには既に `MODEL` JSON（nodes/edges: 構造・プレビュー・Mコード・PQステップ・
  L3セル依存）が埋め込まれている。これを文脈に**ブラウザから直接** Anthropic API を
  呼ぶ。バックエンド・サーバは持たない（**単一ファイル配布を維持**）。
- 採用根拠: ユーザーは「HTMLから質問→回答」かつ「自己完結性」を重視。ローカルサーバや
  MCPは配布性を損なうため、ブラウザ直呼出＋ユーザー貼付キーを選択。

## 3. コンポーネント

### 3.1 パーサ追加（唯一のPython変更）: ノート抽出
新規 `xlsx_flow/parser/notes.py`:
- `extract_notes(ws, ...) -> list[dict]`: 自然言語メモの収集。
  - **長文テキストセル**: data_type が文字列で、長さがしきい値（既定 >=20 文字）以上の
    非数式セル。`{sheet, addr, text}` を返す。
  - **セルコメント**: openpyxl の `cell.comment` があれば `{sheet, addr, text, kind:"comment"}`。
- 上限: ワークブック全体で最大 N 件（既定 60）。超過時は `model.warn` で通知。
- `workbook.analyze` が各シートで呼び、`Model` に集約。

### 3.2 Model 拡張
- `Model` に `notes: list[dict]` フィールドを追加（`to_dict()` に含める）。
- 既存ノード属性は不変。後方互換（notes 空でも従来通り）。

### 3.3 チャットUI（テンプレート: index.html / style.css / app.js）
- 右詳細パネルの下部に**折りたたみチャットパネル**。要素:
  - APIキー入力（type=password）＋「保存」。localStorage キー `xlsxflow_anthropic_key`。
  - メッセージ履歴表示（user/assistant を区別）。
  - 入力欄＋送信ボタン。
  - **サジェストボタン4種**（要望①〜④対応）:
    「全体概要」「このシートは？」「このPQの説明」「メモを整形」。
  - 選択中ノードがあれば、その文脈（例: `sheet:売上集計`）を質問に付与する。
- キー未入力時はキー入力UIのみ表示し、送信不可。

### 3.4 API呼出（ブラウザ→Anthropic直）
- `fetch("https://api.anthropic.com/v1/messages", {...})` POST。
- ヘッダ:
  - `x-api-key: <localStorageのキー>`
  - `anthropic-version: 2023-06-01`
  - `anthropic-dangerous-direct-browser-access: true`
  - `content-type: application/json`
- ボディ:
  - `model: "claude-opus-4-8"`（既定。任意で `claude-haiku-4-5` に切替可能）
  - `system`: 役割指示 ＋ グラウンディング（compact化した `MODEL`）
  - `messages`: 会話履歴（user/assistant 交互）
  - `max_tokens: 4096`、`thinking` は省略（Q&A即応重視）
- MVP は**非ストリーミング**（`stream` なし）。ストリーミングは次フェーズ。

### 3.5 グラウンディングの compact 化
- 送信前に `MODEL` を縮約する `compactModel()`（JS）:
  - cytoscape 用の冗長情報は元々無いが、`m_code` や `preview.cells` 等で肥大しうるため
    各文字列を上限長で切詰め、`notes` は全件（上限済み）。
  - 目的: コンテキスト肥大とコスト抑制。大規模ブックでも破綻しない。

## 4. データフロー

```
ユーザー質問
  → app.js: system(指示 + compactModel(MODEL)) + 会話履歴 を組立
  → fetch POST api.anthropic.com/v1/messages (x-api-key, direct-browser-access)
  → 応答 content[].text を抽出しパネルに追記
```
グラフ・依存・メモは全て `MODEL` に含まれるため、追加のデータ取得は不要。

## 5. セキュリティ / プライバシー

- APIキーは **localStorage のみ**に保存し、**生成HTMLには一切書き込まない**
  （HTMLを共有してもキーは漏れない）。
- パネルに明示文言: 「質問するとワークブックの構造・内容が Anthropic API へ送信されます」。
- **既定オフ（オプトイン）**: キー未入力なら一切送信しない。
- localStorage はオリジン単位（`file://` では実装依存だが実用上保存可）。クリア手段
  （「キー削除」）を用意。

## 6. エラー処理

- キー無 / 空 → 送信不可、キー入力を促す。
- 401（認証）/ 429（レート）/ ネットワーク → パネルにエラーメッセージ表示。グラフ機能は無影響。
- `stop_reason` を確認し、`refusal` 等は本文の有無を見て安全に表示。
- 大規模ブック → `compactModel()` で上限化（送信失敗を予防）。

## 7. テスト

- Python:
  - `notes.extract_notes` のユニットテスト（長文セル検出・コメント取得・しきい値・上限）。
  - `workbook.analyze` が `model.notes` を埋めること、`to_dict()` に含むこと。
  - `gen_sample` にメモセル/コメントを追加（既存アサーションを壊さない範囲で）。
- JS:
  - プロンプト組立 `buildRequest()` / `compactModel()` を純関数として切出し、`node` で
    構文チェック（`node --check`）＋簡易ロジック確認。
  - ブラウザ自動化はこの環境に playwright 不在のため不可。手動確認用にサンプルHTMLを提示。
- 既存テスト（49件）が緑のままであること。

## 8. スコープ外（次フェーズ）

- ストリーミング表示。
- ツール使用（特定セルをオンデマンド取得＝超大規模ブック向け）。
- 会話履歴の永続化、複数ブック横断。

## 9. 受け入れ基準

- `analyze` の静的HTMLは従来通りオフラインで開け、キー未入力ならグラフは全機能動作。
- キー入力後、サジェスト4種と自由入力で、全体概要/シート/PQ・数式/メモ整形の回答が得られる。
- キーはHTMLに書き込まれない。送信前に明示の同意文言が見える。
- 既存テスト緑 ＋ 新規 Python テスト緑、JS は構文チェック通過。
