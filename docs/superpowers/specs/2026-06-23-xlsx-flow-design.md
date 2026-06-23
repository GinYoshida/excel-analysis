# xlsx-flow 設計ドキュメント（フェーズ1: データプロセスの可視化）

- 日付: 2026-06-23
- ステータス: 設計確定（実装計画へ）
- 対象フェーズ: フェーズ1（可視化のみ）。フェーズ2（CSV / Python 変換）は本スペックの対象外。

## 1. 背景と目的

結合セルで列の意味が読み取りにくい、複数シートをまたぐ、PowerQuery を使っている、といった
「追いにくい」Excel ワークブックを、**データフロー図／UML としてインタラクティブに可視化**する
ツールを作る。目的は、ワークブック内でデータがどこから来てどこへ流れるか（外部ソース →
PowerQuery → シート → 他シート）を、人が探索的に理解できるようにすること。

### スコープ外（YAGNI / フェーズ2以降）
- CSV / Python コードへの変換（フェーズ2）
- xlsx の編集・書き戻し
- リアルタイム監視
- 複数ファイル横断の解析

## 2. ユーザーが扱う典型データ（検証用合成サンプルの像）

実ファイルは出さず、こちらで合成 `.xlsx` を生成して検証に使う。代表像:

- シートは 5 枚程度。生データと「値貼り付け」が混在。一部は PowerQuery で処理。
- 結合セルの 3 列表頭。例: 「売上」が 3 列を跨ぎ、配下に「当月 / 前月 / 前月差」のサブ列。
- 罫線で 3 列を囲った擬似ヘッダ（セル結合ではなく枠線で表頭を表現）も混在。
- 縦持ちデータを横持ちへ変換するような加工。
- 複数シートにデータがあり、一部にだけ関数が仕込まれている。

この構造を `samples/gen_sample.py` で再現し、テストの土台にする。

## 3. 出力形式

**自己完結のインタラクティブ HTML** 1 ファイル。
- ズーム / パン
- 粒度切替（後述の L1 / L2 / L3 / PQ のみ）
- ノードクリックで詳細パネル（数式、M コード、判定根拠など）
- サーバ・ネットワーク不要でブラウザ単体で開ける（Cytoscape.js をバンドル）

## 4. アーキテクチャ（A 案: 2 ステージ・パイプライン）

```
.xlsx → [parser] → Model(dataclass) → model.to_json() → [render] → out.html(自己完結)
                                          ↓
                                   out.json も保存(フェーズ2の入力 & デバッグ用)
```

一方向データフロー。パーサは HTML / JS を知らず、レンダラは xlsx を知らない。
両者の唯一の契約が中間モデル（`model.py`）。

### ディレクトリ構成

```
xlsx-flow/
├── xlsx_flow/
│   ├── __init__.py
│   ├── cli.py              # xlsx-flow analyze input.xlsx -o out.html
│   ├── parser/
│   │   ├── workbook.py     # zip展開・全体統括。Sheet/Query/Sourceを集約しModelを返す
│   │   ├── sheets.py       # シート種別判定 + 結合セル→論理ヘッダ復元
│   │   ├── formulas.py     # シート間参照(Sheet2!A1, 名前付き範囲)の抽出
│   │   └── powerquery.py   # DataMashup(M式)の抽出とクエリ間依存の解析
│   ├── model.py            # 中間モデル(dataclass) + JSONシリアライズ
│   └── render/
│       ├── html.py         # モデル→自己完結HTML(テンプレ+JSON埋め込み)
│       └── template/       # index.html / app.js(Cytoscape) / style.css
├── samples/
│   └── gen_sample.py       # 合成xlsxジェネレータ(5シート/結合ヘッダ/PQ/縦横変換)
├── tests/
│   ├── test_sheets.py
│   ├── test_formulas.py
│   ├── test_powerquery.py
│   └── test_model.py
├── pyproject.toml
└── README.md
```

### コンポーネント責務（各 1 目的・独立テスト可能）

- `parser/*`: xlsx を読んで構造化するだけ。openpyxl で読めるもの（シート・セル・結合・数式・
  名前付き範囲）は openpyxl、PowerQuery は xlsx 内の `customXml` / `DataMashup` を直接バイナリ解析。
- `model.py`: パーサとレンダラの唯一の契約。純データ（ノード／エッジ）。
- `render/*`: Model(JSON) を受け取って描くだけ。1 ファイル自己完結 HTML を出す。

### 技術スタック
Python 3.11+ / openpyxl / 標準ライブラリ（zipfile, xml, re）/ pytest。
HTML 側は CDN 非依存で Cytoscape.js をバンドル（オフラインでも開ける）。

## 5. 中間データモデル（パーサ↔レンダラの契約）

グラフ構造。**ノード 4 種・エッジ 2 種**。粒度切替はレンダラ側でノード／エッジ種別を
フィルタして実現（同一モデルを使い回す）。パーサは全粒度ぶんの情報を一度に出す。

### ノード（`Node`）

| 種別 | 説明 | 主な属性 |
|---|---|---|
| `source` | 外部データ源（PQ のファイル/DB/Web 接続） | `name`, `kind`(file/web/db), `path` |
| `query` | PowerQuery の M クエリ | `name`, `loads_to`(出力先), `m_code`(全文) |
| `sheet` | ワークシート | `name`, `sheet_type`, `dimensions`, `has_merged` |
| `column` | シート内の論理列（結合ヘッダ復元後） | `sheet`, `header_path`, `col_range`, `is_formula`, `confidence` |

`sheet_type`: `raw` / `pasted` / `formula` / `pq_output` / `mixed`。

### エッジ（`Edge`）

| 種別 | 意味 | 例 |
|---|---|---|
| `dataflow` | データが流れる（生成・ロード） | source→query, query→query, query→sheet |
| `reference` | 数式参照（読み取り依存） | sheetA.column→sheetB.column, sheet→sheet |

各エッジ: `source_id`, `target_id`, `edge_type`, `granularity`(最小表示レベル), `detail`(数式 / M ステップ等)。

### 粒度レベル（モデルのフィルタ）

- **L1 ワークブック/フロー**: `source`/`query`/`sheet` ノード ＋ `dataflow` ＋ シート単位に丸めた `reference`。
- **L2 テーブル/列**: 上記に `column` を展開、列単位の `reference`。
- **L3 セル/範囲**: `reference` の `detail`（セル/範囲レベルの数式依存）を展開。
- **PQ のみ**: `source`/`query` と `dataflow` だけ。

### ID 規約
`sheet:売上集計`, `col:売上集計/売上/当月`, `query:Query1`, `source:sales.csv`。
安定 ID で差分・クリック詳細を実現。

### JSON 形（抜粋イメージ）

```json
{
  "meta": {"file": "sample.xlsx", "generated_at": "...", "sheet_count": 5},
  "nodes": [
    {"id":"sheet:売上集計","type":"sheet","sheet_type":"formula","has_merged":true},
    {"id":"col:売上集計/売上/前月差","type":"column","sheet":"売上集計","header_path":["売上","前月差"],"is_formula":true}
  ],
  "edges": [
    {"source":"query:Q_売上","target":"sheet:売上集計","edge_type":"dataflow","granularity":"L1"},
    {"source":"col:売上集計/売上/前月差","target":"col:売上集計/売上/当月","edge_type":"reference","granularity":"L2","detail":"=当月-前月"}
  ],
  "warnings": []
}
```

## 6. パーサ中核アルゴリズム

### (a) 結合セル → 論理ヘッダ復元（`sheets.py`）
1. openpyxl の `merged_cells` を取得。横結合（"売上"が 3 列を跨ぐ）と縦結合を区別。
2. ヘッダ行推定: 上部の連続する非数値行をヘッダ帯とみなす（既定最大 3 行、自動検出。
   データ境界は「数値が支配的な行」で判定）。
3. 前方フィル: 横結合の親ラベルを配下の各列へ伝播し、サブラベルと結合して
   `header_path=["売上","当月"]` を構築。
4. 罫線擬似ヘッダ: 結合でなく罫線で囲う表頭は `border` 属性で矩形ブロックを検出する
   ヒューリスティック。確度が低ければ `confidence:"low"` を立て、HTML 側で点線＋「要確認」
   バッジ表示（誤読を断定しない）。

### (b) シート種別判定（`sheets.py`）
セル全走査で `data_type=='f'`（数式）比率を計算:
- 数式 0 件＆値が密 → `raw` か `pasted`（PQ 接続/外部リンク痕跡があれば pasted 寄り）
- 数式あり → `formula`
- PQ の `loads_to` 先 → `pq_output`
- 混在 → `mixed`
判定根拠（数式比率など）を `detail` に残す。

### (c) シート間参照の抽出（`formulas.py`）
各数式を正規表現＋トークナイズで解析し `Sheet2!A1`、`'売上 集計'!B2:B10`、名前付き範囲を抽出。
参照先を (a) の論理列にマップして列レベル `reference` エッジ化（マップ不能ならシート単位に
丸めて L1 へ）。縦持ち→横持ちは「同一シート内で複数の元範囲→1 列」の参照として表現。
循環参照は検出してフラグ。

### (d) PowerQuery 抽出（`powerquery.py`）— openpyxl では取れない
1. xlsx を zip で開き `xl/customXml/item*.xml` 内の DataMashup（Base64＋独自パッケージ）を探す。
2. 復号: ヘッダ＋内部 zip（`Formulas/Section1.m`）から M 式全文を得る。
3. `Section1.m` を `shared <名前> = ...;` 単位でクエリ分割。各本文から他クエリ参照（識別子一致）
   と Source（`Csv.Document` / `Excel.Workbook` / `Web.Contents` / `Sql.Database` 等）を抽出し
   `dataflow` エッジと `source` ノードを生成。
4. クエリ→ロード先シートは `xl/connections.xml` と `queryTable`/`pivotCache` 関係で対応付け。

DataMashup の形式はバージョン差があるため、専用テスト＋確度フラグで守る。抽出失敗時も
クラッシュさせず「PQ 検出（解析一部不可）」として描く。

### (e) 堅牢性方針
各抽出器は失敗を握りつぶさず `Model.warnings[]` に「どのシート/クエリで何が取れなかったか」
を集約 → HTML 上部に警告パネル表示。**部分的にでも必ず図を出す**ことを最優先。

## 7. テスト戦略

- 各抽出器を独立テスト（`test_sheets` / `test_formulas` / `test_powerquery` / `test_model`）。
- 合成 xlsx をテストの土台に: `samples/gen_sample.py` の既知構造に対し
  「結合ヘッダがこう復元される」「このシートは `formula` 判定」
  「Q_売上→売上集計 の dataflow が在る」等をアサート。
- レンダラは「モデル JSON → HTML に必要ノードが埋め込まれる」スモークテスト
  （ブラウザ自動操作はしない、YAGNI）。

## 8. 実装マイルストーン（各 MS の終わりに“見える HTML”が出る）

- **MS1 — 縦切り最小**: 合成 xlsx 生成 ＋ シート一覧と種別だけの L1 図を HTML 出力。
  ここで一度アウトプットを見て方向性確認。
- **MS2 — 結合ヘッダ復元**: L2 の論理列ノード追加。確度フラグ＋警告パネル。
- **MS3 — シート間参照**: `reference` エッジ（列/シート）。L3 詳細（数式）クリック表示。
- **MS4 — PowerQuery**: DataMashup 解析、source/query/dataflow。PQ のみビュー。
- **MS5 — 仕上げ**: 粒度トグル/ズーム/詳細パネルの UX、README、自己完結 HTML バンドル確認。

各 MS 末に HTML を実際に見せ、見せ方の要望を反映しながら進める。

## 9. CLI インターフェース

```
xlsx-flow analyze <input.xlsx> -o <out.html> [--json out.json] [--level L1|L2|L3|pq]
```
`--level` は初期表示粒度の指定（HTML 内では切替可能）。
