# xlsx-flow

結合セル・複数シートまたぎ・PowerQuery を含む Excel ワークブックの
データフローを、自己完結のインタラクティブ HTML として可視化する CLI。

## インストール

    pip install -e ".[dev]"

## 使い方

    xlsx-flow analyze path/to/book.xlsx -o out.html --json out.json

`out.html` をブラウザで開くと、粒度トグル（L1 フロー / L2 列 / L3 セル / PQ のみ）、
ズーム・パン、ノードクリックで詳細パネルが使える。

## 検証用サンプルの生成

    python -m samples.gen_sample sample.xlsx
    xlsx-flow analyze sample.xlsx -o sample.html

## 粒度レベル

- **L1**: 外部ソース → クエリ → シートのデータフロー（シート種別を色分け）
- **L2**: シートを論理列（結合ヘッダ復元）まで展開、列レベルの参照
- **L3**: セル/範囲レベルの数式依存（例 `B3→D3`、クロスシート `生データ!C3→売上集計!B4`）。
  ノード上限を超える大規模ブックでは一部のみ表示し警告を出す
- **PQ のみ**: PowerQuery の M クエリ依存グラフ。クエリは `let` ステップに展開され、
  `Excel.CurrentWorkbook` 参照は実テーブル/名前付き範囲（緑青の「表/範囲」ノード）に解決される

ノード／エッジをクリックすると詳細パネルに実セル内容のプレビュー（値・数式・範囲）、
M ステップ式、参照元の数式などが表示される。

## 注意

`xlsx_flow/render/template/cytoscape.min.js` はオフライン描画のため Cytoscape.js
3.30.2 の UMD ビルドを vendoring している（出力 HTML にインライン展開される）。
更新する場合は `npm pack cytoscape@<version>` で取得した `dist/cytoscape.min.js`
で置き換える。

## スコープ

本ツールはフェーズ1（可視化）。CSV / Python コードへの変換はフェーズ2で対象外。
