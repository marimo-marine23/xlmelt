# xlmelt

**複雑なExcelファイルをAIが読めるJSON/HTMLに変換するCLIツール**

> **Note**: このプロジェクトは初期段階（v0.1.0）です。基本的な変換機能は動作しますが、複雑なレイアウトのExcelファイルでは期待通りに変換できない場合があります。フィードバックや不具合報告は [Issues](https://github.com/marimo-marine23/xlmelt/issues) で歓迎しています。

日本企業に大量に存在する「文書としてのExcel」（Excel方眼紙、レビュー資料、仕様書、議事録など）を、
AIが効率的に読み込める構造化されたJSON/HTMLに変換します。

## 特徴

- **LLM不要で動作** — openpyxlのメタデータ解析だけで文書構造を検出。APIキーもクラウド接続も不要
- **日本のExcel文化に特化** — Excel方眼紙、稟議書、仕様書など日本特有のレイアウトパターンに対応
- **構造を理解する変換** — 見出し・表・キーバリュー・自由テキストを自動で判別し、セマンティックな構造に変換
- **バッチ処理対応** — ディレクトリ内のExcelファイルを一括変換、インデックス自動生成
- **画像・チャート抽出** — Excelに埋め込まれた画像やチャートを自動で抽出・保存
- **変換品質の可視化** — AI読みやすさスコアリングとJSON↔HTML一致性検証

## インストール

```bash
pip install xlmelt
```

### オプション依存関係

```bash
# .xls (Excel 97-2003) ファイルのサポート
pip install xlmelt[xls]

# チャート描画のサポート
pip install xlmelt[charts]

# 全機能
pip install xlmelt[xls,charts]
```

### 開発版をインストールする場合

```bash
git clone https://github.com/marimo-marine23/xlmelt.git
cd xlmelt
pip install -e ".[dev]"
```

## クイックスタート

```bash
# Excelファイルを変換（JSON + HTML を出力）
xlmelt convert input.xlsx -o ./output/

# 変換結果
# output/
# └── input/
#     ├── document.json    ← 構造化JSON
#     ├── document.html    ← セマンティックHTML
#     ├── images/          ← 抽出された画像
#     └── metadata.json    ← 変換メタデータ
```

## 使い方

### `xlmelt convert` — ファイル変換

```bash
# 基本（JSON + HTML を出力）
xlmelt convert input.xlsx -o ./output/

# JSONのみ出力
xlmelt convert input.xlsx -o ./output/ --format json

# HTMLのみ出力（CSSなし＝軽量版）
xlmelt convert input.xlsx -o ./output/ --format html --no-style

# ディレクトリ内の全Excelファイルを一括変換
xlmelt convert ./excel_dir/ -o ./output/

# 画像の抽出をスキップ
xlmelt convert input.xlsx -o ./output/ --images skip

# 標準出力にJSON出力（パイプ・AI連携向け）
xlmelt convert input.xlsx --format json --stdout

# パイプで他ツールに渡す
xlmelt convert input.xlsx --format json --stdout | jq '.document.sheets[0].sections'
```

**オプション一覧:**

| オプション | 値 | デフォルト | 説明 |
|---|---|---|---|
| `-o`, `--output` | パス | `./output` | 出力先ディレクトリ |
| `--format` | `json` / `html` / `both` | `both` | 出力形式 |
| `--images` | `extract` / `skip` | `extract` | 画像の扱い |
| `--no-style` | — | — | HTMLからCSSを除外 |
| `--stdout` | — | — | ファイルに書かず標準出力に出力 |

ディレクトリ一括変換時に2ファイル以上を処理した場合、`index.html`（人間向けリンク一覧）と`manifest.json`（AI向けカタログ）が自動生成されます。

### `xlmelt inspect` — 構造プレビュー

変換前にExcelの文書構造をツリー表示で確認できます。JSON形式での出力にも対応しています。

```bash
# ツリー表示（人間向け）
xlmelt inspect input.xlsx

# JSON出力（AI・プログラム向け）
xlmelt inspect input.xlsx --json
```

出力例（仕様書の場合）:

```
Document: sample_spec
Source: sample_spec.xlsx

Sheet: 仕様書
  [H2] 顧客管理システム　機能仕様書
  [KV] 6 pairs
    文書番号: FS-CRM-001
    版数: 2.0
    作成日: 2026年2月15日
    作成者: システム開発部　田中太郎
  [H3] 改訂履歴
  [TABLE] 3 rows headers=['版数', '日付', '変更者', '変更内容']
  [H2] 第1章　概要
  [H3] 1.1　目的
  [TEXT] 本仕様書は、顧客管理システム（CRM）の機能仕様を定義する...
  [H2] 第2章　機能仕様
  [H3] 2.1　顧客検索機能
  [H4] 2.1.1　検索条件
  [TEXT] 以下の条件による検索が可能であること...
```

### `xlmelt score` — AI読みやすさスコア

変換結果がAIにとってどの程度読みやすいかを定量的にスコアリングします。

```bash
# 単一ファイルのスコアリング
xlmelt score input.xlsx

# ディレクトリ内の全ファイルをスコアリング
xlmelt score ./excel_dir/

# JSON出力（プログラム連携向け）
xlmelt score input.xlsx --json

# レポートファイル出力
xlmelt score ./excel_dir/ --report score_report.md
```

**スコア指標:**

| スコア | Weight | 意味 |
|---|---|---|
| Noise Reduction | 15% | 空セル・装飾のみセルの除去率 |
| Structure Ratio | 30% | セマンティックなセクション（heading/table/kv/list）の比率 |
| Token Efficiency | 25% | 生セルダンプに対するJSON出力の効率 |
| Navigability | 30% | 見出し階層・セクション種類の多様性 |

### `xlmelt verify` — JSON↔HTML一致性検証 + xlsxカバレッジ

変換の品質を3段階で検証します。

```bash
# Excelファイルを指定（変換→検証を一発実行、xlsxカバレッジも自動チェック）
xlmelt verify input.xlsx

# 出力ディレクトリを検証（JSON↔HTML一致性）
xlmelt verify output/sample_spec/

# 出力ディレクトリ + xlsxカバレッジを検証（--xlsxで元ファイルの場所を指定）
xlmelt verify output/ --xlsx samples/output/

# 出力ディレクトリ全体を一括検証（サブディレクトリを自動探索）
xlmelt verify output/

# 検証結果をレポートファイルに出力
xlmelt verify input.xlsx --report verify_report.md
```

**3段階の検証:**

| Phase | 内容 | 検出できること |
|---|---|---|
| Phase 1 | JSON→HTMLの再レンダリング比較 | JSONに情報が欠落していないか |
| Phase 2 | セクション構造チェック | type、title、テーブルセルの整合性 |
| Phase 3 | xlsxセルカバレッジ | 元Excelから変換で漏れたセル |

### `xlmelt index` — インデックス再生成

既存の出力ディレクトリから `index.html` と `manifest.json` を再生成します。

```bash
xlmelt index ./output/
```

`manifest.json` はAIツールでの利用に最適化されたカタログで、各ファイルのシート構成・セクション概要・アウトラインを含みます。AIはこのファイルを1回読むだけで全体像を把握し、必要なファイルだけを選択的に読み込めます。

## 出力フォーマット

### JSON

文書構造をセマンティックなJSONとして出力します。

```json
{
  "xlmelt_version": "0.1.0",
  "schema_version": 1,
  "document": {
    "title": "sample_review",
    "source": "sample_review.xlsx",
    "sheets": [
      {
        "name": "概要",
        "sections": [
          {
            "type": "heading",
            "level": 2,
            "title": "設計レビュー資料",
            "source_range": "R2C1:R2C5",
            "source_range_a1": "A2:E2"
          },
          {
            "type": "key_value",
            "content": {
              "プロジェクト名": "顧客管理システム刷新",
              "対象工程": "基本設計",
              "レビュー日時": "2026年3月8日 14:00-16:00",
              "作成者": "高橋　一郎"
            }
          },
          {
            "type": "table",
            "content": {
              "headers": ["No.", "分類", "重要度", "指摘内容"],
              "rows": [["1", "機能", "重大", "検索条件に..."]]
            }
          }
        ],
        "section_summary": {"heading": 1, "key_value": 1, "table": 1}
      }
    ]
  }
}
```

**検出されるセクションタイプ:**

| type | 検出基準 | 説明 |
|---|---|---|
| `heading` | 大きいフォント、太字、結合セル | 見出し（H1〜H4） |
| `table` | 罫線で囲まれた矩形領域 | 表（ヘッダ行自動検出） |
| `key_value` | 太字ラベル＋値のペア | キーバリューペア |
| `list` | リストマーカー（・、●、1.、①等）で始まる行 | 箇条書き／番号付きリスト |
| `text` | 上記に該当しない文字列領域 | 自由テキスト |
| `image` | 埋め込み画像・チャートPNG | 画像（チャートはmatplotlibで描画） |

### HTML

セマンティックHTMLとして出力します。`h1`〜`h6`、`table`、`dl`（定義リスト）、`p` などの適切なHTML要素にマッピングされます。

## 対応ファイル形式

| 形式 | 対応状況 | 備考 |
|---|---|---|
| `.xlsx` (Excel 2007+) | 対応 | |
| `.xlsm` (マクロ付き) | 対応 | |
| `.xls` (Excel 97-2003) | 対応 | `pip install xlmelt[xls]` が必要 |

## Pythonライブラリとしての利用

CLIだけでなく、Pythonコードから直接利用することもできます。

```python
from xlmelt.core.analyzer import StructureAnalyzer
from xlmelt.output.json_writer import JsonWriter
from xlmelt.output.html_writer import HtmlWriter

# 解析
analyzer = StructureAnalyzer()
doc = analyzer.analyze("input.xlsx")

# JSON出力
json_writer = JsonWriter()
print(json_writer.to_string(doc))

# HTML出力
html_writer = HtmlWriter(include_style=False)
print(html_writer.to_string(doc))
```

## 開発

```bash
# 依存関係のインストール
pip install -e ".[dev]"

# テスト実行
pytest tests/ -v

# サンプルExcelファイルの生成
python samples/generate_samples.py
```

## ライセンス

[MIT License](LICENSE)

## 貢献

[CONTRIBUTING.md](CONTRIBUTING.md) を参照してください。
