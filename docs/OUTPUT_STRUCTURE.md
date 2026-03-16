# 出力構造設計: 画像・複数ファイル対応

## 1. 背景

### Excelファイルの実態
Excel (.xlsx) はZIPアーカイブであり、内部構造は以下の通り:

```
example.xlsx (ZIP)
├── [Content_Types].xml
├── _rels/
├── xl/
│   ├── workbook.xml
│   ├── sharedStrings.xml
│   ├── styles.xml
│   ├── worksheets/
│   │   ├── sheet1.xml
│   │   └── sheet2.xml
│   ├── drawings/
│   │   ├── drawing1.xml    ← 画像の配置情報
│   │   └── drawing2.xml
│   └── media/
│       ├── image1.png      ← 実際の画像ファイル
│       ├── image2.jpeg
│       └── image3.emf
│   └── charts/
│       └── chart1.xml      ← グラフ
└── docProps/
```

### 画像の種類
日本企業のExcelファイルに含まれる画像:
- **スクリーンショット**: システム画面、エラー画面、設計図
- **ロゴ・印影**: 社印、承認印、部門ロゴ
- **図表**: 手書きの構成図、フロー図
- **写真**: 現場写真、製品写真
- **グラフ**: Excelで作成したチャート（XMLとして埋め込まれている）
- **OLEオブジェクト**: 他のOfficeファイルの埋め込み

### 課題
- 画像はJSON/HTMLに「インライン」で含めるのが困難（base64は巨大になる）
- 画像の配置位置（どのセル範囲の上に置かれているか）が構造理解に重要
- グラフはXMLだが、画像化して保存する方が汎用的
- 単一ファイル出力ではなくフォルダ出力が必須

---

## 2. 出力フォルダ構造

### 基本構造
```
入力: report_2024Q3.xlsx
出力: report_2024Q3/
      ├── document.json         ← メイン: 構造化されたJSON
      ├── document.html         ← メイン: セマンティックHTML
      ├── images/               ← 抽出された画像
      │   ├── sheet1_img001.png
      │   ├── sheet1_img002.jpeg
      │   └── sheet2_img001.png
      └── metadata.json         ← 変換メタデータ
```

### バッチ変換時
```
入力: excel_docs/
      ├── report_2024Q3.xlsx
      ├── spec_v2.xlsx
      └── minutes_0301.xlsx

出力: output/
      ├── report_2024Q3/
      │   ├── document.json
      │   ├── document.html
      │   ├── images/
      │   │   ├── sheet1_img001.png
      │   │   └── sheet1_img002.jpeg
      │   └── metadata.json
      ├── spec_v2/
      │   ├── document.json
      │   ├── document.html
      │   └── metadata.json        ← 画像なしならimages/は作らない
      ├── minutes_0301/
      │   ├── document.json
      │   ├── document.html
      │   ├── images/
      │   │   └── sheet1_img001.png
      │   └── metadata.json
      └── index.json                ← バッチ全体のサマリ
```

---

## 3. JSON出力での画像参照

### 画像メタデータをセクションとして表現
```json
{
  "type": "image",
  "title": null,
  "source_range": "B15:F25",
  "content": {
    "path": "images/sheet1_img001.png",
    "original_name": "image1.png",
    "format": "png",
    "size_bytes": 145320,
    "dimensions": {
      "width_px": 640,
      "height_px": 480
    },
    "position": {
      "sheet": "概要",
      "anchor_cell": "B15",
      "span_cols": 5,
      "span_rows": 11
    },
    "alt_text": null,
    "description": null
  }
}
```

### コンテキスト内での位置づけ
画像がどのセクションに属するかを構造的に表現:

```json
{
  "sections": [
    {
      "type": "heading",
      "level": 2,
      "content": "3.2 システム構成図"
    },
    {
      "type": "image",
      "content": {
        "path": "images/sheet1_img003.png",
        "position": {"anchor_cell": "B30"}
      }
    },
    {
      "type": "text",
      "content": "上記の構成図に基づき、各コンポーネントの詳細を以下に示す。"
    }
  ]
}
```

---

## 4. HTML出力での画像参照

### 相対パスで参照
```html
<!-- document.html -->
<h2>3.2 システム構成図</h2>
<figure data-source-range="B30:F40">
  <img src="images/sheet1_img003.png"
       alt="システム構成図"
       width="640" height="480"
       loading="lazy">
  <figcaption>※ 元のExcelファイル: sheet1 B30:F40</figcaption>
</figure>
<p>上記の構成図に基づき、各コンポーネントの詳細を以下に示す。</p>
```

### 画像インライン埋め込みモード（オプション）
単一ファイルHTMLが必要な場合はbase64埋め込みも選択可能:
```bash
# 通常: フォルダ出力（画像は別ファイル）
xlmelt convert input.xlsx -o output/

# 単一HTML: 画像をbase64で埋め込み（ファイルサイズ注意）
xlmelt convert input.xlsx -o output.html --embed-images
```

---

## 5. 画像抽出の実装方針

### 5.1 openpyxlでの画像抽出
```python
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import zipfile

def extract_images(xlsx_path: Path, output_dir: Path) -> list[ImageInfo]:
    """Excelファイルから画像を抽出し、出力ディレクトリに保存"""
    images = []

    # 方法1: openpyxlのdrawingオブジェクトから位置情報を取得
    wb = load_workbook(xlsx_path)
    for sheet in wb.worksheets:
        if sheet._images:
            for idx, img in enumerate(sheet._images):
                # img.anchor で配置位置を取得
                ...

    # 方法2: ZIPとして直接xl/media/を展開（より確実）
    with zipfile.ZipFile(xlsx_path, 'r') as zf:
        media_files = [f for f in zf.namelist() if f.startswith('xl/media/')]
        for media_file in media_files:
            # 画像を抽出して保存
            ...

    return images
```

### 5.2 画像の命名規則
```
images/
├── {sheet_name}_img{連番:03d}.{拡張子}
│
│   例:
├── 概要_img001.png          ← 「概要」シートの1枚目
├── 概要_img002.jpeg         ← 「概要」シートの2枚目
├── レビュー指摘一覧_img001.png  ← 別シートの画像
│
│   シート名にファイル名不正文字がある場合:
├── sheet1_img001.png        ← シート名をサニタイズ
```

### 5.3 対応フォーマット
| 形式 | 対応 | 備考 |
|------|------|------|
| PNG | ◎ | そのまま抽出 |
| JPEG | ◎ | そのまま抽出 |
| GIF | ◎ | そのまま抽出 |
| BMP | ◯ | PNGに変換して保存 |
| EMF/WMF | △ | Windowsメタファイル。PNGへの変換を試みる（Pillow/cairosvgが必要） |
| TIFF | ◯ | PNGに変換して保存 |
| SVG | ◯ | そのまま抽出 |

### 5.4 Excelグラフ（Chart）の扱い
Excelのグラフはxl/charts/にXMLとして保存されている。

**選択肢**:
1. **無視する**（Phase 1）— グラフ再現は複雑すぎる
2. **プレースホルダーとして記録** — 「ここにグラフがあった」という情報だけJSON/HTMLに出力
3. **matplotlibで再描画**（Phase 2以降）— データと設定を読み取って静的画像に変換
4. **Excelのキャッシュ画像を抽出** — Excel内にグラフのキャッシュ画像がある場合がある

**推奨**: Phase 1では選択肢2（プレースホルダー）、Phase 2で選択肢4を試みる

```json
{
  "type": "chart",
  "content": {
    "chart_type": "bar",
    "title": "月別売上推移",
    "source_range": "A1:D12",
    "cached_image": "images/sheet1_chart001.png",
    "note": "Excelグラフから自動抽出。元データはsource_rangeを参照"
  }
}
```

---

## 6. metadata.json

各変換フォルダに含まれる変換メタデータ:

```json
{
  "xlmelt_version": "0.1.0",
  "converted_at": "2026-03-10T23:30:00+09:00",
  "source": {
    "filename": "report_2024Q3.xlsx",
    "size_bytes": 2457600,
    "created": "2024-09-15T10:00:00",
    "modified": "2024-10-01T14:30:00",
    "author": "田中太郎"
  },
  "conversion": {
    "mode": "heuristic",
    "llm_used": false,
    "llm_provider": null,
    "processing_time_ms": 1250
  },
  "content_summary": {
    "sheet_count": 3,
    "total_sections": 15,
    "image_count": 4,
    "chart_count": 1,
    "section_types": {
      "heading": 5,
      "table": 3,
      "key_value": 4,
      "text": 2,
      "image": 4,
      "chart": 1
    }
  },
  "images": [
    {
      "path": "images/概要_img001.png",
      "format": "png",
      "size_bytes": 145320,
      "dimensions": {"width_px": 640, "height_px": 480},
      "sheet": "概要",
      "anchor_cell": "B15"
    }
  ]
}
```

---

## 7. index.json（バッチ処理時）

```json
{
  "xlmelt_version": "0.1.0",
  "batch_converted_at": "2026-03-10T23:30:00+09:00",
  "source_directory": "./excel_docs/",
  "total_files": 3,
  "total_images": 7,
  "documents": [
    {
      "source": "report_2024Q3.xlsx",
      "output_dir": "report_2024Q3/",
      "sheet_count": 3,
      "section_count": 15,
      "image_count": 4
    },
    {
      "source": "spec_v2.xlsx",
      "output_dir": "spec_v2/",
      "sheet_count": 1,
      "section_count": 8,
      "image_count": 0
    },
    {
      "source": "minutes_0301.xlsx",
      "output_dir": "minutes_0301/",
      "sheet_count": 1,
      "section_count": 6,
      "image_count": 3
    }
  ]
}
```

---

## 8. CLI の更新

```bash
# 通常変換 → フォルダが出力される
xlmelt convert input.xlsx -o ./output/
# → ./output/input/ フォルダが生成される

# 出力形式の指定
xlmelt convert input.xlsx -o ./output/ --format json      # JSONのみ
xlmelt convert input.xlsx -o ./output/ --format html      # HTMLのみ
xlmelt convert input.xlsx -o ./output/ --format both      # 両方 (デフォルト)

# 画像の扱い
xlmelt convert input.xlsx -o ./output/ --images extract   # 画像を別ファイルに抽出 (デフォルト)
xlmelt convert input.xlsx -o ./output/ --images embed     # base64埋め込み (単一ファイル)
xlmelt convert input.xlsx -o ./output/ --images skip      # 画像を無視

# 単一ファイル出力 (画像なしの場合、または画像埋め込みの場合)
xlmelt convert input.xlsx -o output.json --images skip    # 拡張子指定で単一ファイル

# バッチ変換
xlmelt convert ./excel_dir/ -o ./output_dir/
# → ./output_dir/{各ファイル名}/ フォルダが生成
# → ./output_dir/index.json が生成
```

---

## 9. セキュリティ上の画像に関する注意点

### 9.1 画像内の機密情報
- 画像にはスクリーンショットや個人情報が含まれうる
- `--images skip` オプションで画像を完全に除外できるようにする
- LLM APIには画像を送信しない（メタデータのみ）

### 9.2 悪意ある画像ファイル
- EMF/WMFは脆弱性の温床として知られる → 変換せずスキップするオプション
- 画像ファイルサイズの上限チェック
- 展開前にZIP内のファイルパスをバリデーション（ZipSlip対策）

### 9.3 出力フォルダのパーミッション
```python
# 出力ディレクトリは700で作成（所有者のみアクセス可能）
output_dir.mkdir(mode=0o700, parents=True, exist_ok=True)
```
