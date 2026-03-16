# Changelog

## [0.1.0] - 2026-03-16

Initial public release.

### Features

- **Excel → JSON/HTML 変換** (`xlmelt convert`)
  - 見出し・表・キーバリュー・リスト・テキスト・画像を自動検出し、セマンティックな構造に変換
  - `.xlsx` / `.xlsm` / `.xls` 対応
  - ディレクトリ一括変換
  - `--stdout` による標準出力モード（パイプ・AI連携向け）
  - 画像・チャート抽出（チャートは matplotlib でPNG描画）

- **構造プレビュー** (`xlmelt inspect`)
  - ツリー表示で文書構造を確認
  - `--json` で機械可読なアウトライン出力

- **AI読みやすさスコア** (`xlmelt score`)
  - 変換前後の読みやすさを定量比較
  - ディレクトリ一括スコアリング
  - レポート出力 (`--report`)

- **変換品質検証** (`xlmelt verify`)
  - JSON↔HTML 一致性チェック
  - xlsx セルカバレッジチェック
  - レポート出力 (`--report`)

- **インデックス生成** (`xlmelt index`)
  - バッチ変換時に `index.html` + `manifest.json` を自動生成
  - 既存出力ディレクトリからの再生成にも対応

### Security

- グリッドサイズ制限（10,000行 × 1,000列）による DoS 防止
- 画像抽出サイズ制限（50MB/ファイル）
- stem 衝突検出による出力上書き防止
