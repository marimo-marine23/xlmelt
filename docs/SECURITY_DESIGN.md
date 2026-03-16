# xlmelt セキュリティ設計書

## 1. 脅威モデル概要

xlmeltは「社内のExcelファイル → JSON/HTML変換」ツールであり、以下の特性からセキュリティが極めて重要:

- **入力**: 社内機密情報を含むExcelファイル（人事、財務、顧客情報等）
- **外部通信**: LLM APIへのデータ送信（オプション）
- **出力**: JSON/HTMLファイル（機密情報がそのまま含まれる）
- **配布**: OSSとして公開（コード自体は誰でも読める）

### 脅威アクター
| アクター | リスク |
|---------|--------|
| 外部攻撃者 | 悪意あるExcelファイルによる攻撃、WebUIへの不正アクセス |
| LLMプロバイダー | APIに送信されたデータの漏洩・学習利用 |
| OSSサプライチェーン | 依存パッケージの脆弱性 |
| 内部ユーザー | 設定ミスによる意図しないデータ露出 |

---

## 2. 最重要: LLM APIへのデータ送信

### 2.1 リスク
LLM APIを使う場合、Excelの内容がAPIプロバイダーのサーバーに送信される。
- 社内機密情報がAPI経由で外部に出る
- プロバイダーによっては学習データに利用される可能性
- ネットワーク経路上での傍受リスク

### 2.2 対策: データ最小化原則

**基本方針**: LLM APIにはセル値の実データを送信しない。構造メタデータのみ送信する。

```
送信するもの (構造メタデータ):
  ✅ セルの型情報 (text/number/date/empty)
  ✅ セル結合パターン (A1:D1 merged)
  ✅ フォントサイズ・太字の有無
  ✅ 罫線パターン
  ✅ 背景色パターン
  ✅ 行/列の相対位置関係
  ✅ テキストの文字数（値そのものではなく）

送信しないもの (実データ):
  ❌ セルの値そのもの
  ❌ 人名、金額、日付等の具体的データ
  ❌ ファイル名、シート名（オプションで匿名化）
```

#### 実装: メタデータサマリ生成
```python
# 実際に送信されるデータのイメージ
{
  "sheet_structure": {
    "rows": 45,
    "cols": 12,
    "regions": [
      {
        "range": "A1:L1",
        "type": "merged_cell",
        "font_size": 16,
        "bold": true,
        "text_length": 15,
        "content_type": "japanese_text"
      },
      {
        "range": "A3:B3",
        "type": "cell_pair",
        "left_font": {"bold": true, "size": 11},
        "right_font": {"bold": false, "size": 11},
        "pattern": "label_value"
      },
      {
        "range": "A6:L20",
        "type": "bordered_region",
        "has_header_row": true,
        "header_bg_color": "#4472C4",
        "col_count": 12,
        "row_count": 15
      }
    ]
  }
}
```

### 2.3 ユーザーへの透明性

#### 送信前確認
```bash
# --dry-run で送信されるデータを事前に確認可能
xlmelt convert input.xlsx -o output.json --llm claude --dry-run

# 出力例:
# [LLM] 以下のメタデータがAPI に送信されます:
# - シート構造情報: 3シート分
# - 推定トークン数: ~1,200 tokens
# - 推定コスト: ~$0.003
# - セルの実データ: 送信されません
# 続行しますか? [y/N]
```

#### --no-content モードと --allow-content モード
```bash
# デフォルト: メタデータのみ送信 (安全)
xlmelt convert input.xlsx --llm claude

# 精度向上のため内容送信を明示的に許可する場合
xlmelt convert input.xlsx --llm claude --allow-content

# --allow-content 使用時の警告:
# ⚠ WARNING: --allow-content が指定されています
# セルの内容がLLM APIに送信されます
# 機密情報が含まれていないことを確認してください
# API プロバイダーのデータポリシー:
#   Anthropic: https://www.anthropic.com/policies/privacy
#   OpenAI: https://openai.com/policies/api-data-usage
```

### 2.4 ローカルLLM推奨

機密性が高い環境では、OllamaなどのローカルLLMを推奨:

```bash
# ローカルLLM使用 (データは一切外部に出ない)
xlmelt convert input.xlsx --llm ollama --model llama3

# 設定ファイルでデフォルトをローカルに
# ~/.xlmelt/config.toml
# [llm]
# provider = "ollama"
# model = "llama3"
```

---

## 3. 入力ファイルの安全性

### 3.1 悪意あるExcelファイルへの対策

Excelファイルにはマクロ、外部リンク、OLEオブジェクト等が含まれうる。

| 脅威 | 対策 |
|------|------|
| VBAマクロ | openpyxlはマクロを実行しない (読み込みのみ) → 安全 |
| 外部データ接続 | パースせず無視する |
| OLEオブジェクト | 画像以外は無視、画像もメタデータのみ抽出 |
| XML External Entity (XXE) | openpyxlのdefusedxml依存で対策済み |
| Zip爆弾 (.xlsxはZIP) | ファイルサイズ上限チェック (デフォルト100MB) |
| パストラバーサル | 出力パスのバリデーション |
| 数式インジェクション | 数式は評価せず文字列として扱う |

#### 実装: 入力バリデーション
```python
class InputValidator:
    MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB
    ALLOWED_EXTENSIONS = {'.xlsx', '.xls', '.xlsm', '.xlsb'}

    def validate(self, file_path: Path) -> ValidationResult:
        # 1. 拡張子チェック
        # 2. ファイルサイズチェック
        # 3. マジックバイトチェック (ZIP/OLE2シグネチャ)
        # 4. 展開後サイズチェック (zip bomb対策)
        ...
```

### 3.2 出力ファイルの安全性

#### HTML出力時のXSS対策
Excelセルの内容をHTMLに出力する際、XSSを防止:

```python
# セル値をHTML出力する際は必ずエスケープ
import html
safe_value = html.escape(cell_value)

# ただし、WebUI上でのプレビューではCSPヘッダーも設定
# Content-Security-Policy: default-src 'self'; script-src 'self'
```

#### JSON出力時のインジェクション対策
```python
# JSON出力は標準ライブラリのjson.dumpsを使用
# ensure_ascii=False で日本語をそのまま出力
# ただしJSON内にHTMLタグ等が含まれる場合の注意を文書化
```

---

## 4. WebUI セキュリティ

### 4.1 基本方針
WebUIは **localhost のみ** でリッスン。外部公開は想定しない。

```python
# デフォルト: localhostのみ
uvicorn.run(app, host="127.0.0.1", port=8080)

# ⚠ --host 0.0.0.0 は明示的なフラグでのみ許可し、警告を表示
# xlmelt ui --host 0.0.0.0
# WARNING: 外部からのアクセスが可能になります
# 認証なしでの公開は推奨しません
```

### 4.2 セキュリティヘッダー
```python
# FastAPIミドルウェアで設定
@app.middleware("http")
async def security_headers(request, call_next):
    response = await call_next(request)
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Content-Security-Policy"] = "default-src 'self'; style-src 'self' 'unsafe-inline'"
    response.headers["X-XSS-Protection"] = "1; mode=block"
    return response
```

### 4.3 ファイルアクセス制御
WebUI経由でのファイル操作は、指定されたワークディレクトリ内に制限:

```python
# パストラバーサル対策
def safe_resolve(base_dir: Path, user_path: str) -> Path:
    resolved = (base_dir / user_path).resolve()
    if not resolved.is_relative_to(base_dir.resolve()):
        raise SecurityError("パストラバーサルが検出されました")
    return resolved
```

---

## 5. APIキー管理

### 5.1 保存方法
```toml
# ~/.xlmelt/config.toml
# ファイル権限: 600 (owner read/write only)

[llm.anthropic]
api_key = "sk-ant-..."

[llm.openai]
api_key = "sk-..."
```

### 5.2 安全策
```python
# 設定ファイル作成時にパーミッション設定
import os
config_path = Path.home() / ".xlmelt" / "config.toml"
config_path.parent.mkdir(mode=0o700, exist_ok=True)
config_path.touch(mode=0o600)

# 環境変数からの読み込みも対応 (CI/CD向け)
# EXCELLENS_ANTHROPIC_API_KEY
# EXCELLENS_OPENAI_API_KEY

# APIキーをログに出力しない
# APIキーをエラーメッセージに含めない
# APIキーを出力JSONに含めない
```

### 5.3 .gitignore テンプレート
プロジェクトの.gitignoreに以下を推奨として文書化:
```
# xlmelt
.xlmelt/
*.xlsx    # (必要に応じて)
output/
```

---

## 6. 依存パッケージのセキュリティ

### 6.1 最小依存原則
不必要な依存を避け、攻撃対象面を最小化:

```
# 必須依存 (Phase 1)
openpyxl >= 3.1.0      # Excelパース (defusedxml内蔵)
click >= 8.0            # CLI
pydantic >= 2.0         # データバリデーション

# オプション依存
fastapi                 # WebUI使用時のみ
uvicorn                 # WebUI使用時のみ
anthropic               # Claude API使用時のみ
openai                  # OpenAI API使用時のみ
```

### 6.2 脆弱性スキャン
```bash
# CI/CDで自動チェック
pip-audit               # 既知の脆弱性チェック
safety check            # 代替ツール

# GitHub Dependabotの有効化
# Snykの導入検討
```

### 6.3 サプライチェーン攻撃対策
- pyproject.toml でバージョン範囲を適切に制限
- ロックファイル (pip-compile/poetry.lock) をリポジトリに含める
- 主要依存のハッシュ固定を検討

---

## 7. ログとプライバシー

### 7.1 ログポリシー
```python
# ログに含めてよいもの
logger.info(f"Processing file: {file_path.name}")  # ファイル名はOK
logger.info(f"Sheet '{sheet_name}' has {row_count} rows")
logger.info(f"Detected {len(sections)} sections")

# ログに含めてはいけないもの
# ❌ セルの値
# ❌ APIキー
# ❌ ファイルのフルパス (ユーザー名が含まれうる)
# ❌ LLM APIのレスポンス全文
```

### 7.2 テレメトリ
**収集しない**。OSSとして信頼を得るため、テレメトリ/分析機能は一切実装しない。

---

## 8. セキュリティチェックリスト

### リリース前チェック
- [ ] 入力バリデーション (ファイルサイズ、拡張子、マジックバイト)
- [ ] HTML出力のXSSエスケープテスト
- [ ] パストラバーサルテスト
- [ ] APIキーがログ/出力に漏洩しないことの確認
- [ ] WebUIのlocalhostバインド確認
- [ ] 設定ファイルのパーミッション確認
- [ ] 依存パッケージの脆弱性スキャン
- [ ] LLMへのデータ送信内容の確認 (dry-run テスト)
- [ ] zip bomb対策テスト
- [ ] 数式インジェクション対策テスト

### 定期チェック
- [ ] Dependabot/pip-auditの警告対応
- [ ] LLMプロバイダーのデータポリシー変更監視
- [ ] セキュリティレポートへの対応プロセス (SECURITY.md)

---

## 9. SECURITY.md (リポジトリ用)

リポジトリのルートに以下を配置:

```markdown
# Security Policy

## Reporting a Vulnerability
If you discover a security vulnerability, please report it via:
- Email: [security contact]
- GitHub Security Advisories (private)

Please do NOT open a public issue for security vulnerabilities.

## Scope
- Excel file parsing vulnerabilities
- XSS in HTML output
- API key exposure
- Data leakage to LLM APIs
- Path traversal
- WebUI security

## Response
We aim to respond within 48 hours and release patches for critical issues within 7 days.
```
