# SKILLS_LIBRARY.md — Anthropic公式スキル保管庫

このファイルはAnthropicが提供する公式スキルの日本語リファレンス集。
SKILL.mdとは別管理。エージェントが特定の能力を必要とするときにここを参照する。

---

## 参照ルール

- **SKILL.md**：UNLIDプロジェクト専用の常時ロードスキル（エージェント定義・ブランドカラー等）
- **SKILLS_LIBRARY.md（このファイル）**：Anthropic公式スキルの保管庫。以下の状況で参照すること：
  - ユーザーがドキュメント生成（PDF・Word・Excel・PowerPoint）を依頼したとき
  - コードのテスト・品質改善を依頼されたとき
  - デザイン・ビジュアル系の出力を求められたとき
  - 上記以外でも「このタスクに適したスキルが保管庫にあるかも」と判断したとき

---

## 1. バンドルスキル（Claude Code標準搭載）

スラッシュコマンドで直接呼び出し可能。インストール不要。

| スキル | コマンド | 概要 | 主な用途 |
|---|---|---|---|
| batch | `/batch <指示>` | コードベース全体への大規模変更を並列実行。最大30タスクに分解し各タスクを独立エージェントが担当 | 大規模リファクタリング・一括変換 |
| claude-api | `/claude-api` | Claude APIリファレンスをロード（Python・TS・Go等）。ツール使用・ストリーミング・バッチ処理をカバー | Claude API実装時 |
| debug | `/debug [説明]` | 現在のセッションのデバッグログを解析し問題を診断 | セッション不具合の調査 |
| loop | `/loop [間隔] <プロンプト>` | 指定間隔でプロンプトを繰り返し実行 | デプロイ監視・PR定期確認 |
| simplify | `/simplify [フォーカス]` | 変更コードを3並列エージェントがレビューし品質・効率・再利用性の問題を修正 | コードレビュー・品質改善 |

---

## 2. ドキュメント処理スキル

### pdf — PDFドキュメント操作

**インストール：** `npx skills add anthropics/claude-code --skill pdf`

**できること：**
- PDFからテキスト・表を抽出
- 新規PDF作成
- 複数PDFの結合・分割
- フォームの読み取り・記入

**使い方の例：**
```
このPDFからテーブルのデータを抽出して
契約書のPDFをMarkdownに変換して
複数のPDFを1つにまとめて
```

---

### docx — Wordドキュメント操作

**インストール：** `npx skills add anthropics/claude-code --skill docx`

**できること：**
- .docxファイルの作成・編集
- 変更追跡・コメントの処理
- 書式設定の保持・変換
- テキスト抽出

**使い方の例：**
```
この内容をWordドキュメントとして保存して
Markdownをdocx形式に変換して
```

---

### pptx — PowerPointプレゼンテーション操作

**インストール：** `npx skills add anthropics/claude-code --skill pptx`

**できること：**
- PowerPointの作成・編集・解析
- レイアウト・テンプレートの適用
- チャート・グラフの生成
- スライドの自動生成

**使い方の例：**
```
この事業計画書をPowerPointに変換して
5枚のピッチデッキを作って
```

---

### xlsx — Excelスプレッドシート操作

**インストール：** `npx skills add anthropics/claude-code --skill xlsx`

**できること：**
- Excelファイルの作成・編集
- データ解析・集計
- グラフ生成
- 複数シートの管理

**使い方の例：**
```
この売上データをExcelにまとめて
損益シミュレーション表をxlsx形式で作成して
```

---

## 3. 開発・技術スキル

### webapp-testing — Webアプリテスト自動化

**インストール：** `npx skills add anthropics/claude-code --skill webapp-testing`

**できること：**
- Playwrightを使ったE2Eテスト作成
- UIのスクリーンショット取得・比較
- フォーム操作・ナビゲーションテスト

**使い方の例：**
```
このWebアプリのログインフローをテストして
```

---

### mcp-builder — MCPサーバー構築

**インストール：** `npx skills add anthropics/claude-code --skill mcp-builder`

**できること：**
- Model Context Protocol（MCP）サーバーのスキャフォールド生成
- ツール定義・プロンプト設計の支援
- 既存APIのMCPラッパー作成

**使い方の例：**
```
このAPIをMCPサーバーとして公開して
```

---

### claude-api（スキル版） — Claude API統合

**インストール：** コード内で`anthropic`をimportすると自動発動

**できること：**
- Claude APIの実装支援（Python・TypeScript・Go・Ruby・C#・PHP・Java）
- ツール使用・ストリーミング・バッチ処理のコード生成
- Agent SDKの使い方ガイド

---

## 4. デザイン・クリエイティブスキル

### frontend-design — フロントエンドデザイン

**インストール：** `npx skills add anthropics/claude-code --skill frontend-design`
**（277K+インストール、最多人気スキル）**

**できること：**
- UIコンポーネントのデザインパターン適用
- レスポンシブレイアウト設計
- アクセシビリティ対応
- CSSアーキテクチャの最適化

**使い方の例：**
```
このコンポーネントをより洗練されたデザインにして
```

---

### canvas-design — キャンバスデザイン

**インストール：** `npx skills add anthropics/claude-code --skill canvas-design`

**できること：**
- HTML5 Canvasを使ったビジュアル生成
- チャート・グラフの描画
- アニメーション作成

---

### algorithmic-art — アルゴリズムアート生成

**インストール：** `npx skills add anthropics/claude-code --skill algorithmic-art`

**できること：**
- ジェネレーティブアートの生成
- パターン・フラクタル・数理的ビジュアルの作成

---

### theme-factory — テーマ自動生成

**インストール：** `npx skills add anthropics/claude-code --skill theme-factory`

**できること：**
- 配色・タイポグラフィの一貫したテーマ生成
- CSS変数・デザイントークンの出力
- ブランドカラーからテーマ展開

---

## 5. ビジネス・コミュニケーションスキル

### internal-comms — 社内コミュニケーション文書

**インストール：** `npx skills add anthropics/claude-code --skill internal-comms`

**できること：**
- ステータスレポートの作成
- 社内ニュースレターの執筆
- FAQ・マニュアルの作成
- アナウンスメントの構造化

**使い方の例：**
```
先週の進捗を社内レポート形式でまとめて
```

---

### brand-guidelines — ブランドガイドライン適用

**インストール：** `npx skills add anthropics/claude-code --skill brand-guidelines`

**できること：**
- ブランドカラー・フォントの一貫適用
- スタイルガイドに沿ったコンテンツ生成
- ブランドチェックリストの確認

---

### doc-coauthoring — ドキュメント共同執筆

**インストール：** `npx skills add anthropics/claude-code --skill doc-coauthoring`

**できること：**
- 共同編集ワークフローの管理
- ドキュメントのレビュー・改善サイクル支援

---

## 6. ユーティリティスキル

### skill-creator — 新スキル作成ツール

**インストール：** `npx skills add anthropics/claude-code --skill skill-creator`

**できること：**
- 対話形式で新しいスキルを設計
- SKILL.mdのテンプレート生成
- スキルの説明・frontmatterの最適化

**使い方の例：**
```
/skill-creator
→ 対話形式でスキルの目的・動作・制約を定義→ SKILL.mdが自動生成される
```

---

### slack-gif-creator — Slack用GIF作成

**インストール：** `npx skills add anthropics/claude-code --skill slack-gif-creator`

**できること：**
- Slack投稿用のGIFアニメ生成
- テキストアニメーション・リアクションGIF作成

---

## 7. スキルのインストール方法

```bash
# 個別インストール
npx skills add anthropics/claude-code --skill <スキル名>

# 例：PDFスキルをインストール
npx skills add anthropics/claude-code --skill pdf

# グローバル（全プロジェクト共通）にインストール
npx skills add anthropics/claude-code --skill pdf --global
```

インストール後は `/pdf` または「このPDFを〜して」と話しかけると自動発動。

---

## 8. カスタムスキルの作り方（簡易リファレンス）

```yaml
# ~/.claude/skills/<skill-name>/SKILL.md の構造

---
name: skill-name                    # スラッシュコマンド名
description: いつ・何をするスキルか  # Claude が自動判断に使う
disable-model-invocation: true      # true = 手動呼び出しのみ
allowed-tools: Read, Grep, Bash     # 使用許可するツール
context: fork                       # fork = サブエージェントで実行
---

ここにスキルの指示を書く（Markdown形式）

$ARGUMENTS で呼び出し時の引数を参照できる
```

**保存場所による適用範囲：**

| 場所 | パス | 適用範囲 |
|---|---|---|
| 個人（全プロジェクト共通） | `~/.claude/skills/<name>/SKILL.md` | 自分の全プロジェクト |
| プロジェクト専用 | `.claude/skills/<name>/SKILL.md` | そのプロジェクトのみ |
| エンタープライズ | 管理者設定ファイル | 組織全員 |
