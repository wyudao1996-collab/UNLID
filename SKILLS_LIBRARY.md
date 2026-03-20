# SKILLS_LIBRARY.md — Anthropic公式スキル保管庫

Anthropicが提供する公式スキルの日本語リファレンス集。
SKILL.mdとは別管理。エージェントが特定の能力を必要とするときにここを参照する。

**最終更新：2026-03-20 ／ 公式スキル17種 + バンドル4種 + パートナー7種 = 全28種収録**

---

## 参照ルール

以下の状況では必ずこのファイルを確認すること：

- PDF・Word・Excel・PowerPointの生成・操作を依頼されたとき
- コードのテスト・品質改善を依頼されたとき
- デザイン・ビジュアル系の出力を求められたとき
- Notion・Jira・Figmaなど外部サービスとの連携を求められたとき
- 「このタスクに適したスキルが保管庫にあるか確認したい」と判断したとき

---

## 1. バンドルスキル（Claude Code標準搭載・インストール不要）

スラッシュコマンドで直接呼び出し可能。

| スキル | コマンド | 概要 | 主な用途 |
|---|---|---|---|
| **batch** | `/batch <指示>` | コードベース全体への大規模変更を最大30タスクに分解し各タスクを独立エージェントが並列実行 | 大規模リファクタリング・一括変換 |
| **debug** | `/debug [説明]` | 現在のセッションのデバッグログを解析し問題を診断 | セッション不具合の調査 |
| **loop** | `/loop [間隔] <プロンプト>` | 指定間隔でプロンプトを繰り返し実行 | デプロイ監視・PR定期確認 |
| **simplify** | `/simplify [フォーカス]` | 変更コードを3並列エージェントがレビューし品質・効率・再利用性の問題を修正 | コードレビュー・品質改善 |

---

## 2. ドキュメント処理スキル（4種）

> ⚠️ これら4スキルは source-available（非OSSライセンス）。参照用として公開。

### pdf — PDFドキュメント操作

**インストール：** `npx skills add anthropics/claude-code --skill pdf`

**できること：**
- PDFからテキスト・表を抽出（OCR対応）
- 新規PDF作成・既存PDF編集
- 複数PDFの結合・分割
- フォームの読み取り・記入
- 透かし・暗号化

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
- 数式の動的管理・データ解析・集計
- グラフ生成
- 複数シートの管理

**使い方の例：**
```
この売上データをExcelにまとめて
損益シミュレーション表をxlsx形式で作成して
```

---

## 3. 開発・技術スキル（4種）

### webapp-testing — Webアプリテスト自動化

**インストール：** `npx skills add anthropics/claude-code --skill webapp-testing`

**できること：**
- PlaywrightによるE2Eテスト自動化
- UIのスクリーンショット取得・比較
- DOMインスペクション・ブラウザログ確認
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
- Python（FastMCP）・TypeScript（MCP SDK）両対応
- 既存APIのMCPラッパー作成

**使い方の例：**
```
このAPIをMCPサーバーとして公開して
```

---

### claude-api — Claude API統合

**インストール：** `npx skills add anthropics/claude-code --skill claude-api`

**できること：**
- Claude APIの実装支援（Python・TypeScript・Go・Ruby・C#・PHP・Java）
- ツール使用・ストリーミング・バッチ処理のコード生成
- Agent SDKの使い方ガイド

**自動発動条件：** コードが `anthropic` / `@anthropic-ai/sdk` / `claude_agent_sdk` をインポートするとき

---

### skill-creator — 新スキル作成ツール

**インストール：** `npx skills add anthropics/claude-code --skill skill-creator`

**できること：**
- 対話形式で新しいスキルを設計
- SKILL.mdのテンプレート生成
- スキルの説明・frontmatterの最適化
- ベンチマーク評価ツール付き

**使い方の例：**
```
/skill-creator
→ 対話形式でスキルの目的・動作・制約を定義 → SKILL.mdが自動生成される
```

---

## 4. デザイン・クリエイティブスキル（5種）

### frontend-design — フロントエンドデザイン

**インストール：** `npx skills add anthropics/claude-code --skill frontend-design`
**（277K+インストール、最多人気スキル）**

**できること：**
- 「AIスロップ（汎用的で凡庸なUI）」を避ける高品質デザイン
- React + Tailwind + shadcn/ui による実装
- レスポンシブ・アクセシビリティ対応
- ジェネリックな配色（Ariel系・紫グラデ等）を禁止した品質基準

**使い方の例：**
```
このコンポーネントをより洗練されたデザインにして
ダッシュボードのUIを作って
```

---

### web-artifacts-builder — インタラクティブWebアーティファクト

**インストール：** `npx skills add anthropics/claude-code --skill web-artifacts-builder`

**できること：**
- React 18 + TypeScript + Tailwind + shadcn/ui によるインタラクティブなWebアーティファクト作成
- 複合コンポーネントの設計・実装
- プロトタイプの高速生成

**使い方の例：**
```
インタラクティブなデータ可視化コンポーネントを作って
```

---

### canvas-design — キャンバスデザイン

**インストール：** `npx skills add anthropics/claude-code --skill canvas-design`

**できること：**
- PNG・PDF形式のビジュアルアセット生成
- デザイン哲学に基づく構図設計
- グラフィック要素の作成

---

### algorithmic-art — アルゴリズムアート生成

**インストール：** `npx skills add anthropics/claude-code --skill algorithmic-art`

**できること：**
- p5.jsによるジェネレーティブアート生成
- シードランダム性・フローフィールド・パーティクルシステム
- パターン・フラクタル・数理的ビジュアルの作成

---

### theme-factory — テーマ自動生成

**インストール：** `npx skills add anthropics/claude-code --skill theme-factory`

**できること：**
- 10種のプリセットテーマを適用
- カスタムテーマの生成
- CSS変数・デザイントークンの出力
- ブランドカラーからのテーマ展開

---

## 5. ビジネス・コミュニケーションスキル（3種）

### internal-comms — 社内コミュニケーション文書

**インストール：** `npx skills add anthropics/claude-code --skill internal-comms`

**できること：**
- 3Pアップデート・社内ニュースレター・FAQレスポンス
- ステータスレポート・インシデントレポートの標準フォーマット生成

**使い方の例：**
```
先週の進捗を社内レポート形式でまとめて
```

---

### brand-guidelines — ブランドガイドライン適用

**インストール：** `npx skills add anthropics/claude-code --skill brand-guidelines`

**できること：**
- Anthropicブランドカラー・タイポグラフィ（Poppins/Lora）をアーティファクトに自動適用
- スタイルガイドに沿ったコンテンツ生成

---

### doc-coauthoring — ドキュメント共同執筆

**インストール：** `npx skills add anthropics/claude-code --skill doc-coauthoring`

**できること：**
- 共同編集ワークフローの管理
- 3イテレーション後に変更なければ削除確認という品質ゲート付き
- ドキュメントの反復改善サポート

---

## 6. ユーティリティスキル（1種）

### slack-gif-creator — Slack用GIF作成

**インストール：** `npx skills add anthropics/claude-code --skill slack-gif-creator`

**できること：**
- Slackのサイズ制限に最適化されたGIFアニメーション生成
- テキストアニメーション・リアクションGIF作成

---

## 7. パートナー公式スキル（7種）

Anthropic認定パートナーが提供する公式スキル。

| パートナー | スキル概要 | インストール |
|---|---|---|
| **Notion** | ページ作成・データベース操作・コンテンツ管理の自動化 | `/plugin install notion@anthropic-agent-skills` |
| **Asana** | タスク・プロジェクト管理・担当割り当て・進捗追跡の自動化 | `/plugin install asana@anthropic-agent-skills` |
| **Atlassian** | Jira（課題管理）・Confluence（ドキュメント）との統合 | `/plugin install atlassian@anthropic-agent-skills` |
| **Canva** | デザインテンプレート操作・ビジュアルコンテンツ作成連携 | `/plugin install canva@anthropic-agent-skills` |
| **Figma** | UIデザインファイルの読み取り・コンポーネント情報取得 | `/plugin install figma@anthropic-agent-skills` |
| **Sentry** | エラー監視・バグトラッキング・インシデント管理の自動化 | `/plugin install sentry@anthropic-agent-skills` |
| **Zapier** | 5,000+アプリとのワークフロー自動化・トリガー設計 | `/plugin install zapier@anthropic-agent-skills` |

---

## 8. スキルのインストール方法まとめ

```bash
# 個別インストール（プロジェクト専用）
npx skills add anthropics/claude-code --skill <スキル名>

# グローバルインストール（全プロジェクト共通）
npx skills add anthropics/claude-code --skill <スキル名> --global

# パートナースキル（プラグイン経由）
/plugin install <パートナー名>@anthropic-agent-skills

# 例
npx skills add anthropics/claude-code --skill pdf
npx skills add anthropics/claude-code --skill frontend-design --global
```

---

## 9. カスタムスキルの作り方（簡易リファレンス）

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

---

## 参照元

- [github.com/anthropics/skills](https://github.com/anthropics/skills) — 公式リポジトリ
- [github.com/anthropics/skills/tree/main/skills](https://github.com/anthropics/skills/tree/main/skills) — スキル一覧
