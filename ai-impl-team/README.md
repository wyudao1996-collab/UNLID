# /ai-impl-team — 外部AI実装支援 9体エージェントチーム

## 設計の3原則（必ず守ること）

1. **Maker-Checker 原則**：すべての成果物は必ず別エージェントがレビューする
2. **ループ上限ルール**：同一フェーズの修正は最大3回。超えたら人間にエスカレーション
3. **Git-First 原則**：フェーズ完了時に即コミット。いつでもロールバック可能な状態を維持する

---

## 開発ループ

```
@requirements → @architect → (@cx コピー確認) → @ux
→ ★GATE#1 顧客承認(@cx) → @dev → @reviewer → @qa
→ ★GATE#2 顧客承認(@cx) → @docs → @pm（納品最終確認）
→ 追加要件があればループ先頭に戻る
```

**補足：**
- `(@cx コピー確認)`：@architect がブリーフを渡した後、@ux 着手前に @cx が画面コピーを確定させる
- `@pm（納品最終確認）`：@docs が4ファイル完成後、@pm が全成果物の揃いを確認してから納品する

---

## プロジェクト標準フォルダ構造

```
[client_name]/
├── requirements.md        # @requirements が作成
├── architecture.md        # @architect が作成
├── feedback_[n].md        # @cx が作成（n=ループ番号）
├── review_report.md       # @reviewer が作成
├── qa_report.md           # @qa が作成
├── loop_counter.md        # @pm が管理
├── mockup/[機能名].html   # @ux が作成
├── src/[機能名].py        # @dev が作成
├── tests/test_[機能名].py
└── docs/
    ├── user_manual.md
    ├── technical_spec.md
    ├── troubleshooting.md
    └── maintenance.md
```

---

## エージェント一覧

| エージェント | 役割 | 詳細 |
|---|---|---|
| @pm | プロジェクト司令塔 | [agents/pm.md](agents/pm.md) |
| @requirements | 要件定義アナリスト | [agents/requirements.md](agents/requirements.md) |
| @architect | ソリューションアーキテクト | [agents/architect.md](agents/architect.md) |
| @ux | UX / モックアップデザイナー | [agents/ux.md](agents/ux.md) |
| @dev | 開発者 | [agents/dev.md](agents/dev.md) |
| @reviewer | コードレビュアー | [agents/reviewer.md](agents/reviewer.md) |
| @qa | QAエンジニア | [agents/qa.md](agents/qa.md) |
| @docs | テクニカルライター | [agents/docs.md](agents/docs.md) |
| @cx | 顧客対応スペシャリスト | [agents/cx.md](agents/cx.md) |
