# Skill Creator — 他リポジトリへのインストール手順

このスキルは **UNLID リポジトリ（公開）** でホストされています。  
他の Claude Code プロジェクトから以下の方法でインストール・参照できます。

---

## 方法 A：`.skill` ファイルで直接インストール（推奨）

**.skill ファイルは全リソースを含む完全なパッケージ**です。
エージェント定義（agents/）・評価スクリプト（scripts/）・ビューワー（eval-viewer/）がすべて同梱されています。

### インストールコマンド

```bash
claude install https://raw.githubusercontent.com/wyudao1996-collab/UNLID/main/skill-creator/skill-creator.skill
```

インストール後は `/skill-development` トリガーワードで自動発動します。

---

## 方法 B：ローカルに clone して参照

```bash
# リポジトリをクローン（または submodule として追加）
git clone https://github.com/wyudao1996-collab/UNLID.git /tmp/UNLID

# skill-creator だけコピー
cp -r /tmp/UNLID/skill-creator ~/.claude/skills/
```

---

## 方法 C：CLAUDE.md に URL 参照を追加

他プロジェクトの `CLAUDE.md` に以下を追記するだけで、
Claude がそのプロジェクトで作業するたびに Skill Creator の指示を自動読み込みします。

```markdown
## スキルクリエイター参照

@https://raw.githubusercontent.com/wyudao1996-collab/UNLID/main/skill-creator/SKILL.md
```

> **注意**：この方法では SKILL.md 本体のみ読み込まれます。
> `agents/`・`scripts/`・`eval-viewer/` などのバンドルリソースは
> 実行時に別途 URL 参照が必要です（方法 A の .skill インストールを推奨）。

---

## バンドルリソースの個別 URL（方法 C 使用時の補足参照先）

| ファイル | raw URL |
|---|---|
| SKILL.md（メイン） | `https://raw.githubusercontent.com/wyudao1996-collab/UNLID/main/skill-creator/SKILL.md` |
| agents/grader.md | `https://raw.githubusercontent.com/wyudao1996-collab/UNLID/main/skill-creator/agents/grader.md` |
| agents/analyzer.md | `https://raw.githubusercontent.com/wyudao1996-collab/UNLID/main/skill-creator/agents/analyzer.md` |
| agents/comparator.md | `https://raw.githubusercontent.com/wyudao1996-collab/UNLID/main/skill-creator/agents/comparator.md` |
| references/schemas.md | `https://raw.githubusercontent.com/wyudao1996-collab/UNLID/main/skill-creator/references/schemas.md` |

---

## バージョン固定（本番利用時の推奨）

最新 main を追随せずコミット SHA で固定したい場合：

```bash
# 例：特定コミットの .skill を指定
claude install https://raw.githubusercontent.com/wyudao1996-collab/UNLID/<COMMIT_SHA>/skill-creator/skill-creator.skill
```

現在の最新コミット SHA は以下で確認：

```bash
curl -s https://api.github.com/repos/wyudao1996-collab/UNLID/commits/main | jq '.sha'
```

---

## 更新方法

```bash
# 最新版に更新
claude install https://raw.githubusercontent.com/wyudao1996-collab/UNLID/main/skill-creator/skill-creator.skill --force
```
