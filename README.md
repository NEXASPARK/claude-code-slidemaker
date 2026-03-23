# Claude Code Slidemaker

Claude Code のカスタムスラッシュコマンドで、ブランドテンプレートに沿った PPTX スライドを自動生成するプロジェクトです。

## 概要

2つのカスタムスラッシュコマンドで、テンプレート生成からスライド作成までを自動化します。

- **`/custom-template`** — 既存の .pptx を解析し、テーマ定数とテンプレート生成スクリプトを自動生成
- **`/create-slide`** — Markdown の原稿からブランドテンプレートに沿ったスライドを自動作成

## プロジェクト構成

```
claude-code-slidemaker/
├── .claude/
│   └── commands/
│       ├── custom-template.md    ← テンプレート生成コマンド
│       └── create-slide.md       ← スライド作成コマンド
├── input/
│   └── information.md            ← スライドに載せる情報（原稿）
├── template/
│   ├── input/                    ← 元となる .pptx を配置
│   ├── theme.js                  ← カラー・フォント・レイアウト定数
│   ├── generate-template.js      ← テンプレート生成スクリプト
│   └── template.pptx             ← 生成されたテンプレート
└── output/                       ← 完成したスライドの出力先
```

## 前提条件

- [Claude Code](https://docs.anthropic.com/en/docs/claude-code) がインストール済みであること
- [Anthropic 公式の pptx スキル](https://github.com/anthropics/skills/tree/main/skills/pptx) がインストール済みであること
- Node.js がインストール済みであること

## 使い方

### 1. テンプレートを作成する

既存の .pptx ファイルからテンプレートを自動生成します。

1. `template/input/` に元となる .pptx ファイルを配置
2. Claude Code で `/custom-template` を実行
3. `template/theme.js` と `template/generate-template.js` が自動更新される

### 2. スライドを作成する

Markdown の原稿からスライドを自動生成します。

1. `input/information.md` にスライドに載せたい情報を Markdown で記述
2. Claude Code で `/create-slide` を実行
3. `output/` ディレクトリにスライドが生成される

## スライドレイアウト

19種類のレイアウトを用意しており、コンテンツに応じて自動選択されます。

| 内容 | レイアウト |
|------|-----------|
| 表紙 | タイトルスライド |
| 章の切り替え | セクション区切り / 青アクセント |
| 箇条書き・説明 | 標準コンテンツ / ヘッダーなし |
| 図表 | 画像中央 / テキスト＋図 |
| 対比・並列 | 2カラム / 3カラム / 比較 |
| 数値・KPI | 主要指標 |
| 手順・フロー | プロセス・手順 |
| 機能・特徴一覧 | アイコングリッド |
| 引用・強調 | 引用・キーメッセージ |
| 目次 | アジェンダ |
| データ表 | テーブル |
| メンバー紹介 | チーム紹介 |
| 締め | クロージング |

## カスタマイズ

### テーマの変更

`template/theme.js` を編集してカラー・フォント・レイアウトを変更できます。変更後は以下を実行してテンプレートを再生成してください。

```bash
node template/generate-template.js
```

### 新しいテンプレートの適用

新しい .pptx を `template/input/` に配置して `/custom-template` を再実行すると、テーマとテンプレートが自動で更新されます。
