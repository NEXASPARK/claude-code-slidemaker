---
description: template/input/ の .pptx を解析して generate-template.js を更新する
allowed-tools: [Read, Glob, Grep, Bash, Write, Edit, Agent]
---

# カスタムテンプレート生成

`template/input/` に配置された .pptx ファイルを XML に展開・解析し、`template/generate-template.js` と `template/theme.js` を更新します。

## 処理手順

### ステップ1: 入力ファイルの確認

1. `template/input/` ディレクトリ内の .pptx ファイルを探す
2. .pptx ファイルが見つからない場合は、ユーザーに `template/input/` に .pptx を配置するよう案内して終了する
3. 複数ある場合は一覧を表示し、どれを使うか確認する

### ステップ2: XML に展開

1. 選択された .pptx ファイルを unzip で展開する
   ```
   mkdir -p template/input/unpacked
   unzip -o template/input/{ファイル名}.pptx -d template/input/unpacked
   ```

### ステップ3: XML を解析

以下のファイルを読み取り、デザイン情報とスライド構造を解析する。

**テーマ情報（並列で読み取る）:**
- `ppt/theme/theme1.xml` — カラーパレット、フォント定義
- `ppt/theme/theme2.xml` — サブテーマ（存在する場合）
- `ppt/presentation.xml` — スライドサイズ
- `ppt/slideMasters/slideMaster1.xml` — マスタースライド

**全スライド（並列で読み取る）:**
- `ppt/slides/slide1.xml` 〜 全スライド
- `ppt/slides/_rels/` 以下の .rels ファイル（画像参照の確認）

各スライドから以下を抽出する:
- レイアウトタイプ（タイトル、コンテンツ、セクション区切り等）
- テキスト内容、フォント、サイズ（EMU → pt 変換: ÷ 12700）、色、太字/斜体
- 図形の位置（EMU → inches 変換: ÷ 914400）、サイズ、塗り色、線色
- 画像参照
- 背景色

### ステップ4: theme.js を更新

解析結果をもとに `template/theme.js` を更新する。

抽出すべき定数:
- **COLORS**: スライド内で使われている全色（ブランドカラー、テキスト色、背景色、ボーダー色）
- **FONTS**: 使用フォントファミリー（見出し用、本文用、コード用）
- **SIZE**: フォントサイズ（タイトル、H2、H3、本文、小テキスト、キャプション、注釈、統計数字）
- **SLIDE**: スライド寸法（W, H）
- **LAYOUT**: マージン、ヘッダー高さ、コンテンツ領域の位置・サイズ

既存の `theme.js` のエクスポート構造（`COLORS`, `FONTS`, `SIZE`, `SLIDE`, `LAYOUT`, `makeShadow`, `makeCardShadow`）は維持すること。

### ステップ5: generate-template.js を更新

解析した全スライドの構造を PptxGenJS のコードとして `template/generate-template.js` に反映する。

**守るべきルール:**
- 既存の `sample-generate-template.js` のコード構造（ヘルパー関数 `addHeaderBar`, `addPageNumber`, `addFooterNote`、アイコン処理等）を踏襲する
- `theme.js` の定数を `require("./theme")` で読み込んで使用する
- スライド内のテキストは `XXXX` プレフィックス付きのプレースホルダーにする（例: `"XXXX スライドタイトル"`）
- 出力先は `path.join(__dirname, "template.pptx")`
- ロゴが SVG 埋め込みの場合はそのまま再現、テキストベースの場合はテキストとして再現

### ステップ6: テンプレートを生成して確認

1. `node template/generate-template.js` を実行する
2. エラーがあれば修正して再実行する
3. 成功したら `template/template.pptx` が生成されたことを報告する

### ステップ7: クリーンアップ

1. `template/input/unpacked/` ディレクトリを削除する
   ```
   rm -rf template/input/unpacked
   ```

## 完了報告

処理完了後、以下を報告する:
- 検出されたスライド枚数とレイアウト構成
- theme.js の変更点（色、フォント等の差分）
- generate-template.js の変更点
- template.pptx の生成結果
