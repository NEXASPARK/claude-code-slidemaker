# Classmethod PPTX テンプレート

PptxGenJS 用テンプレートです。

## ファイル構成

| ファイル | 用途 |
|--------|------|
| `template.pptx` | 生成済みテンプレート（18レイアウト） |
| `theme.js` | テーマ定数（カラー、フォント、サイズ、レイアウト） |
| `generate-template.js` | テンプレート生成スクリプト |

## スライドレイアウト一覧

### レイアウト

| # | レイアウト名 | 説明 |
|---|-----------|------|
| 1 | **タイトルスライド** | ロゴ、タイトル、サブタイトル、日付（グレー背景） |
| 2 | **セクション区切り** | 青アクセントバー + セクションタイトル（グレー背景） |
| 3 | **標準コンテンツ** | ヘッダーバー + 見出し + 箇条書き |
| 4 | **ヘッダーなし** | フルコンテンツ（ヘッダーバーなし） |
| 5 | **画像中央** | ヘッダー + 中央に画像プレースホルダー |
| 6 | **テキスト＋図（右）** | 左にテキスト、右にグレー画像エリア |
| 7 | **テキスト＋図（左）** | 左にグレー画像エリア、右にテキスト |
| 8 | **2カラム** | 縦線で分割された2列レイアウト |
| 9 | **3カラム** | 縦線で分割された3列レイアウト |
| 10 | **主要指標** | 3枚のカードに大きな数字 + アイコン + ラベル |
| 11 | **プロセス・手順** | 4ステップのタイムライン（番号付き丸 + 接続線） |
| 12 | **比較（Before/After）** | 2枚のカードで比較表示 |
| 13 | **アイコングリッド** | 2×3 のアイコン + タイトル + 説明 |
| 14 | **引用・キーメッセージ** | 引用符 + 大きなテキスト + 発言者情報 |
| 15 | **アジェンダ** | 番号付き丸 + 項目リスト |
| 16 | **テーブル** | ヘッダーバー + データテーブル |
| 17 | **青アクセント** | 青背景に白テキスト（強調セクション） |
| 18 | **チーム紹介** | 4人分のメンバーカード（写真プレースホルダー + 名前 + 役職 + 紹介文） |
| 19 | **クロージング** | ロゴ + 挨拶 + 連絡先情報 |

## デザインシステム

### カラーパレット

| 用途 | 色名 | HEX |
|------|-----|-----|
| ブランドブルー | BLUE | `2C67E5` |
| ブランドレッド | RED | `DF3756` |
| メインテキスト | GRAY_MEDIUM | `434343` |
| サブテキスト | GRAY_LIGHT | `595959` |
| 標準背景 | BG_DEFAULT | `FFFFFF` |
| アクセント背景 | BG_ACCENT | `F3F3F3` |
| ボーダー | BORDER | `D9D9D9` |

### フォント・サイズ

| 要素 | フォント | サイズ |
|------|---------|-------|
| タイトル | Noto Sans JP Bold | 36pt |
| H2見出し | Noto Sans JP Bold | 28pt |
| H3見出し | Noto Sans JP Bold | 24pt |
| 本文 | Noto Sans JP | 16pt |
| 小テキスト | Noto Sans JP | 14pt |
| キャプション | Noto Sans JP | 12pt |
| 統計数字 | Noto Sans JP Bold | 60pt |

## 使い方

### 新規作成: PptxGenJS でゼロから作成

`theme.js` のテーマ定数を利用してスライドを作成します。グラフ、テーブル、自由なレイアウトなど柔軟に対応できます。

```javascript
const { COLORS, FONTS, SIZE, SLIDE, LAYOUT } = require("./template/theme");
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

const slide = pres.addSlide();
slide.addText("タイトル", {
  x: 0.5, y: 0.08, w: 9, h: 0.6,
  fontSize: SIZE.TITLE - 4, fontFace: FONTS.HEADING,
  color: COLORS.GRAY_MEDIUM, bold: true,
});
```

### 既存スライドの修正

PptxGenJS で作成したスライドの場合は、生成スクリプト（.js）を編集して再実行してください。

PptxGenJS 以外で作成された .pptx ファイルを修正する場合は、pptx スキルの XML 編集ワークフローを使用します。

```bash
# 1. 展開
python .claude/skills/pptx/scripts/office/unpack.py existing.pptx unpacked/

# 2. 各スライドの XML を編集

# 3. クリーン＆パック
python .claude/skills/pptx/scripts/clean.py unpacked/
python .claude/skills/pptx/scripts/office/pack.py unpacked/ output.pptx --original existing.pptx
```

## テンプレートの再生成

テーマを変更した場合：

```bash
node template/generate-template.js
```
