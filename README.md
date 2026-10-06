# create-slides

YAML のスライド設定から PowerPoint (`.pptx`) を生成する TypeScript / Node.js CLI です。  
1 つの YAML ファイル = 1 枚のスライド。共通テンプレートを参照し、画像はプロジェクト相対パスで指定します。

A config-driven PPTX generator: one YAML file per slide, optional shared template, project-relative images.

## 必要環境 / Requirements

- Node.js 18+
- npm

## インストール / Install

```bash
npm install
```

コンパイル済みバイナリを使う場合:

```bash
npm run build
npx create-slides --help
```

開発時は `tsx` 経由の `npm run build-slides` が簡単です。

## 使い方 / Usage

```bash
npm run build-slides -- --slides slides --template templates/default.yaml --out dist/deck.pptx
```

同等の呼び出し:

```bash
npx tsx src/cli.ts --slides slides --template templates/default.yaml --out dist/deck.pptx
```

| オプション | 既定値 | 説明 |
| --- | --- | --- |
| `--slides <dir>` | `slides` | スライド YAML が入ったディレクトリ |
| `--template <file>` | `templates/default.yaml` | `template: default` のときに使うテンプレート |
| `--out <file>` | `dist/deck.pptx` | 出力 `.pptx` |
| `--root <dir>` | cwd（`package.json` を上方向に探索） | プロジェクトルート。画像パスの基準 |

スライドの読み込み順は **ファイル名のソート順** です（`01-title.yaml`, `02-overview.yaml`, …）。

存在しないファイル、壊れた YAML、スキーマ違反、欠けている画像は、パス付きのエラーで終了します。

## ディレクトリ構成 / Layout

```
create-slides/
  templates/
    default.yaml      # 共通の色・フォント（フォールバック）
    title.yaml        # 表紙
    content.yaml      # 本文
    conclusion.yaml   # まとめ / クロージング
  slides/
    01-title.yaml
    02-overview.yaml
    …
    08-conclusion.yaml
  assets/
    logo.png
  examples/
    sample-deck.pptx
  src/
    cli.ts
    schema.ts
    load.ts
    render.ts
    shapes.ts
    pptx.ts
  package.json
  tsconfig.json
  README.md
```

## 設定モデル / Config model

### テンプレート

既定フォントは **Yu Gothic UI**（`fontFace` も同じ文字列）、背景は白〜クリームの明るい色です。

| ファイル | `template:` | 用途 |
| --- | --- | --- |
| `templates/default.yaml` | `default` | サイズ・パレット・フッター。CLI の `--template` 既定 |
| `templates/title.yaml` | `title` | 表紙（左アクセント＋フッター） |
| `templates/content.yaml` | `content` | 本文（ヘッダー帯＋フッター）。見出しは各スライド側 |
| `templates/conclusion.yaml` | `conclusion` | まとめ |

各テンプレートで定義できる主な項目:

- `size.width` / `size.height` — インチ。未指定時はワイドスクリーン **13.333 × 7.5**
- `background` — スライド背景色
- `fonts.default` / `fonts.heading` — 既定は `Yu Gothic UI`
- `lang` — テキストの言語タグ（サンプルは `ja-JP`）
- `colors` — 名前付きパレット。要素の `color` / `fill` から参照
- `title` / `author` / `subject` — プレゼンテーションのメタデータ
- `elements` — そのテンプレートを使うスライドに先に描画される共有要素
- フッター右のページ番号 — テキストに `{{page}} / {{pages}}` を書くと、描画時に「現在ページ / 総ページ」（1 始まり）へ置換されます。別名: `{{currentPage}}` / `{{totalPages}}`

サンプルの title / content / conclusion / default は、フッター左にデッキ名、右寄せで `1 / 8` 形式のページ番号を置きます。

### スライド YAML

```yaml
template: content   # title / content / conclusion / default、またはパス
elements:
  - type: text
    text: 見出し
    x: 0.45
    y: 0.06
    w: 12.4
    h: 0.5
    fontSize: 26
    bold: true
    color: text
```

- `template: default` は CLI の `--template` を使います
- 名前だけ (`content`) なら `templates/content.yaml` を探します
- パス (`templates/foo.yaml`) はプロジェクトルートからの相対です
- **スライド側の値がテンプレートを上書き**します（`size` / `background` / `fonts` / `colors`）
- `elements` は **テンプレート要素 + スライド要素**（テンプレートが下、スライドが上）

色は `#0B6BCB` / `0B6BCB` / パレット名（`accent`）のいずれでも指定できます。座標 `x,y,w,h` の単位はインチです（pptxgenjs のパーセント文字列 `"50%"` も可）。

### `type: text`

| フィールド | 内容 |
| --- | --- |
| `text` | 文字列（`\n` で改行） |
| `x,y,w,h` | 位置とサイズ |
| `fontSize` | pt |
| `fontFace` | 未指定時はテンプレートの `fonts`（既定 `Yu Gothic UI`） |
| `color` | 文字色 |
| `align` | `left` / `center` / `right` / `justify` |
| `valign` | `top` / `middle` / `bottom` |
| `bold` / `italic` / `underline` | boolean |
| `fill` | テキストボックス背景（図形と同じ透明度の指定が可能） |
| `margin` | 余白 |
| `lang` | 未指定時はテンプレートの `lang` |

### `type: image`

| フィールド | 内容 |
| --- | --- |
| `src` | プロジェクトルート相対パス（例: `assets/logo.png`） |
| `x,y,w,h` | 位置とサイズ |
| `sizing` | `contain` / `cover` / `crop`、または `{ type, w, h, x, y }` |
| `rotate` / `rounding` / `transparency` / `altText` | 任意 |

### `type: shape`

| フィールド | 内容 |
| --- | --- |
| `shape` | 下表のフレンドリー名、または pptxgenjs の `ShapeType` 名 |
| `x,y,w,h` | 位置とサイズ |
| `fill` | 色文字列（`#RRGGBB` / `#RRGGBBAA`）、または `{ color, opacity?, transparency? }` |
| `line` | 色文字列、または `{ color, width, dashType, beginArrowType, endArrowType, opacity?, transparency? }` |
| `text` | 図形内テキスト（pptxgenjs の `addText` + `shape`） |
| `rectRadius` | `roundRect` の丸み（0–1） |
| `rotate` / `flipH` / `flipV` | 任意 |
| テキスト系 | `fontSize`, `color`, `align`, `valign`, `bold`, `lang` など |

`line.beginArrowType` / `endArrowType` は `none` / `arrow` / `diamond` / `oval` / `stealth` / `triangle` です。

### 透明度 / Transparency

図形の `fill` と `line`（テキストボックスの `fill` も同じ）で透明度を指定できます。省略時は不透明です。`opacity` と `transparency` は同時に指定できません。

| 書き方 | 意味 |
| --- | --- |
| `opacity: 0.4` | 不透明度 0–1（1 = 不透明、0 = 透明） |
| `transparency: 60` | pptxgenjs のパーセント（0 = 不透明、100 = 透明） |
| `#0B6BCB80` | 8 桁 HEX。末尾 2 桁がアルファ（`00`–`FF`） |

```yaml
- type: shape
  shape: ellipse
  x: 1
  y: 1
  w: 2
  h: 2
  fill:
    color: accent
    opacity: 0.4
  line:
    color: "#0B6BCBAA"
    width: 1.5
```

## シェイプ一覧 / Shapes

フレンドリー名は大文字小文字・空白・ハイフンを無視してマッチします（`rounded rectangle` = `roundRect`）。

| YAML の `shape` | pptxgenjs `ShapeType` |
| --- | --- |
| `rect`, `rectangle` | `rect` |
| `roundRect`, `roundedRect`, `rounded rectangle` | `roundRect` |
| `ellipse`, `oval`, `circle` | `ellipse` |
| `line` | `line` |
| `lineArrow` | `line` + `endArrowType: arrow` |
| `lineDoubleArrow` | `line` + 両端矢印 |
| `rightArrow`, `arrow`, `arrow-right` | `rightArrow` |
| `leftArrow`, `arrow-left` | `leftArrow` |
| `upArrow`, `arrow-up` | `upArrow` |
| `downArrow`, `arrow-down` | `downArrow` |
| `leftRightArrow` | `leftRightArrow` |
| `upDownArrow` | `upDownArrow` |
| `notchedRightArrow` | `notchedRightArrow` |
| `stripedRightArrow` | `stripedRightArrow` |
| `chevron` | `chevron` |
| `triangle`, `isosceles`, `isoscelesTriangle` | `triangle`（二等辺三角形） |
| `rtTriangle`, `rightTriangle` | `rtTriangle`（直角三角形） |
| `diamond` | `diamond` |

上記以外でも、[pptxgenjs `ShapeType`](https://github.com/gitbrent/PptxGenJS/blob/master/src/core-enums.ts) に存在する名前（`hexagon`, `star5`, `flowChartDecision` など）はそのまま使えます。存在しない名前はエラーになります。

三角形: pptxgenjs の公式名は `triangle`（二等辺）と `rtTriangle`（直角）です。`isoscelesTriangle` という ShapeType はなく、旧 enum `ISOSCELES_TRIANGLE` が `triangle` に対応します。

ライン矢印は `shape: line` に `line.endArrowType: arrow` を付けるか、`lineArrow` エイリアスを使います。ブロック矢印は `rightArrow` 系の ShapeType です。

## サンプル / Sample

同梱のサンプルは、表紙 → 本文複数枚 → まとめ、の RISC-V 解説デッキ（日本語）です。本文図に rect / roundRect / ellipse / 矢印 / triangle を使っています。

| スライド | テンプレート | 内容 |
| --- | --- | --- |
| `01-title.yaml` | `title` | 表紙 |
| `02-overview.yaml` | `content` | RISC-V とは |
| `03-principles.yaml` | `content` | 設計思想 |
| `04-base-isa.yaml` | `content` | ベース整数 ISA とレジスタ |
| `05-formats.yaml` | `content` | 命令フォーマット |
| `06-extensions.yaml` | `content` | 標準拡張 |
| `07-privilege.yaml` | `content` | 特権レベル |
| `08-conclusion.yaml` | `conclusion` | まとめ |

生成:

```bash
npm install
npm run build-slides -- --slides slides --template templates/default.yaml --out dist/deck.pptx
```

成功すると `dist/deck.pptx` ができます（`dist/` は `.gitignore` 対象です）。同じコマンドで生成したサンプルを `examples/sample-deck.pptx` に同梱しています。

本文の見出しは content テンプレートのヘッダー帯の上に、各 YAML で置いてください（例: `y: 0.06`, `h: 0.5`）。

## 開発 / Development

```bash
npm run typecheck
npm run build
```
