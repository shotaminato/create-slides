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

スライドの読み込み順は **ファイル名のソート順** です（`01-title.yaml`, `02-shapes.yaml`, …）。

存在しないファイル、壊れた YAML、スキーマ違反、欠けている画像は、パス付きのエラーで終了します。

## ディレクトリ構成 / Layout

```
create-slides/
  templates/
    default.yaml
  slides/
    01-title.yaml
    02-shapes.yaml
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

### テンプレート (`templates/default.yaml`)

既定値を定義します。

- `size.width` / `size.height` — インチ。未指定時はワイドスクリーン **13.333 × 7.5**
- `background` — スライド背景色（`#RRGGBB` またはパレット名）
- `fonts.default` / `fonts.heading`
- `colors` — 名前付きパレット。要素の `color` / `fill` から参照できる
- `title` / `author` / `subject` — プレゼンテーションのメタデータ
- `elements` — 全スライドに先に描画される共有要素（フッターなど）

### スライド YAML

```yaml
template: default   # または other / templates/other.yaml
background: "0F172A"
elements:
  - type: text
    text: Hello
    x: 0.5
    y: 1
    w: 12
    h: 1
    fontSize: 32
    color: text
    align: center
```

- `template: default` は CLI の `--template` を使います
- 名前だけ (`theme`) なら `templates/theme.yaml` を探します
- パス (`templates/foo.yaml`) はプロジェクトルートからの相対です
- **スライド側の値がテンプレートを上書き**します（`size` / `background` / `fonts` / `colors`）
- `elements` は **テンプレート要素 + スライド要素**（テンプレートが下、スライドが上）

色は `#38BDF8` / `38BDF8` / パレット名（`accent`）のいずれでも指定できます。座標 `x,y,w,h` の単位はインチです（pptxgenjs のパーセント文字列 `"50%"` も可）。

### `type: text`

| フィールド | 内容 |
| --- | --- |
| `text` | 文字列（`\n` で改行） |
| `x,y,w,h` | 位置とサイズ |
| `fontSize` | pt |
| `fontFace` | 未指定時はテンプレートの `fonts` |
| `color` | 文字色 |
| `align` | `left` / `center` / `right` / `justify` |
| `valign` | `top` / `middle` / `bottom` |
| `bold` / `italic` / `underline` | boolean |
| `fill` | テキストボックス背景 |
| `margin` | 余白 |

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
| `fill` | 色文字列、または `{ color, transparency }` |
| `line` | 色文字列、または `{ color, width, dashType, beginArrowType, endArrowType }` |
| `text` | 図形内テキスト（pptxgenjs の `addText` + `shape`） |
| `rectRadius` | `roundRect` の丸み（0–1） |
| `rotate` / `flipH` / `flipV` | 任意 |
| テキスト系 | `fontSize`, `color`, `align`, `valign`, `bold` など |

`line.beginArrowType` / `endArrowType` は `none` / `arrow` / `diamond` / `oval` / `stealth` / `triangle` です。

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

同梱データ:

- `templates/default.yaml` — ワイドスクリーン、色、フッター
- `slides/01-title.yaml` — テキスト + 画像
- `slides/02-shapes.yaml` — rect / roundRect / ellipse / triangle / 矢印
- `assets/logo.png`

生成:

```bash
npm install
npm run build-slides -- --slides slides --template templates/default.yaml --out dist/deck.pptx
```

成功すると `dist/deck.pptx` ができます（`dist/` は `.gitignore` 対象です）。同じコマンドで生成したサンプルを `examples/sample-deck.pptx` に同梱しています。

## 開発 / Development

```bash
npm run typecheck
npm run build
```
