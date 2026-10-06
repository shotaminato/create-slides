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

スライドの読み込み順は **ファイル名のソート順** です（`01-title.yaml`, …）。

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
    02-open-isa.yaml
    03-isa-blocks.yaml
    04-timeline.yaml
    05-rva23-year.yaml
    06-market-soc.yaml
    07-market-ip.yaml
    08-conclusion.yaml
  assets/
    logo.png
    riscv-vs-arm-license-model.png
    riscv-isa-building-blocks.png
    riscv-timeline-2010-2026.png
    riscv-shd-soc-units-2022-2031.png
    riscv-shd-soc-revenue-2022-2031.png
    riscv-shd-ip-revenue-2022-2031.png
    riscv-slide-conclusion.png
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
- `background` — スライドマスターの背景色
- `fonts.default` / `fonts.heading` — 既定は `Yu Gothic UI`
- `lang` — テキストの言語タグ（サンプルは `ja-JP`）
- `colors` — 名前付きパレット。要素の `color` / `fill` から参照
- `title` / `author` / `subject` — プレゼンテーションのメタデータ（スライド見出しではない）
- `elements` — 共通クローム **または** スロット。`id` のない要素（または `master: true` / `role: chrome`）は **スライドマスター**（pptxgenjs `defineSlideMaster`）へ。`id` 付きはスロットとして各スライドにマージする（後述）
- `pageNumber` — フッター右の `current / total`。pptxgenjs の `slideNumber` は現在ページのフィールドだけなので、描画時に差し込みます。`text` の既定は `{{page}} / {{pages}}`（1 始まり）。別名: `{{currentPage}}` / `{{totalPages}}`。`align` の既定は `right`

サンプルの title / content / conclusion / default は、マスターに背景とフッター帯を置き、フッター左にラベル、右寄せで `1 / N` 形式のページ番号を描画します。**見出しテキストはマスターに焼き込みません**（スロット＋各スライド YAML）。

### スライド YAML

```yaml
template: content   # title / content / conclusion / default、またはパス
elements:
  - id: heading
    text: 見出し
```

- `template: default` は CLI の `--template` を使います
- 名前だけ (`content`) なら `templates/content.yaml` を探します
- パス (`templates/foo.yaml`) はプロジェクトルートからの相対です
- **スライド側の値がテンプレートを上書き**します（`size` / `background` / `fonts` / `colors`）。`background` をスライドで指定したときだけ、マスター背景をスライド側で上書きします
- テンプレートの **chrome**（`id` なし）はマスターへ。**スロット**（`id` あり）とスライドの `elements` はそのスライドだけ。`pageNumber` は描画時に各スライドへ載せます

色は `#0B6BCB` / `0B6BCB` / パレット名（`accent`）のいずれでも指定できます。座標 `x,y,w,h` の単位はインチです（pptxgenjs のパーセント文字列 `"50%"` も可）。

### スロット / Template slots

タイトルやまとめのように、位置・フォントサイズはテンプレート側に置き、スライドは `text` だけ変える、という書き方ができます。同じ `id` の要素は **テンプレートをベースにスライド側をディープマージ**（衝突したキーはスライド優先）し、結果は **スライド上** に描画します（マスターには焼き込みません。テキストがスライドごとに違うためです）。

```yaml
# templates/title.yaml（抜粋）
elements:
  - type: shape          # id なし → スライドマスター（クローム）
    shape: rect
    x: 0
    y: 0
    w: 0.16
    h: 7.5
    fill: accent
  - id: title            # スロット → 各スライドで text を上書き
    type: text
    x: 0.7
    y: 2.15
    w: 11.5
    h: 1.2
    fontSize: 40
    bold: true
    color: text
    wrap: false
  - id: subtitle
    type: text
    x: 0.7
    y: 4.05
    w: 12.0
    h: 0.5
    fontSize: 22
    color: accent
    wrap: false
```

```yaml
# slides/01-title.yaml
template: title
elements:
  - id: title
    text: RISC-V 入門＆市場動向（2026年10月）
  - id: subtitle
    text: 基礎からSHD予測まで／出典付き
```

まとめスライドも同様です。`templates/conclusion.yaml` の `id: statement` に `wrap: false` とバナー形状を置き、スライドは 1 行だけ渡します。`結言` のようなラベルはテンプレートにもスライドにも置きません。

```yaml
# slides/08-conclusion.yaml
template: conclusion
elements:
  - id: statement
    text: 2026年、RISC-VはRVA23実機とAIで成長期へ。ただしオープンISA≠無償シリコン、対Armはこれから。
```

マージ規則:

| テンプレート側 | スライド側 | 結果 |
| --- | --- | --- |
| `id` あり（かつ `master: true` / `role: chrome` ではない） | 同じ `id` | ディープマージして **スライド** に描画。スライドのキーが勝つ |
| `id` なし、または `master: true` / `role: chrome` | — | 今日どおり **スライドマスター**（ヘッダー帯・静的フッターなど） |
| スロット | 一致する `id` なし | テンプレートの既定のままスライドへ。`type: text` で `text` が空なら描画しない |
| — | 一致するテンプレート `id` なし | 従来どおりスライド末尾に追加（完全な要素である必要あり） |
| `pageNumber` | 任意 | 従来どおり描画時に `{{page}} / {{pages}}` を差し込み |

`fill` や `line` のような入れ子オブジェクトは中までマージします。プリミティブと配列は置き換えです。

本文テンプレートは `id: heading` と任意の `id: headingNote` を持っています。見出しの座標と 28pt はテンプレート側です。

### スライドマスター / Slide masters

pptxgenjs の `defineLayout`（サイズ）と `defineSlideMaster`（クローム）を使います。テンプレートごとに 1 つのマスターを作り、`addSlide({ masterName })` で参照します。ネイティブのマスターオブジェクトは `rect` / `line` / `text` / `image` です。`ellipse` や `roundRect` などは `{ text, options.shape }` 経由でマスターに載せます。

PowerPoint のスライド番号フィールドは現在ページだけなので、`n / total` はマスターではなく描画時のテキストです。

Title text stays in slide YAML and is not fixed on the master.

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
| `wrap` | テキスト折り返し。`false` で 1 行のまま（バナー向け） |
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
| テキスト系 | `fontSize`, `color`, `align`, `valign`, `bold`, `lang`, `wrap` など |

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

同梱のサンプルは、RISC-V の入門と SHD Group 2026 市場予測を出典付きでまとめた **8 枚**の日本語デッキです。表紙・本文・まとめはスロットで見出しを差し、図版 PNG を `assets/` から埋めます。

| スライド | テンプレート | 内容 |
| --- | --- | --- |
| `01-title.yaml` | `title` | 表紙。スロット `title` / `subtitle`。フッターは調査日 2026-10-06 |
| `02-open-isa.yaml` | `content` | オープンな ISA。ライセンス比較図 |
| `03-isa-blocks.yaml` | `content` | 基本 ISA＋拡張＋プロファイル |
| `04-timeline.yaml` | `content` | 2010→2026 マイルストーン |
| `05-rva23-year.yaml` | `content` | 2026年＝RVA23 ハードウェア元年（チャートなし） |
| `06-market-soc.yaml` | `content` | SoC 出荷数・売上（SHD Tables 13–14） |
| `07-market-ip.yaml` | `content` | CPU IP 売上 2024 $205M → 2031 $1.91B、CAGR 39.7% |
| `08-conclusion.yaml` | `conclusion` | 1 行まとめ（`wrap: false`）。「結言」ラベルなし |

生成:

```bash
npm install
npm run build-slides -- --slides slides --template templates/default.yaml --out dist/deck.pptx
```

成功すると `dist/deck.pptx` ができます（`dist/` は `.gitignore` 対象です）。同じコマンドで生成したサンプルを `examples/sample-deck.pptx` に同梱しています。

見出しの位置とサイズはテンプレートのスロット側です。本文カードは 18pt、出典フッターは 16pt、ページ番号は 18pt。1 行に収めたいテキストは `wrap: false` を付けます。

## 開発 / Development

```bash
npm run typecheck
npm run build
```
