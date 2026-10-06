import fs from "node:fs";
import path from "node:path";
import {
  resolveColor,
  SlideConfigError,
  type FillConfig,
  type ImageElement,
  type LineConfig,
  type ResolvedSlide,
  type ShapeElement,
  type SlideElement,
  type TextElement,
} from "./schema.js";
import { createPresentation, type LineProps, type PptxSlide } from "./pptx.js";
import { resolveShapeName } from "./shapes.js";

export interface RenderOptions {
  root: string;
  outFile: string;
  title?: string;
  author?: string;
  subject?: string;
}

type Palette = Record<string, string>;

function toFill(
  fill: FillConfig | undefined,
  palette: Palette,
): { color: string; transparency?: number } | undefined {
  if (fill === undefined) return undefined;
  if (typeof fill === "string") {
    const color = resolveColor(fill, palette);
    return color ? { color } : undefined;
  }
  const color = resolveColor(fill.color, palette);
  if (!color) return undefined;
  return { color, transparency: fill.transparency };
}

function toLine(
  line: LineConfig | undefined,
  palette: Palette,
  defaults?: {
    beginArrowType?: "none" | "arrow" | "diamond" | "oval" | "stealth" | "triangle";
    endArrowType?: "none" | "arrow" | "diamond" | "oval" | "stealth" | "triangle";
  },
): LineProps | undefined {
  const fromConfig: LineProps = {};
  if (typeof line === "string") {
    const color = resolveColor(line, palette);
    if (color) fromConfig.color = color;
  } else if (line) {
    const color = resolveColor(line.color, palette);
    if (color) fromConfig.color = color;
    if (line.width !== undefined) fromConfig.width = line.width;
    if (line.dashType) fromConfig.dashType = line.dashType;
    if (line.beginArrowType) fromConfig.beginArrowType = line.beginArrowType;
    if (line.endArrowType) fromConfig.endArrowType = line.endArrowType;
    if (line.transparency !== undefined) fromConfig.transparency = line.transparency;
  }

  const merged: LineProps = {
    ...defaults,
    ...fromConfig,
  };

  return Object.keys(merged).length > 0 ? merged : undefined;
}

function fontFace(el: { fontFace?: string }, slide: ResolvedSlide, heading = false): string | undefined {
  return el.fontFace ?? (heading ? slide.fonts.heading : undefined) ?? slide.fonts.default;
}

function addText(slideObj: PptxSlide, el: TextElement, slide: ResolvedSlide): void {
  const color = resolveColor(el.color, slide.colors, "text");
  const fill = toFill(el.fill, slide.colors);
  slideObj.addText(el.text ?? "", {
    x: el.x,
    y: el.y,
    w: el.w,
    h: el.h,
    fontFace: fontFace(el, slide, true),
    fontSize: el.fontSize,
    color,
    align: el.align,
    valign: el.valign,
    bold: el.bold,
    italic: el.italic,
    underline: el.underline,
    margin: el.margin ?? 0,
    fill,
  });
}

function addImage(
  slideObj: PptxSlide,
  el: ImageElement,
  source: string,
  root: string,
): void {
  const abs = path.resolve(root, el.src);
  if (!fs.existsSync(abs)) {
    throw new SlideConfigError(
      `Image not found: ${el.src} (resolved to ${abs}, referenced from ${source})`,
    );
  }
  if (!fs.statSync(abs).isFile()) {
    throw new SlideConfigError(`Image path is not a file: ${el.src} (referenced from ${source})`);
  }

  let sizing: { type: "contain" | "cover" | "crop"; w: number; h: number; x?: number; y?: number } | undefined;
  if (typeof el.sizing === "string") {
    const w = typeof el.w === "number" ? el.w : undefined;
    const h = typeof el.h === "number" ? el.h : undefined;
    sizing = { type: el.sizing, w: w ?? 1, h: h ?? 1 };
  } else if (el.sizing) {
    sizing = {
      type: el.sizing.type,
      w: el.sizing.w ?? (typeof el.w === "number" ? el.w : 1),
      h: el.sizing.h ?? (typeof el.h === "number" ? el.h : 1),
      x: el.sizing.x,
      y: el.sizing.y,
    };
  }

  slideObj.addImage({
    path: abs,
    x: el.x,
    y: el.y,
    w: el.w,
    h: el.h,
    sizing,
    rotate: el.rotate,
    rounding: el.rounding,
    transparency: el.transparency,
    altText: el.altText,
  });
}

function addShape(slideObj: PptxSlide, el: ShapeElement, slide: ResolvedSlide): void {
  const resolved = resolveShapeName(el.shape);
  const shapeType = resolved.type;
  const fill = toFill(el.fill, slide.colors);
  const line = toLine(el.line, slide.colors, resolved.lineDefaults);
  const color = resolveColor(el.color, slide.colors, "text");

  const common = {
    x: el.x,
    y: el.y,
    w: el.w,
    h: el.h,
    fill,
    line,
    rectRadius: el.rectRadius,
    rotate: el.rotate,
    flipH: el.flipH,
    flipV: el.flipV,
  };

  if (el.text !== undefined && el.text !== "") {
    slideObj.addText(el.text, {
      ...common,
      shape: shapeType,
      fontFace: fontFace(el, slide),
      fontSize: el.fontSize,
      color,
      align: el.align ?? "center",
      valign: el.valign ?? "middle",
      bold: el.bold,
      italic: el.italic,
      underline: el.underline,
      margin: el.margin ?? 4,
    });
    return;
  }

  slideObj.addShape(shapeType, common);
}

function renderElement(
  slideObj: PptxSlide,
  el: SlideElement,
  slide: ResolvedSlide,
  root: string,
): void {
  switch (el.type) {
    case "text":
      addText(slideObj, el, slide);
      return;
    case "image":
      addImage(slideObj, el, slide.source, root);
      return;
    case "shape":
      addShape(slideObj, el, slide);
      return;
  }
}

export async function renderDeck(
  slides: ResolvedSlide[],
  options: RenderOptions,
): Promise<string> {
  if (slides.length === 0) {
    throw new SlideConfigError("No slides to render.");
  }

  const pptx = createPresentation();
  const { width, height } = slides[0].size;
  pptx.defineLayout({ name: "CREATE_SLIDES", width, height });
  pptx.layout = "CREATE_SLIDES";

  if (options.title) pptx.title = options.title;
  if (options.author) pptx.author = options.author;
  if (options.subject) pptx.subject = options.subject;

  for (const slide of slides) {
    const slideObj = pptx.addSlide();
    const bg = resolveColor(slide.background, slide.colors, "background");
    if (bg) {
      slideObj.background = { color: bg };
    }
    if (slide.notes) {
      slideObj.addNotes(slide.notes);
    }
    for (const el of slide.elements) {
      renderElement(slideObj, el, slide, options.root);
    }
  }

  const outFile = path.resolve(options.outFile);
  fs.mkdirSync(path.dirname(outFile), { recursive: true });
  await pptx.writeFile({ fileName: outFile });
  return outFile;
}
