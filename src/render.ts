import fs from "node:fs";
import path from "node:path";
import {
  resolveColor,
  resolveFillPaint,
  resolvePaint,
  lineTransparency,
  SlideConfigError,
  type FillConfig,
  type ImageElement,
  type LineConfig,
  type ResolvedSlide,
  type ShapeElement,
  type SlideElement,
  type SlideMasterDef,
  type TextElement,
} from "./schema.js";
import { createPresentation, type LineProps, type MasterObject, type PptxSlide } from "./pptx.js";
import { resolveShapeName } from "./shapes.js";

export interface RenderOptions {
  root: string;
  outFile: string;
  title?: string;
  author?: string;
  subject?: string;
  masters?: SlideMasterDef[];
}

type Palette = Record<string, string>;

interface PageContext {
  page: number;
  pages: number;
}

/** Substitute {{page}} / {{pages}} (and currentPage / totalPages aliases) at render time. */
export function applyPagePlaceholders(text: string, ctx: PageContext): string {
  return text
    .replaceAll("{{currentPage}}", String(ctx.page))
    .replaceAll("{{totalPages}}", String(ctx.pages))
    .replaceAll("{{page}}", String(ctx.page))
    .replaceAll("{{pages}}", String(ctx.pages));
}

function toFill(
  fill: FillConfig | undefined,
  palette: Palette,
): { color: string; transparency?: number } | undefined {
  if (fill === undefined) return undefined;
  return resolveFillPaint(fill, palette);
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
    const paint = resolvePaint(line, palette);
    if (paint) {
      fromConfig.color = paint.color;
      if (paint.transparency !== undefined) fromConfig.transparency = paint.transparency;
    }
  } else if (line) {
    if (line.color) {
      const paint = resolveFillPaint(
        { color: line.color, opacity: line.opacity, transparency: line.transparency },
        palette,
      );
      if (paint) {
        fromConfig.color = paint.color;
        if (paint.transparency !== undefined) fromConfig.transparency = paint.transparency;
      }
    } else {
      const t = lineTransparency(line);
      if (t !== undefined) fromConfig.transparency = t;
    }
    if (line.width !== undefined) fromConfig.width = line.width;
    if (line.dashType) fromConfig.dashType = line.dashType;
    if (line.beginArrowType) fromConfig.beginArrowType = line.beginArrowType;
    if (line.endArrowType) fromConfig.endArrowType = line.endArrowType;
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

function addText(slideObj: PptxSlide, el: TextElement, slide: ResolvedSlide, page: PageContext): void {
  const color = resolveColor(el.color, slide.colors, "text");
  const fill = toFill(el.fill, slide.colors);
  slideObj.addText(applyPagePlaceholders(el.text ?? "", page), {
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
    lang: el.lang ?? slide.lang,
  });
}

function addImage(
  slideObj: PptxSlide,
  el: ImageElement,
  source: string,
  root: string,
): void {
  const abs = resolveImagePath(el, source, root);

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

function addShape(slideObj: PptxSlide, el: ShapeElement, slide: ResolvedSlide, page: PageContext): void {
  const { type: shapeType, opts: common } = shapeCommonProps(el, slide.colors);
  const color = resolveColor(el.color, slide.colors, "text");

  if (el.text !== undefined && el.text !== "") {
    slideObj.addText(applyPagePlaceholders(el.text, page), {
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
      lang: el.lang ?? slide.lang,
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
  page: PageContext,
): void {
  switch (el.type) {
    case "text":
      addText(slideObj, el, slide, page);
      return;
    case "image":
      addImage(slideObj, el, slide.source, root);
      return;
    case "shape":
      addShape(slideObj, el, slide, page);
      return;
  }
}

function shapeCommonProps(
  el: ShapeElement,
  palette: Palette,
): { type: string; opts: Record<string, unknown> } {
  const resolved = resolveShapeName(el.shape);
  const fill = toFill(el.fill, palette);
  const line = toLine(el.line, palette, resolved.lineDefaults);
  const opts: Record<string, unknown> = {
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
  return { type: resolved.type, opts };
}

function resolveImagePath(el: ImageElement, source: string, root: string): string {
  const abs = path.resolve(root, el.src);
  if (!fs.existsSync(abs)) {
    throw new SlideConfigError(
      `Image not found: ${el.src} (resolved to ${abs}, referenced from ${source})`,
    );
  }
  if (!fs.statSync(abs).isFile()) {
    throw new SlideConfigError(`Image path is not a file: ${el.src} (referenced from ${source})`);
  }
  return abs;
}

/**
 * Map a YAML element to a pptxgenjs `defineSlideMaster` object.
 * Native master keys are only `rect` / `line` / `text` / `image`.
 * Other shapes (ellipse, roundRect, arrows, …) go through `{ text, options.shape }`.
 */
function elementToMasterObject(
  el: SlideElement,
  master: SlideMasterDef,
  root: string,
): MasterObject {
  const slideLike: ResolvedSlide = {
    source: master.source,
    masterName: master.name,
    size: { width: 0, height: 0 },
    fonts: master.fonts,
    colors: master.colors,
    lang: master.lang,
    elements: [],
  };

  if (el.type === "text") {
    const color = resolveColor(el.color, master.colors, "text");
    const fill = toFill(el.fill, master.colors);
    return {
      text: {
        text: el.text ?? "",
        options: {
          x: el.x,
          y: el.y,
          w: el.w,
          h: el.h,
          fontFace: fontFace(el, slideLike, true),
          fontSize: el.fontSize,
          color,
          align: el.align,
          valign: el.valign,
          bold: el.bold,
          italic: el.italic,
          underline: el.underline,
          margin: el.margin ?? 0,
          fill,
          lang: el.lang ?? master.lang,
        },
      },
    };
  }

  if (el.type === "image") {
    const abs = resolveImagePath(el, master.source, root);
    return {
      image: {
        path: abs,
        x: el.x,
        y: el.y,
        w: el.w,
        h: el.h,
        rotate: el.rotate,
        rounding: el.rounding,
        transparency: el.transparency,
        altText: el.altText,
      },
    };
  }

  const { type, opts } = shapeCommonProps(el, master.colors);
  const color = resolveColor(el.color, master.colors, "text");
  const hasText = el.text !== undefined && el.text !== "";
  if (!hasText && type === "rect") {
    return { rect: opts };
  }
  if (!hasText && type === "line") {
    return { line: opts };
  }
  return {
    text: {
      text: el.text ?? "",
      options: {
        ...opts,
        shape: type,
        fontFace: fontFace(el, slideLike),
        fontSize: el.fontSize,
        color,
        align: el.align ?? "center",
        valign: el.valign ?? "middle",
        bold: el.bold,
        italic: el.italic,
        underline: el.underline,
        margin: el.margin ?? 4,
        lang: el.lang ?? master.lang,
      },
    },
  };
}

function defineMasters(
  pptx: ReturnType<typeof createPresentation>,
  masters: SlideMasterDef[],
  root: string,
): void {
  const names = new Set<string>();
  for (const master of masters) {
    if (names.has(master.name)) {
      throw new SlideConfigError(`Duplicate slide master name: ${master.name}`);
    }
    names.add(master.name);
    const bg = resolveColor(master.background, master.colors, "background");
    pptx.defineSlideMaster({
      title: master.name,
      background: bg ? { color: bg } : undefined,
      objects: master.elements.map((el) => elementToMasterObject(el, master, root)),
    });
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

  const masters = options.masters ?? [];
  defineMasters(pptx, masters, options.root);
  const masterNames = new Set(masters.map((m) => m.name));

  const pages = slides.length;
  for (let i = 0; i < slides.length; i += 1) {
    const slide = slides[i];
    const page = { page: i + 1, pages };
    if (slide.masterName && !masterNames.has(slide.masterName)) {
      throw new SlideConfigError(
        `Unknown slide master "${slide.masterName}" (from ${slide.source})`,
      );
    }
    const slideObj = pptx.addSlide(
      slide.masterName ? { masterName: slide.masterName } : undefined,
    );
    const bg = resolveColor(slide.backgroundOverride, slide.colors, "background");
    if (slide.backgroundOverride && bg) {
      slideObj.background = { color: bg };
    }
    if (slide.notes) {
      slideObj.addNotes(slide.notes);
    }
    for (const el of slide.elements) {
      renderElement(slideObj, el, slide, options.root, page);
    }
  }

  const outFile = path.resolve(options.outFile);
  fs.mkdirSync(path.dirname(outFile), { recursive: true });
  await pptx.writeFile({ fileName: outFile });
  return outFile;
}
