import { z } from "zod";

/** Inches (number) or pptxgenjs percent strings such as "50%". */
const coord = z.union(
  [
    z.number(),
    z.string().regex(/^-?\d+(\.\d+)?%?$/, "Expected inches (number) or a percent string like 50%"),
  ],
  {
    errorMap: (issue, ctx) => {
      if (issue.code === "invalid_union") {
        return { message: "Required: inches (number) or a percent string like 50%" };
      }
      return { message: ctx.defaultError };
    },
  },
);

const hexOrName = z.string().min(1);
const opacityUnit = z.number().min(0).max(1);
const transparencyPercent = z.number().min(0).max(100);

const opacityXorTransparency = {
  opacity: opacityUnit.optional(),
  transparency: transparencyPercent.optional(),
};

function rejectOpacityAndTransparency(
  val: { opacity?: number; transparency?: number },
  ctx: z.RefinementCtx,
): void {
  if (val.opacity !== undefined && val.transparency !== undefined) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "Use either opacity (0–1) or transparency (0–100), not both",
      path: ["opacity"],
    });
  }
}

const fillObject = z
  .object({
    color: hexOrName,
    ...opacityXorTransparency,
  })
  .superRefine(rejectOpacityAndTransparency);

export const fillSchema = z.union([hexOrName, fillObject]);

const dashType = z.enum([
  "solid",
  "dash",
  "dashDot",
  "lgDash",
  "lgDashDot",
  "lgDashDotDot",
  "sysDash",
  "sysDot",
]);

const arrowHead = z.enum(["none", "arrow", "diamond", "oval", "stealth", "triangle"]);

const lineObject = z
  .object({
    color: hexOrName.optional(),
    width: z.number().min(0).optional(),
    dashType: dashType.optional(),
    beginArrowType: arrowHead.optional(),
    endArrowType: arrowHead.optional(),
    ...opacityXorTransparency,
  })
  .superRefine(rejectOpacityAndTransparency);

export const lineSchema = z.union([hexOrName, lineObject]);

const alignH = z.enum(["left", "center", "right", "justify"]);
const alignV = z.enum(["top", "middle", "bottom"]);

const box = z.object({
  x: coord,
  y: coord,
  w: coord.optional(),
  h: coord.optional(),
});

const textStyle = z.object({
  text: z.string().optional(),
  fontFace: z.string().optional(),
  fontSize: z.number().positive().optional(),
  color: hexOrName.optional(),
  align: alignH.optional(),
  valign: alignV.optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  margin: z.number().optional(),
  lang: z.string().optional(),
  /** pptxgenjs `wrap`. false keeps a one-liner from wrapping. */
  wrap: z.boolean().optional(),
});

const slotRole = z.enum(["chrome", "slot"]);

const slotMeta = {
  /** Same id on a slide deep-merges onto this element and draws on the slide (not the master). */
  id: z.string().min(1).optional(),
  /** Force this element onto the slide master even if it has an `id`. */
  master: z.boolean().optional(),
  /** `chrome` = master; `slot` = per-slide (requires `id`). */
  role: slotRole.optional(),
};

export const textElementSchema = box.merge(textStyle).extend({
  type: z.literal("text"),
  fill: fillSchema.optional(),
  ...slotMeta,
});

const imageSizingObject = z.object({
  type: z.enum(["contain", "cover", "crop"]),
  w: z.number().optional(),
  h: z.number().optional(),
  x: z.number().optional(),
  y: z.number().optional(),
});

export const imageElementSchema = box.extend({
  type: z.literal("image"),
  src: z.string().min(1),
  sizing: z.union([z.enum(["contain", "cover", "crop"]), imageSizingObject]).optional(),
  rotate: z.number().optional(),
  rounding: z.boolean().optional(),
  transparency: z.number().min(0).max(100).optional(),
  altText: z.string().optional(),
  ...slotMeta,
});

export const shapeElementSchema = box.merge(textStyle).extend({
  type: z.literal("shape"),
  /** Friendly alias (rect, triangle, isoscelesTriangle, …) or a pptxgenjs ShapeType name. */
  shape: z.string().min(1),
  fill: fillSchema.optional(),
  line: lineSchema.optional(),
  rectRadius: z.number().min(0).max(1).optional(),
  rotate: z.number().optional(),
  flipH: z.boolean().optional(),
  flipV: z.boolean().optional(),
  ...slotMeta,
});

export const elementSchema = z.discriminatedUnion("type", [
  textElementSchema,
  imageElementSchema,
  shapeElementSchema,
]);

/**
 * Per-slide slot override. `id` is required; geometry/type come from the matching
 * template element. Slide keys win on conflict. Also used for extra complete elements
 * that happen to include an `id`.
 */
export const slotOverrideSchema = z.object({
  id: z.string().min(1),
  type: z.enum(["text", "image", "shape"]).optional(),
  master: z.boolean().optional(),
  role: slotRole.optional(),
  x: coord.optional(),
  y: coord.optional(),
  w: coord.optional(),
  h: coord.optional(),
  text: z.string().optional(),
  fontFace: z.string().optional(),
  fontSize: z.number().positive().optional(),
  color: hexOrName.optional(),
  align: alignH.optional(),
  valign: alignV.optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  margin: z.number().optional(),
  lang: z.string().optional(),
  wrap: z.boolean().optional(),
  fill: fillSchema.optional(),
  src: z.string().min(1).optional(),
  sizing: z.union([z.enum(["contain", "cover", "crop"]), imageSizingObject]).optional(),
  rotate: z.number().optional(),
  rounding: z.boolean().optional(),
  transparency: z.number().min(0).max(100).optional(),
  altText: z.string().optional(),
  shape: z.string().min(1).optional(),
  line: lineSchema.optional(),
  rectRadius: z.number().min(0).max(1).optional(),
  flipH: z.boolean().optional(),
  flipV: z.boolean().optional(),
});

/** Full element or a partial override of a template slot (`id` + fields to change). */
export const slideElementSchema = z.union([elementSchema, slotOverrideSchema]);

export const fontsSchema = z.object({
  default: z.string().optional(),
  heading: z.string().optional(),
});

export const sizeSchema = z.object({
  width: z.number().positive(),
  height: z.number().positive(),
});

/**
 * Footer page number overlay. Injected per slide at render time because
 * pptxgenjs `slideNumber` is a PowerPoint current-page field only (not `n / total`).
 * Default text is `{{page}} / {{pages}}`.
 */
export const pageNumberSchema = box.merge(textStyle);

const themeBaseSchema = z.object({
  size: sizeSchema.optional(),
  background: hexOrName.optional(),
  fonts: fontsSchema.optional(),
  colors: z.record(z.string(), z.string()).optional(),
  title: z.string().optional(),
  author: z.string().optional(),
  subject: z.string().optional(),
  lang: z.string().optional(),
  pageNumber: pageNumberSchema.optional(),
});

export const templateSchema = themeBaseSchema.extend({
  /**
   * Template elements. Those with `id` (and not `master: true` / `role: chrome`)
   * are slots merged onto each slide. The rest are slide-master chrome.
   */
  elements: z.array(elementSchema).optional(),
});

/** Alias: template-shaped theme fields (full elements only). */
export const themeFieldsSchema = templateSchema;

export const slideFileSchema = themeBaseSchema.extend({
  template: z.string().optional(),
  notes: z.string().optional(),
  /** Complete elements and/or `{ id, ... }` overrides of template slots. */
  elements: z.array(slideElementSchema).optional(),
});

export type Coord = z.infer<typeof coord>;
export type FillConfig = z.infer<typeof fillSchema>;
export type LineConfig = z.infer<typeof lineSchema>;
export type TextElement = z.infer<typeof textElementSchema>;
export type ImageElement = z.infer<typeof imageElementSchema>;
export type ShapeElement = z.infer<typeof shapeElementSchema>;
export type SlideElement = z.infer<typeof elementSchema>;
export type SlotOverride = z.infer<typeof slotOverrideSchema>;
export type SlideInputElement = z.infer<typeof slideElementSchema>;
export type TemplateConfig = z.infer<typeof templateSchema>;
export type SlideFileConfig = z.infer<typeof slideFileSchema>;
export type PageNumberConfig = z.infer<typeof pageNumberSchema>;

export function isMasterChrome(el: {
  id?: string;
  master?: boolean;
  role?: "chrome" | "slot";
}): boolean {
  if (el.master === true) return true;
  if (el.role === "chrome") return true;
  if (el.role === "slot") return false;
  return !el.id;
}

export interface SlideMasterDef {
  /** pptxgenjs `defineSlideMaster` title / `addSlide({ masterName })`. */
  name: string;
  source: string;
  background?: string;
  fonts: { default?: string; heading?: string };
  colors: Record<string, string>;
  lang?: string;
  elements: SlideElement[];
}

export interface ResolvedSlide {
  source: string;
  masterName: string;
  size: { width: number; height: number };
  /** Slide-YAML background override. Master background is used when omitted. */
  backgroundOverride?: string;
  fonts: { default?: string; heading?: string };
  colors: Record<string, string>;
  title?: string;
  author?: string;
  subject?: string;
  lang?: string;
  notes?: string;
  elements: SlideElement[];
}

/** Matches `{{page}}` / `{{pages}}` / `{{currentPage}}` / `{{totalPages}}`. */
export const PAGE_PLACEHOLDER_RE = /\{\{(?:page|pages|currentPage|totalPages)\}\}/;

export function elementHasPagePlaceholder(el: SlideElement): boolean {
  if (el.type === "text" || el.type === "shape") {
    return PAGE_PLACEHOLDER_RE.test(el.text ?? "");
  }
  return false;
}

export function pageNumberToElement(cfg: PageNumberConfig): TextElement {
  return {
    type: "text",
    text: cfg.text && cfg.text.length > 0 ? cfg.text : "{{page}} / {{pages}}",
    x: cfg.x,
    y: cfg.y,
    w: cfg.w,
    h: cfg.h,
    fontFace: cfg.fontFace,
    fontSize: cfg.fontSize,
    color: cfg.color,
    align: cfg.align ?? "right",
    valign: cfg.valign ?? "middle",
    bold: cfg.bold,
    italic: cfg.italic,
    underline: cfg.underline,
    margin: cfg.margin,
    lang: cfg.lang,
    wrap: cfg.wrap,
  };
}

export const DEFAULT_SIZE = { width: 13.333, height: 7.5 } as const;

export class SlideConfigError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SlideConfigError";
  }
}

export function formatZodError(error: z.ZodError, file: string): string {
  const lines = error.issues.map((issue) => {
    const path = issue.path.length > 0 ? issue.path.join(".") : "(root)";
    return `  - ${path}: ${issue.message}`;
  });
  return `Schema validation failed in ${file}:\n${lines.join("\n")}`;
}

export interface ParsedPaint {
  /** 6-digit RGB hex, no `#` */
  color: string;
  /** pptxgenjs transparency percent 0–100. Omitted when fully opaque. */
  transparency?: number;
}

/** 0 = invisible, 1 = opaque → pptxgenjs transparency (0 = opaque, 100 = invisible). */
export function opacityToTransparency(opacity: number): number | undefined {
  const t = Math.round((1 - opacity) * 100);
  if (t <= 0) return undefined;
  return Math.min(100, t);
}

function alphaByteToTransparency(alpha: number): number | undefined {
  return opacityToTransparency(alpha / 255);
}

/**
 * Parse `#RGB`, `#RRGGBB`, `#RRGGBBAA` (and 4-digit `#RGBA`).
 * Named tokens that are not hex are left to the caller.
 */
export function parseHexPaint(value: string): ParsedPaint | undefined {
  const hex = value.trim().startsWith("#") ? value.trim().slice(1) : value.trim();
  if (/^[0-9a-fA-F]{3}$/.test(hex)) {
    return {
      color: hex
        .split("")
        .map((c) => c + c)
        .join("")
        .toUpperCase(),
    };
  }
  if (/^[0-9a-fA-F]{4}$/.test(hex)) {
    const rgb = hex
      .slice(0, 3)
      .split("")
      .map((c) => c + c)
      .join("")
      .toUpperCase();
    const alpha = parseInt(hex[3] + hex[3], 16);
    return { color: rgb, transparency: alphaByteToTransparency(alpha) };
  }
  if (/^[0-9a-fA-F]{6}$/.test(hex)) {
    return { color: hex.toUpperCase() };
  }
  if (/^[0-9a-fA-F]{8}$/.test(hex)) {
    return {
      color: hex.slice(0, 6).toUpperCase(),
      transparency: alphaByteToTransparency(parseInt(hex.slice(6, 8), 16)),
    };
  }
  return undefined;
}

export function resolvePaint(
  value: string | undefined,
  palette: Record<string, string>,
): ParsedPaint | undefined {
  if (!value) return undefined;
  const token = value.trim();
  const raw = token in palette ? palette[token] : token;
  const parsed = parseHexPaint(raw);
  if (parsed) return parsed;
  const fallback = parseHexPaint(token);
  return fallback;
}

function pickTransparency(
  fromHex: number | undefined,
  opacity: number | undefined,
  transparency: number | undefined,
): number | undefined {
  if (transparency !== undefined) {
    return transparency <= 0 ? undefined : transparency;
  }
  if (opacity !== undefined) {
    return opacityToTransparency(opacity);
  }
  return fromHex;
}

export function resolveFillPaint(
  fill: { color: string; opacity?: number; transparency?: number } | string,
  palette: Record<string, string>,
): ParsedPaint | undefined {
  if (typeof fill === "string") {
    return resolvePaint(fill, palette);
  }
  const parsed = resolvePaint(fill.color, palette);
  if (!parsed) return undefined;
  return {
    color: parsed.color,
    transparency: pickTransparency(parsed.transparency, fill.opacity, fill.transparency),
  };
}

export function lineTransparency(
  line: { opacity?: number; transparency?: number },
  fromHex?: number,
): number | undefined {
  return pickTransparency(fromHex, line.opacity, line.transparency);
}

/** Strip `#`, expand 3-digit hex, uppercase. 8-digit hex drops alpha. Named keys are left as-is. */
export function normalizeColorToken(value: string): string {
  const parsed = parseHexPaint(value);
  if (parsed) return parsed.color;
  return value.trim();
}

export function resolveColor(
  value: string | undefined,
  palette: Record<string, string>,
  fallback?: string,
): string | undefined {
  const raw = value ?? fallback;
  if (!raw) return undefined;
  const token = raw.trim();
  if (token in palette) {
    return normalizeColorToken(palette[token]);
  }
  return normalizeColorToken(token);
}
