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
});

export const textElementSchema = box.merge(textStyle).extend({
  type: z.literal("text"),
  fill: fillSchema.optional(),
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
});

export const elementSchema = z.discriminatedUnion("type", [
  textElementSchema,
  imageElementSchema,
  shapeElementSchema,
]);

export const fontsSchema = z.object({
  default: z.string().optional(),
  heading: z.string().optional(),
});

export const sizeSchema = z.object({
  width: z.number().positive(),
  height: z.number().positive(),
});

export const themeFieldsSchema = z.object({
  size: sizeSchema.optional(),
  background: hexOrName.optional(),
  fonts: fontsSchema.optional(),
  colors: z.record(z.string(), z.string()).optional(),
  title: z.string().optional(),
  author: z.string().optional(),
  subject: z.string().optional(),
  lang: z.string().optional(),
  elements: z.array(elementSchema).optional(),
});

export const templateSchema = themeFieldsSchema;

export const slideFileSchema = themeFieldsSchema.extend({
  template: z.string().optional(),
  notes: z.string().optional(),
});

export type Coord = z.infer<typeof coord>;
export type FillConfig = z.infer<typeof fillSchema>;
export type LineConfig = z.infer<typeof lineSchema>;
export type TextElement = z.infer<typeof textElementSchema>;
export type ImageElement = z.infer<typeof imageElementSchema>;
export type ShapeElement = z.infer<typeof shapeElementSchema>;
export type SlideElement = z.infer<typeof elementSchema>;
export type TemplateConfig = z.infer<typeof templateSchema>;
export type SlideFileConfig = z.infer<typeof slideFileSchema>;

export interface ResolvedSlide {
  source: string;
  size: { width: number; height: number };
  background?: string;
  fonts: { default?: string; heading?: string };
  colors: Record<string, string>;
  title?: string;
  author?: string;
  subject?: string;
  lang?: string;
  notes?: string;
  elements: SlideElement[];
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
