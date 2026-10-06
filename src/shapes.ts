import { SlideConfigError } from "./schema.js";
import { createPresentation } from "./pptx.js";

/**
 * Friendly YAML names → pptxgenjs `ShapeType` values.
 * Keys are matched after lowercasing and stripping spaces / hyphens / underscores.
 * Any official pptxgenjs ShapeType name is also accepted as-is.
 */
const ALIASES: Record<string, string> = {
  rect: "rect",
  rectangle: "rect",

  roundrect: "roundRect",
  roundedrect: "roundRect",
  roundedrectangle: "roundRect",

  ellipse: "ellipse",
  oval: "ellipse",
  circle: "ellipse",

  line: "line",
  lineinv: "lineInv",

  rightarrow: "rightArrow",
  arrowright: "rightArrow",
  arrow: "rightArrow",
  blockarrow: "rightArrow",

  leftarrow: "leftArrow",
  arrowleft: "leftArrow",

  uparrow: "upArrow",
  arrowup: "upArrow",

  downarrow: "downArrow",
  arrowdown: "downArrow",

  leftrightarrow: "leftRightArrow",
  updownarrow: "upDownArrow",

  notchedrightarrow: "notchedRightArrow",
  stripedrightarrow: "stripedRightArrow",
  chevron: "chevron",

  // pptxgenjs ShapeType.triangle is an isosceles triangle.
  // There is no ShapeType named isoscelesTriangle; the legacy enum is ISOSCELES_TRIANGLE = "triangle".
  triangle: "triangle",
  isosceles: "triangle",
  isoscelestriangle: "triangle",

  // Right triangle (pptxgenjs ShapeType.rtTriangle / legacy RIGHT_TRIANGLE)
  rttriangle: "rtTriangle",
  righttriangle: "rtTriangle",

  diamond: "diamond",
};

export interface ResolvedShape {
  /** pptxgenjs ShapeType string, e.g. "roundRect" */
  type: string;
  /** Extra line props implied by aliases such as lineArrow */
  lineDefaults?: {
    beginArrowType?: "none" | "arrow" | "diamond" | "oval" | "stealth" | "triangle";
    endArrowType?: "none" | "arrow" | "diamond" | "oval" | "stealth" | "triangle";
  };
}

let shapeTypeSet: Set<string> | undefined;

function officialShapeTypes(): Set<string> {
  if (!shapeTypeSet) {
    const inst = createPresentation();
    shapeTypeSet = new Set(Object.values(inst.ShapeType));
  }
  return shapeTypeSet;
}

function normalizeKey(name: string): string {
  return name.trim().toLowerCase().replace(/[-_\s]/g, "");
}

/**
 * Special aliases that are not themselves ShapeType values.
 * `lineArrow` uses ShapeType.line plus an arrow head on the line.
 */
function specialAlias(key: string): ResolvedShape | undefined {
  if (key === "linearrow" || key === "arrowline") {
    return { type: "line", lineDefaults: { endArrowType: "arrow" } };
  }
  if (key === "linedoublearrow" || key === "doublelinearrow") {
    return { type: "line", lineDefaults: { beginArrowType: "arrow", endArrowType: "arrow" } };
  }
  return undefined;
}

export function resolveShapeName(name: string): ResolvedShape {
  const trimmed = name.trim();
  const key = normalizeKey(trimmed);
  const special = specialAlias(key);
  if (special) return special;

  const aliased = ALIASES[key];
  const candidate = aliased ?? trimmed;
  const official = officialShapeTypes();

  if (official.has(candidate)) {
    return { type: candidate };
  }

  // Allow camelCase / PascalCase variants of official names (e.g. RightArrow → rightArrow)
  for (const value of official) {
    if (normalizeKey(value) === key) {
      return { type: value };
    }
  }

  const aliasList = [
    "rect / rectangle",
    "roundRect / roundedRect / rounded rectangle",
    "ellipse / oval / circle",
    "line",
    "lineArrow",
    "lineDoubleArrow",
    "rightArrow / leftArrow / upArrow / downArrow",
    "leftRightArrow / upDownArrow",
    "triangle / isoscelesTriangle",
    "rtTriangle / rightTriangle",
  ].join(", ");

  throw new SlideConfigError(
    `Unknown shape "${name}". Use a friendly name (${aliasList}) or any pptxgenjs ShapeType (see README).`,
  );
}

export function listOfficialShapeTypes(): string[] {
  return [...officialShapeTypes()].sort();
}
