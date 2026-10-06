import fs from "node:fs";
import path from "node:path";
import yaml from "js-yaml";
import { ZodError } from "zod";
import {
  DEFAULT_SIZE,
  elementHasPagePlaceholder,
  elementSchema,
  formatZodError,
  isMasterChrome,
  pageNumberToElement,
  slideFileSchema,
  SlideConfigError,
  templateSchema,
  type ResolvedSlide,
  type SlideElement,
  type SlideFileConfig,
  type SlideInputElement,
  type SlideMasterDef,
  type TemplateConfig,
} from "./schema.js";

export interface LoadOptions {
  root: string;
  slidesDir: string;
  defaultTemplate: string;
}

function readText(filePath: string): string {
  try {
    return fs.readFileSync(filePath, "utf8");
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "ENOENT") {
      throw new SlideConfigError(`File not found: ${filePath}`);
    }
    throw new SlideConfigError(`Cannot read file: ${filePath}\n${(err as Error).message}`);
  }
}

function parseYamlFile<T>(
  filePath: string,
  schema: { parse: (data: unknown) => T },
): T {
  const raw = readText(filePath);
  let data: unknown;
  try {
    data = yaml.load(raw, { filename: filePath });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    throw new SlideConfigError(`Invalid YAML in ${filePath}:\n${message}`);
  }

  if (data === undefined || data === null) {
    throw new SlideConfigError(`YAML file is empty: ${filePath}`);
  }

  try {
    return schema.parse(data);
  } catch (err) {
    if (err instanceof ZodError) {
      throw new SlideConfigError(formatZodError(err, filePath));
    }
    throw err;
  }
}

export function findProjectRoot(cwd: string): string {
  let dir = path.resolve(cwd);
  while (true) {
    if (fs.existsSync(path.join(dir, "package.json"))) {
      return dir;
    }
    const parent = path.dirname(dir);
    if (parent === dir) {
      return path.resolve(cwd);
    }
    dir = parent;
  }
}

/** `default` → CLI template; a bare name → templates/<name>.yaml; otherwise a path from root. */
export function resolveTemplatePath(
  spec: string | undefined,
  defaultTemplate: string,
  root: string,
): string {
  const value = spec?.trim() || "default";
  if (value === "default") {
    return path.resolve(root, defaultTemplate);
  }
  if (value.endsWith(".yaml") || value.endsWith(".yml") || value.includes("/") || value.includes("\\")) {
    return path.resolve(root, value);
  }
  return path.resolve(root, "templates", `${value}.yaml`);
}

function allocateMasterName(used: Set<string>, templatePath: string): string {
  const base = path.basename(templatePath, path.extname(templatePath)) || "master";
  if (!used.has(base)) {
    used.add(base);
    return base;
  }
  let n = 2;
  while (used.has(`${base}-${n}`)) n += 1;
  const name = `${base}-${n}`;
  used.add(name);
  return name;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Deep-merge plain objects; slide override wins on conflicting keys. Arrays / primitives replace. */
export function deepMergeRecord(
  base: Record<string, unknown>,
  override: Record<string, unknown>,
): Record<string, unknown> {
  const out: Record<string, unknown> = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value === undefined || value === null) continue;
    const prev = out[key];
    if (isPlainObject(prev) && isPlainObject(value)) {
      out[key] = deepMergeRecord(prev, value);
    } else {
      out[key] = value;
    }
  }
  return out;
}

export function splitTemplateElements(
  template: TemplateConfig,
  templateSource: string,
): {
  masterElements: SlideElement[];
  overlayElements: SlideElement[];
  slots: SlideElement[];
} {
  const masterElements: SlideElement[] = [];
  const overlayElements: SlideElement[] = [];
  const slots: SlideElement[] = [];
  const seen = new Set<string>();

  for (const el of template.elements ?? []) {
    if (el.role === "slot" && !el.id) {
      throw new SlideConfigError(
        `Template slot is missing id in ${templateSource} (role: slot requires id)`,
      );
    }
    if (isMasterChrome(el)) {
      if (elementHasPagePlaceholder(el)) overlayElements.push(el);
      else masterElements.push(el);
      continue;
    }
    const id = el.id as string;
    if (seen.has(id)) {
      throw new SlideConfigError(`Duplicate template slot id "${id}" in ${templateSource}`);
    }
    seen.add(id);
    slots.push(el);
  }

  if (template.pageNumber) {
    overlayElements.push(pageNumberToElement(template.pageNumber));
  }
  return { masterElements, overlayElements, slots };
}

/**
 * Template slots (template order) merged with matching slide `{ id }` overrides,
 * then unmatched slide elements appended.
 */
export function mergeSlotElements(
  templateSlots: SlideElement[],
  slideElements: SlideInputElement[],
  source: string,
): SlideElement[] {
  const slotById = new Map<string, SlideElement>();
  for (const slot of templateSlots) {
    if (slot.id) slotById.set(slot.id, slot);
  }

  const overrideById = new Map<string, Record<string, unknown>>();
  const extras: SlideInputElement[] = [];

  for (const raw of slideElements) {
    if (raw.id && slotById.has(raw.id)) {
      if (overrideById.has(raw.id)) {
        throw new SlideConfigError(`Duplicate slot id "${raw.id}" in ${source}`);
      }
      overrideById.set(raw.id, raw as Record<string, unknown>);
    } else {
      extras.push(raw);
    }
  }

  const mergedSlots = templateSlots.map((slot) => {
    const override = slot.id ? overrideById.get(slot.id) : undefined;
    const merged = override
      ? deepMergeRecord(slot as unknown as Record<string, unknown>, override)
      : (slot as unknown as Record<string, unknown>);
    const parsed = elementSchema.safeParse(merged);
    if (!parsed.success) {
      throw new SlideConfigError(
        `Slot "${slot.id ?? "?"}" in ${source} is incomplete after merging template + slide override:\n` +
          formatZodError(parsed.error, source),
      );
    }
    return parsed.data;
  }).filter((el) => {
    // Optional text slots (e.g. headingNote) are omitted when the slide leaves them empty.
    if (el.type === "text" && !(el.text && el.text.length > 0)) return false;
    return true;
  });

  const extraElements = extras.map((raw) => {
    const parsed = elementSchema.safeParse(raw);
    if (!parsed.success) {
      const id = raw.id ? ` (id: ${raw.id})` : "";
      throw new SlideConfigError(
        `Element${id} in ${source} is not a complete slide element and does not match a template slot.\n` +
          formatZodError(parsed.error, source),
      );
    }
    return parsed.data;
  });

  return [...mergedSlots, ...extraElements];
}

function mergeSlide(
  template: TemplateConfig,
  slots: SlideElement[],
  overlayElements: SlideElement[],
  slide: SlideFileConfig,
  source: string,
  masterName: string,
): ResolvedSlide {
  const size = {
    width: slide.size?.width ?? template.size?.width ?? DEFAULT_SIZE.width,
    height: slide.size?.height ?? template.size?.height ?? DEFAULT_SIZE.height,
  };

  const pageNumber = slide.pageNumber
    ? pageNumberToElement(slide.pageNumber)
    : undefined;
  const overlay = pageNumber
    ? overlayElements.filter((el) => !elementHasPagePlaceholder(el)).concat(pageNumber)
    : overlayElements;

  return {
    source,
    masterName,
    size,
    backgroundOverride: slide.background,
    fonts: { ...template.fonts, ...slide.fonts },
    colors: { ...template.colors, ...slide.colors },
    title: slide.title ?? template.title,
    author: slide.author ?? template.author,
    subject: slide.subject ?? template.subject,
    lang: slide.lang ?? template.lang,
    notes: slide.notes,
    elements: [...overlay, ...mergeSlotElements(slots, slide.elements ?? [], source)],
  };
}

export function listSlideFiles(slidesDir: string): string[] {
  if (!fs.existsSync(slidesDir)) {
    throw new SlideConfigError(`Slide directory not found: ${slidesDir}`);
  }
  const stat = fs.statSync(slidesDir);
  if (!stat.isDirectory()) {
    throw new SlideConfigError(`Slide path is not a directory: ${slidesDir}`);
  }

  const files = fs
    .readdirSync(slidesDir)
    .filter((name) => name.endsWith(".yaml") || name.endsWith(".yml"))
    .sort((a, b) => a.localeCompare(b, "en", { numeric: true }));

  if (files.length === 0) {
    throw new SlideConfigError(`No YAML slide files found in ${slidesDir}`);
  }

  return files.map((name) => path.join(slidesDir, name));
}

interface CachedTemplate {
  config: TemplateConfig;
  master: SlideMasterDef;
  slots: SlideElement[];
  overlayElements: SlideElement[];
}

export function loadDeck(options: LoadOptions): {
  slides: ResolvedSlide[];
  masters: SlideMasterDef[];
  title?: string;
  author?: string;
  subject?: string;
} {
  const slidesDir = path.resolve(options.root, options.slidesDir);
  const slidePaths = listSlideFiles(slidesDir);
  const templateCache = new Map<string, CachedTemplate>();
  const usedMasterNames = new Set<string>();
  const masters: SlideMasterDef[] = [];

  const loadTemplate = (filePath: string): CachedTemplate => {
    const cached = templateCache.get(filePath);
    if (cached) return cached;
    if (!fs.existsSync(filePath)) {
      throw new SlideConfigError(`Template file not found: ${filePath}`);
    }
    const config = parseYamlFile(filePath, templateSchema);
    const templateSource = path.relative(options.root, filePath) || filePath;
    const { masterElements, overlayElements, slots } = splitTemplateElements(
      config,
      templateSource,
    );
    const master: SlideMasterDef = {
      name: allocateMasterName(usedMasterNames, filePath),
      source: templateSource,
      background: config.background,
      fonts: { ...config.fonts },
      colors: { ...config.colors },
      lang: config.lang,
      elements: masterElements,
    };
    const entry = { config, master, slots, overlayElements };
    templateCache.set(filePath, entry);
    masters.push(master);
    return entry;
  };

  const slides = slidePaths.map((slidePath) => {
    const slide = parseYamlFile(slidePath, slideFileSchema);
    const templatePath = resolveTemplatePath(slide.template, options.defaultTemplate, options.root);
    const loaded = loadTemplate(templatePath);
    const relative = path.relative(options.root, slidePath) || slidePath;
    return mergeSlide(
      loaded.config,
      loaded.slots,
      loaded.overlayElements,
      slide,
      relative,
      loaded.master.name,
    );
  });

  const first = slides[0];
  return {
    slides,
    masters,
    title: first?.title,
    author: first?.author,
    subject: first?.subject,
  };
}
