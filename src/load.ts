import fs from "node:fs";
import path from "node:path";
import yaml from "js-yaml";
import { ZodError } from "zod";
import {
  DEFAULT_SIZE,
  formatZodError,
  slideFileSchema,
  SlideConfigError,
  templateSchema,
  type ResolvedSlide,
  type SlideElement,
  type SlideFileConfig,
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

function mergeElements(templateEls: SlideElement[] | undefined, slideEls: SlideElement[] | undefined): SlideElement[] {
  return [...(templateEls ?? []), ...(slideEls ?? [])];
}

function mergeSlide(template: TemplateConfig, slide: SlideFileConfig, source: string): ResolvedSlide {
  const size = {
    width: slide.size?.width ?? template.size?.width ?? DEFAULT_SIZE.width,
    height: slide.size?.height ?? template.size?.height ?? DEFAULT_SIZE.height,
  };

  return {
    source,
    size,
    background: slide.background ?? template.background,
    fonts: { ...template.fonts, ...slide.fonts },
    colors: { ...template.colors, ...slide.colors },
    title: slide.title ?? template.title,
    author: slide.author ?? template.author,
    subject: slide.subject ?? template.subject,
    lang: slide.lang ?? template.lang,
    notes: slide.notes,
    elements: mergeElements(template.elements, slide.elements),
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

export function loadDeck(options: LoadOptions): {
  slides: ResolvedSlide[];
  title?: string;
  author?: string;
  subject?: string;
} {
  const slidesDir = path.resolve(options.root, options.slidesDir);
  const slidePaths = listSlideFiles(slidesDir);
  const templateCache = new Map<string, TemplateConfig>();

  const loadTemplate = (filePath: string): TemplateConfig => {
    const cached = templateCache.get(filePath);
    if (cached) return cached;
    if (!fs.existsSync(filePath)) {
      throw new SlideConfigError(`Template file not found: ${filePath}`);
    }
    const parsed = parseYamlFile(filePath, templateSchema);
    templateCache.set(filePath, parsed);
    return parsed;
  };

  const slides = slidePaths.map((slidePath) => {
    const slide = parseYamlFile(slidePath, slideFileSchema);
    const templatePath = resolveTemplatePath(slide.template, options.defaultTemplate, options.root);
    const template = loadTemplate(templatePath);
    const relative = path.relative(options.root, slidePath) || slidePath;
    return mergeSlide(template, slide, relative);
  });

  const first = slides[0];
  return {
    slides,
    title: first?.title,
    author: first?.author,
    subject: first?.subject,
  };
}
