#!/usr/bin/env node
import path from "node:path";
import { SlideConfigError } from "./schema.js";
import { findProjectRoot, loadDeck } from "./load.js";
import { renderDeck } from "./render.js";

interface CliArgs {
  help?: boolean;
  slides: string;
  template: string;
  out: string;
  root?: string;
}

function printHelp(): void {
  const text = `
create-slides — YAML configs to PowerPoint (.pptx)

Usage:
  npx tsx src/cli.ts [options]
  npm run build-slides -- [options]
  create-slides [options]          (after npm run build)

Options:
  --slides <dir>       Directory of slide YAML files (default: slides)
  --template <file>    Default template YAML (default: templates/default.yaml)
  --out <file>         Output .pptx path (default: dist/deck.pptx)
  --root <dir>         Project root (default: cwd, walking up to package.json)
  -h, --help           Show this help

Slide files are loaded in sorted filename order (01-title.yaml, …).
Each slide may set template: default (or a name/path). Slide fields override the template.
Template elements with an id are slots: the slide can override just text (and other keys).
`.trim();
  console.log(text);
}

function parseArgs(argv: string[]): CliArgs {
  const out: CliArgs = {
    slides: "slides",
    template: "templates/default.yaml",
    out: "dist/deck.pptx",
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    const next = argv[i + 1];
    const take = (): string => {
      if (!next || next.startsWith("-")) {
        throw new SlideConfigError(`Option ${arg} requires a value.`);
      }
      i += 1;
      return next;
    };

    switch (arg) {
      case "-h":
      case "--help":
        out.help = true;
        break;
      case "--slides":
        out.slides = take();
        break;
      case "--template":
        out.template = take();
        break;
      case "--out":
        out.out = take();
        break;
      case "--root":
        out.root = take();
        break;
      default:
        throw new SlideConfigError(`Unknown option: ${arg}\nUse --help to see available options.`);
    }
  }

  return out;
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    printHelp();
    return;
  }

  const root = args.root ? path.resolve(args.root) : findProjectRoot(process.cwd());
  const deck = loadDeck({
    root,
    slidesDir: args.slides,
    defaultTemplate: args.template,
  });

  const outFile = path.resolve(root, args.out);
  const written = await renderDeck(deck.slides, {
    root,
    outFile,
    title: deck.title,
    author: deck.author,
    subject: deck.subject,
    masters: deck.masters,
  });

  console.log(`Wrote ${deck.slides.length} slide(s) → ${written}`);
  if (deck.masters.length > 0) {
    console.log(`  masters: ${deck.masters.map((m) => m.name).join(", ")}`);
  }
  for (const slide of deck.slides) {
    console.log(`  - ${slide.source} [${slide.masterName}] (${slide.elements.length} elements)`);
  }
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error(`Error: ${message}`);
  process.exitCode = 1;
});
