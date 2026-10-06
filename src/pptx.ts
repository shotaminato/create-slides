import pptxgen from "pptxgenjs";

/**
 * pptxgenjs ships as CJS. Under Node16 ESM resolution the default export
 * is a constructor at runtime, but the bundled .d.ts is not constructable.
 */
type PptxCtor = new () => PptxPres;

export type ArrowHead = "none" | "arrow" | "diamond" | "oval" | "stealth" | "triangle";

export interface LineProps {
  color?: string;
  width?: number;
  dashType?: string;
  beginArrowType?: ArrowHead;
  endArrowType?: ArrowHead;
  transparency?: number;
}

export interface FillProps {
  color: string;
  transparency?: number;
}

export interface PptxSlide {
  background: { color?: string };
  addNotes(notes: string): void;
  addText(text: string, opts: Record<string, unknown>): void;
  addImage(opts: Record<string, unknown>): void;
  addShape(shapeName: string, opts?: Record<string, unknown>): void;
}

export interface PptxPres {
  ShapeType: Record<string, string>;
  layout: string;
  title: string;
  author: string;
  subject: string;
  defineLayout(layout: { name: string; width: number; height: number }): void;
  addSlide(): PptxSlide;
  writeFile(props: { fileName: string }): Promise<string>;
}

export function createPresentation(): PptxPres {
  return new (pptxgen as unknown as PptxCtor)();
}
