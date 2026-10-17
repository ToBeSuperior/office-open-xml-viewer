// OMML equation preparation for the PPTX slide renderer: MathJax raster
// cache and per-slide equation warm-up. Split from renderer.ts (#1857).

import type { Slide, SlideElement, TextBody } from './types';
import {
  mathToMathML,
  rasterizeMathSvg,
  isOoxmlDecodedImageLimitError,
  OoxmlResourceLimitError,
} from '@silurus/ooxml-core';
import type { MathNode, MathRenderer, RasterizedMathSvg } from '@silurus/ooxml-core';

// ===== Text layout helpers =====

// The run-underline painter `drawUnderline` (ECMA-376 §20.1.10.82
// ST_TextUnderlineType) was hoisted to core (`@silurus/ooxml-core`
// `drawUnderline`) so the docx renderer can share the exact same geometry /
// dash dispatch; it is imported at the top of this file and its behaviour is
// byte-identical to the former local copy.

// ── Math (OMML) rendering ──────────────────────────────────────────────────
// Equations are converted to SVG by MathJax once, cached by their MathNode[]
// reference (stable from parse through render), then drawn as images inline.
// Mirrors the docx renderer's pipeline so the typesetting is identical.
interface MathRender {
  /** The equation rasterized as opaque black glyphs on transparent. */
  raster: RasterizedMathSvg;
  /** baseline-relative extents in em (1em = the equation's font size in px). */
  widthEm: number;
  ascentEm: number;
  descentEm: number;
  /** Per-colour tinted copies (the black `img` recoloured via source-in). */
  tinted: Map<string, CanvasImageSource>;
}
export const mathRenders = new WeakMap<MathNode[], MathRender>();
export const failedMath = new WeakSet<MathNode[]>();

/** All currently supported text-bearing slide elements, including table cells. */
export function* elementTextBodies(element: SlideElement): Generator<TextBody> {
  if (element.type === 'shape' && element.textBody) yield element.textBody;
  else if (element.type === 'table') {
    for (const row of element.rows) for (const cell of row.cells) {
      if (cell.textBody) yield cell.textBody;
    }
  }
}

/** Gather every math run reachable from a slide's shapes and table cells. */
function collectSlideMathRuns(slide: Slide): { nodes: MathNode[]; display: boolean }[] {
  const found: { nodes: MathNode[]; display: boolean }[] = [];
  for (const element of slide.elements) for (const body of elementTextBodies(element)) {
    for (const para of body.paragraphs) for (const run of para.runs) {
      if (run.type === 'math') found.push({ nodes: run.nodes, display: run.display });
    }
  }
  return found;
}

/**
 * Pre-render every equation on the slide to an image (async), so the
 * synchronous layout/draw passes can place them by reading cached extents.
 * Conversion failures retain their authored fallback or a visible diagnostic.
 * An absent optional engine remains an explicit caller opt-out.
 */
export async function prepareSlideMath(slide: Slide, math: MathRenderer): Promise<void> {
  const runs = collectSlideMathRuns(slide);
  if (runs.length === 0) return;
  try { await math.loadMathJax(); }
  catch (error) {
    if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error)) throw error;
    for (const r of runs) if (!mathRenders.has(r.nodes)) failedMath.add(r.nodes);
    return;
  }
  for (const r of runs) {
    if (mathRenders.has(r.nodes)) continue;
    try {
      const out = await math.mathMLToSvg(mathToMathML(r.nodes, r.display));
      const raster = await rasterizeMathSvg(out, '#000000');
      failedMath.delete(r.nodes);
      mathRenders.set(r.nodes, {
        raster,
        widthEm: out.widthEm,
        ascentEm: out.ascentEm,
        descentEm: out.descentEm,
        tinted: new Map(),
      });
    } catch (error) {
      // Resource governance is fail-closed; only conversion failures select
      // an authored preview or the unavailable-equation diagnostic.
      if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error)) throw error;
      failedMath.add(r.nodes);
    }
  }
}
