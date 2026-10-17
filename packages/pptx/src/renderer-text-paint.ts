// PPTX text-run painting helpers: highlight boxes, underline / strikethrough
// decorations, and inline equation drawing. Split from renderer.ts (#1857).

import {
  crispOffset,
  resolveFill as resolveFillCore,
  EMU_PER_PT as PT_TO_EMU,
  cachedTintedMathRaster,
  highlightBox,
  drawUnderline,
  DECORATION_LINE_WIDTH_EM,
  STRIKE_RISE_EM,
  paintOptionalImagePlaceholder,
} from '@silurus/ooxml-core';
import { emuToPx } from './renderer-context';
import { mathRenders } from './renderer-math';
import type { LayoutSegment } from './renderer-text-layout';

/**
 * Paint a run's text-highlight (marker) box behind the glyphs.
 * ECMA-376 §21.1.2.3.4 — `<a:rPr><a:highlight>`. Called before the glyphs are
 * drawn so they (and any underline / strikethrough) sit on top, and before any
 * shadow is set on the context so the box itself isn't shadowed. `width` is the
 * glyph advance computed by the caller (it differs between the normal and
 * tab-stop paths only by the justification stretch added to it). The vertical
 * band comes from the shared `highlightBox` helper. Restore the resolved glyph
 * paint, which may be a CanvasPattern, before the subsequent fillText.
 */
export function paintHighlight(
  ctx: CanvasRenderingContext2D,
  x: number,
  baseline: number,
  width: number,
  fontPx: number,
  highlight: string,
  glyphPaint: string | CanvasGradient | CanvasPattern,
): void {
  const { top, height } = highlightBox(baseline, fontPx);
  ctx.fillStyle = highlight;
  ctx.fillRect(x, top, width, height);
  ctx.fillStyle = glyphPaint;
}

/** DrawingML underline (§20.1.10.82, §21.1.2.3.9). Horizontal runs use
 * their drawing size; stacked text supplies its column's common size/frame. */
export function paintRunUnderline(ctx: CanvasRenderingContext2D, seg: LayoutSegment,
  x: number, baseline: number, width: number, glyphPaint: string | CanvasGradient | CanvasPattern,
  scale: number, dpr: number, drawSizePx = seg.drawSizePx ?? seg.sizePx): void {
  if (seg.underline && !seg.underlineLineNoFill
    && (!seg.noFill || seg.underlineFill || seg.underlineColor || seg.underlineLine?.fill)) {
    // §21.1.2.3.13 uFillTx follows the actual glyph fill (including
    // pattFill), while §21.1.2.3.12 uFill supplies its own fill choice.
    // Office PDF shows a separately chosen underline pattern on the
    // same slide-aligned grid.
    const underlinePaint = seg.underlineFill
      ? resolveFillCore(seg.underlineFill, ctx, x, baseline, width, drawSizePx * DECORATION_LINE_WIDTH_EM,
          0, scale * PT_TO_EMU) ?? 'rgba(0,0,0,0)'
      : seg.underlineLine?.fill
        ? resolveFillCore(seg.underlineLine.fill, ctx, x, baseline, width,
            drawSizePx * DECORATION_LINE_WIDTH_EM, 0, scale * PT_TO_EMU) ?? 'rgba(0,0,0,0)'
        : seg.underlineColor ?? glyphPaint;
    drawUnderline(ctx, x, baseline, width, drawSizePx,
      underlinePaint, seg.underlineStyle, dpr,
      seg.underlineLine?.width ? emuToPx(seg.underlineLine.width, scale) : undefined);
  }
}

export function paintRunStrike(ctx: CanvasRenderingContext2D, seg: LayoutSegment,
  x: number, baseline: number, width: number, glyphPaint: string | CanvasGradient | CanvasPattern,
  dpr: number): void {
  const drawSizePx = seg.drawSizePx ?? seg.sizePx;
  if (seg.strikethrough && !seg.noFill) {
    const lineW = Math.max(1, drawSizePx * DECORATION_LINE_WIDTH_EM);
    ctx.strokeStyle = glyphPaint;
    ctx.lineWidth = lineW;
    ctx.setLineDash([]);
    // Crispness nudge (see crispOffset): the strike is a horizontal stroke;
    // an odd device-pixel width is centered on one device row by snapping its
    // y onto the nearest crisp device position (otherwise it straddles two
    // rows → blurry). Snap each line from its own y.
    // Strike rise and width: core library policy (see text/underline.ts).
    const yMid = baseline - drawSizePx * STRIKE_RISE_EM;
    if (seg.strikeDouble) {
      // Two parallel lines straddling the standard strike position, each
      // 0.9 line widths from it (library policy; ST_TextStrikeType
      // §20.1.10.79 `dblStrike` gives no spacing and no Office measurement
      // is recorded).
      const offset = lineW * 0.9;
      const yUp = yMid - offset;
      const yDn = yMid + offset;
      ctx.beginPath();
      ctx.moveTo(x, yUp + crispOffset(yUp, lineW, dpr));
      ctx.lineTo(x + width, yUp + crispOffset(yUp, lineW, dpr));
      ctx.moveTo(x, yDn + crispOffset(yDn, lineW, dpr));
      ctx.lineTo(x + width, yDn + crispOffset(yDn, lineW, dpr));
      ctx.stroke();
    } else {
      const sy = yMid + crispOffset(yMid, lineW, dpr);
      ctx.beginPath();
      ctx.moveTo(x, sy);
      ctx.lineTo(x + width, sy);
      ctx.stroke();
    }
  }
}

/** Horizontal text uses the same frame for both decorations. */
export function paintRunDecorations(ctx: CanvasRenderingContext2D, seg: LayoutSegment,
  x: number, baseline: number, width: number, glyphPaint: string | CanvasGradient | CanvasPattern,
  scale: number, dpr: number): void {
  paintRunUnderline(ctx, seg, x, baseline, width, glyphPaint, scale, dpr);
  paintRunStrike(ctx, seg, x, baseline, width, glyphPaint, dpr);
}

/** Reuse the unavailable-content painter for a failed equation with no MCE
 * preview. This is library failure policy, not a substitute Office equation. */
export function paintEquation(ctx: CanvasRenderingContext2D, style: LayoutSegment, x: number, y: number) {
  const equation = style.math!;
  const render = equation.unavailable ? undefined : mathRenders.get(equation.nodes);
  const height = equation.ascent + equation.descent;
  if (!(equation.width > 0 && height > 0)) return;
  if (render) ctx.drawImage(cachedTintedMathRaster(render, style.color), x, y, equation.width, height);
  else if (equation.fallbackText) {
    ctx.font = style.font;
    ctx.fillStyle = style.color;
    const baseline = ctx.textBaseline;
    const align = ctx.textAlign;
    ctx.textBaseline = 'top'; ctx.textAlign = 'left';
    ctx.fillText(equation.fallbackText, x, y);
    ctx.textBaseline = baseline; ctx.textAlign = align;
  } else if (equation.unavailable) {
    paintOptionalImagePlaceholder(ctx, 'equation', { x, y, width: equation.width, height });
  }
}
