// XLSX cell comment indicator (the red corner triangle). Split from
// renderer.ts (#1857).

import { parseA1 } from './a1.js';

/** Parse an A1-style cell reference to 1-based row/col. Aliased to the shared
 *  {@link parseA1} so the renderer, data-validation and the comment popup all
 *  agree on ref parsing (the shared form also tolerates `$`-absolute markers). */
export const parseA1Ref = parseA1;

/**
 * Draw Excel's comment marker — a small filled triangle in the top-right
 * corner of the cell — coloured like Excel's default red indicator. Scales
 * with cell size but is clamped so it stays legible at small zoom.
 */
export function drawCommentMarker(
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
): void {
  const size = Math.max(4, Math.min(8, Math.min(w, h) * 0.18));
  ctx.save();
  ctx.fillStyle = '#D40000';
  ctx.beginPath();
  ctx.moveTo(x + w - size, y);
  ctx.lineTo(x + w, y);
  ctx.lineTo(x + w, y + size);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}
