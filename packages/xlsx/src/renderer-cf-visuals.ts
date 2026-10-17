// XLSX conditional-format visuals: data bars and icon sets. Split from
// renderer.ts (#1857).

import { hexToRgba } from '@silurus/ooxml-core';

/** Convert a stored column-width value (ECMA-376 §18.3.1.13 `<col width>`, in
 *  "number of characters" = max digit widths) to CSS pixels.
 *
 *  This is the spec's file→pixel formula verbatim:
 *    `Truncate(((256 * width + Truncate(128 / MDW)) / 256) * MDW)`
 *  Note both truncations: the `Truncate(128 / MDW)` constant is computed and
 *  truncated *before* it is folded into the numerator (§18.3.1.13), then the
 *  whole expression is truncated to an integer pixel. Excel stores integer
 *  pixel column widths, so this yields exactly the width Excel renders. */
/**
 * Fill a data-bar rectangle. Excel 2010+ dataBars default to a horizontal
 * gradient (`x14:dataBar@gradient="1"`): solid color on the left, fading
 * to an ~85%-tinted-to-white version on the right. We render with alpha
 * stops rather than literally mixing toward white so underlying cell
 * background (including zebra-striping or fills) shows through. With
 * `gradient="0"` the bar is drawn as a flat solid color.
 */
export function fillDataBar(
  ctx: CanvasRenderingContext2D,
  color: string,
  x: number, y: number, w: number, h: number,
  gradient: boolean,
): void {
  if (w <= 0 || h <= 0) return;
  if (gradient) {
    const grad = ctx.createLinearGradient(x, y, x + w, y);
    grad.addColorStop(0, hexToRgba(color, 0.85));
    grad.addColorStop(1, hexToRgba(color, 0.15));
    ctx.fillStyle = grad;
  } else {
    ctx.fillStyle = hexToRgba(color);
  }
  ctx.fillRect(x, y, w, h);
}

// ────────────────────────────────────────────────────────────────
// Icon Set drawing
// ────────────────────────────────────────────────────────────────
const ICON_COLORS_3 = ['#FF0000', '#FFFF00', '#00B050'];
const ICON_COLORS_4 = ['#FF0000', '#FF6600', '#FFFF00', '#00B050'];
const ICON_COLORS_5 = ['#FF0000', '#FF6600', '#FFFF00', '#92D050', '#00B050'];

export function drawCfIcon(ctx: CanvasRenderingContext2D, name: string, index: number, x: number, y: number, sz: number): void {
  if (name === 'NoIcons') return;
  const safeName = name || '3TrafficLights1';
  const nIcons = parseInt(safeName[0]) || 3;
  const palette = nIcons === 5 ? ICON_COLORS_5 : nIcons === 4 ? ICON_COLORS_4 : ICON_COLORS_3;
  const color = palette[Math.max(0, Math.min(index, palette.length - 1))];
  ctx.save();
  ctx.fillStyle = color;
  if (safeName.includes('Arrow')) {
    const half = sz / 2;
    ctx.beginPath();
    if (index === nIcons - 1) {
      ctx.moveTo(x + half, y); ctx.lineTo(x + sz, y + sz); ctx.lineTo(x, y + sz);
    } else if (index === 0) {
      ctx.moveTo(x, y); ctx.lineTo(x + sz, y); ctx.lineTo(x + half, y + sz);
    } else {
      ctx.moveTo(x, y + sz * 0.3); ctx.lineTo(x + sz, y + half); ctx.lineTo(x, y + sz * 0.7);
    }
    ctx.closePath();
    ctx.fill();
  } else if (safeName === '3Signs' && index < 2) {
    // ECMA-376 §18.18.42 3Signs: Excel draws a red diamond, a yellow
    // triangle and a green circle (observed in Excel's PDF output).
    ctx.beginPath();
    if (index === 0) {
      ctx.moveTo(x + sz / 2, y); ctx.lineTo(x + sz, y + sz / 2);
      ctx.lineTo(x + sz / 2, y + sz); ctx.lineTo(x, y + sz / 2);
    } else {
      ctx.moveTo(x + sz / 2, y); ctx.lineTo(x + sz, y + sz); ctx.lineTo(x, y + sz);
    }
    ctx.closePath();
    ctx.fill();
  } else if (safeName.includes('Flag')) {
    ctx.beginPath();
    ctx.moveTo(x, y); ctx.lineTo(x + sz, y); ctx.lineTo(x, y + sz);
    ctx.closePath();
    ctx.fill();
  } else {
    ctx.beginPath();
    ctx.arc(x + sz / 2, y + sz / 2, sz / 2, 0, Math.PI * 2);
    ctx.fill();
  }
  ctx.restore();
}
