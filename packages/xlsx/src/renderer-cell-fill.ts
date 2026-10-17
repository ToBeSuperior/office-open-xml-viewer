// XLSX cell pattern and gradient fills. Split from renderer.ts (#1857).

import { ownValue } from '@silurus/ooxml-core/internal/own-key';
import type { CellFill, GradientFillSpec } from './types.js';
import { createAuxCanvas, hexToRgba } from '@silurus/ooxml-core';

/**
 * Fractional fg coverage for an ECMA-376 ST_PatternType (§18.8.22). Values are
 * derived from the spec's verbal descriptions — gray125 = 12.5% fg on bg,
 * gray0625 = 6.25%, mediumGray ≈ 50%, darkGray ≈ 75%, lightGray ≈ 25%. Hatch
 * variants (darkHorizontal etc.) approximate their average ink density.
 * Unknown values default to 1 so they render as solid fg.
 */
function patternCoverage(pt: string): number {
  switch (pt) {
    case 'solid':         return 1;
    case 'darkGray':      return 0.75;
    case 'mediumGray':    return 0.50;
    case 'lightGray':     return 0.25;
    case 'gray125':       return 0.125;
    case 'gray0625':      return 0.0625;
    // Directional hatches: visual ink ratio is roughly 50/50 at default cell
    // sizes; without a true hatch tile we blend at 50% so the cell reads as a
    // middle-tone fill rather than a solid fg block.
    case 'darkHorizontal':
    case 'darkVertical':
    case 'darkDown':
    case 'darkUp':
    case 'darkGrid':
    case 'darkTrellis':   return 0.5;
    case 'lightHorizontal':
    case 'lightVertical':
    case 'lightDown':
    case 'lightUp':
    case 'lightGrid':
    case 'lightTrellis':  return 0.25;
    default:              return 1;
  }
}

/**
 * Cache of (pattern, fg, bg, ctx scale) → CanvasPattern. Keyed by a compound
 * string so the same pattern across thousands of cells only builds the tile
 * once *for a given destination scale*. The ctx scale is part of the key
 * because pat.setTransform() pre-bakes the inverse-scale transform into the
 * pattern, and we don't want a Storybook scale=1.5 cell to reuse a cached
 * pattern from a previous scale=2 render.
 */
const PATTERN_CACHE = new Map<string, CanvasPattern | null>();

/**
 * Bitmaps for ECMA-376 ST_PatternType (§18.8.22). Each entry is a square tile
 * — the row count (rows.length) is the tile size in pixels, and bit
 * (size - 1) of each row is the leftmost pixel. A `1` bit places fg on the
 * tile, `0` leaves bg.
 *
 * Mixing tile sizes lets us pick the smallest square that exactly fits each
 * pattern's natural pitch: gray and diagonal families fit in 8×8, but the
 * horizontal / vertical / grid families need a 12×12 tile so dark and light
 * variants can share the same line pitch (3 px) while differing only in
 * stroke thickness (2 px vs 1 px). At 8×8 this would have forced a 4-px
 * pitch on dark vs 2-px on light — visibly different line counts instead of
 * the matched-count, different-thickness pair Excel renders.
 *
 * Values follow the long-standing Office pattern set; the spec only names
 * the patterns, never their pixel geometry.
 *
 * Coverage targets:
 *   gray125    12.5%   sparse dot  (1 dot per 8 pixels)
 *   gray0625    6.25%   very sparse (1 dot per 16 pixels)
 *   lightGray  25%      checker dot (1 in 4)
 *   mediumGray 50%      full checker
 *   darkGray   75%      inverse of lightGray
 */
const PATTERN_BITMAPS: Record<string, number[]> = {
  // ── 8×8 — gray-family dot patterns ────────────────────────────────────
  // Constructed as a single dot motif scaled up tier-by-tier — each pattern
  // is the previous one with extra dots interleaved at the same 4-row pitch.
  // gray0625 is the reference seed (4 dots / 64 ≈ 6%) and the cascade roughly
  // doubles density at each step, so all five tiers read as the same stipple
  // texture at progressively higher coverage. Some moiré at intermediate
  // zooms is acceptable per user preference; the visual continuity wins.
  gray0625:   [0b10000000, 0b00000000, 0b00001000, 0b00000000, 0b10000000, 0b00000000, 0b00001000, 0b00000000], // ≈ 6%
  gray125:    [0b10001000, 0b00000000, 0b00100010, 0b00000000, 0b10001000, 0b00000000, 0b00100010, 0b00000000], // ≈ 12%
  lightGray:  [0b10101010, 0b00000000, 0b01010101, 0b00000000, 0b10101010, 0b00000000, 0b01010101, 0b00000000], // ≈ 25%
  mediumGray: [0b10101010, 0b01010101, 0b10101010, 0b01010101, 0b10101010, 0b01010101, 0b10101010, 0b01010101], // ≈ 50%
  // 75% — distribute the empty pixels evenly across rows so the cell reads
  // as a solid darker grey instead of alternating bands.
  darkGray:   [0b01110111, 0b11011101, 0b01110111, 0b11011101, 0b01110111, 0b11011101, 0b01110111, 0b11011101], // ≈ 75%

  // ── 12×12 — horizontal / vertical (matched line count) ───────────────
  // dark* and light* share the same 4-line-per-tile count and 3-px pitch;
  // they differ only in stroke thickness — dark uses a 2-px bar and a
  // 1-px gap, light uses a 1-px line and a 2-px gap. Per user feedback
  // the visual rule is "B19 has the same number of horizontal lines as
  // B18, just thinner", so the matched-count construction is restored
  // here. The earlier "more lines in light" 8×8 attempt was Excel-wrong.
  darkHorizontal: [
    0b111111111111, 0b111111111111, 0b000000000000,
    0b111111111111, 0b111111111111, 0b000000000000,
    0b111111111111, 0b111111111111, 0b000000000000,
    0b111111111111, 0b111111111111, 0b000000000000,
  ],
  lightHorizontal: [
    0b111111111111, 0b000000000000, 0b000000000000,
    0b111111111111, 0b000000000000, 0b000000000000,
    0b111111111111, 0b000000000000, 0b000000000000,
    0b111111111111, 0b000000000000, 0b000000000000,
  ],
  // darkVertical: 2-col bars at cols 0,1 / 3,4 / 6,7 / 9,10
  //   bits set per row: 11,10 + 8,7 + 5,4 + 2,1 = 0b110110110110 = 0xDB6
  darkVertical:   Array(12).fill(0xDB6),
  // lightVertical: 1-col lines at cols 0 / 3 / 6 / 9
  //   bits set per row: 11 + 8 + 5 + 2 = 0b100100100100 = 0x924
  lightVertical:  Array(12).fill(0x924),

  // darkGrid: Excel renders this as a fine 2×2 checkerboard, not a thick
  // horizontal+vertical lattice. Each black/white "cell" is 2×2 source
  // pixels, so the cell reads as a clear "市松模様" rather than a solid
  // 50% gray (the 1×1 checkerboard mediumGray we already ship).
  darkGrid:  [0b11001100, 0b11001100, 0b00110011, 0b00110011, 0b11001100, 0b11001100, 0b00110011, 0b00110011],
  // lightGrid: a sparse grid with 4-px pitch — horizontal lines at rows
  // 0 and 4, vertical lines at cols 0 and 4. Grid cells (white squares
  // between the lines) are 3×3 source pixels, matching Excel's rendering
  // of lightGrid where the lattice is clearly larger than the dense
  // version and a true "格子" pattern is recognisable.
  lightGrid: [0b11111111, 0b10001000, 0b10001000, 0b10001000, 0b11111111, 0b10001000, 0b10001000, 0b10001000],

  // ── 8×8 — diagonals & trellis ─────────────────────────────────────────
  // darkDown / darkUp: 2-px-wide diagonal stripes every 4 cells.
  // lightDown / lightUp: 1-px-wide diagonal stripes at the same 4-cell pitch.
  darkDown:   [0b11001100, 0b01100110, 0b00110011, 0b10011001, 0b11001100, 0b01100110, 0b00110011, 0b10011001],
  lightDown:  [0b10001000, 0b01000100, 0b00100010, 0b00010001, 0b10001000, 0b01000100, 0b00100010, 0b00010001],
  darkUp:     [0b00110011, 0b01100110, 0b11001100, 0b10011001, 0b00110011, 0b01100110, 0b11001100, 0b10011001],
  lightUp:    [0b00010001, 0b00100010, 0b01000100, 0b10001000, 0b00010001, 0b00100010, 0b01000100, 0b10001000],
  // darkTrellis = darkDown | darkUp.
  darkTrellis:  [0b11111111, 0b01100110, 0b11111111, 0b10011001, 0b11111111, 0b01100110, 0b11111111, 0b10011001],
  // lightTrellis = lightDown | lightUp.
  lightTrellis: [0b10011001, 0b01100110, 0b01100110, 0b10011001, 0b10011001, 0b01100110, 0b01100110, 0b10011001],
};

/**
 * Build a repeating 8x8 tile for an ECMA-376 preset pattern. Returns null
 * for unknown pattern names so the caller can fall back to a flat colour.
 */
function hatchPattern(
  ctx: CanvasRenderingContext2D,
  pt: string,
  fgHex: string,
  bgHex: string,
): CanvasPattern | null {
  // Read the destination context's effective scale so the offscreen tile is
  // sized at the destination's device-pixel resolution. Drawing the source
  // pre-scaled keeps each "1 bit" at exactly one destination *device* pixel
  // when the rendering context is doing ctx.scale(dpr * cs) for hi-DPI /
  // Storybook zoom — without this the crisp 1-px source bits get linearly
  // resampled to 1.5–2 device px and either chunky (integer scale) or
  // blurred-out-of-existence (non-integer scale, the prior pat.setTransform
  // approach lost lightHorizontal entirely at scale=1.5).
  const t = ctx.getTransform();
  const sx = Math.max(1, Math.round(Math.hypot(t.a, t.b)));
  const sy = Math.max(1, Math.round(Math.hypot(t.c, t.d)));
  const key = `${pt}|${fgHex}|${bgHex}|${sx}|${sy}`;
  if (PATTERN_CACHE.has(key)) return PATTERN_CACHE.get(key)!;

  const rows = ownValue(PATTERN_BITMAPS, pt);
  if (!rows) {
    PATTERN_CACHE.set(key, null);
    return null;
  }

  // Square tile — size inferred from the row count. Bit (size-1) is
  // leftmost so the binary literals read left-to-right at any tile width.
  // The offscreen canvas is sized at sx × tileSize device pixels so each
  // "1 bit" becomes a sx × sy rect when the destination context's scale
  // multiplies it back. Result: the source bit lands on exactly one
  // destination device pixel regardless of the user's CSS zoom.
  const tile = rows.length;
  const off = createAuxCanvas(tile, tile);
  if (!off) { PATTERN_CACHE.set(key, null); return null; }
  const octx = off.getContext('2d') as
    | CanvasRenderingContext2D
    | OffscreenCanvasRenderingContext2D
    | null;
  if (!octx) { PATTERN_CACHE.set(key, null); return null; }

  octx.fillStyle = hexToRgba(bgHex);
  octx.fillRect(0, 0, tile, tile);
  octx.fillStyle = hexToRgba(fgHex);
  for (let y = 0; y < tile; y++) {
    const row = rows[y];
    for (let x = 0; x < tile; x++) {
      if (row & (1 << (tile - 1 - x))) octx.fillRect(x, y, 1, 1);
    }
  }

  const pat = ctx.createPattern(off, 'repeat');
  // Pre-bake an inverse-scale matrix only when the scale rounds to an
  // integer (1, 2, 3, …). Fractional inverses (e.g. 1/1.5 ≈ 0.67) trigger
  // the canvas's bilinear pattern resampler and smear 1-px features into
  // a uniform half-tone — which is exactly what wiped out lightHorizontal
  // at Storybook scale=1.5. For non-integer ctx scales we leave the tile
  // unscaled and accept ~1.5-device-px bits (slightly chunky but crisp).
  if (pat && typeof DOMMatrix !== 'undefined' && (sx >= 2 || sy >= 2)) {
    const m = new DOMMatrix();
    m.scaleSelf(1 / sx, 1 / sy);
    pat.setTransform(m);
  }
  PATTERN_CACHE.set(key, pat);
  return pat;
}

/**
 * Paint a cell background according to its <patternFill> / <gradientFill>.
 *
 * ECMA-376 §18.8.20 / §18.8.22 specifies fg/bg defaults: when the colour
 * children are absent, fg defaults to the system foreground (black) and bg
 * to the system background (white). Without those defaults the directional
 * hatches that Excel emits with no explicit colours (`<patternFill
 * patternType="darkHorizontal"/>` etc.) would render as nothing because
 * the prior gate required a non-null fgColor.
 *
 * Returns true when the cell was painted so the caller can short-circuit
 * its tableStyle / banded fallbacks.
 */
export function paintCellPatternFill(
  ctx: CanvasRenderingContext2D,
  fill: CellFill,
  x: number, y: number, w: number, h: number,
): boolean {
  if (fill.gradient && fill.gradient.stops.length > 0) {
    ctx.fillStyle = buildGradientFill(ctx, fill.gradient, x, y, w, h);
    ctx.fillRect(x, y, w, h);
    return true;
  }
  const pt = fill.patternType;
  if (!pt || pt === 'none') return false;
  const fg = fill.fgColor ?? '000000';
  const bg = fill.bgColor ?? 'FFFFFF';
  if (pt === 'solid') {
    ctx.fillStyle = hexToRgba(fg);
    ctx.fillRect(x, y, w, h);
    return true;
  }
  const hatch = hatchPattern(ctx, pt, fg, bg);
  if (hatch) {
    ctx.fillStyle = hatch;
  } else {
    const coverage = patternCoverage(pt);
    ctx.fillStyle = coverage >= 1
      ? hexToRgba(fg)
      : blendHex(fg, bg, coverage);
  }
  ctx.fillRect(x, y, w, h);
  return true;
}

/**
 * Build a Canvas gradient object for an xlsx `<gradientFill>`. Linear uses
 * the degree attribute (0° = left→right, 90° = top→bottom). Path gradients
 * radiate from a rectangular inner bounds defined by left/right/top/bottom
 * as fractions of the cell.
 */
function buildGradientFill(
  ctx: CanvasRenderingContext2D,
  g: GradientFillSpec,
  x: number, y: number, w: number, h: number,
): CanvasGradient {
  let grad: CanvasGradient;
  if (g.gradientType === 'path') {
    // Use the inner rectangle's center as the radial origin; radius spans to
    // the farthest cell corner so stop=1 always reaches a cell edge.
    const cxg = x + w * (g.left + (1 - g.right - g.left) / 2);
    const cyg = y + h * (g.top + (1 - g.bottom - g.top) / 2);
    const r = Math.hypot(Math.max(cxg - x, x + w - cxg), Math.max(cyg - y, y + h - cyg));
    grad = ctx.createRadialGradient(cxg, cyg, 0, cxg, cyg, r);
  } else {
    // Linear: rotate around the cell's center and extend to the bounds.
    const rad = (g.degree * Math.PI) / 180;
    const cxg = x + w / 2;
    const cyg = y + h / 2;
    const ext = (Math.abs(Math.cos(rad)) * w + Math.abs(Math.sin(rad)) * h) / 2;
    grad = ctx.createLinearGradient(
      cxg - Math.cos(rad) * ext, cyg - Math.sin(rad) * ext,
      cxg + Math.cos(rad) * ext, cyg + Math.sin(rad) * ext,
    );
  }
  for (const stop of g.stops) {
    const pos = Math.min(1, Math.max(0, stop.position));
    grad.addColorStop(pos, hexToRgba(stop.color));
  }
  return grad;
}

/** Linear-interpolate two #RRGGBB values in RGB space at coverage fg weight. */
function blendHex(fgHex: string, bgHex: string, fgCoverage: number): string {
  const fh = fgHex.replace('#', '');
  const bh = bgHex.replace('#', '');
  const fr = parseInt(fh.slice(0, 2), 16);
  const fg = parseInt(fh.slice(2, 4), 16);
  const fb = parseInt(fh.slice(4, 6), 16);
  const br = parseInt(bh.slice(0, 2), 16);
  const bg = parseInt(bh.slice(2, 4), 16);
  const bb = parseInt(bh.slice(4, 6), 16);
  const c = Math.min(1, Math.max(0, fgCoverage));
  const r = Math.round(fr * c + br * (1 - c));
  const g = Math.round(fg * c + bg * (1 - c));
  const b = Math.round(fb * c + bb * (1 - c));
  return `rgb(${r},${g},${b})`;
}
