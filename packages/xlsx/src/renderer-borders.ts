// XLSX cell and merged-cell border painting. Split from renderer.ts (#1857).

import type { Styles, Cell, Border, BorderEdge } from './types.js';
import { crispOffset, xlsxBorderDashArray, hexToRgba } from '@silurus/ooxml-core';
import { resolveXf } from './renderer-cell-text.js';

// ────────────────────────────────────────────────────────────────
// Border drawing
// ────────────────────────────────────────────────────────────────
/**
 * Overlay any CF-rule border edges on top of the cell's base border. CF
 * borders win per-edge where they set a style (e.g. a red left+right for a
 * "today" column marker replaces the underlying edge only, leaving top/bottom
 * from the base style intact).
 */
/**
 * Resolve the outer border of a merged range. Excel keeps each constituent
 * cell's own style; the merged rectangle's outer edges come from the cells
 * on those edges (e.g. the right edge from the rightmost column's `right`),
 * not from the anchor cell alone. Without this the right border of an
 * `E2:F2` merge — stored on F2 — goes missing.
 */
export function resolveMergeBorder(
  anchorBorder: Border,
  anchorRow: number,
  anchorCol: number,
  rightCol: number,
  bottomRow: number,
  cellMap: Map<string, Cell>,
  styles: Styles,
): Border {
  if (rightCol === anchorCol && bottomRow === anchorRow) return anchorBorder;
  const edgeBorder = (r: number, c: number): Border | null => {
    if (r === anchorRow && c === anchorCol) return null;
    const cell = cellMap.get(`${r}:${c}`);
    if (!cell) return null;
    return resolveXf(styles, cell.styleIndex ?? 0).border;
  };
  const rightB  = edgeBorder(anchorRow, rightCol);
  const bottomB = edgeBorder(bottomRow, anchorCol);
  const cornerB = edgeBorder(bottomRow, rightCol);
  const pick = (primary: BorderEdge | null | undefined, ...rest: Array<BorderEdge | null | undefined>): BorderEdge | null => {
    if (primary?.style) return primary;
    for (const r of rest) if (r?.style) return r;
    return primary ?? null;
  };
  return {
    left:         anchorBorder.left,
    top:          anchorBorder.top,
    right:        pick(rightB?.right,   cornerB?.right,   anchorBorder.right),
    bottom:       pick(bottomB?.bottom, cornerB?.bottom,  anchorBorder.bottom),
    diagonalUp:   anchorBorder.diagonalUp ?? null,
    diagonalDown: anchorBorder.diagonalDown ?? null,
  };
}

export function mergeBorders(base: Border, overlay: Border | undefined): Border {
  if (!overlay) return base;
  const pick = (a: BorderEdge | null | undefined, b: BorderEdge | null | undefined): BorderEdge | null =>
    (b && b.style) ? b : (a ?? null);
  return {
    left:         pick(base.left,         overlay.left),
    right:        pick(base.right,        overlay.right),
    top:          pick(base.top,          overlay.top),
    bottom:       pick(base.bottom,       overlay.bottom),
    diagonalUp:   pick(base.diagonalUp,   overlay.diagonalUp),
    diagonalDown: pick(base.diagonalDown, overlay.diagonalDown),
  };
}

/**
 * Endpoint offsets for the two rails of an OOXML `double` cell border.
 *
 * Excel only extends/trims a rail at a corner that actually has a
 * perpendicular border to join.  A standalone bottom border, for example,
 * keeps both rails at the full cell width; applying the boxed-cell corner
 * treatment there makes the inner (upper) rail visibly too short.
 */
export function doubleBorderRailEndpoints(
  start: number,
  end: number,
  startJoined: boolean,
  endJoined: boolean,
  off = 1,
): { outerStart: number; outerEnd: number; innerStart: number; innerEnd: number } {
  return {
    outerStart: startJoined ? start - off : start,
    outerEnd: endJoined ? end + off : end,
    innerStart: startJoined ? start + off : start,
    innerEnd: endJoined ? end - off : end,
  };
}

export function renderBorder(
  ctx: CanvasRenderingContext2D,
  border: Border,
  x: number, y: number, w: number, h: number,
  dpr = 1,
): void {
  type EdgeRef = {
    edge: BorderEdge | null | undefined;
    x1: number; y1: number; x2: number; y2: number;
    /** 'h' = horizontal (top/bottom), 'v' = vertical (left/right),
     *  'd' = diagonal (no doubling support). */
    kind: 'h' | 'v' | 'd';
  };
  const edges: EdgeRef[] = [
    { edge: border.top,         x1: x,     y1: y,     x2: x + w, y2: y,     kind: 'h' },
    { edge: border.bottom,      x1: x,     y1: y + h, x2: x + w, y2: y + h, kind: 'h' },
    { edge: border.left,        x1: x,     y1: y,     x2: x,     y2: y + h, kind: 'v' },
    { edge: border.right,       x1: x + w, y1: y,     x2: x + w, y2: y + h, kind: 'v' },
    { edge: border.diagonalUp,  x1: x,     y1: y + h, x2: x + w, y2: y,     kind: 'd' },
    { edge: border.diagonalDown,x1: x,     y1: y,     x2: x + w, y2: y + h, kind: 'd' },
  ];
  for (const { edge, x1, y1, x2, y2, kind } of edges) {
    if (!edge || !edge.style || edge.style === 'none') continue;
    const color = edge.color ? hexToRgba(edge.color) : '#000000';
    // ECMA-376 §18.18.3 ST_BorderStyle "double": two parallel thin lines
    // with a small gap. Drawn as a 1-px line on either side of the cell
    // edge so the pair reads as a ~3-px-thick double rule, matching Excel.
    // The outer line extends past the cell corners by `off` so adjacent
    // doubled edges close cleanly; the inner line is *shortened* by `off`
    // on each end so the two perpendicular inner lines meet exactly at the
    // inner corner (without the trim they cross past the corner forming a
    // small "+" overhang). Diagonals are handled separately further down.
    if (edge.style === 'double' && kind === 'd') {
      // Two parallel diagonal lines offset perpendicular to the diagonal
      // direction. Excel renders <diagonal style="double"/> the same way it
      // does horizontal/vertical doubles — two thin lines with a small gap.
      ctx.strokeStyle = color;
      ctx.lineWidth = 1;
      ctx.setLineDash([]);
      const dx = x2 - x1;
      const dy = y2 - y1;
      const len = Math.hypot(dx, dy);
      // Perpendicular unit vector. ±off shifts the line ~1 logical px to either side.
      const off = 1;
      const px = (-dy / len) * off;
      const py = (dx / len) * off;
      ctx.beginPath();
      ctx.moveTo(x1 + px, y1 + py); ctx.lineTo(x2 + px, y2 + py);
      ctx.moveTo(x1 - px, y1 - py); ctx.lineTo(x2 - px, y2 - py);
      ctx.stroke();
      continue;
    }
    if (edge.style === 'double' && kind !== 'd') {
      // NB: core's `fillDoubleBorder` (used by docx) is intentionally NOT used
      // here. Excel's double is two symmetric 1px strokes at ±1px with corner
      // extension/trim so a double BOX closes cleanly; the core fill helper's
      // floored-thirds band has a +0.5-device centring bias and no corner
      // handling, which measured ~0.3% further from the Excel references
      // (sample-11) when trialled. These are different conventions, not
      // duplicated math — keep the Excel-faithful stroke model.
      ctx.strokeStyle = color;
      ctx.lineWidth = 1;
      ctx.setLineDash([]);
      const off = 1;
      ctx.beginPath();
      if (kind === 'h') {
        const isTop = y1 === y;
        // Outer line extends past the corners by `off` so adjacent doubled
        // edges close; the inner line is shortened by `off` on each end so
        // perpendicular inner lines meet exactly at the inner corner. In the
        // two-pass order each side draws its own double fully (no fill erases
        // it), so no inherited-redraw "swap" is needed.
        const outerY = isTop ? y - off : y + h + off;
        const innerY = isTop ? y + off : y + h - off;
        const rails = doubleBorderRailEndpoints(
          x,
          x + w,
          !!border.left?.style && border.left.style !== 'none',
          !!border.right?.style && border.right.style !== 'none',
          off,
        );
        ctx.moveTo(rails.outerStart, outerY); ctx.lineTo(rails.outerEnd, outerY);
        ctx.moveTo(rails.innerStart, innerY); ctx.lineTo(rails.innerEnd, innerY);
      } else {
        const isLeft = x1 === x;
        const outerX = isLeft ? x - off : x + w + off;
        const innerX = isLeft ? x + off : x + w - off;
        const rails = doubleBorderRailEndpoints(
          y,
          y + h,
          !!border.top?.style && border.top.style !== 'none',
          !!border.bottom?.style && border.bottom.style !== 'none',
          off,
        );
        ctx.moveTo(outerX, rails.outerStart); ctx.lineTo(outerX, rails.outerEnd);
        ctx.moveTo(innerX, rails.innerStart); ctx.lineTo(innerX, rails.innerEnd);
      }
      ctx.stroke();
      continue;
    }
    ctx.beginPath();
    ctx.strokeStyle = color;
    // Logical-px width (thin=1, medium=2, thick=3). ctx.scale(dpr,dpr) scales it
    // to device px, so a thin border is 2 device px at dpr=2 — matching Excel's
    // measured on-screen weight. Do NOT divide by dpr (that halved it to 1 px).
    const lw = borderStyleWidth(edge.style);
    ctx.lineWidth = lw;
    const dash = borderStyleDash(edge.style);
    ctx.setLineDash(dash);
    // Crispness snap (see crispOffset): snap the stroke to the nearest crisp
    // device position derived from its own coordinate — the x of a vertical
    // edge, the y of a horizontal edge. An odd device-width stroke lands on a
    // pixel midpoint; an even device width snaps to an integer boundary (0 when
    // already aligned). Diagonals can't be pixel-aligned, so no offset.
    const dpx = kind === 'v' ? crispOffset(x1, lw, dpr) : 0;
    const dpy = kind === 'h' ? crispOffset(y1, lw, dpr) : 0;
    ctx.moveTo(x1 + dpx, y1 + dpy);
    ctx.lineTo(x2 + dpx, y2 + dpy);
    ctx.stroke();
    ctx.setLineDash([]);
  }
}

function borderStyleWidth(style: string): number {
  // Widths follow Excel's pt convention so medium and thick read as visibly
  // distinct from thin (and from each other) — thin=1pt, medium=2pt, thick=3pt.
  // Earlier values (1.5 / 2) compressed medium and thick to roughly the same
  // 2-row antialiased band, which made e.g. sample-27 row 9 appear to have
  // top (medium) and bottom (thick) of equal weight.
  switch (style) {
    case 'thick': return 3;
    case 'medium': case 'mediumDashed': case 'mediumDashDot': case 'mediumDashDotDot': case 'slantDashDot': return 2;
    case 'hair': return 0.5;
    default: return 1;
  }
}

/**
 * ECMA-376 §18.18.3 ST_BorderStyle dash families → a static-pixel `setLineDash`
 * pattern. Thin wrapper over core's shared `xlsxBorderDashArray` (which owns the
 * §18.18.3 relative table); Excel cell borders use a constant pixel cadence
 * regardless of the (sub-pixel) hairline width, so unlike docx/pptx this does
 * NOT scale with the stroked width. Returns `[]` for solid styles.
 */
function borderStyleDash(style: string): number[] {
  return xlsxBorderDashArray(style);
}

/**
 * Visual precedence per ECMA-376 ST_BorderStyle. Higher = stronger / more
 * visually prominent. Excel's behaviour at a shared edge between two adjacent
 * cells that both define a border is to render the stronger one — this lets
 * the renderer pick a single edge to draw instead of stacking two strokes
 * (which compounds antialiasing and visually thickens the line).
 */
function borderPrecedence(style: string | null | undefined): number {
  switch (style) {
    case 'double': return 13;
    case 'thick': return 12;
    case 'medium': return 11;
    case 'mediumDashed': return 10;
    case 'mediumDashDot': return 9;
    case 'slantDashDot': return 8;
    case 'mediumDashDotDot': return 7;
    case 'thin': return 6;
    case 'dashed': return 5;
    case 'dashDot': return 4;
    case 'dashDotDot': return 3;
    case 'dotted': return 2;
    case 'hair': return 1;
    default: return 0;
  }
}

export function pickStrongerEdge(
  a: BorderEdge | null | undefined,
  b: BorderEdge | null | undefined,
): BorderEdge | null {
  const aP = borderPrecedence(a?.style);
  const bP = borderPrecedence(b?.style);
  if (aP === 0 && bP === 0) return null;
  if (aP >= bP) return a ?? null;
  return b ?? null;
}
