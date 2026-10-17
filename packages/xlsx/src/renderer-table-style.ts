// XLSX table-style resolution: per-cell table style map, header/total and
// stripe colours, overlay borders, and table font dxfs. Split from
// renderer.ts (#1857).

import type { Worksheet, Styles, Border, CellXf, Dxf } from './types.js';
import {
  assertCoordinateRangeArea,
  setCoordinateIndexValue,
  type CoordinateIndexIdentity,
} from './renderer-coordinate-index.js';
import { getWorksheetPolicy } from './worksheet-policy-context.js';

// ────────────────────────────────────────────────────────────────
// Excel Table style overlays (ECMA-376 §18.5 / §18.8.83)
// ────────────────────────────────────────────────────────────────
// Two distinct paths:
//
// 1. CUSTOM styles (defined in the file's `<tableStyles>` block, §18.5.1.2):
//    rendered strictly from the dxfs of their declared `<tableStyleElement>`s.
//    A custom style contributes ONLY what its elements define — if none carry
//    a border, Excel draws no table-level border (the visible structure comes
//    from theme borders baked into each cell `xf`, which we render separately).
//    `isCustom` cells therefore never get accent synthesis.
//
// 2. BUILT-IN style names (`TableStyleLight18`, …) whose definitions are NOT
//    in the file. We don't yet ship the preset catalog, so we approximate them
//    from a single "accent" color: bold header + banded fills + horizontal
//    rules, so those files render with visible structure rather than blank
//    ranges. This is an approximation, kept until a real catalog ships
//    (post-1.0).
export interface TableCellStyle {
  accent: string;
  /** `true` for a custom `<tableStyle>` — disables all accent approximation. */
  isCustom: boolean;
  isHeader: boolean;
  isTotals: boolean;
  /** `true` when this is a banded data row that should get the stripe fill. */
  isBanded: boolean;
  isFirstCol: boolean;
  isLastCol: boolean;
  isTopEdge: boolean;
  isBottomEdge: boolean;
  /** Dxf for the whole-table element of a custom `<tableStyle>`
   *  (ECMA-376 §18.8.83). Border/fill apply to every cell as a base layer. */
  wholeTableDxf?: number;
  /** Dxf for the header-row element of a custom `<tableStyle>`. Provides
   *  header fill, font color/weight, and vertical separators. */
  headerRowDxf?: number;
  /** Dxf for the total-row element. */
  totalRowDxf?: number;
  /** Dxf for the first-column element (when `showFirstColumn`). */
  firstColumnDxf?: number;
  /** Dxf for the last-column element (when `showLastColumn`). */
  lastColumnDxf?: number;
  /** Dxf for the stripe (band1=odd, band2=even) that applies to this row when
   *  `showRowStripes` is set; undefined when this row is not a stripe. */
  stripeDxf?: number;
}

export function buildTableStyleMap(worksheet: Worksheet): Map<string, TableCellStyle> {
  const map = new Map<string, TableCellStyle>();
  const identity = coordinateIndexIdentity(
    worksheet,
    'worksheet-table-style-index',
    'expand-styled-table-coordinates',
  );
  for (const t of worksheet.tables ?? []) {
    // ECMA-376 §18.5.1.4: empty styleName means the table has "None" style —
    // no visual table formatting overlay should be applied. Cell xf borders
    // and fills are still rendered as defined (per §18.8.45); only the
    // table-style overlay (banded fills, separator rules, etc.) is skipped.
    if (!t.styleName) continue;
    const { top, bottom, left, right } = t.range;
    assertCoordinateRangeArea(t.range, identity);
    const accent = t.accentColor || '#808080';
    const isCustom = !!t.isCustom;
    const hdr = Math.max(0, t.headerRowCount ?? 1);
    const tot = Math.max(0, t.totalsRowCount ?? 0);
    const headerEnd = top + hdr - 1;
    const totalsStart = bottom - tot + 1;
    for (let r = top; r <= bottom; r++) {
      const isHeader = hdr > 0 && r <= headerEnd;
      const isTotals = tot > 0 && r >= totalsStart;
      const dataIdx = (!isHeader && !isTotals) ? (r - headerEnd - 1) : -1;
      // Row banding stripes alternate band1 (odd) / band2 (even) over the data
      // region (§18.18.93). For built-in approximation we only paint the odd
      // stripe via `isBanded`; for custom styles we pick the matching dxf.
      const isStripeRow = t.showRowStripes && dataIdx >= 0;
      const stripeDxf = isStripeRow
        ? (dataIdx % 2 === 1 ? t.band1HorizontalDxf : t.band2HorizontalDxf)
        : undefined;
      for (let c = left; c <= right; c++) {
        setCoordinateIndexValue(map, `${r}:${c}`, {
          accent,
          isCustom,
          isHeader,
          isTotals,
          isBanded: t.showRowStripes && dataIdx >= 0 && dataIdx % 2 === 1,
          isFirstCol: t.showFirstColumn && c === left,
          isLastCol: t.showLastColumn && c === right,
          isTopEdge: r === top,
          isBottomEdge: r === bottom,
          wholeTableDxf: t.wholeTableDxf,
          headerRowDxf: t.headerRowDxf,
          totalRowDxf: t.totalRowDxf,
          firstColumnDxf: t.firstColumnDxf,
          lastColumnDxf: t.lastColumnDxf,
          stripeDxf,
        }, identity);
      }
    }
  }
  return map;
}

/**
 * Result of resolving a table cell's overlay *border* (ECMA-376 §18.8.83).
 *  - `none`   — draw nothing extra (custom style with no border dxf, or a
 *               built-in cell that has no rule on this edge).
 *  - `dxf`    — draw `border` exactly as the resolved dxf defines it.
 *  - `accent` — built-in approximation only: a synthesized accent-colored rule
 *               under the cell (plus the table top edge when `topEdge`).
 */
export type TableOverlayBorder =
  | { kind: 'none' }
  | { kind: 'dxf'; border: Border }
  | { kind: 'accent'; color: string; lineWidth: number; topEdge: boolean };

/**
 * Decide the table-style overlay border for one cell.
 *
 * `dxfWhole` / `dxfHeader` are the already-resolved dxfs for the cell's
 * wholeTable / headerRow elements (undefined when the element is absent).
 * `colIndex` is the cell's column index *within the table* (0 = first column),
 * used to draw the table's outer-left edge only on the leftmost column.
 *
 * Custom styles (`ts.isCustom`) draw only from these dxfs and never synthesize
 * accent rules: when no border dxf is present they contribute nothing. Built-in
 * style names keep the accent approximation.
 */
export function tableOverlayBorder(
  ts: TableCellStyle,
  dxfWhole: Dxf | undefined,
  dxfHeader: Dxf | undefined,
  colIndex: number,
): TableOverlayBorder {
  const horiz = dxfWhole?.border?.horizontal;
  const wtTop = dxfWhole?.border?.top;
  const wtBot = dxfWhole?.border?.bottom;
  const wtLeft = dxfWhole?.border?.left;
  const wtRight = dxfWhole?.border?.right;
  const hdrBot = dxfHeader?.border?.bottom;
  const hdrTop = dxfHeader?.border?.top;
  const hasDxfBorder = !!(horiz || wtTop || wtBot || wtLeft || wtRight || hdrBot || hdrTop);

  if (hasDxfBorder) {
    // Compose per-edge from the table-style hierarchy
    // (wholeTable < headerRow), §18.8.83. The inner `horizontal` rule fills
    // top/bottom for interior rows; outer edges come from the table extents.
    const overlay: Border = { left: null, right: null, top: null, bottom: null };
    if (ts.isTopEdge) overlay.top = wtTop ?? null;
    else if (horiz) overlay.top = horiz;
    if (ts.isHeader && hdrBot) overlay.bottom = hdrBot;
    else if (ts.isBottomEdge) overlay.bottom = wtBot ?? null;
    else if (horiz) overlay.bottom = horiz;
    if (ts.isFirstCol || colIndex === 0) overlay.left = wtLeft ?? null;
    if (ts.isLastCol) overlay.right = wtRight ?? null;
    return { kind: 'dxf', border: overlay };
  }

  // No border dxf. A custom style contributes nothing here (§18.5.1.2):
  // Excel draws no table-level border, only the theme borders baked into the
  // cell xf. Built-in names fall through to the accent approximation.
  if (ts.isCustom) return { kind: 'none' };
  return { kind: 'accent', color: ts.accent, lineWidth: ts.isHeader ? 1.5 : 1, topEdge: ts.isTopEdge };
}

export function stripeColorFor(accent: string): string {
  // Library approximation, not the spec rule: a 20% accent / 80% white mix for
  // banded rows of a built-in table style whose stripe dxf is not resolved.
  // The normative definitions are the built-in table styles shipped with
  // ECMA-376 Part 1 (SpreadsheetML presetTableStyles.xml: per-style dxfs with
  // theme colour + tint, applied per §18.8.40 tableStyle / §18.5.1.5
  // tableStyleInfo); no Excel measurement backs the 20% value. Resolving those
  // definitions changes table rendering and is tracked as follow-up work.
  const hex = accent.replace('#', '');
  if (hex.length < 6) return '#F2F2F2';
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const mix = (ch: number) => Math.round(ch * 0.2 + 255 * 0.8);
  const toHex = (v: number) => v.toString(16).padStart(2, '0').toUpperCase();
  return `#${toHex(mix(r))}${toHex(mix(g))}${toHex(mix(b))}`;
}

export function coordinateIndexIdentity(
  worksheet: Worksheet,
  resource: string,
  operation: string,
): CoordinateIndexIdentity {
  return {
    resource,
    operation,
    limit: getWorksheetPolicy(worksheet).maxCoordinateIndexEntries,
  };
}

/** The table-element dxf that supplies a table cell's font (header / total /
 *  last / first column / stripe / whole table), by the same §18.8.83 element
 *  hierarchy the main paint path uses. */
export function tableFontDxfFor(tableStyle: TableCellStyle | undefined, styles: Styles): Dxf | undefined {
  if (!tableStyle) return undefined;
  const dxfList = styles.dxfs ?? [];
  return tableStyle.isHeader
    ? dxfList[tableStyle.headerRowDxf ?? -1]
    : tableStyle.isTotals
      ? dxfList[tableStyle.totalRowDxf ?? -1]
      : tableStyle.isLastCol && tableStyle.lastColumnDxf != null
        ? dxfList[tableStyle.lastColumnDxf]
        : tableStyle.isFirstCol && tableStyle.firstColumnDxf != null
          ? dxfList[tableStyle.firstColumnDxf]
          : tableStyle.stripeDxf != null
            ? dxfList[tableStyle.stripeDxf]
            : dxfList[tableStyle.wholeTableDxf ?? -1];
}

/** A table-element dxf's font color, unless the cell's font color is its own
 *  formatting (`ownFontColor`, resolved by the parser against the Normal
 *  style): Excel draws that color over the table style's. */
export function tableStyleFontColor(tableFontDxf: Dxf | undefined, xf: CellXf): string | null {
  return xf.ownFontColor ? null : tableFontDxf?.font?.color ?? null;
}
