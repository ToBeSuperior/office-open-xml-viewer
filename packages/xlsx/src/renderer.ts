
import { inheritCellFontLanguage } from './internal/cell-font-language.js';
import type { CjkLang } from '@silurus/ooxml-core';
import type { PivotCellFormat } from './pivot-style.js';
import type {
  Worksheet,
  Styles,
  Cell,
  CellFont,
  ViewportRange,
  RenderViewportOptions,
  XlsxTextRunInfo,
} from './types.js';
import type { ChartThreeDRenderer, ChartRegionMapRenderer } from '@silurus/ooxml-core';
import {
  crispOffset,
  renderSparkline,
  hexToRgba,
  verticalTrLongMark,
  verticalVertGlyphReachable,
  type SparklineModel,
} from '@silurus/ooxml-core';
import { formatCellValueWithColor } from './number-format.js';
import { type CfContext, evaluateCf } from './conditional-format.js';
import type { CfDiagnosticCollector } from './cf-diagnostics.js';
import { cellBaseRtl } from './bidi-line.js';
import { formatA1 } from './a1.js';
import { drawStackedVerticalChar } from './vertical-text.js';
import { setCoordinateIndexValue } from './renderer-coordinate-index.js';
import { MAX_WORKSHEET_COL } from './internal/grid-geometry.js';
import type { GridAxisGeometry } from './internal/grid-axis-geometry.js';
import { inheritWorksheetPolicy } from './worksheet-policy-context.js';
import {
  colWidthToPx,
  pxToColWidth,
  rowHeightToPx,
  pxToRowHeight,
} from './internal/grid-metrics.js';
import {
  rtlMirrorX,
  getGridGeometryForWorksheet,
  bindXlsxOfficeFontRoutes,
  pinXlsxGridGeometry,
  HEADER_W,
  HEADER_H,
  DEFAULT_FONT_FAMILY,
} from './renderer-font-metrics.js';
import { drawCommentMarker } from './renderer-comment-marker.js';
import { paintCellPatternFill } from './renderer-cell-fill.js';
import { fillDataBar, drawCfIcon } from './renderer-cf-visuals.js';
import {
  resolveXf,
  effectiveCellStyleIndex,
  buildFont,
  generalHorizontalAlignment,
  drawHashFill,
  drawWrappedRichText,
  drawWrappedPlainText,
  drawNonWrapRichText,
  singleLineVerticalAnchor,
  paintCellText,
  vMetricPx,
  drawTextDecoLine,
  measureInFont,
  drawPhoneticBand,
} from './renderer-cell-text.js';
import {
  coordinateIndexIdentity,
  tableFontDxfFor,
  tableStyleFontColor,
  stripeColorFor,
  tableOverlayBorder,
  type TableCellStyle,
} from './renderer-table-style.js';
import { getSheetRenderCache, sheetRenderCache } from './renderer-sheet-cache.js';
import {
  layeredCellFont,
  styleFontColor,
  richRunsWithTableColor,
} from './renderer-auto-row-height.js';
import { renderAnchoredDrawings } from './renderer-drawings.js';
import {
  resolveMergeBorder,
  mergeBorders,
  renderBorder,
  pickStrongerEdge,
} from './renderer-borders.js';
import { renderSlicers } from './renderer-slicers.js';

// renderer.ts keeps its historical export surface (worksheet, viewer, render
// orchestrator, and tests import from here); the implementations live in the
// renderer-* modules split out for #1857.
export { drawSlicerFrame } from './renderer-slicers.js';
export { doubleBorderRailEndpoints } from './renderer-borders.js';
export {
  worksheetHasUncachedMath,
  prepareWorksheetMath,
  drawShapeText,
  imageCacheKey,
} from './renderer-drawings.js';
export {
  applyAutoRowHeights,
  hasPreparedAutoRowHeights,
  derivedAutoRowHeights,
  markAutoRowHeightsPrepared,
  invalidateAutoRowHeights,
} from './renderer-auto-row-height.js';
export { invalidateSheetRenderCache, getSheetRenderCache } from './renderer-sheet-cache.js';
export { buildTableStyleMap, tableOverlayBorder } from './renderer-table-style.js';
export type { TableCellStyle, TableOverlayBorder } from './renderer-table-style.js';
export {
  resolvedThemeCellFont,
  drawPhoneticBand,
  wrapParagraphLines,
  layoutRichTextLines,
  generalHorizontalAlignment,
  drawResolvedRichLine,
  drawNonWrapRichText,
  drawWrappedPlainText,
  drawWrappedRichText,
} from './renderer-cell-text.js';
export type { RichCellGeom } from './renderer-cell-text.js';
export {
  bindXlsxOfficeFontRoutes,
  bindXlsxWorksheetOfficeFontRoutes,
  pinXlsxGridGeometry,
  cssTailFor,
  fontStackFor,
  HEADER_W,
  HEADER_H,
  rtlMirrorX,
  sheetAnchoredRectX,
  computeMdw,
  getMdwForWorksheet,
  getGridGeometryForWorksheet,
} from './renderer-font-metrics.js';

export { colWidthToPx, pxToColWidth, rowHeightToPx, pxToRowHeight };

// Thin line drawn between frozen and scrollable areas
const FREEZE_LINE_COLOR = '#7a7a7a';

function colToLetter(col: number): string {
  let result = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    col = Math.floor((col - 1) / 26);
  }
  return result;
}

interface RenderContext {
  worksheet: Worksheet;
  styles: Styles;
  cellMap: Map<string, Cell>;
  mergeAnchorMap: Map<string, { totalW: number; totalH: number; right: number; bottom: number }>;
  mergeSkipSet: Set<string>;
  cfContext: CfContext;
  colWidths: number[];
  rowHeights: number[];
  colAxis: GridAxisGeometry;
  rowAxis: GridAxisGeometry;
  frozenColWidths: number[];
  frozenRowHeights: number[];
  frozenW: number;
  frozenH: number;
  startRow: number;
  startCol: number;
  cs: number;
  dpr: number;
  autoFilterCells: Set<string>;
  hyperlinkMap: Map<string, string>;
  /** row:col keys for cells that carry a comment; renderer draws a small
   *  red triangle in the top-right corner (ECMA-376 §18.7.3 commentList). */
  commentCells: Set<string>;
  /** row:col → table-style overlay (bold header, banded rows, borders). */
  tableStyleMap: Map<string, TableCellStyle>;
  pivotStyleMap: Map<string, PivotCellFormat>;
  /** row:col → render-ready SparklineModel for cells that host an
   *  `x14:sparkline`. Built once at viewport start by flattening the
   *  parser's SparklineGroup + per-cell Sparkline pair. */
  sparklineMap: Map<string, SparklineModel>;
  /** Off-screen cells whose unwrapped text reaches this viewport through empty
   *  neighbours. These anchors bypass the ordinary cell-rectangle cull so the
   *  existing overflow clip can paint the still-visible portion of the text. */
  overflowTextAnchors: Set<string>;
  /** Max Digit Width resolved for the workbook's default font, `<fonts>[0]`
   *  (ECMA-376 §18.3.1.13). Used by `colWidthToPx` to convert character-
   *  unit column widths into pixels. */
  mdw: number;
  onTextRun?: (info: XlsxTextRunInfo) => void;
  /** Sheet-level right-to-left grid mirror (ECMA-376 §18.3.1.87
   *  `<sheetView rightToLeft>`). When true the whole grid is mirrored:
   *  column A sits at the right edge, columns flow right-to-left, and the
   *  row-number header strip sits on the right of the canvas. Implemented
   *  by remapping each cell's canvas x to `canvasW - cx - cellW` (mirror
   *  about the cell-area band) — glyphs themselves are NOT flipped, and
   *  cell-level left/right alignment stays physical (Excel behavior). */
  rtl: boolean;
  /** Logical canvas width (device-independent px). Needed by the RTL
   *  mirror; identical to the value the top-level render fn computes. */
  canvasW: number;
  threeD?: ChartThreeDRenderer;
  regionMap?: ChartRegionMapRenderer;
  /** Operation-local CF diagnostics of this render (#1547). Passed only to
   *  the visible-frame evaluations in renderQuadrant; the overflow probe and
   *  auto row-height precomputation are off-screen and deliberately do not
   *  report. */
  cfDiagnostics?: CfDiagnosticCollector;
}

function drawAutoFilterArrow(ctx: CanvasRenderingContext2D, cx: number, cy: number, cw: number, ch: number): void {
  const sz = Math.max(6, Math.round(Math.min(cw, ch) * 0.45));
  const x = cx + cw - sz - 1;
  const y = cy + ch - sz - 1;
  ctx.save();
  ctx.fillStyle = '#D0D0D0';
  ctx.fillRect(x, y, sz, sz);
  ctx.fillStyle = '#444444';
  const tri = sz * 0.55;
  const tx = x + (sz - tri) / 2;
  const ty = y + (sz - tri * 0.5) / 2;
  ctx.beginPath();
  ctx.moveTo(tx, ty);
  ctx.lineTo(tx + tri, ty);
  ctx.lineTo(tx + tri / 2, ty + tri * 0.5);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

// ────────────────────────────────────────────────────────────────
// Render one rectangular region of cells
// ────────────────────────────────────────────────────────────────
function renderQuadrant(
  ctx: CanvasRenderingContext2D,
  rc: RenderContext,
  startRow: number, startCol: number,
  colWidths: number[], rowHeights: number[],
  colIndices: readonly number[], rowIndices: readonly number[],
  pixOffsetX: number, pixOffsetY: number,
  originX: number, originY: number,
  clipX: number, clipY: number, clipW: number, clipH: number,
  cjkFallback?: CjkLang,
): void {
  if (clipW <= 0 || clipH <= 0) return;

  const { styles, cellMap, mergeAnchorMap, mergeSkipSet, cfContext, cs, dpr } = rc;
  const numCols = colWidths.length;
  const numRows = rowHeights.length;
  const visibleRows = new Set(rowIndices);
  const visibleCols = new Set(colIndices);

  // RTL grid mirror (ECMA-376 §18.3.1.87). Maps a left-anchored cell rect
  // [x, x+w] to its mirror [canvasW - x - w, canvasW - x] within the
  // cell-area band. Because the LTR cell area is [hw, canvasW] and the RTL
  // cell area is [0, canvasW - hw] (header strip moves to the right), the
  // transform for any cell/quadrant collapses to `canvasW - x - w`. Applied
  // to every column x (so fills, borders, gridlines, text positions, merge
  // spans and the clip rect all mirror) while leaving glyphs un-flipped and
  // cell-level left/right alignment physical.
  const mirrorX = (x: number, w: number): number => rc.rtl ? rtlMirrorX(x, w, rc.canvasW) : x;

  // Canvas x for each column
  const colXs: number[] = [];
  let x = -pixOffsetX;
  for (let ci = 0; ci < numCols; ci++) {
    colXs.push(x);
    x += colWidths[ci];
  }

  // Canvas y for each row
  const rowYs: number[] = [];
  let y = -pixOffsetY;
  for (let ri = 0; ri < numRows; ri++) {
    rowYs.push(y);
    y += rowHeights[ri];
  }

  ctx.save();
  ctx.beginPath();
  ctx.rect(mirrorX(clipX, clipW), clipY, clipW, clipH);
  ctx.clip();

  // Deferred text drawing. Excel renders cell text on top of *all* cell
  // backgrounds, so a left-aligned overflow into an adjacent empty cell with
  // a white fill stays visible. If we drew fill+text per cell in one pass,
  // the next cell's fill would overpaint the previous cell's overflow.
  const textTasks: Array<() => void> = [];

  // Deferred merged-cell border drawing. Cell gridlines (#d0d0d0) are stroked
  // interleaved with cell borders inside the grid loop, at each cell's right +
  // bottom edge. For a 1×1 cell that ordering is harmless (within a row the
  // left neighbour is visited — and draws its right gridline — before the cell
  // strokes its own left border over it). A merged anchor, however, draws its
  // border ONCE at full span height/width during the anchor's row, while the
  // cells in the column to the left of the covered rows are visited in LATER
  // rows and stroke their right gridline over the merge's already-drawn left
  // border — shaving a device column off it (sample-30 B14:B23 left edge read
  // thinner/greyer below the anchor). Excel treats an explicit cell border as
  // replacing the gridline on that edge, so the merge perimeter must sit above
  // every gridline. Defer it to a pass flushed after the whole grid loop (but
  // before text, which stays on top — same z-order as the per-cell path).
  const mergeBorderTasks: Array<() => void> = [];

  // Deferred NON-merged cell borders + table-style overlays. Two-pass painting:
  // every cell fill (and the gray sheet gridlines) is laid down in the main
  // loop (pass 1); all styled cell borders run here (pass 2), AFTER every fill.
  // This removes the latent "a neighbour's fill over-paints a shared border"
  // bug structurally — once borders are deferred past all fills, no fill can
  // erase them — so the old `invertedTop`/`invertedLeft` inherit-and-repair
  // mechanism is gone (only genuine conflict resolution via `pickStrongerEdge`
  // remains). Mirrors docx PR #545 and the pptx renderer's two-pass order.
  const borderTasks: Array<() => void> = [];

  // Pre-pass: merge cells whose anchor lies outside this viewport quadrant but whose
  // span overlaps it (e.g. scrolled past the anchor row/col, or the anchor is in a
  // frozen quadrant while we are rendering the scrollable quadrant).
  for (const mc of rc.worksheet.mergeCells ?? []) {
    const aRow = mc.top, aCol = mc.left;
    // If anchor is within the main loop range, skip — handled normally below.
    if (visibleRows.has(aRow) && visibleCols.has(aCol)) continue;
    // Skip if merge span has no overlap with this viewport.
    if (!rowIndices.some((row) => row >= mc.top && row <= mc.bottom)) continue;
    if (!colIndices.some((col) => col >= mc.left && col <= mc.right)) continue;

    const info = rc.mergeAnchorMap.get(`${aRow}:${aCol}`);
    if (!info) continue;

    // Sparse cumulative axes keep far off-screen anchors O(log custom bands)
    // instead of walking every intervening row/column.
    let aCx = originX - pixOffsetX
      + rc.colAxis.offsetOf(aCol) - rc.colAxis.offsetOf(startCol);
    let aCy = originY - pixOffsetY
      + rc.rowAxis.offsetOf(aRow) - rc.rowAxis.offsetOf(startRow);

    const cW = info.totalW, cH = info.totalH;
    aCx = mirrorX(aCx, cW);
    const key = `${aRow}:${aCol}`;
    const cell = rc.cellMap.get(key);
    const { font, fill, border, xf } = resolveXf(
      styles, effectiveCellStyleIndex(rc.worksheet, cell, aCol),
    );
    const cf = evaluateCf(cell, aRow, aCol, cfContext, styles.dxfs ?? [], rc.cfDiagnostics);
    const effectiveFill = cf.fill ?? fill;
    // Same layers as the main path's merged anchor: PivotTable style fill
    // under the cell's own fill, and table / PivotTable font beneath CF.
    // (Table fills and banding are not composed here: Excel does not allow
    // merged cells inside a Table, see the table overlay note below.)
    const tableStyle = rc.tableStyleMap.get(key);
    const pivotFormat = rc.pivotStyleMap.get(key);

    if (!paintCellPatternFill(ctx, effectiveFill, aCx, aCy, cW, cH) && pivotFormat?.fill) {
      paintCellPatternFill(ctx, pivotFormat.fill, aCx, aCy, cW, cH);
    }
    if (cf.dataBar && cf.dataBar.ratio > 0) {
      const bInset = 2;
      const bW = Math.max(0, (cW - bInset * 2) * cf.dataBar.ratio);
      fillDataBar(ctx, cf.dataBar.color, aCx + bInset, aCy + bInset, bW, cH - bInset * 2, cf.dataBar.gradient);
    }
    // Defer this off-screen-anchored merge's border too: its in-viewport span
    // is exposed to the same gridline overpaint (left-column cells in the
    // covered rows stroke their right gridline over it in the main loop). Note
    // this pre-pass draws its own text *inline* below (the main loop defers
    // text via `textTasks`), so the deferred border ends up above this cell's
    // text rather than below it — visible only if a *thick* border grazed the
    // padding-inset, clipped text of an off-screen-anchored merge. Routing this
    // pre-pass text through `textTasks` for full gridline<border<text parity is
    // a follow-up; the current order is visually inert.
    const mergedBorder = resolveMergeBorder(border, aRow, aCol, info.right, info.bottom, rc.cellMap, styles);
    const preBorder = mergeBorders(mergedBorder, cf.border);
    mergeBorderTasks.push(() => renderBorder(ctx, preBorder, aCx, aCy, cW, cH, dpr));

    if (!cell) continue;
    const formatted = formatCellValueWithColor(cell, styles, cf.numFmt, rc.worksheet.date1904);
    const text = formatted.text;
    if (!text || (text === '0' && rc.worksheet.showZeros === false)) continue;

    const fontForDraw = layeredCellFont(font, cf, tableStyle, styles, pivotFormat);
    ctx.font = buildFont(ctx, fontForDraw, cs, cjkFallback, text);
    const hyperlinkUrl = rc.hyperlinkMap.get(key);
    // Colour precedence (as the main path): hyperlink theme colour >
    // conditional-formatting font colour > number-format section colour
    // ([Red] etc., §18.8.30) > table / PivotTable style colour > the cell's
    // own font colour (a cell's own colour beats the table's; see
    // tableStyleFontColor).
    const tableFontColor = styleFontColor(tableFontDxfFor(tableStyle, styles), pivotFormat, xf);
    const textColor = hyperlinkUrl
      ? '#0563C1'
      : (cf.fontColor ?? formatted.color ?? tableFontColor ?? font.color);
    ctx.fillStyle = textColor ? hexToRgba(textColor) : '#000000';

    const paddingX = 3, paddingY = 2;
    const alignH = xf.alignH ?? generalHorizontalAlignment(cell.value.type);
    const alignV = xf.alignV ?? 'bottom';
    // Indent: ECMA-376 §18.8.1 alignment@indent — one level indents by 3
    // character widths (MDW) of the workbook's normal-style font.
    const indentPx = xf.indent ? Math.round(xf.indent * 3 * rc.mdw) : 0;
    const leftPad = paddingX + (alignH === 'left' || !xf.alignH ? indentPx : 0);

    ctx.save();
    ctx.beginPath();
    ctx.rect(aCx, aCy, cW, cH);
    ctx.clip();

    if (formatted.fill) {
      drawHashFill(ctx, aCx, aCy, cW, cH, alignV, fontForDraw, cs);
      ctx.restore();
      continue;
    }

    let textX: number;
    if (alignH === 'right') { textX = aCx + cW - paddingX; ctx.textAlign = 'right'; }
    else if (alignH === 'center') { textX = aCx + cW / 2; ctx.textAlign = 'center'; }
    else { textX = aCx + leftPad; ctx.textAlign = 'left'; }

    // Text layout mirrors the main loop's merged-anchor path (see the
    // `xf.wrapText` branches below): wrapped cells split into lines via
    // wrapTextLines / layoutRichTextLines and flow vertically per alignV.
    // Using the same code keeps the off-screen-anchor pre-pass and the
    // in-viewport-anchor path identical, so a merged cell renders the same
    // text whether or not its top-left anchor cell is scrolled out of view.
    // Rich runs take the table style's font color by the measured run rule
    // (richRunsWithTableColor); runs without <rPr> through the base font.
    const tableRunColor = tableStyleFontColor(tableFontDxfFor(tableStyle, styles), xf);
    const rawRuns = cell.value.type === 'text' ? cell.value.runs : undefined;
    const runs = rawRuns && richRunsWithTableColor(rawRuns, tableRunColor);
    const richBaseFont = tableRunColor != null ? { ...fontForDraw, color: tableRunColor } : fontForDraw;
    const hasRichText = runs && runs.length > 0;

    if (xf.wrapText && hasRichText) {
      // Same helper as the in-viewport wrap path so an off-screen-anchored merge
      // renders identical wrapped rich text — per-run fonts, super/subscript,
      // underline/strike, and bidi (previously this pre-pass drew only plain
      // per-segment fonts).
      drawWrappedRichText(
        ctx, runs, richBaseFont,
        { alignH, alignV, cx: aCx, cy: aCy, cellW: cW, cellH: cH, leftPad, paddingX, paddingY },
        cs, dpr, { fontColor: cf.fontColor, readingOrder: xf.readingOrder }, cjkFallback
      );
    } else if (xf.wrapText) {
      drawWrappedPlainText(
        ctx,
        text,
        textX,
        fontForDraw,
        { alignH, alignV, cx: aCx, cy: aCy, cellW: cW, cellH: cH, leftPad, paddingX, paddingY },
        cs,
      );
    } else if (hasRichText) {
      // Non-wrap rich text — same helper as the in-viewport path so an
      // off-screen-anchored merge renders identical per-run text (single line, or
      // multiple lines on a hard break) instead of joined base-font text.
      drawNonWrapRichText(
        ctx, runs, richBaseFont,
        { alignH, alignV, cx: aCx, cy: aCy, cellW: cW, cellH: cH, leftPad, paddingX, paddingY },
        cs, dpr, { fontColor: cf.fontColor, readingOrder: xf.readingOrder }, cjkFallback
      );
    } else {
      const { baseline, textY } = singleLineVerticalAnchor({
        alignH, alignV, cx: aCx, cy: aCy, cellW: cW, cellH: cH, leftPad, paddingX, paddingY,
      });
      ctx.textBaseline = baseline;
      paintCellText(ctx, text, textX, textY, fontForDraw, cs);
    }
    ctx.restore();
  }

  for (let ri = 0; ri < numRows; ri++) {
    const rowIndex = rowIndices[ri];
    const cy = originY + rowYs[ri];
    const ch = rowHeights[ri];
    const rowOutsideClip = cy + ch <= clipY || cy >= clipY + clipH;
    if (rowOutsideClip) {
      // A merge anchor can remain in the materialized row-band list after its
      // own row rectangle has scrolled behind the header. Keep that row alive
      // when the full merged height still intersects the pane; the per-cell
      // cull below performs the corresponding exact rectangle check.
      const mergedRowIntersects = colIndices.some((colIndex) => {
        const merge = mergeAnchorMap.get(`${rowIndex}:${colIndex}`);
        return merge != null && cy + merge.totalH > clipY && cy < clipY + clipH;
      });
      if (!mergedRowIntersects) continue;
    }

    // Pre-compute centerContinuous ranges in this row. ECMA-376 §18.18.40:
    // a contiguous run of cells whose alignment is `centerContinuous` is
    // treated as a single visual span — Excel hides the default gridlines
    // *and* the explicit cell borders inside the run, so the run reads as
    // one merged-looking span with only the outer perimeter visible.
    //
    // A run is bounded by either (a) a non-centerContinuous cell, or
    // (b) another centerContinuous cell that itself carries a value —
    // the spec says spanned cells reference the same style id while only
    // the anchor holds the displayed value. Two anchors in a row therefore
    // mean two adjacent runs, with the border between them visible.
    const suppressRightGridCol = new Set<number>();
    const suppressLeftGridCol = new Set<number>();
    let runStart = -1;
    const closeRun = (endExclusive: number) => {
      if (runStart >= 0 && endExclusive - runStart >= 2) {
        for (let k = runStart; k < endExclusive - 1; k++) suppressRightGridCol.add(k);
        for (let k = runStart + 1; k < endExclusive; k++) suppressLeftGridCol.add(k);
      }
      runStart = -1;
    };
    for (let ci = 0; ci <= numCols; ci++) {
      let isCC = false;
      let hasValue = false;
      if (ci < numCols) {
        const ckey = `${rowIndex}:${colIndices[ci]}`;
        if (!mergeSkipSet.has(ckey) && !mergeAnchorMap.has(ckey)) {
          const c = cellMap.get(ckey);
          const cXf = resolveXf(
            styles, effectiveCellStyleIndex(rc.worksheet, c, colIndices[ci]),
          ).xf;
          isCC = cXf.alignH === 'centerContinuous';
          hasValue = !!(c && c.value && c.value.type !== 'empty');
        }
      }
      if (!isCC) {
        closeRun(ci);
      } else if (hasValue && runStart >= 0 && ci > runStart) {
        // New anchor: close previous run [runStart, ci) and start a fresh one.
        closeRun(ci);
        runStart = ci;
      } else if (runStart < 0) {
        runStart = ci;
      }
    }

    for (let ci = 0; ci < numCols; ci++) {
      const colIndex = colIndices[ci];
      const ltrCx = originX + colXs[ci];
      const cw = colWidths[ci];
      // Cull against the (un-mirrored) clip band — visibility is preserved
      // by the mirror, so the LTR test is correct for both directions.
      const key = `${rowIndex}:${colIndex}`;
      if (mergeSkipSet.has(key)) continue;

      // Use the authored merge rectangle for virtualization. A merge anchor's
      // own row/column band may remain materialized while that single band is
      // already fully clipped; Excel keeps painting the anchor until the full
      // merged range has left the pane.
      const mergeInfo = mergeAnchorMap.get(key);
      const cellW = mergeInfo ? mergeInfo.totalW : cw;
      const cellH = mergeInfo ? mergeInfo.totalH : ch;
      const horizontallyOutside = ltrCx + cellW <= clipX || ltrCx >= clipX + clipW;
      const verticallyOutside = cy + cellH <= clipY || cy >= clipY + clipH;
      if (
        verticallyOutside ||
        (horizontallyOutside && !rc.overflowTextAnchors.has(key))
      ) continue;
      // Mirror the cell's canvas x for RTL. Uses the *full* drawn width
      // (merge span included) so the rect lands at the correct mirror band.
      const cx = mirrorX(ltrCx, cellW);

      const cell = cellMap.get(key);
      const styleIndex = effectiveCellStyleIndex(rc.worksheet, cell, colIndex);
      const { font, fill, border, xf } = resolveXf(styles, styleIndex);
      const cf = evaluateCf(cell, rowIndex, colIndex, cfContext, styles.dxfs ?? [], rc.cfDiagnostics);
      const effectiveFill = cf.fill ?? fill;
      const tableStyle = rc.tableStyleMap.get(key);
      // Custom `<tableStyle>` dxfs (ECMA-376 §18.8.83). When present, they
      // drive fills / font color and borders strictly from the declared
      // elements — no accent fallback for custom styles.
      const dxfList = styles.dxfs ?? [];
      const dxfAt = (i?: number) => (i != null ? dxfList[i] : undefined);
      const tsDxfWhole = dxfAt(tableStyle?.wholeTableDxf);
      const tsDxfHeader = dxfAt(tableStyle?.headerRowDxf);
      const tsDxfTotal = dxfAt(tableStyle?.totalRowDxf);
      const tsDxfFirstCol = dxfAt(tableStyle?.firstColumnDxf);
      const tsDxfLastCol = dxfAt(tableStyle?.lastColumnDxf);
      const tsDxfStripe = dxfAt(tableStyle?.stripeDxf);
      // Per-cell table fill resolved from the element hierarchy
      // (§18.8.83: wholeTable < band < column < header/total). The later a
      // layer appears the higher its precedence; we pick the most specific
      // fill that defines a fgColor.
      const tableFillDxf =
        (tableStyle?.isHeader && tsDxfHeader?.fill?.fgColor) ? tsDxfHeader :
        (tableStyle?.isTotals && tsDxfTotal?.fill?.fgColor) ? tsDxfTotal :
        (tableStyle?.isLastCol && tsDxfLastCol?.fill?.fgColor) ? tsDxfLastCol :
        (tableStyle?.isFirstCol && tsDxfFirstCol?.fill?.fgColor) ? tsDxfFirstCol :
        (tsDxfStripe?.fill?.fgColor) ? tsDxfStripe :
        (!tableStyle?.isHeader && !tableStyle?.isTotals && tsDxfWhole?.fill?.fgColor) ? tsDxfWhole :
        undefined;

      // Background fill (base or CF override). ECMA-376 §18.8.22 ST_PatternType.
      // - solid/gray*: blend fgColor with bgColor at the pattern's fg coverage.
      // - directional hatches (dark/light Horizontal/Vertical/Down/Up/Grid/
      //   Trellis): render via a small repeating tile using createPattern so
      //   the hatch actually shows, rather than approximating as a blend.
      let hasCellBackground = paintCellPatternFill(ctx, effectiveFill, cx, cy, cellW, cellH);
      // PivotTable style fill (§18.8.41) under the cell's own fill.
      const pivotFormat = rc.pivotStyleMap.get(key);
      if (!hasCellBackground && pivotFormat?.fill) {
        hasCellBackground = paintCellPatternFill(ctx, pivotFormat.fill, cx, cy, cellW, cellH);
      }
      if (hasCellBackground) {
        // own fill painted; tableStyle fallbacks intentionally skipped
      } else if (tableStyle && tableFillDxf?.fill?.fgColor) {
        // Custom or built-in: a resolved table-element dxf fill wins.
        ctx.fillStyle = hexToRgba(tableFillDxf.fill.fgColor);
        ctx.fillRect(cx, cy, cellW, cellH);
        hasCellBackground = true;
      } else if (tableStyle && !tableStyle.isCustom && tableStyle.isBanded) {
        // Accent-tint banding is an approximation for built-in style names
        // only. Custom styles with no stripe dxf get no banding fill (Excel
        // draws none — §18.5.1.2).
        ctx.fillStyle = stripeColorFor(tableStyle.accent);
        ctx.fillRect(cx, cy, cellW, cellH);
        hasCellBackground = true;
      }

      // Comment indicator triangle — drawn above fill but below borders so
      // borders still read cleanly around the cell edge.
      if (rc.commentCells.has(key)) {
        drawCommentMarker(ctx, cx, cy, cellW, cellH);
      }

      // DataBar (drawn inside the cell, left-anchored). Excel 2010+ renders
      // these with a horizontal gradient by default.
      if (cf.dataBar && cf.dataBar.ratio > 0) {
        const barInset = 2;
        const barW = Math.max(0, (cellW - barInset * 2) * cf.dataBar.ratio);
        fillDataBar(ctx, cf.dataBar.color, cx + barInset, cy + barInset, barW, cellH - barInset * 2, cf.dataBar.gradient);
      }

      // Sparkline (Office 2010 x14:sparklineGroup). Drawn after the cell
      // background but before borders / text so borders frame the sparkline
      // and any cell text overlays it (matches Excel's z-order, and lets
      // a label like "Trend" share the same cell as the sparkline).
      const sparkModel = rc.sparklineMap.get(key);
      if (sparkModel) {
        renderSparkline(ctx, { x: cx, y: cy, w: cellW, h: cellH }, sparkModel);
      }

      // Grid lines – draw only right + bottom edges once per cell (avoids double-drawing at
      // shared cell boundaries). crispOffset snaps each line onto the device
      // pixel grid so we get a crisp 1-device-pixel result at any dpr (a fixed
      // 0.5/dpr offset blurs at even device widths, e.g. dpr=3).
      // Skipped when the sheet has `<sheetView showGridLines="0">` (View →
      // Gridlines unchecked; ECMA-376 §18.3.1.83).
      if (rc.worksheet.isChartSheet !== true
          && rc.worksheet.showGridlines !== false
          && !hasCellBackground) {
        ctx.strokeStyle = '#d0d0d0';
        ctx.lineWidth = 0.5;
        ctx.beginPath();
        if (!suppressRightGridCol.has(ci)) {
          const rx = cx + cellW + crispOffset(cx + cellW, 0.5, dpr);  // right edge
          ctx.moveTo(rx, cy);
          ctx.lineTo(rx, cy + cellH);
        }
        const by = cy + cellH + crispOffset(cy + cellH, 0.5, dpr);    // bottom edge
        ctx.moveTo(cx, by);
        ctx.lineTo(cx + cellW, by);
        if (ri === 0) {                            // top edge for first row
          const ty = cy + crispOffset(cy, 0.5, dpr);
          ctx.moveTo(cx, ty);
          ctx.lineTo(cx + cellW, ty);
        }
        if (ci === 0) {                            // left edge for first column
          const lx = cx + crispOffset(cx, 0.5, dpr);
          ctx.moveTo(lx, cy);
          ctx.lineTo(lx, cy + cellH);
        }
        ctx.stroke();
      }

      // Cell borders (base + any CF borders overlaid via per-edge merge).
      // For merged anchors, combine the anchor's border with the right/bottom
      // edges from the constituent cells on those edges (ECMA-376 §18.3.1.55).
      const baseBorder = mergeInfo
        ? resolveMergeBorder(border, rowIndex, colIndex, mergeInfo.right, mergeInfo.bottom, cellMap, styles)
        : border;
      let mergedBorder = mergeBorders(baseBorder, cf.border);
      // PivotTable style edges (§18.8.41) fill the edges the cell leaves unset.
      if (pivotFormat) {
        mergedBorder = {
          ...mergedBorder,
          top: mergedBorder.top?.style ? mergedBorder.top : (pivotFormat.top ?? mergedBorder.top),
          bottom: mergedBorder.bottom?.style ? mergedBorder.bottom : (pivotFormat.bottom ?? mergedBorder.bottom),
          left: mergedBorder.left?.style ? mergedBorder.left : (pivotFormat.left ?? mergedBorder.left),
          right: mergedBorder.right?.style ? mergedBorder.right : (pivotFormat.right ?? mergedBorder.right),
        };
      }
      // centerContinuous: hide internal vertical borders so the run reads as
      // one visual span (matches Excel — see precompute block above).
      if (suppressRightGridCol.has(ci) || suppressLeftGridCol.has(ci)) {
        mergedBorder = {
          ...mergedBorder,
          left: suppressLeftGridCol.has(ci) ? null : mergedBorder.left,
          right: suppressRightGridCol.has(ci) ? null : mergedBorder.right,
        };
      }
      // Shared-edge conflict resolution (ECMA-376 §18.18.3). Two adjacent cells
      // share the boundary along their common row / column edge. In the two-pass
      // order each cell's border closure runs after every fill, so the upper
      // cell's `bottom` (and the left cell's `right`) is never erased — it is
      // always drawn by that neighbour. So for an INTERIOR cell we only touch our
      // own top / left when *we* also define an edge there that conflicts, and
      // render the stronger of the two. Adopting a neighbour's edge we don't
      // define would double-stroke it (two strokes compounding to an over-dark
      // rule — sample-27) or paint a spurious line over a fill-less cell
      // (sample-6). This replaces the old fill-repair (`paintedFill` gate +
      // `invertedTop` / `invertedLeft`), which is gone.
      //
      // EXCEPTION — the quadrant's first rendered row/column (`ri`/`ci` === 0):
      // the upper / left neighbour lies OUTSIDE this quadrant (the viewport has
      // no top/left overscan — viewer.ts walks forward from startRow/startCol
      // with a +2 buffer only at the bottom/right), so when scrolled it is never
      // iterated and never strokes its own facing edge. There we adopt the
      // neighbour's edge even when our own side is unset, so a boundary authored
      // only as the neighbour's bottom/right still shows at the viewport edge.
      const aboveCell = cellMap.get(`${rowIndex - 1}:${colIndex}`);
      const aboveBottom = aboveCell
        ? resolveXf(
            styles, effectiveCellStyleIndex(rc.worksheet, aboveCell, colIndex),
          ).border.bottom
        : null;
      if (aboveBottom?.style && (ri === 0 || mergedBorder.top?.style)) {
        mergedBorder = { ...mergedBorder, top: pickStrongerEdge(mergedBorder.top, aboveBottom) };
      }
      if (!suppressLeftGridCol.has(ci)) {
        // Skip when the left edge was deliberately suppressed for a
        // centerContinuous run (ECMA-376 §18.18.40) — otherwise the
        // neighbour's xf.right re-introduces the internal vertical we just hid.
        const leftCell = cellMap.get(`${rowIndex}:${colIndex - 1}`);
        const leftRight = leftCell
          ? resolveXf(
              styles, effectiveCellStyleIndex(rc.worksheet, leftCell, colIndex - 1),
            ).border.right
          : null;
        if (leftRight?.style && (ci === 0 || mergedBorder.left?.style)) {
          mergedBorder = { ...mergedBorder, left: pickStrongerEdge(mergedBorder.left, leftRight) };
        }
      }
      // Excel Table style overlay (ECMA-376 §18.8.83). Custom styles draw only
      // their dxf-defined borders; built-in style names fall back to a
      // synthesized accent rule. Drawn on top of the cell border so an
      // empty-border data cell still shows the table structure the style
      // defines. None-style tables produce no `tableStyleMap` entry, so this is
      // skipped. (Excel forbids merging cells inside a structured Table, so a
      // merged anchor never carries a `tableStyle`; for the merged branch only
      // the AutoFilter arrow can apply.)
      const overlay = tableStyle
        ? tableOverlayBorder(tableStyle, tsDxfWhole, tsDxfHeader, colIndex)
        : null;
      const drawArrow = rc.autoFilterCells.has(key);
      // Paint the base cell border, then the table overlay on top of it, then
      // the AutoFilter arrow — the original inline order, kept intact whether
      // this runs deferred (non-merged) or inline (merged anchor).
      const paintOverlayAndArrow = () => {
        if (overlay) {
          if (overlay.kind === 'dxf') {
            renderBorder(ctx, overlay.border, cx, cy, cellW, cellH, dpr);
          } else if (overlay.kind === 'accent') {
            const hp = 0.5 / dpr;
            ctx.strokeStyle = overlay.color;
            ctx.lineWidth = overlay.lineWidth;
            ctx.beginPath();
            // perimeter inset (-hp) kept literal: snapping could push it
            // outside the cell's bottom edge; dpr>=3 edge accepted
            ctx.moveTo(cx, cy + cellH - hp);
            ctx.lineTo(cx + cellW, cy + cellH - hp);
            if (overlay.topEdge) {
              const ty = cy + crispOffset(cy, overlay.lineWidth, dpr);
              ctx.moveTo(cx, ty);
              ctx.lineTo(cx + cellW, ty);
            }
            ctx.stroke();
          }
        }
        // AutoFilter dropdown indicator
        if (drawArrow) {
          drawAutoFilterArrow(ctx, cx, cy, cw, cellH);
        }
      };

      if (mergeInfo) {
        // Merged anchors draw their (full-span) border above all gridlines —
        // see `mergeBorderTasks`. Snapshot the per-iteration `let` binding
        // (`mergedBorder` is reassigned by the conflict-resolution above) so
        // the deferred closure draws the value as of *this* anchor; cx/cy/cellW/
        // cellH are already block `const` and safe to capture directly. The
        // overlay/arrow paint inline (pass 1) here, matching the prior order
        // where the deferred merge border sat above them.
        const mb = mergedBorder;
        mergeBorderTasks.push(() => renderBorder(ctx, mb, cx, cy, cellW, cellH, dpr));
        paintOverlayAndArrow();
      } else {
        // Non-merged borders are deferred to pass 2 (after every fill) so a
        // later neighbour's fill can never over-paint this cell's edges. The
        // overlay + arrow ride in the SAME closure, after the base border, so
        // their mutual base-border < overlay < arrow order is preserved exactly
        // — only the whole trio shifts above every cell's gray gridline.
        const mb = mergedBorder;
        borderTasks.push(() => {
          renderBorder(ctx, mb, cx, cy, cellW, cellH, dpr);
          paintOverlayAndArrow();
        });
      }

      if (!cell) continue;
      const formatted = formatCellValueWithColor(cell, styles, cf.numFmt, rc.worksheet.date1904);
      const text = formatted.text;
      if (!text || (text === '0' && rc.worksheet.showZeros === false)) continue;

      textTasks.push(() => {
      // The table-element dxf that applies to this cell's font (header/total/
      // column/stripe/wholeTable), resolved by the same §18.8.83 hierarchy as
      // the fill above. Used for both bold and color of custom styles.
      const tableFontDxf =
        (tableStyle?.isHeader) ? tsDxfHeader :
        (tableStyle?.isTotals) ? tsDxfTotal :
        (tableStyle?.isLastCol && tsDxfLastCol) ? tsDxfLastCol :
        (tableStyle?.isFirstCol && tsDxfFirstCol) ? tsDxfFirstCol :
        (tsDxfStripe) ? tsDxfStripe :
        (tableStyle ? tsDxfWhole : undefined);
      // Built-in style names: synthesize bold header/total rows (approximation).
      // Custom styles: bold comes strictly from the element dxf font (§18.5.1.2).
      const tableBold = tableStyle
        ? (tableStyle.isCustom
            ? !!tableFontDxf?.font?.bold
            : (tableStyle.isHeader || tableStyle.isTotals))
        : false;
      // CF is the top layer: a matched rule's toggle, on or off, overrides
      // the cell, table and PivotTable formatting beneath it (CfResult).
      const effectiveBold = cf.fontBold ?? (font.bold || tableBold || !!pivotFormat?.bold);
      const effectiveItalic = cf.fontItalic ?? (font.italic || !!pivotFormat?.italic);
      const effectiveUnderline = cf.fontUnderline ?? (font.underline || !!pivotFormat?.underline);
      const effectiveStrike = cf.fontStrike ?? (font.strike || !!pivotFormat?.strike);
      const fontForDraw: CellFont = (
        effectiveBold !== font.bold || effectiveItalic !== font.italic ||
        effectiveUnderline !== font.underline || effectiveStrike !== font.strike
      )
        ? { ...font, bold: effectiveBold, italic: effectiveItalic, underline: effectiveUnderline, strike: effectiveStrike }
        : font;
      ctx.font = buildFont(ctx, fontForDraw, cs, cjkFallback, text);
      const hyperlinkUrl = rc.hyperlinkMap.get(key);
      // Table-style element dxfs can override font color (ECMA-376 §18.8.83),
      // following the same element hierarchy as the fill/bold above.
      const tableFontColor = styleFontColor(tableFontDxf, pivotFormat, xf);
      // Colour precedence: hyperlink > conditional-formatting font colour >
      // number-format section colour ([Red] etc., §18.8.30) > table-style dxf
      // colour > the cell's own font colour (a cell's own colour beats the
      // table's; see tableStyleFontColor).
      const textColor = hyperlinkUrl
        ? '#0563C1'
        : (cf.fontColor ?? formatted.color ?? tableFontColor ?? font.color);
      ctx.fillStyle = textColor ? hexToRgba(textColor) : '#000000';

      const paddingX = 3;
      const paddingY = 2;
      const isNumeric = cell.value.type === 'number';
      const alignH = xf.alignH ?? generalHorizontalAlignment(cell.value.type);
      const alignV = xf.alignV ?? 'bottom';
      // Indent: ECMA-376 §18.8.1 alignment@indent — one level indents by 3
      // character widths (MDW) of the workbook's normal-style font.
      const indentPx = xf.indent ? Math.round(xf.indent * 3 * rc.mdw) : 0;
      // IconSet: reserve space on the left for the icon
      const iconSz = cf.iconSet ? Math.max(8, Math.round(Math.min(cellW, cellH) * 0.55)) : 0;
      const iconPad = iconSz > 0 ? iconSz + 4 : 0;
      const leftPad = paddingX + (alignH === 'left' || !xf.alignH ? indentPx : 0) + iconPad;

      // ECMA-376 §18.18.40 ST_HorizontalAlignment value `centerContinuous`:
      // text is centered across the **selection range** — the leftmost cell
      // with content + adjacent empty cells to the right that also carry
      // `centerContinuous`. Cells stay independent (unlike merge), but the
      // centering uses the combined width. Walk right collecting empty
      // centerContinuous neighbours so we know how wide to center over.
      let centerContinuousW = cellW;
      let centerContinuousLastCi = ci;
      if (alignH === 'centerContinuous' && !mergeInfo) {
        for (let oci = ci + 1; oci < numCols; oci++) {
          const adjKey = `${rowIndex}:${colIndices[oci]}`;
          if (mergeSkipSet.has(adjKey) || mergeAnchorMap.has(adjKey)) break;
          const adjCell = cellMap.get(adjKey);
          if (adjCell && adjCell.value.type !== 'empty') break;
          const adjStyleIndex = effectiveCellStyleIndex(
            rc.worksheet, adjCell, colIndices[oci],
          );
          const adjXf = resolveXf(styles, adjStyleIndex).xf;
          if (adjXf.alignH !== 'centerContinuous') break;
          centerContinuousW += colWidths[oci];
          centerContinuousLastCi = oci;
        }
      }
      const centerContinuousX = rc.rtl
        ? cx - (centerContinuousW - cellW)
        : cx;

      // Text overflow into adjacent empty cells is observed Excel display
      // behavior when `wrapText=false`; OOXML persists the cell's alignment
      // and wrapping state but does not specify this paint-time spill rule.
      // Left-aligned text flows rightward, right-aligned flows leftward, and
      // centered text splits evenly. We only extend the clip rect; the text
      // itself is still drawn once.
      // Stops at merge-cell boundaries, non-empty cells, and iconSet-left
      // overrun (since an icon sits inside this cell's left padding).
      let drawX = alignH === 'centerContinuous' ? centerContinuousX : cx;
      let drawW = alignH === 'centerContinuous' ? centerContinuousW : cellW;
      // Excel only overflows text into adjacent empty cells; numeric values
      // that don't fit are rendered as "####" (they never spill). Cells
      // containing hard line breaks render multi-line in place, so they don't
      // overflow either. centerContinuous overflows symmetrically across the
      // selection range — text stays centered on the original range, but the
      // clip rect extends into adjacent empty cells when the text is wider.
      const hasHardBreak = text.includes('\n');
      if (!mergeInfo && !xf.wrapText && !xf.textRotation && !isNumeric && !hasHardBreak) {
        const textW = ctx.measureText(text).width;
        const isCenterCont = alignH === 'centerContinuous';
        const textPx = isCenterCont ? textW + paddingX * 2 : textW + leftPad + paddingX;
        const containerW = isCenterCont ? centerContinuousW : cellW;
        if (textPx > containerW) {
          const overflow = textPx - containerW;
          let extendRight = 0;
          let extendLeft = 0;
          if (alignH === 'right') {
            extendLeft = overflow;
          } else if (alignH === 'center' || isCenterCont) {
            extendLeft = overflow / 2;
            extendRight = overflow / 2;
          } else {
            extendRight = overflow;
          }
          // Resolve the logical range to its physical outside neighbours once.
          // In an RTL sheet (§18.3.1.87), increasing column indices move left,
          // so the range's physical right edge is `ci` and its physical left
          // edge is `centerContinuousLastCi`.
          const physicalRightStep = rc.rtl ? -1 : 1;
          const physicalRightStartOci = rc.rtl
            ? ci - 1
            : centerContinuousLastCi + 1;
          const physicalLeftStep = rc.rtl ? 1 : -1;
          const physicalLeftStartOci = rc.rtl
            ? centerContinuousLastCi + 1
            : ci - 1;
          if (extendRight > 0) {
            let budget = extendRight;
            for (let oci = physicalRightStartOci;
              oci >= 0 && oci < numCols && budget > 0;
              oci += physicalRightStep) {
              const adjKey = `${rowIndex}:${colIndices[oci]}`;
              if (mergeSkipSet.has(adjKey) || mergeAnchorMap.has(adjKey)) break;
              const adjCell = cellMap.get(adjKey);
              if (adjCell && adjCell.value.type !== 'empty') break;
              drawW += colWidths[oci];
              budget -= colWidths[oci];
            }
          }
          if (extendLeft > 0) {
            let budget = extendLeft;
            for (let oci = physicalLeftStartOci;
              oci >= 0 && oci < numCols && budget > 0;
              oci += physicalLeftStep) {
              const adjKey = `${rowIndex}:${colIndices[oci]}`;
              if (mergeSkipSet.has(adjKey) || mergeAnchorMap.has(adjKey)) break;
              const adjCell = cellMap.get(adjKey);
              if (adjCell && adjCell.value.type !== 'empty') break;
              drawX -= colWidths[oci];
              drawW += colWidths[oci];
              budget -= colWidths[oci];
            }
          }
        }
      }

      // ECMA-376 §18.18.40 horizontal alignment:
      // - `fill`: repeat the cell content until the cell fills horizontally.
      //   We resolve to the repeated string here; layout is then standard
      //   left-aligned.
      // - `distributed`: distribute characters evenly across the cell width
      //   so the first char hugs the left edge and the last hugs the right.
      //   Implemented via Canvas `letterSpacing` (a string CSS length).
      // - `justify`: full-line justification. Multi-line distribution would
      //   need the wrap engine; for the single-line case we fall back to
      //   `distributed` if the text is short, else left-aligned (matches
      //   what Excel does when justify cannot expand a one-line string).
      // Anything we don't recognise falls through to left-align (general).
      let drawText = text;
      let letterSpacingPx = 0;
      if (alignH === 'fill' && !isNumeric && text.length > 0) {
        const innerW = Math.max(1, cellW - paddingX * 2);
        const oneW = ctx.measureText(text).width;
        if (oneW > 0 && oneW < innerW) {
          const reps = Math.max(1, Math.floor(innerW / oneW));
          drawText = text.repeat(reps);
        }
      }
      if (alignH === 'distributed' || (alignH === 'justify' && !xf.wrapText && !hasHardBreak)) {
        const innerW = Math.max(1, cellW - paddingX * 2);
        const naturalW = ctx.measureText(drawText).width;
        const gaps = Math.max(1, [...drawText].length - 1);
        if (naturalW < innerW) {
          letterSpacingPx = Math.max(0, (innerW - naturalW) / gaps);
        }
      }

      let textX: number;
      let textAlign: CanvasTextAlign;
      if (alignH === 'right') {
        textX = cx + cellW - paddingX;
        textAlign = 'right';
      } else if (alignH === 'center') {
        textX = cx + cellW / 2;
        textAlign = 'center';
      } else if (alignH === 'centerContinuous') {
        // Center on the original selection range; drawX/drawW already covers
        // the range (and any overflow extension into empty neighbours).
        textX = centerContinuousX + centerContinuousW / 2;
        textAlign = 'center';
      } else if (alignH === 'distributed' || (alignH === 'justify' && !xf.wrapText && !hasHardBreak)) {
        // Distribute characters evenly: first hugs left edge, last hugs right.
        textX = cx + paddingX;
        textAlign = 'left';
      } else {
        textX = cx + leftPad;
        textAlign = 'left';
      }

      const rotation = xf.textRotation ?? 0;
      const isStacked = rotation === 255;
      const isRotated = rotation > 0 && rotation !== 255;

      // Draw icon set icon (before text clip block)
      if (cf.iconSet && iconSz > 0) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(cx, cy, cellW, cellH);
        ctx.clip();
        drawCfIcon(ctx, cf.iconSet.name, cf.iconSet.index, cx + 2, cy + (cellH - iconSz) / 2, iconSz);
        ctx.restore();
      }

      ctx.save();
      ctx.beginPath();
      ctx.rect(drawX, cy, drawW, cellH);
      ctx.clip();

      if (formatted.fill) {
        drawHashFill(ctx, cx, cy, cellW, cellH, alignV, fontForDraw, cs);
        ctx.restore();
        return;
      }

      // Stacked text (textRotation=255): draw each character on its own line
      if (isStacked) {
        const charH = vMetricPx(font.size, cs, 1.1);
        // One stacked slot per CODE POINT (the draw loop iterates code points), so
        // an astral character reserves one slot, keeping center/bottom anchoring
        // correct (issue #790 codex review, finding 3).
        const stackedCount = [...text].length;
        const totalH = stackedCount * charH;
        let charY = alignV === 'top' ? cy + paddingY
          : alignV === 'center' ? cy + (cellH - totalH) / 2
          : cy + cellH - totalH - paddingY;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'top';
        const paintStacked = (value: string, x: number, y: number) =>
          paintCellText(ctx, value, x, y, fontForDraw, cs);
        for (const ch of text) {
          // UAX#50 per-glyph orientation (issue #790): CJK/Latin stay upright,
          // 、。 substitute their vertical form, fullwidth brackets substitute
          // their U+FE3x form, and ー rotates 90° to a vertical bar.
          const cp = ch.codePointAt(0) ?? 0;
          const vertCapable =
            verticalTrLongMark(cp) && verticalVertGlyphReachable(ctx, cp);
          drawStackedVerticalChar(ctx, ch, cx + cellW / 2, charY, charH, vertCapable, paintStacked);
          charY += charH;
        }
        ctx.restore();
        return;
      }

      // Rotated text: translate to cell center, rotate, draw, restore
      if (isRotated) {
        const angleRad = rotation <= 90
          ? -(rotation * Math.PI / 180)
          : ((rotation - 90) * Math.PI / 180);
        ctx.translate(cx + cellW / 2, cy + cellH / 2);
        ctx.rotate(angleRad);
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        paintCellText(ctx, text, 0, 0, fontForDraw, cs);
        ctx.restore();
        return;
      }

      // shrinkToFit: scale context horizontally if text is wider than cell
      if (xf.shrinkToFit) {
        const textW = ctx.measureText(text).width;
        const availW = cellW - leftPad - paddingX;
        if (textW > availW && textW > 0) {
          const scale = availW / textW;
          const pivotX = alignH === 'right' ? cx + cellW - paddingX
            : alignH === 'center' ? cx + cellW / 2
            : cx + leftPad;
          ctx.transform(scale, 0, 0, 1, pivotX * (1 - scale), 0);
        }
      }

      ctx.textAlign = textAlign;
      // ECMA-376 §18.18.39 distributed / §18.18.40 readingOrder. Canvas
      // `letterSpacing` (CSS length string, supported in modern browsers)
      // implements distributed by stretching the gap between glyphs;
      // `direction` flips the writing direction for RTL languages.
      if (letterSpacingPx > 0) {
        try { (ctx as CanvasRenderingContext2D & { letterSpacing: string }).letterSpacing = `${letterSpacingPx}px`; } catch { /* ignore */ }
      }
      // §18.8.1 readingOrder: 1 = LTR, 2 = RTL, 0/absent = Context (first
      // strong character decides). Always set the direction explicitly — the
      // previous explicit-only ladder leaked the prior cell's direction into
      // Context cells.
      try {
        (ctx as CanvasRenderingContext2D & { direction: 'ltr' | 'rtl' }).direction =
          cellBaseRtl(xf.readingOrder, text) ? 'rtl' : 'ltr';
      } catch { /* ignore */ }

      // Rich text: draw each run with its own font. Rich runs take the table
      // style's font color by the measured run rule (richRunsWithTableColor).
      const tableRunColor = tableStyleFontColor(tableFontDxf, xf);
      const rawRuns = cell.value.type === 'text' ? cell.value.runs : undefined;
      const runs = rawRuns && richRunsWithTableColor(rawRuns, tableRunColor);
      const richBaseFont = tableRunColor != null ? { ...fontForDraw, color: tableRunColor } : fontForDraw;
      const hasRichText = runs && runs.length > 0;

      if (xf.wrapText && hasRichText) {
        // Rich text with wrapping — shared with the off-screen pre-pass.
        drawWrappedRichText(
          ctx, runs, richBaseFont,
          { alignH, alignV, cx, cy, cellW, cellH, leftPad, paddingX, paddingY },
          cs, dpr, { fontColor: cf.fontColor, readingOrder: xf.readingOrder }, cjkFallback
        );
      } else if (xf.wrapText) {
        drawWrappedPlainText(
          ctx,
          text,
          textX,
          fontForDraw,
          { alignH, alignV, cx, cy, cellW, cellH, leftPad, paddingX, paddingY },
          cs,
        );
      } else if (hasRichText) {
        // Non-wrap rich text: per-run fonts, honoring hard breaks (Alt+Enter LF;
        // ECMA-376 §18.8.1 — Excel renders breaks even with wrapText off). The
        // shared helper draws a break-free value as a single alignV-anchored line
        // and a value with breaks as multiple lines, keeping this in-viewport
        // path and the off-screen-anchor pre-pass identical.
        drawNonWrapRichText(
          ctx, runs, richBaseFont,
          { alignH, alignV, cx, cy, cellW, cellH, leftPad, paddingX, paddingY },
          cs, dpr, { fontColor: cf.fontColor, readingOrder: xf.readingOrder }, cjkFallback
        );
      } else {
        // ECMA-376 §18.4.14 vertAlign / ST_VerticalAlignRun §22.9.2.17 —
        // cell-level super/subscript: render the glyphs at ~65% size, shifted
        // off the baseline so the cell still reads at the right vertical band.
        // §18.4.14 mandates the size reduction; the exact ratio/offset is
        // implementation-defined (ratios match Office's visual output).
        const cellVertAlign = fontForDraw.vertAlign;
        const baseSizePxOrig = vMetricPx(font.size, cs);
        let vaYShift = 0;
        if (cellVertAlign === 'superscript') vaYShift = -Math.round(baseSizePxOrig * 0.35);
        else if (cellVertAlign === 'subscript') vaYShift = Math.round(baseSizePxOrig * 0.10);
        const drawFont: CellFont = cellVertAlign
          ? { ...fontForDraw, size: fontForDraw.size * 0.65 }
          : fontForDraw;
        if (cellVertAlign) {
          ctx.font = buildFont(ctx, drawFont, cs, cjkFallback, text);
        }

        // Measure once for both underline and strike
        let overlayMetrics: TextMetrics | null = null;
        const measureOverlay = () => overlayMetrics ??= ctx.measureText(text);
        const overlayX = () => {
          const tW = Math.min(measureOverlay().width, drawW - leftPad - paddingX);
          return {
            x: alignH === 'right' ? cx + cellW - paddingX - tW
              : alignH === 'center' ? cx + cellW / 2 - tW / 2
              : cx + leftPad,
            width: tW,
          };
        };
        const sizePx = vMetricPx(drawFont.size, cs);

        if (fontForDraw.underline || hyperlinkUrl) {
          const { x: ux, width: tW } = overlayX();
          const uy = (alignV === 'top'
            ? cy + paddingY + sizePx + 1
            : alignV === 'center'
              ? cy + cellH / 2 + Math.round(sizePx * 0.55)
              : cy + cellH - paddingY + 1) + vaYShift;
          const stroke = hyperlinkUrl ? '#0563C1' : (textColor ? hexToRgba(textColor) : '#000000');
          const dbl = fontForDraw.underlineStyle === 'double' || fontForDraw.underlineStyle === 'doubleAccounting';
          drawTextDecoLine(ctx, ux, ux + tW, uy, stroke, dbl, dpr);
        }
        if (fontForDraw.strike) {
          const { x: sx, width: tW } = overlayX();
          // Strike line sits roughly at the x-height mid-line (~45% up from baseline)
          const syBase = (alignV === 'top'
            ? cy + paddingY + Math.round(sizePx * 0.5)
            : alignV === 'center'
              ? cy + cellH / 2
              : cy + cellH - paddingY - Math.round(sizePx * 0.35)) + vaYShift;
          const sy = syBase + crispOffset(syBase, 0.5, dpr);
          ctx.save();
          ctx.strokeStyle = textColor ? hexToRgba(textColor) : '#000000';
          ctx.lineWidth = 0.5;
          ctx.beginPath(); ctx.moveTo(sx, sy); ctx.lineTo(sx + tW, sy); ctx.stroke();
          ctx.restore();
        }
        // Hard line breaks (\n from Alt+Enter) render as multiple lines even
        // when wrapText is false — this matches Excel's behavior.
        if (text.includes('\n')) {
          const lines = text.split('\n');
          const lineH = vMetricPx(font.size, cs, 1.2);
          const totalTextH = lines.length * lineH;
          let startY: number;
          if (alignV === 'top') { startY = cy + paddingY; ctx.textBaseline = 'top'; }
          else if (alignV === 'center') { startY = cy + (cellH - totalTextH) / 2; ctx.textBaseline = 'top'; }
          else { startY = cy + cellH - totalTextH - paddingY; ctx.textBaseline = 'top'; }
          for (let li = 0; li < lines.length; li++) {
            paintCellText(ctx, lines[li], textX, startY + li * lineH + vaYShift, drawFont, cs);
          }
        } else {
          const { baseline, textY } = singleLineVerticalAnchor({
            alignH, alignV, cx, cy, cellW, cellH, leftPad, paddingX, paddingY,
          });
          ctx.textBaseline = baseline;
          paintCellText(ctx, drawText, textX, textY + vaYShift, drawFont, cs);
        }
      }

      // ECMA-376 §18.4.6 / §18.4.3 furigana: when the cell opts in (`ph="1"`)
      // and its String Item carries phonetic runs, stamp the reading across the
      // top of the cell over the base glyphs. Drawn inside the cell clip (so an
      // over-tall reading clips like Excel) and skipped for stacked/rotated
      // text (which return early above) and hard-break multi-line values
      // (furigana over wrapped lines is not a shape Excel produces for
      // phonetic cells). `baseLeftX` follows where the base text is drawn:
      // left/general anchors at the text start; center/right shift by the
      // measured base width.
      const phRuns = cell.value.type === 'text' ? cell.value.phoneticRuns : undefined;
      if (cell.showPhonetic && phRuns && phRuns.length > 0 && !text.includes('\n')) {
        const baseFontStr = buildFont(ctx, fontForDraw, cs, cjkFallback, text);
        const baseTextW = measureInFont(ctx, text, baseFontStr);
        let baseLeftX: number;
        if (alignH === 'right') baseLeftX = cx + cellW - paddingX - baseTextW;
        else if (alignH === 'center') baseLeftX = cx + cellW / 2 - baseTextW / 2;
        else baseLeftX = cx + leftPad;
        const phColor = textColor ? hexToRgba(textColor) : '#000000';
        drawPhoneticBand(
          ctx,
          phRuns,
          cell.value.type === 'text' ? cell.value.phoneticPr : undefined,
          text,
          baseFontStr,
          styles,
          baseLeftX,
          cy,
          cs,
          phColor, cjkFallback
        );
      }

      ctx.restore();

      if (text && rc.onTextRun) {
        rc.onTextRun({
          sheetName: rc.worksheet.name,
          cellRef: formatA1(rowIndex, colIndex),
          text,
          x: cx,
          y: cy,
          width: cellW,
          height: cellH,
          row: rowIndex,
          col: colIndex,
        });
      }
      });
    }
  }

  // Two-pass flush. Pass 1 (the grid loop above) laid down every cell fill and
  // the gray sheet gridlines; pass 2 paints all styled borders on top, so no
  // fill can ever over-paint a neighbour's shared edge (the bug this removes).
  // Order within pass 2: non-merged borders first (they were inline = earlier
  // in the old per-cell path), then merged-anchor borders (always deferred =
  // later), then text last (stays on top of every border). This preserves the
  // prior gridline < non-merged border < merged border < text z-order.
  for (const task of borderTasks) task();

  for (const task of mergeBorderTasks) task();

  for (const task of textTasks) task();

  ctx.restore();
}

interface OverflowColumnOverscan {
  startCol: number;
  endCol: number;
  anchorKeys: Set<string>;
}

function precedingColumn(cols: readonly number[], boundary: number): number | undefined {
  let low = 0;
  let high = cols.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (cols[middle] < boundary) low = middle + 1;
    else high = middle;
  }
  return low > 0 ? cols[low - 1] : undefined;
}

function followingColumn(cols: readonly number[], boundary: number): number | undefined {
  let low = 0;
  let high = cols.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (cols[middle] <= boundary) low = middle + 1;
    else high = middle;
  }
  return low < cols.length ? cols[low] : undefined;
}

function mergeBlocksOverflow(
  worksheet: Worksheet,
  row: number,
  left: number,
  right: number,
): boolean {
  return (worksheet.mergeCells ?? []).some((merge) => (
    row >= merge.top && row <= merge.bottom &&
    merge.right >= left && merge.left <= right
  ));
}

/**
 * Find off-screen text anchors whose authored, unwrapped display text still
 * intersects the visible column band.
 *
 * OOXML stores alignment and wrapping (§18.8.1), but Excel's spill into empty
 * neighbours is a paint-time behavior. Virtualization must therefore cull by
 * the resulting text paint bounds rather than by the value cell's bounds. We
 * inspect only the nearest non-empty cell on each side of each visible row:
 * any earlier cell is necessarily blocked by that nearer value. A candidate is
 * admitted only when its measured overflow crosses the sheet-space viewport
 * boundary, so this does not introduce an arbitrary overscan column count.
 */
function virtualizedTextOverflowOverscan(
  ctx: CanvasRenderingContext2D,
  worksheet: Worksheet,
  styles: Styles,
  cellMap: Map<string, Cell>,
  nonEmptyColsByRow: Map<number, readonly number[]>,
  cfContext: CfContext,
  tableStyleMap: Map<string, TableCellStyle>,
  pivotStyleMap: Map<string, PivotCellFormat>,
  mergeAnchorMap: Map<string, { totalW: number; totalH: number; right: number; bottom: number }>,
  colAxis: GridAxisGeometry,
  rowAxis: GridAxisGeometry,
  rowIndices: readonly number[],
  visibleStartCol: number,
  visibleEndCol: number,
  firstScrollableCol: number,
  cs: number,
  mdw: number,
  rtl: boolean,
  cjkFallback?: CjkLang,
): OverflowColumnOverscan {
  let startCol = visibleStartCol;
  let endCol = visibleEndCol;
  const anchorKeys = new Set<string>();

  const reaches = (
    row: number,
    col: number,
    logicalDirection: 'lower' | 'higher',
    distanceToViewport: number,
  ): boolean => {
    const key = `${row}:${col}`;
    const cell = cellMap.get(key);
    if (!cell || cell.value.type === 'empty' || mergeAnchorMap.has(key)) return false;

    const { font, xf } = resolveXf(styles, cell.styleIndex ?? 0);
    const cf = evaluateCf(cell, row, col, cfContext, styles.dxfs ?? []);
    const formatted = formatCellValueWithColor(cell, styles, cf.numFmt, worksheet.date1904);
    const text = formatted.text;
    if (
      !text ||
      (text === '0' && worksheet.showZeros === false) ||
      cell.value.type === 'number' ||
      xf.wrapText ||
      !!xf.textRotation ||
      text.includes('\n')
    ) return false;

    const tableStyle = tableStyleMap.get(key);
    const tableFontDxfId = tableStyle?.isHeader
      ? tableStyle.headerRowDxf
      : tableStyle?.isTotals
        ? tableStyle.totalRowDxf
        : tableStyle?.isLastCol && tableStyle.lastColumnDxf != null
          ? tableStyle.lastColumnDxf
          : tableStyle?.isFirstCol && tableStyle.firstColumnDxf != null
            ? tableStyle.firstColumnDxf
            : tableStyle?.stripeDxf ?? tableStyle?.wholeTableDxf;
    const tableFontDxf = tableFontDxfId != null ? styles.dxfs?.[tableFontDxfId] : undefined;
    const builtInTableBold = !!tableStyle && !tableStyle.isCustom && (tableStyle.isHeader || tableStyle.isTotals);
    const pivotFormat = pivotStyleMap.get(key);
    const effectiveBold = cf.fontBold ?? (font.bold || builtInTableBold || !!tableFontDxf?.font?.bold
      || !!pivotFormat?.bold);
    const effectiveItalic = cf.fontItalic ?? (font.italic || !!pivotFormat?.italic);
    const effectiveFont = (effectiveBold !== font.bold || effectiveItalic !== font.italic)
      ? { ...font, bold: effectiveBold, italic: effectiveItalic }
      : font;
    ctx.font = buildFont(ctx, effectiveFont, cs, cjkFallback, text);

    const alignH = xf.alignH ?? generalHorizontalAlignment(cell.value.type);
    const paddingX = 3;
    const indentPx = xf.indent ? Math.round(xf.indent * 3 * mdw) : 0;
    const cellW = colAxis.sizeOf(col);
    const cellH = rowAxis.sizeOf(row);
    const iconSize = cf.iconSet ? Math.max(8, Math.round(Math.min(cellW, cellH) * 0.55)) : 0;
    const iconPad = iconSize > 0 ? iconSize + 4 : 0;
    const leftPad = paddingX + (alignH === 'left' || !xf.alignH ? indentPx : 0) + iconPad;
    const textPx = ctx.measureText(text).width + leftPad + paddingX;
    const overflow = Math.max(0, textPx - cellW);
    if (overflow <= 0) return false;

    const physicalDirection = rtl
      ? (logicalDirection === 'higher' ? 'left' : 'right')
      : (logicalDirection === 'higher' ? 'right' : 'left');
    const extension = alignH === 'center' || alignH === 'centerContinuous'
      ? overflow / 2
      : (physicalDirection === 'left' ? alignH === 'right' : alignH !== 'right')
        ? overflow
        : 0;
    return extension > distanceToViewport;
  };

  for (const row of rowIndices) {
    const nonEmptyCols = nonEmptyColsByRow.get(row);
    if (!nonEmptyCols) continue;

    const before = precedingColumn(nonEmptyCols, visibleStartCol);
    const visibleStartCell = cellMap.get(`${row}:${visibleStartCol}`);
    if (
      before != null &&
      before >= firstScrollableCol &&
      (!visibleStartCell || visibleStartCell.value.type === 'empty') &&
      !mergeBlocksOverflow(worksheet, row, before + 1, visibleStartCol) &&
      reaches(
        row,
        before,
        'higher',
        colAxis.offsetOf(visibleStartCol) - colAxis.offsetOf(before + 1),
      )
    ) {
      startCol = Math.min(startCol, before);
      anchorKeys.add(`${row}:${before}`);
    }

    const after = followingColumn(nonEmptyCols, visibleEndCol);
    const visibleEndCell = cellMap.get(`${row}:${visibleEndCol}`);
    if (
      after != null &&
      after <= MAX_WORKSHEET_COL &&
      (!visibleEndCell || visibleEndCell.value.type === 'empty') &&
      !mergeBlocksOverflow(worksheet, row, visibleEndCol, after - 1) &&
      reaches(
        row,
        after,
        'lower',
        colAxis.offsetOf(after) - colAxis.offsetOf(visibleEndCol + 1),
      )
    ) {
      endCol = Math.max(endCol, after);
      anchorKeys.add(`${row}:${after}`);
    }
  }

  return { startCol, endCol, anchorKeys };
}

/** Reuse viewport-independent indexes for a render-local worksheet projection.
 * The projection may differ only in row/column sizing and outline flags; cells,
 * merges, formatting, tables, links, comments, and sparklines remain source-owned. */
export function inheritSheetRenderCache(source: Worksheet, projection: Worksheet): void {
  inheritCellFontLanguage(source, projection);
  inheritWorksheetPolicy(source, projection);
  sheetRenderCache.set(projection, getSheetRenderCache(source));
}

export function renderViewport(
  ctx: CanvasRenderingContext2D,
  worksheet: Worksheet,
  styles: Styles,
  viewport: ViewportRange,
  opts: RenderViewportOptions = {},
  cjkFallback?: CjkLang,
): void {
  bindXlsxOfficeFontRoutes(ctx, worksheet, opts.officeFontRoutes, opts.googleSubstitutes === true);
  pinXlsxGridGeometry(worksheet, opts.authoritativeMdw);
  const dpr = opts.dpr ?? 1;
  const cs = opts.cellScale ?? 1;
  const chartSheet = worksheet.isChartSheet === true;
  // Resolve MDW once per render — workbook-wide value derived from the
  // default font, `<fonts>[0]` (ECMA-376 §18.3.1.13).
  const geometry = getGridGeometryForWorksheet(worksheet);
  const mdw = geometry.maximumDigitWidth;
  const canvasW = ctx.canvas.width / dpr;
  const canvasH = ctx.canvas.height / dpr;

  ctx.clearRect(0, 0, canvasW, canvasH);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvasW, canvasH);

  // Scaled pixel helper: apply cellScale to all cell/header dimensions
  const sp = (px: number) => Math.round(px * cs);
  // Chart sheets have no worksheet grid or row/column headers. Their DrawingML
  // absolute anchors are measured from the chart-sheet canvas origin.
  const hw = chartSheet ? 0 : sp(HEADER_W);  // scaled header column width
  const hh = chartSheet ? 0 : sp(HEADER_H);  // scaled header row height

  // Viewport rows and columns are 1-based. The grid bands already clamp a
  // start below 1 to the first row/column; clamp it once here so anchored
  // drawings (positioned from the same start) stay registered with the grid
  // instead of shifting by one default column width / row height.
  const { rows: numRows, cols: numCols } = viewport;
  const startRow = Math.max(1, Math.trunc(viewport.row) || 1);
  const startCol = Math.max(1, Math.trunc(viewport.col) || 1);
  const scrollOffsetX = (opts.scrollOffsetX ?? 0) * cs;
  const scrollOffsetY = (opts.scrollOffsetY ?? 0) * cs;
  const { col: colAxis, row: rowAxis } = geometry.axesAtScale(cs);
  // A freeze count may legally cover the entire worksheet. Only materialize
  // bands that can reach this canvas; the last one may extend past the edge and
  // naturally reduces the scrollable quadrant to zero.
  const frozenColBands = colAxis.bandsToCover(
    1,
    chartSheet ? 0 : (opts.freezeCols ?? 0),
    Math.max(0, canvasW - hw),
  );
  const frozenRowBands = rowAxis.bandsToCover(
    1,
    chartSheet ? 0 : (opts.freezeRows ?? 0),
    Math.max(0, canvasH - hh),
  );
  const freezeRows = frozenRowBands.at(-1)?.index ?? 0;
  const freezeCols = frozenColBands.at(-1)?.index ?? 0;

  // ── Compute frozen area pixel sizes (scaled) ─────────────────
  const frozenColIndices = frozenColBands.map(({ index }) => index);
  const frozenRowIndices = frozenRowBands.map(({ index }) => index);
  const frozenColWidths = frozenColBands.map(({ size }) => size);
  const frozenRowHeights = frozenRowBands.map(({ size }) => size);
  const frozenW = frozenColWidths.reduce((s, w) => s + w, 0);
  const frozenH = frozenRowHeights.reduce((s, h) => s + h, 0);

  // ── Scrollable col/row pixel widths (scaled) ─────────────────
  const scrollColBands = colAxis.bandsToCover(
    startCol,
    Math.min(16_384, startCol + numCols - 1),
  );
  const scrollRowBands = rowAxis.bandsToCover(
    startRow,
    Math.min(1_048_576, startRow + numRows - 1),
  );
  const scrollColIndices = scrollColBands.map(({ index }) => index);
  const scrollRowIndices = scrollRowBands.map(({ index }) => index);
  const scrollColWidths = scrollColBands.map(({ size }) => size);
  const scrollRowHeights = scrollRowBands.map(({ size }) => size);

  // ── Viewport-independent lookups (memoized per Worksheet) ────
  const {
    cellMap, cfContext, mergeSkipSet, autoFilterCells,
    hyperlinkMap, commentCells, tableStyleMap, pivotStyleMap, sparklineMap,
    nonEmptyColsByRow,
  } = getSheetRenderCache(worksheet);

  // Merge anchor sizes are cellScale-scaled, so they stay per-frame.
  const mergeAnchorMap = new Map<string, { totalW: number; totalH: number; right: number; bottom: number }>();
  const mergeAnchorIdentity = coordinateIndexIdentity(
    worksheet,
    'worksheet-merge-anchor-index',
    'index-merge-anchor-coordinates',
  );
  for (const mc of worksheet.mergeCells ?? []) {
    const totalW = colAxis.offsetOf(mc.right + 1) - colAxis.offsetOf(mc.left);
    const totalH = rowAxis.offsetOf(mc.bottom + 1) - rowAxis.offsetOf(mc.top);
    setCoordinateIndexValue(
      mergeAnchorMap,
      `${mc.top}:${mc.left}`,
      { totalW, totalH, right: mc.right, bottom: mc.bottom },
      mergeAnchorIdentity,
    );
  }

  // Keep off-screen value cells alive only when their actual unwrapped text
  // paint bounds cross into the visible scrollable band. The ordinary
  // scrollColBands remain unchanged for headers and viewport bookkeeping;
  // renderScrollColBands add only the measured spill anchors and the empty
  // columns through which they paint.
  const visibleStartCol = scrollColIndices[0] ?? startCol;
  const visibleEndCol = scrollColIndices.at(-1) ?? Math.min(16_384, startCol + numCols - 1);
  const overflowRows = [...new Set([...frozenRowIndices, ...scrollRowIndices])];
  const overflowOverscan = virtualizedTextOverflowOverscan(
    ctx,
    worksheet,
    styles,
    cellMap,
    nonEmptyColsByRow,
    cfContext,
    tableStyleMap,
    pivotStyleMap,
    mergeAnchorMap,
    colAxis,
    rowAxis,
    overflowRows,
    visibleStartCol,
    visibleEndCol,
    freezeCols + 1,
    cs,
    mdw,
    worksheet.rightToLeft === true, cjkFallback
  );
  const renderScrollColBands = colAxis.bandsToCover(
    overflowOverscan.startCol,
    overflowOverscan.endCol,
  );
  const renderScrollColIndices = renderScrollColBands.map(({ index }) => index);
  const renderScrollColWidths = renderScrollColBands.map(({ size }) => size);
  const renderScrollOffsetX = scrollOffsetX
    + colAxis.offsetOf(visibleStartCol)
    - colAxis.offsetOf(overflowOverscan.startCol);

  const rc: RenderContext = {
    worksheet, styles, cellMap, mergeAnchorMap, mergeSkipSet, cfContext,
    colWidths: renderScrollColWidths,
    rowHeights: scrollRowHeights,
    colAxis,
    rowAxis,
    frozenColWidths, frozenRowHeights,
    frozenW, frozenH,
    startRow, startCol,
    cs,
    dpr,
    autoFilterCells,
    hyperlinkMap,
    commentCells,
    tableStyleMap,
    pivotStyleMap,
    sparklineMap,
    overflowTextAnchors: overflowOverscan.anchorKeys,
    mdw,
    onTextRun: opts.onTextRun,
    rtl: worksheet.rightToLeft === true,
    canvasW,
    threeD: opts.threeD,
    cfDiagnostics: opts.cfDiagnostics,
  };

  // Canvas areas for each quadrant
  const cellAreaX = hw;
  const cellAreaY = hh;
  const scrollAreaX = cellAreaX + frozenW;
  const scrollAreaY = cellAreaY + frozenH;
  const scrollAreaW = Math.max(0, canvasW - scrollAreaX);
  const scrollAreaH = Math.max(0, canvasH - scrollAreaY);

  // ── Q1: frozen rows × frozen cols ───────────────────────────
  if (freezeRows > 0 && freezeCols > 0) {
    renderQuadrant(ctx, rc,
      1, 1, frozenColWidths, frozenRowHeights,
      frozenColIndices, frozenRowIndices,
      0, 0,
      cellAreaX, cellAreaY,
      cellAreaX, cellAreaY, frozenW, frozenH, cjkFallback
    );
  }

  // ── Q2: frozen rows × scrollable cols ───────────────────────
  if (freezeRows > 0) {
    renderQuadrant(ctx, rc,
      1, overflowOverscan.startCol, renderScrollColWidths, frozenRowHeights,
      renderScrollColIndices, frozenRowIndices,
      renderScrollOffsetX, 0,
      scrollAreaX, cellAreaY,
      scrollAreaX, cellAreaY, scrollAreaW, frozenH, cjkFallback
    );
  }

  // ── Q3: scrollable rows × frozen cols ───────────────────────
  if (freezeCols > 0) {
    renderQuadrant(ctx, rc,
      startRow, 1, frozenColWidths, scrollRowHeights,
      frozenColIndices, scrollRowIndices,
      0, scrollOffsetY,
      cellAreaX, scrollAreaY,
      cellAreaX, scrollAreaY, frozenW, scrollAreaH, cjkFallback
    );
  }

  // ── Q4: scrollable rows × scrollable cols (main area) ───────
  if (!chartSheet) {
    renderQuadrant(ctx, rc,
      startRow, overflowOverscan.startCol, renderScrollColWidths, scrollRowHeights,
      renderScrollColIndices, scrollRowIndices,
      renderScrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH, cjkFallback
    );
  }

  // ── Anchored DrawingML objects, in authored z-order ───────────
  renderAnchoredDrawings(
    ctx, worksheet, colAxis, rowAxis, opts.loadedImages, cs,
    startRow, startCol, scrollOffsetX, scrollOffsetY,
    scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH,
    worksheet.rightToLeft === true, canvasW, opts.threeD, opts.regionMap, opts.chartEx, cjkFallback
  );

  // ── Anchored slicers (Office 2010+ pivot/table filter buttons) ──
  if (!chartSheet && worksheet.slicers && worksheet.slicers.length > 0) {
    renderSlicers(
      ctx, worksheet, colAxis, rowAxis, cs,
      startRow, startCol,
      scrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaW, scrollAreaH,
      worksheet.rightToLeft === true, canvasW,
    );
  }

  // ── Row/col headers (drawn last, always on top) ──────────────
  if (!chartSheet) {
    renderHeaders(ctx, canvasW, canvasH,
      startRow, startCol, numRows, numCols,
      scrollColWidths, scrollRowHeights,
      scrollColIndices, scrollRowIndices,
      scrollOffsetX, scrollOffsetY,
      frozenColWidths, frozenRowHeights,
      frozenColIndices, frozenRowIndices,
      frozenW, frozenH,
      hw, hh, cs, dpr,
      opts.selectedRowRange ?? null,
      opts.selectedColRange ?? null,
      worksheet.rightToLeft === true,
      opts.chromeColors,
    );
  }

  // ── Freeze pane separator lines ──────────────────────────────
  // RTL mirrors the vertical freeze divider to the left edge of the frozen
  // band (which now anchors at the right). The horizontal divider is
  // unaffected by the x-mirror but its row-header end moves to the right.
  const rtl = worksheet.rightToLeft === true;
  // crispOffset snaps the 0.5-px divider onto the device grid (same convention
  // as gridlines/headers). A fixed 0.5/dpr offset blurred at even device
  // widths (e.g. dpr=3).
  if (freezeRows > 0) {
    ctx.save();
    ctx.strokeStyle = FREEZE_LINE_COLOR;
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    const dy = scrollAreaY + crispOffset(scrollAreaY, 0.5, dpr);
    // Excel carries the separator through the row-number header so the frozen
    // boundary stays visible at the left (LTR) or right (RTL) edge as well as
    // across the cell area.
    ctx.moveTo(0, dy);
    ctx.lineTo(canvasW, dy);
    ctx.stroke();
    ctx.restore();
  }
  if (freezeCols > 0) {
    ctx.save();
    ctx.strokeStyle = FREEZE_LINE_COLOR;
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    // Divider sits between the frozen band and the scroll band. In LTR that
    // is at scrollAreaX = hw + frozenW; mirrored that becomes
    // canvasW - scrollAreaX.
    const dividerBase = rtl ? canvasW - scrollAreaX : scrollAreaX;
    const dividerX = dividerBase + crispOffset(dividerBase, 0.5, dpr);
    ctx.moveTo(dividerX, hh);
    ctx.lineTo(dividerX, canvasH);
    ctx.stroke();
    ctx.restore();
  }
}

// ────────────────────────────────────────────────────────────────
// Headers
// ────────────────────────────────────────────────────────────────
function renderHeaders(
  ctx: CanvasRenderingContext2D,
  canvasW: number, canvasH: number,
  startRow: number, startCol: number,
  numRows: number, numCols: number,
  scrollColWidths: number[], scrollRowHeights: number[],
  scrollColIndices: readonly number[], scrollRowIndices: readonly number[],
  scrollOffsetX: number, scrollOffsetY: number,
  frozenColWidths: number[], frozenRowHeights: number[],
  frozenColIndices: readonly number[], frozenRowIndices: readonly number[],
  frozenW: number, frozenH: number,
  hw: number, hh: number, cs: number, dpr: number,
  selectedRowRange: { start: number; end: number; strong: boolean } | null,
  selectedColRange: { start: number; end: number; strong: boolean } | null,
  rtl: boolean,
  chromeColors: import('./types.js').XlsxChromeColors | undefined,
): void {
  const HEADER_BG = chromeColors?.surface ?? '#f8f9fa';
  const HEADER_BG_SUBTLE = chromeColors?.mutedSurface ?? '#e8eaed';
  const HEADER_BG_STRONG = chromeColors?.selectedSurface ?? '#caddf6';
  const HEADER_BORDER = chromeColors?.border ?? '#c8ccd0';
  const HEADER_BORDER_STRONG = chromeColors?.accent ?? '#5b9bd5';
  const HEADER_TEXT = chromeColors?.text ?? '#444';

  const colBg = (col: number): string => {
    if (!selectedColRange || col < selectedColRange.start || col > selectedColRange.end) return HEADER_BG;
    return selectedColRange.strong ? HEADER_BG_STRONG : HEADER_BG_SUBTLE;
  };
  const colBorder = (col: number): string => {
    if (!selectedColRange || col < selectedColRange.start || col > selectedColRange.end) return HEADER_BORDER;
    return selectedColRange.strong ? HEADER_BORDER_STRONG : HEADER_BORDER;
  };
  const rowBg = (row: number): string => {
    if (!selectedRowRange || row < selectedRowRange.start || row > selectedRowRange.end) return HEADER_BG;
    return selectedRowRange.strong ? HEADER_BG_STRONG : HEADER_BG_SUBTLE;
  };
  const rowBorder = (row: number): string => {
    if (!selectedRowRange || row < selectedRowRange.start || row > selectedRowRange.end) return HEADER_BORDER;
    return selectedRowRange.strong ? HEADER_BORDER_STRONG : HEADER_BORDER;
  };
  const headerFontSize = Math.max(1, Math.round(11 * cs));
  const HEADER_FONT = `${headerFontSize}px ${DEFAULT_FONT_FAMILY}`;
  const scrollAreaX = hw + frozenW;
  const scrollAreaY = hh + frozenH;
  const hp = 0.5 / dpr;  // half device-pixel offset for 1dp crisp lines

  // RTL grid mirror (ECMA-376 §18.3.1.87). The cell-area band mirrors about
  // canvasW: a left-anchored col-header rect [x, x+w] maps to its mirror.
  // The corner box and the row-number strip move from the left edge
  // ([0, hw]) to the right edge ([canvasW - hw, canvasW]).
  const mirrorX = (x: number, w: number): number => rtl ? rtlMirrorX(x, w, canvasW) : x;
  const cornerX = rtl ? canvasW - hw : 0;  // x-origin of the corner / row-header strip

  // Corner — draw only the edges that belong to the spreadsheet grid. The
  // canvas perimeter is deliberately left unpainted: framing the viewer is the
  // host container's responsibility, so an app-supplied border never doubles
  // up with renderer chrome.
  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(cornerX, 0, hw, hh);
  ctx.strokeStyle = HEADER_BORDER;
  ctx.lineWidth = 0.5;
  ctx.beginPath();
  const cornerGridX = rtl
    ? cornerX + crispOffset(cornerX, 0.5, dpr)
    : cornerX + hw - hp;
  ctx.moveTo(cornerGridX, 0); ctx.lineTo(cornerGridX, hh);            // row-header / grid divider
  ctx.moveTo(cornerX, hh - hp); ctx.lineTo(cornerX + hw, hh - hp);    // bottom
  ctx.stroke();

  ctx.font = HEADER_FONT;
  ctx.fillStyle = HEADER_TEXT;

  // Helper: draw one column header cell.
  // The inter-column separator is the LEADING (left) edge at +hp, so it lands on
  // the SAME device pixel as the data grid's vertical column boundary (gridlines
  // and cell borders draw the shared boundary at cx+hp). Drawing it as this
  // cell's own left edge — rather than the previous cell's trailing right edge at
  // cx+cw-hp — both aligns it with the data column line (the strip used to sit 1
  // device px to the left) AND survives the next cell's fillRect (which starts at
  // cx+cw and never touches cx+hp), so no separate fill/stroke pass is needed.
  // The bottom line stays as the grid-facing header boundary. The canvas-top
  // perimeter is owned by the host and is not painted here.
  const drawColHeader = (col: number, ltrCx: number, cw: number) => {
    const cx = mirrorX(ltrCx, cw);
    ctx.fillStyle = colBg(col);
    ctx.fillRect(cx, 0, cw, hh);
    ctx.strokeStyle = colBorder(col);
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    const colSepX = cx + crispOffset(cx, 0.5, dpr);  // column boundary, aligns with data grid
    ctx.moveTo(colSepX, 0);           ctx.lineTo(colSepX, hh);       // left = column boundary (aligns with data grid)
    ctx.moveTo(cx, hh - hp);          ctx.lineTo(cx + cw, hh - hp);  // bottom (strip perimeter, inset -hp kept literal)
    ctx.stroke();
    ctx.fillStyle = HEADER_TEXT;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(colToLetter(col), cx + cw / 2, hh / 2);
  };

  // Helper: draw one row header cell.
  // The inter-row separator is the LEADING (top) edge at +hp, so it lands on the
  // SAME device pixel as the data grid's horizontal row boundary (gridlines and
  // cell borders draw the shared boundary at cy+hp). Drawing it as this cell's
  // own top edge — rather than the previous cell's trailing bottom edge at
  // cy+ch-hp — aligns it with the data row line (the strip used to sit 1 device
  // px above) AND survives the next cell's fillRect (which starts at cy+ch and
  // never touches cy+hp), so no separate fill/stroke pass is needed. The
  // Only the grid-facing vertical edge remains. The outer vertical canvas edge
  // is framing chrome and therefore belongs to the host container.
  const drawRowHeader = (row: number, cy: number, ch: number) => {
    const rx = cornerX;  // left edge of the row-header strip (mirrors to the right in RTL)
    ctx.fillStyle = rowBg(row);
    ctx.fillRect(rx, cy, hw, ch);
    ctx.strokeStyle = rowBorder(row);
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    const rowSepY = cy + crispOffset(cy, 0.5, dpr);   // row boundary, aligns with data grid
    const rowGridX = rtl
      ? rx + crispOffset(rx, 0.5, dpr)
      : rx + hw - hp;
    ctx.moveTo(rowGridX, cy);      ctx.lineTo(rowGridX, cy + ch);       // row-header / grid divider
    ctx.moveTo(rx, rowSepY);       ctx.lineTo(rx + hw, rowSepY);         // top = row boundary (aligns with data grid)
    ctx.stroke();
    ctx.fillStyle = HEADER_TEXT;
    ctx.textBaseline = 'middle';
    const pad = Math.max(2, Math.round(4 * cs));
    if (rtl) {
      // Strip is on the right; the grid sits to its left, so anchor the
      // number toward the grid-facing (left) edge.
      ctx.textAlign = 'left';
      ctx.fillText(String(row), rx + pad, cy + ch / 2);
    } else {
      ctx.textAlign = 'right';
      ctx.fillText(String(row), rx + hw - pad, cy + ch / 2);
    }
  };

  // Frozen col headers (no h-scroll, fixed positions)
  if (frozenColWidths.length > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(mirrorX(hw, frozenW), 0, frozenW, hh);
    ctx.clip();
    let cx = hw;
    for (let ci = 0; ci < frozenColWidths.length; ci++) {
      drawColHeader(frozenColIndices[ci], cx, frozenColWidths[ci]);
      cx += frozenColWidths[ci];
    }
    ctx.restore();
  }

  // Scrollable col headers
  ctx.save();
  ctx.beginPath();
  ctx.rect(mirrorX(scrollAreaX, canvasW - scrollAreaX), 0, canvasW - scrollAreaX, hh);
  ctx.clip();
  let cx = scrollAreaX - scrollOffsetX;
  for (let ci = 0; ci < scrollColWidths.length; ci++) {
    const cw = scrollColWidths[ci];
    if (cx + cw > scrollAreaX && cx < canvasW) {
      drawColHeader(scrollColIndices[ci], cx, cw);
    }
    cx += cw;
  }
  ctx.restore();

  // Frozen row headers (no v-scroll)
  if (frozenRowHeights.length > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(cornerX, hh, hw, frozenH);
    ctx.clip();
    let cy = hh;
    for (let ri = 0; ri < frozenRowHeights.length; ri++) {
      drawRowHeader(frozenRowIndices[ri], cy, frozenRowHeights[ri]);
      cy += frozenRowHeights[ri];
    }
    ctx.restore();
  }

  // Scrollable row headers
  ctx.save();
  ctx.beginPath();
  ctx.rect(cornerX, scrollAreaY, hw, canvasH - scrollAreaY);
  ctx.clip();
  let cy = scrollAreaY - scrollOffsetY;
  for (let ri = 0; ri < scrollRowHeights.length; ri++) {
    const ch = scrollRowHeights[ri];
    if (cy + ch > scrollAreaY && cy < canvasH) {
      drawRowHeader(scrollRowIndices[ri], cy, ch);
    }
    cy += ch;
  }
  ctx.restore();

}
