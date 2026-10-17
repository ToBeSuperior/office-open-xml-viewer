// XLSX automatic row heights derived from cell text. Split from renderer.ts
// (#1857).

import type { CjkLang } from '@silurus/ooxml-core';
import type { PivotCellFormat } from './pivot-style.js';
import type { Worksheet, Styles, Cell, CellFont, CellXf, Dxf, Run } from './types.js';
import { formatCellValueWithColor } from './number-format.js';
import { type CfResult, evaluateCf } from './conditional-format.js';
import { GridGeometry } from './internal/grid-geometry.js';
import { rowResizeRanges, rowResizeContains } from './internal/worksheet-size-context.js';
import { rowHeightToPx, pxToRowHeight } from './internal/grid-metrics.js';
import { DEFAULT_FONT_SIZE, getGridGeometryForWorksheet } from './renderer-font-metrics.js';
import {
  resolveXf,
  buildFont,
  generalHorizontalAlignment,
  vMetricPx,
  layoutRichTextLines,
  wrapTextLines,
  richHardBreakLineMetrics,
  applyRunFont,
} from './renderer-cell-text.js';
import { tableFontDxfFor, type TableCellStyle } from './renderer-table-style.js';
import { getSheetRenderCache } from './renderer-sheet-cache.js';

/** Worksheets whose absent row heights have already been resolved against the
 * loaded document fonts and effective column widths. Kept outside the public
 * model: parsed `row.height` / `rowHeights` continue to mean authored OOXML,
 * while viewer-owned projections may carry the derived display heights. */
interface AutoRowHeightState {
  /** Derived point heights keyed by 1-based row. Used to distinguish an
   * automatic value from a later user resize when column widths change. */
  derived: Map<number, number>;
}

const autoRowHeightState = new WeakMap<Worksheet, AutoRowHeightState>();

/** The table or PivotTable style font color for a cell, or null when the
 *  cell's font color is its own formatting. Measured in Excel for both: a
 *  PivotTable cell given Automatic, RGB black or red keeps that color, while
 *  one given the theme "Black, Text 1" (authored like Normal) shows the
 *  PivotTable style's color. */
export function styleFontColor(
  tableFontDxf: Dxf | undefined,
  pivotFormat: PivotCellFormat | undefined,
  xf: CellXf,
): string | null {
  if (xf.ownFontColor) return null;
  return tableFontDxf?.font?.color ?? pivotFormat?.fontColor ?? null;
}

/** Rich-text runs as Excel draws them under a table style font color
 *  (measured): a run whose <rPr> color is confirmed to be authored like the
 *  Normal style's (`normalColor`) takes the table color; any other run with
 *  <rPr> keeps its own color (automatic when it has none). Runs without <rPr>
 *  take the base font, whose color the caller sets to the table color.
 *  `tableColor` null (no table color, or the cell's font color is its own)
 *  leaves the runs unchanged. */
export function richRunsWithTableColor(runs: Run[], tableColor: string | null): Run[] {
  if (tableColor == null) return runs;
  return runs.map((run) => (run.font?.normalColor
    ? { ...run, font: { ...run.font, color: tableColor } }
    : run));
}

/** The cell font with its bold/italic/underline/strike composed from every
 *  formatting layer: cell style, table style and PivotTable style beneath,
 *  and conditional formatting on top, whose defined toggle (on or off) wins
 *  (CfResult). Matches the main paint path. */
export function layeredCellFont(
  base: CellFont,
  cf: CfResult,
  tableStyle: TableCellStyle | undefined,
  styles: Styles,
  pivotFormat: PivotCellFormat | undefined,
): CellFont {
  const tableFontDxf = tableFontDxfFor(tableStyle, styles);
  const tableBold = tableStyle
    ? tableStyle.isCustom
      ? !!tableFontDxf?.font?.bold
      : tableStyle.isHeader || tableStyle.isTotals
    : false;
  const bold = cf.fontBold ?? (base.bold || tableBold || !!pivotFormat?.bold);
  const italic = cf.fontItalic ?? (base.italic || !!pivotFormat?.italic);
  const underline = cf.fontUnderline ?? (base.underline || !!pivotFormat?.underline);
  const strike = cf.fontStrike ?? (base.strike || !!pivotFormat?.strike);
  return bold === base.bold && italic === base.italic &&
    underline === base.underline && strike === base.strike
    ? base
    : { ...base, bold, italic, underline, strike };
}

function requiredAutoCellHeightPx(
  ctx: CanvasRenderingContext2D,
  cell: Cell,
  text: string,
  font: CellFont,
  xf: CellXf,
  cellWidthPx: number,
  mdw: number,
  iconSet: CfResult['iconSet'],
  currentRowHeightPx: number,
  cjkFallback?: CjkLang,
): number {
  const paddingX = 3;
  const paddingY = 2;
  const alignH = xf.alignH ?? generalHorizontalAlignment(cell.value.type);
  const indentPx = xf.indent ? Math.round(xf.indent * 3 * mdw) : 0;
  // Keep the wrapping width identical to paint: icon-set cells reserve their
  // current icon square plus 4px. The icon itself scales with the row height,
  // so the row solver below iterates this monotonic dependency to a fixed point.
  const iconSize = iconSet
    ? Math.max(8, Math.round(Math.min(cellWidthPx, currentRowHeightPx) * 0.55))
    : 0;
  const iconPad = iconSize > 0 ? iconSize + 4 : 0;
  const leftPad = paddingX + (alignH === 'left' || !xf.alignH ? indentPx : 0) + iconPad;
  const availableWidth = Math.max(1, cellWidthPx - leftPad - paddingX);
  const runs = cell.value.type === 'text' ? cell.value.runs : undefined;
  const hasRichText = !!runs?.length;
  const rotation = xf.textRotation ?? 0;

  ctx.font = buildFont(ctx, font, 1, cjkFallback, text);
  if (rotation === 255) {
    // Paint deliberately uses the compact 1.1 slot without a document-family
    // design-line floor for stacked glyphs; auto-fit must use the same metric.
    return [...text].length * vMetricPx(font.size, 1, 1.1) + paddingY * 2;
  }
  if (rotation > 0) {
    const angle = rotation <= 90
      ? rotation * Math.PI / 180
      : (rotation - 90) * Math.PI / 180;
    const lineHeight = vMetricPx(font.size, 1, 1.2);
    const textWidth = ctx.measureText(text.replace(/\n/g, ' ')).width;
    return Math.abs(Math.sin(angle)) * textWidth
      + Math.abs(Math.cos(angle)) * lineHeight
      + paddingY * 2;
  }

  let lineHeights: number[];
  if (xf.wrapText && hasRichText) {
    lineHeights = layoutRichTextLines(ctx, runs, font, 1, availableWidth, cjkFallback)
      .map((line) => vMetricPx(line.maxFontSize, 1, 1.2));
  } else if (xf.wrapText) {
    ctx.font = buildFont(ctx, font, 1, cjkFallback, text);
    const lineHeight = vMetricPx(font.size, 1, 1.2);
    lineHeights = wrapTextLines(ctx, text, availableWidth).map(() => lineHeight);
  } else if (hasRichText) {
    lineHeights = richHardBreakLineMetrics(runs, font, 1).map((line) => line.heightPx);
  } else {
    const lineCount = Math.max(1, text.split('\n').length);
    const lineHeight = vMetricPx(font.size, 1, 1.2);
    lineHeights = Array.from({ length: lineCount }, () => lineHeight);
  }

  const textHeight = lineHeights.reduce((sum, height) => sum + height, 0);
  // One ordinary 11pt line fits Excel's 15pt default row: its design line box
  // plus the bottom inset is exactly 20 CSS px. Multi-line blocks need both
  // top and bottom insets because they use the wrapped top-baseline painter.
  return textHeight + (lineHeights.length > 1 ? paddingY * 2 : paddingY);
}

function cellCanGrowAutomaticRow(
  cell: Cell,
  font: CellFont,
  xf: CellXf,
  defaultFontLineHeightPx: number,
): boolean {
  if (xf.wrapText || (xf.textRotation ?? 0) > 0) return true;
  if (cell.value.type === 'text') {
    if (cell.value.text.includes('\n')) return true;
    for (const run of cell.value.runs ?? []) {
      const runFont = applyRunFont(font, run);
      if (vMetricPx(runFont.size, 1, 1.2) > defaultFontLineHeightPx) {
        return true;
      }
    }
  }
  // Excel's authored/default row height already accommodates a Normal-style
  // single line. Comparing that design line box with the rounded CSS-pixel row
  // height can spuriously grow every ordinary row by one pixel (for example a
  // 14.25pt default row rounds to 19px while Calibri 11pt plus cell inset is
  // 20px). Only a font whose design line box exceeds the workbook Normal font
  // can make an otherwise unwrapped, single-line cell eligible for auto-fit.
  return vMetricPx(font.size, 1, 1.2) > defaultFontLineHeightPx;
}

/**
 * Resolve Excel's display-time row auto-fit for rows that omit `row@ht`.
 *
 * ECMA-376 Part 1 §18.3.1.73 defines `customHeight` only as the fact that a row
 * was manually sized and defines `ht` as the point height; it does not specify
 * an auto-fit formula. Office behavior therefore supplies the limited
 * compatibility contract here: explicit/hidden heights and a manually set
 * sheet default win, merged cells do not drive auto-fit, and unmerged cells use
 * the exact same Canvas wrapping, rich-run line metrics, font fallback, and
 * rotation geometry as cell paint. The sheet-default distinction was verified
 * with an Office boundary workbook that varied only sheetFormatPr@customHeight.
 * The function mutates only a viewer/render-owned worksheet projection.
 */
export function applyAutoRowHeights(
  ctx: CanvasRenderingContext2D,
  worksheet: Worksheet,
  styles: Styles,
  cjkFallback?: CjkLang,
): boolean {
  if (autoRowHeightState.has(worksheet) || worksheet.isChartSheet) return false;
  if (worksheet.defaultRowHeightCustom === true) {
    autoRowHeightState.set(worksheet, { derived: new Map() });
    return false;
  }

  const geometry = getGridGeometryForWorksheet(worksheet);
  const defaultHeightPx = rowHeightToPx(worksheet.defaultRowHeight);
  // Row auto-fit compares against the Normal style's font; column MDW uses
  // the default font (`<fonts>[0]`), which may differ.
  const defaultFontLineHeightPx = vMetricPx(
    worksheet.normalFontSize ?? worksheet.defaultFontSize ?? DEFAULT_FONT_SIZE,
    1,
    1.2,
  );
  const { cfContext, mergeAnchorSet, mergeSkipSet, tableStyleMap, pivotStyleMap } =
    getSheetRenderCache(worksheet);
  const derived = new Map<number, number>();
  let changed = false;
  ctx.save();
  try {
    for (const row of worksheet.rows) {
      // `height !== null` is the parser-preserved `row@ht` fact. A pre-existing
      // view override on an otherwise automatic row (resize/worker projection)
      // is also authoritative for this projection.
      if (
        row.hidden ||
        row.customHeight === true ||
        row.height !== null ||
        rowResizeContains(rowResizeRanges(worksheet), row.index) ||
        Object.hasOwn(worksheet.rowHeights, row.index)
      ) continue;
      let requiredPx = defaultHeightPx;
      const iconMeasurements: Array<{
        cell: Cell;
        text: string;
        font: CellFont;
        xf: CellXf;
        cellWidthPx: number;
        iconSet: CfResult['iconSet'];
      }> = [];
      for (const cell of row.cells) {
        const key = `${row.index}:${cell.col}`;
        if (
          cell.value.type === 'empty' ||
          mergeAnchorSet.has(key) ||
          mergeSkipSet.has(key)
        ) continue;
        const { font, xf } = resolveXf(styles, cell.styleIndex ?? 0);
        // The common spreadsheet case (ordinary unwrapped default-font cells)
        // cannot exceed the default row. Avoid number formatting, CF lookup,
        // and Canvas measurement for those cells so one-time auto-fit remains
        // O(cells) with a small constant even on a dense virtualized sheet.
        if (!cellCanGrowAutomaticRow(cell, font, xf, defaultFontLineHeightPx)) continue;
        const cf = evaluateCf(cell, row.index, cell.col, cfContext, styles.dxfs ?? []);
        const measuredFont = layeredCellFont(
          font, cf, tableStyleMap.get(key), styles, pivotStyleMap.get(key),
        );
        const formatted = formatCellValueWithColor(cell, styles, cf.numFmt, worksheet.date1904);
        const text = formatted.text;
        if (!text || (text === '0' && worksheet.showZeros === false)) continue;
        const cellWidthPx = geometry.col.sizeOf(cell.col);
        requiredPx = Math.max(requiredPx, requiredAutoCellHeightPx(
          ctx,
          cell,
          text,
          measuredFont,
          xf,
          cellWidthPx,
          geometry.maximumDigitWidth,
          cf.iconSet,
          Math.ceil(requiredPx), cjkFallback
        ));
        if (cf.iconSet) {
          iconMeasurements.push({ cell, text, font: measuredFont, xf, cellWidthPx, iconSet: cf.iconSet });
        }
      }
      // Icon size depends on the final row height used by paint. Re-evaluate
      // icon-bearing cells until their reserved width and row height agree.
      // The relation is monotonic and saturates once the icon reaches the cell
      // width. The hard bound prevents malformed input from creating an
      // unbounded loop; the final conservative pass uses that saturated width.
      let converged = iconMeasurements.length === 0;
      for (let iteration = 0; !converged && iteration < 32; iteration++) {
        // Paint receives the integer CSS-pixel row size reconstructed from the
        // derived point height. Resolve every icon against that same candidate,
        // not the pre-rounded floating measurement, or a 0.1px difference can
        // cross the icon's round() boundary and add an unmeasured wrap line.
        const candidateHeightPx = Math.ceil(requiredPx);
        let nextRequiredPx = requiredPx;
        for (const measurement of iconMeasurements) {
          nextRequiredPx = Math.max(nextRequiredPx, requiredAutoCellHeightPx(
            ctx,
            measurement.cell,
            measurement.text,
            measurement.font,
            measurement.xf,
            measurement.cellWidthPx,
            geometry.maximumDigitWidth,
            measurement.iconSet,
            candidateHeightPx, cjkFallback
          ));
        }
        requiredPx = nextRequiredPx;
        converged = Math.ceil(requiredPx) === candidateHeightPx;
      }
      if (!converged) {
        for (const measurement of iconMeasurements) {
          requiredPx = Math.max(requiredPx, requiredAutoCellHeightPx(
            ctx,
            measurement.cell,
            measurement.text,
            measurement.font,
            measurement.xf,
            measurement.cellWidthPx,
            geometry.maximumDigitWidth,
            measurement.iconSet,
            measurement.cellWidthPx, cjkFallback
          ));
        }
      }
      const roundedPx = Math.ceil(requiredPx);
      if (roundedPx > defaultHeightPx) {
        const pointHeight = pxToRowHeight(roundedPx);
        worksheet.rowHeights[row.index] = pointHeight;
        derived.set(row.index, pointHeight);
        changed = true;
      }
    }
  } finally {
    ctx.restore();
  }
  autoRowHeightState.set(worksheet, { derived });
  if (changed) GridGeometry.invalidate(worksheet);
  return changed;
}

export function hasPreparedAutoRowHeights(worksheet: Worksheet): boolean {
  return autoRowHeightState.has(worksheet);
}

/** @internal Viewer/worker transport of display-derived heights. The returned
 * map is immutable by convention; authored/manual row sizes are deliberately
 * absent so they remain a separate override class. */
export function derivedAutoRowHeights(worksheet: Worksheet): ReadonlyMap<number, number> {
  return autoRowHeightState.get(worksheet)?.derived ?? new Map();
}

/** @internal Mark a render-local worker projection whose derived row heights
 * were already supplied by its owning viewer. */
export function markAutoRowHeightsPrepared(worksheet: Worksheet): void {
  if (!autoRowHeightState.has(worksheet)) {
    autoRowHeightState.set(worksheet, { derived: new Map() });
  }
}

/** Drop display-derived row heights before refitting after a view-only column
 * resize. A row manually resized after auto-fit is intentionally preserved:
 * its current value no longer equals the derived value recorded here. */
export function invalidateAutoRowHeights(
  worksheet: Worksheet,
  preserveRows: Iterable<number> = [],
): void {
  const state = autoRowHeightState.get(worksheet);
  if (!state) return;
  const preserved = new Set(preserveRows);
  let changed = false;
  for (const [row, derived] of state.derived) {
    if (rowResizeContains(rowResizeRanges(worksheet), row)) continue;
    if (preserved.has(row)) continue;
    if (worksheet.rowHeights[row] !== derived) continue;
    delete worksheet.rowHeights[row];
    changed = true;
  }
  autoRowHeightState.delete(worksheet);
  if (changed) GridGeometry.invalidate(worksheet);
}
