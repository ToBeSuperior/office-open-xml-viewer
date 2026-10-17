// XLSX per-worksheet render cache (comment, table-style, and sparkline maps).
// Split from renderer.ts (#1857).

import { buildPivotStyleMap, type PivotCellFormat } from './pivot-style.js';
import type { Worksheet, Cell } from './types.js';
import type { SparklineModel } from '@silurus/ooxml-core';
import { type CfContext, compileCf } from './conditional-format.js';
import {
  addCoordinateIndexEntry,
  assertCoordinateRangeArea,
  buildCellCoordinateIndex,
  setCoordinateIndexValue,
} from './renderer-coordinate-index.js';
import { getWorksheetPolicy } from './worksheet-policy-context.js';
import { parseA1Ref } from './renderer-comment-marker.js';
import {
  coordinateIndexIdentity,
  buildTableStyleMap,
  type TableCellStyle,
} from './renderer-table-style.js';

/** Flatten the worksheet's parsed `sparklineGroups` into a per-cell render
 *  model. Each Sparkline inherits its group's formatting; min/max are
 *  computed from the values when the group's `*AxisType` is `individual`,
 *  shared across the group when `group`, or taken from `manualMin/Max` when
 *  `custom`. The renderer can then look up `row:col` and call core's
 *  `renderSparkline` without further work. */
function buildSparklineMap(worksheet: Worksheet): Map<string, SparklineModel> {
  const map = new Map<string, SparklineModel>();
  const identity = coordinateIndexIdentity(
    worksheet,
    'worksheet-sparkline-index',
    'index-sparkline-coordinates',
  );
  for (const g of worksheet.sparklineGroups ?? []) {
    // Group-wide min/max if needed.
    let groupMin = Infinity, groupMax = -Infinity;
    if (g.minAxisType === 'group' || g.maxAxisType === 'group') {
      for (const sl of g.sparklines) {
        for (const v of sl.values) {
          if (typeof v === 'number') {
            if (v < groupMin) groupMin = v;
            if (v > groupMax) groupMax = v;
          }
        }
      }
      if (!isFinite(groupMin) || !isFinite(groupMax)) {
        groupMin = 0; groupMax = 1;
      }
    }
    for (const sl of g.sparklines) {
      const numeric = sl.values.filter((v): v is number => typeof v === 'number');
      const indMin = numeric.length ? Math.min(...numeric) : 0;
      const indMax = numeric.length ? Math.max(...numeric) : 1;
      const min = g.minAxisType === 'custom' && typeof g.manualMin === 'number'
        ? g.manualMin
        : g.minAxisType === 'group' ? groupMin : indMin;
      const max = g.maxAxisType === 'custom' && typeof g.manualMax === 'number'
        ? g.manualMax
        : g.maxAxisType === 'group' ? groupMax : indMax;
      setCoordinateIndexValue(map, `${sl.row}:${sl.col}`, {
        kind: g.kind,
        values: sl.values,
        min,
        max,
        displayEmptyCellsAs: (g.displayEmptyCellsAs === 'zero' || g.displayEmptyCellsAs === 'span')
          ? g.displayEmptyCellsAs
          : 'gap',
        displayXAxis: g.displayXAxis,
        lineWeight: g.lineWeight,
        markers: g.markers,
        high: g.high,
        low: g.low,
        first: g.first,
        last: g.last,
        negative: g.negative,
        colorSeries: g.colorSeries,
        colorNegative: g.colorNegative,
        colorAxis: g.colorAxis,
        colorMarkers: g.colorMarkers,
        colorFirst: g.colorFirst,
        colorLast: g.colorLast,
        colorHigh: g.colorHigh,
        colorLow: g.colorLow,
      }, identity);
    }
  }
  return map;
}

// ────────────────────────────────────────────────────────────────
// Main render function
// ────────────────────────────────────────────────────────────────
/** Viewport-independent lookups derived purely from a Worksheet. The workbook
 *  caches one Worksheet object per sheet, so this WeakMap hits on every scroll
 *  frame and only misses on a sheet switch / re-parse — avoiding a full-sheet
 *  cell-Map rebuild and a conditional-formatting recompile per frame. */
interface SheetRenderCache {
  rowCount: number;
  /** Bound per-index entry limit these indexes were built under. */
  coordinateIndexLimit: number;
  cellMap: Map<string, Cell>;
  nonEmptyColsByRow: Map<number, readonly number[]>;
  cfContext: CfContext;
  mergeAnchorSet: Set<string>;
  mergeSkipSet: Set<string>;
  autoFilterCells: Set<string>;
  hyperlinkMap: Map<string, string>;
  commentCells: Set<string>;
  tableStyleMap: Map<string, TableCellStyle>;
  pivotStyleMap: Map<string, PivotCellFormat>;
  sparklineMap: Map<string, SparklineModel>;
}
export const sheetRenderCache = new WeakMap<Worksheet, SheetRenderCache>();

/** @internal A provisional row graph grows between renders; its coordinate
 * indexes must be rebuilt after every accepted chunk. */
export function invalidateSheetRenderCache(worksheet: Worksheet): void {
  sheetRenderCache.delete(worksheet);
}

export function getSheetRenderCache(worksheet: Worksheet): SheetRenderCache {
  const coordinateIndexLimit = getWorksheetPolicy(worksheet).maxCoordinateIndexEntries;
  const cached = sheetRenderCache.get(worksheet);
  if (
    cached &&
    cached.rowCount === worksheet.rows.length &&
    cached.coordinateIndexLimit === coordinateIndexLimit
  ) {
    return cached;
  }
  if (cached) sheetRenderCache.delete(worksheet);

  const cellIdentity = coordinateIndexIdentity(
    worksheet,
    'worksheet-cell-index',
    'index-worksheet-cells',
  );
  const cellMap = buildCellCoordinateIndex(worksheet.rows, cellIdentity);
  const nonEmptyColsByRow = new Map<number, readonly number[]>();
  for (const row of worksheet.rows) {
    const cols = row.cells
      .filter((cell) => cell.value.type !== 'empty')
      .map((cell) => cell.col)
      .sort((a, b) => a - b);
    if (cols.length > 0) nonEmptyColsByRow.set(row.index, cols);
  }

  // Merge skip-set is pure topology (no scaled sizes) so it is cacheable; the
  // anchor map's pixel sizes depend on cellScale and stay per-frame.
  const mergeSkipSet = new Set<string>();
  const mergeAnchorSet = new Set<string>();
  const mergeAnchorIdentity = coordinateIndexIdentity(
    worksheet,
    'worksheet-merge-anchor-index',
    'index-merge-anchor-coordinates',
  );
  const mergeIdentity = coordinateIndexIdentity(
    worksheet,
    'worksheet-merge-skip-index',
    'expand-merged-cell-coordinates',
  );
  for (const mc of worksheet.mergeCells ?? []) {
    addCoordinateIndexEntry(mergeAnchorSet, `${mc.top}:${mc.left}`, mergeAnchorIdentity);
    assertCoordinateRangeArea(mc, mergeIdentity, 1);
    for (let r = mc.top; r <= mc.bottom; r++) {
      for (let c = mc.left; c <= mc.right; c++) {
        if (r === mc.top && c === mc.left) continue;
        addCoordinateIndexEntry(mergeSkipSet, `${r}:${c}`, mergeIdentity);
      }
    }
  }

  const autoFilterCells = new Set<string>();
  if (worksheet.autoFilter) {
    const af = worksheet.autoFilter;
    const filterIdentity = coordinateIndexIdentity(
      worksheet,
      'worksheet-auto-filter-index',
      'expand-auto-filter-coordinates',
    );
    assertCoordinateRangeArea(
      { top: af.top, bottom: af.top, left: af.left, right: af.right },
      filterIdentity,
    );
    for (let c = af.left; c <= af.right; c++) {
      addCoordinateIndexEntry(autoFilterCells, `${af.top}:${c}`, filterIdentity);
    }
  }

  const hyperlinkMap = new Map<string, string>();
  const hyperlinkIdentity = coordinateIndexIdentity(
    worksheet,
    'worksheet-hyperlink-index',
    'index-hyperlink-coordinates',
  );
  for (const hl of worksheet.hyperlinks ?? []) {
    if (hl.url) setCoordinateIndexValue(hyperlinkMap, `${hl.row}:${hl.col}`, hl.url, hyperlinkIdentity);
  }

  const commentCells = new Set<string>();
  const commentIdentity = coordinateIndexIdentity(
    worksheet,
    'worksheet-comment-index',
    'index-comment-coordinates',
  );
  for (const ref of worksheet.commentRefs ?? []) {
    const parsed = parseA1Ref(ref);
    if (parsed) addCoordinateIndexEntry(commentCells, `${parsed.row}:${parsed.col}`, commentIdentity);
  }

  const entry: SheetRenderCache = {
    rowCount: worksheet.rows.length,
    coordinateIndexLimit,
    cellMap,
    nonEmptyColsByRow,
    cfContext: compileCf(worksheet, cellMap),
    mergeAnchorSet,
    mergeSkipSet,
    autoFilterCells,
    hyperlinkMap,
    commentCells,
    tableStyleMap: buildTableStyleMap(worksheet),
    pivotStyleMap: buildPivotStyleMap(worksheet),
    sparklineMap: buildSparklineMap(worksheet),
  };
  sheetRenderCache.set(worksheet, entry);
  return entry;
}
