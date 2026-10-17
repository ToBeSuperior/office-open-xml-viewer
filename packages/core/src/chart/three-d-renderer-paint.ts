// 3-D chart scene types and paint planning: primitive budgets, percent
// category scales, stacked buckets, and per-datum paint resolution. Split
// from three-d-renderer.ts (#1857); reached only through the optional
// three-d renderer chunk.

import type { ChartDataPointOverride, ChartModel, ChartSeries } from '../types/chart.js';
import type { Fill } from '../types/common.js';
import { chartThreeDDatumFillDecision } from './three-d-datum-style.js';
import { markerPaintComponents } from './marker-style.js';
import { EMU_PER_PT, PT_TO_PX } from '../units.js';
import {
  chartThreeDSurfacePaint,
  chartStyleDirectLineDecision,
  chartStyleDirectNoLineDecision,
  chartStyleLineCascade,
} from './style-paint.js';
import {
  MAX_CHART_PAINT_COMPONENTS,
  MAX_CHART_PAINT_RECIPE_COMPONENTS,
} from './resource-limits.js';
import {
  chartDataPointStyleRole,
  chartSeriesVariesByPoint,
  chartStyleDashChoice,
  rawLinkedChartStyleRole,
} from './effective-style.js';

export const PALETTE = ['4472C4', 'ED7D31', '70AD47', 'A5A5A5', 'FFC000', '5B9BD5'] as const;
const MAX_THREE_D_PAINT_RECIPE_COMPONENTS = MAX_CHART_PAINT_RECIPE_COMPONENTS;
const MAX_THREE_D_PAINT_COMPONENTS = MAX_CHART_PAINT_COMPONENTS;
export const SUPPORTED_CARTESIAN_THREE_D_TYPES = new Set([
  'line', 'stackedLine', 'stackedLinePct',
  'area', 'stackedArea', 'stackedAreaPct',
  'clusteredBar', 'clusteredBarH',
  'stackedBar', 'stackedBarH', 'stackedBarPct', 'stackedBarHPct',
]);

/** Office renders authored 3-D mesh edges in chart/sheet display pixels, not
 * with the 96-DPI typography expansion used for text. The supplied Excel
 * reference has `a:ln@w=12700` (1pt) and a one-pixel solid at 100% zoom across
 * all six equivalent 3-D bar charts. Keep that observed material-edge rule
 * separate from 2-D DrawingML strokes and scale it only with viewer zoom. */
export function threeDMeshOutlineWidthPx(lineWidthEmu: number, ptToPx: number): number {
  const zoom = Number.isFinite(ptToPx) && ptToPx > 0 ? ptToPx / PT_TO_PX : 1;
  const points = Number.isFinite(lineWidthEmu) && lineWidthEmu >= 0
    ? lineWidthEmu / EMU_PER_PT : 0;
  return Math.max(0.25, points) * zoom;
}

/** Group stack geometry in one pass before resolving shared internal caps.
 * The parser/render preflight bounds the item count, but rescanning the whole
 * primitive array once per category still turns a valid wide chart into
 * quadratic CPU work. Invalid public-model category indexes are ignored. */
export function bucketThreeDStackItems<T extends { readonly categoryIndex: number }>(
  items: readonly T[],
  categoryCount: number,
): T[][] {
  const boundedCount = Number.isSafeInteger(categoryCount) && categoryCount > 0
    ? categoryCount : 0;
  const buckets = Array.from({ length: boundedCount }, () => [] as T[]);
  for (const item of items) {
    const categoryIndex = item.categoryIndex;
    if (Number.isSafeInteger(categoryIndex)
      && categoryIndex >= 0
      && categoryIndex < boundedCount) {
      buckets[categoryIndex].push(item);
    }
  }
  return buckets;
}

export interface Point { x: number; y: number }

export interface SceneFace {
  points: readonly Point[];
  color: string;
  /** One already-resolved paint shared by every face of the same datum.
   * `null` is authored noFill/unresolved and suppresses the automatic solid. */
  paint?: string | CanvasGradient | CanvasPattern | null;
  paintRole: 'fill' | 'outline';
  shade: number;
  cameraDepth: number;
  cameraDepths?: readonly number[];
  cameraWeights?: readonly number[];
  outline: boolean;
  outlineColor?: string;
  outlineWidth?: number;
  outlineDash?: string;
  outlineCap?: CanvasLineCap;
  outlineJoin?: CanvasLineJoin;
}

export interface MeshOutlineStyle {
  readonly color: string;
  readonly width: number;
  readonly dash: readonly number[];
  readonly cap: CanvasLineCap;
  readonly join: CanvasLineJoin;
}

export interface ScenePrimitiveBudget {
  remaining: number;
  exceeded: boolean;
}

export const colorFor = (index: number, series?: ChartSeries): string =>
  `#${series?.color ?? PALETTE[index % PALETTE.length]}`;

export interface ThreeDDatumPaint {
  color: string;
  fill: Fill | null | undefined;
  lineColor: string | null;
  lineFill: Fill | null | undefined;
  lineWidthEmu: number | null | undefined;
  lineDash: string | null | undefined;
  lineCustomDash: ChartModel['plotAreaLineCustomDash'];
  lineCap: CanvasLineCap;
  lineJoin: CanvasLineJoin;
}

function threeDPaintsWithinBudget<T>(
  items: readonly T[],
  fillFor: (item: T) => Fill | null | undefined,
  lineFor: (item: T) => Fill | null | undefined,
): boolean {
  let total = 0;
  for (const item of items) {
    for (const recipe of [fillFor(item), lineFor(item)]) {
      if (recipe == null) continue;
      const components = markerPaintComponents(recipe);
      if ((recipe.fillType === 'gradient'
          && components > MAX_THREE_D_PAINT_RECIPE_COMPONENTS)
        || components > MAX_THREE_D_PAINT_COMPONENTS - total) return false;
      total += components;
    }
  }
  return true;
}

interface ThreeDPercentCategoryScale {
  maxMagnitude: number;
  scaledTotal: number;
}

/** Plan percent denominators once per category in O(series × categories).
 * Dividing by the category maximum first keeps finite Number.MAX_VALUE inputs
 * proportional without overflowing their absolute-value sum. */
export function planThreeDPercentCategoryScales(
  chart: ChartModel,
  categoryCount: number,
): ThreeDPercentCategoryScale[] {
  const scales = Array.from({ length: categoryCount }, () => ({
    maxMagnitude: 0,
    scaledTotal: 0,
  }));
  for (const series of chart.series) {
    for (let index = 0; index < categoryCount; index++) {
      const candidate = series.values[index];
      if (candidate != null && Number.isFinite(candidate)) {
        scales[index].maxMagnitude = Math.max(
          scales[index].maxMagnitude,
          Math.abs(candidate),
        );
      }
    }
  }
  for (const series of chart.series) {
    for (let index = 0; index < categoryCount; index++) {
      const candidate = series.values[index];
      const maxMagnitude = scales[index].maxMagnitude;
      if (candidate != null && Number.isFinite(candidate) && maxMagnitude > 0) {
        scales[index].scaledTotal += Math.abs(candidate) / maxMagnitude;
      }
    }
  }
  return scales;
}

/** Convert one value into the plotted percent-stack domain from the bounded
 * category plan shared by preflight, extent planning, and geometry paint. */
export function threeDPlottedValue(
  chart: ChartModel,
  seriesIndex: number,
  categoryIndex: number,
  percentScales: readonly ThreeDPercentCategoryScale[] | undefined,
): number {
  const raw = chart.series[seriesIndex]?.values[categoryIndex] ?? 0;
  if (!Number.isFinite(raw)) return 0;
  if (!percentScales) return raw;
  const { maxMagnitude, scaledTotal } = percentScales[categoryIndex]
    ?? { maxMagnitude: 0, scaledTotal: 0 };
  if (!(maxMagnitude > 0)) return 0;
  return scaledTotal > 0 ? raw / maxMagnitude / scaledTotal * 100 : 0;
}

export function threeDSafeAdd(left: number, right: number): number {
  const sum = left + right;
  if (Number.isFinite(sum)) return sum;
  return right < 0 ? -Number.MAX_VALUE : Number.MAX_VALUE;
}

/** Preflight every structured recipe that the selected optional 3-D family
 * can consume before title, plot, wall, or datum paint begins. This keeps a
 * public model or linked style with an oversized gradient atomic instead of
 * leaving a partially painted chart behind. */
export function threeDChartPaintsWithinBudget(chart: ChartModel): boolean {
  const paints: Array<{
    fill: Fill | null | undefined;
    line: Fill | null | undefined;
  }> = [];
  for (const [surface, role] of [
    [chart.threeD?.floor, 'floor'],
    [chart.threeD?.sideWall, 'wall'],
    [chart.threeD?.backWall, 'wall'],
  ] as const) {
    const paint = chartThreeDSurfacePaint(chart, surface, role);
    paints.push({ fill: paint.fill, line: paint.line });
  }

  if (chart.chartType === 'pie') {
    const series = chart.series[0];
    if (series) {
      const overrides = new Map(
        series.dataPointOverrides?.map(point => [point.idx, point]) ?? [],
      );
      for (let index = 0; index < series.values.length; index++) {
        const value = series.values[index];
        if (value == null || !Number.isFinite(value) || Math.abs(value) <= 0) continue;
        const paint = threeDDatumPaint(chart, series, overrides.get(index), index, 0);
        paints.push({ fill: paint.fill, line: paint.lineFill });
      }
    }
  } else {
    const bars = chart.chartType === 'clusteredBar'
      || chart.chartType === 'clusteredBarH'
      || chart.chartType.startsWith('stackedBar');
    if (bars) {
      const stacked = chart.chartType.startsWith('stacked');
      const percent = chart.chartType.endsWith('Pct');
      const zeroBlanks = stacked || chart.dispBlanksAs === 'zero';
      const logarithmic = chart.valAxisLogBase != null
        && Number.isFinite(chart.valAxisLogBase) && chart.valAxisLogBase >= 2;
      const categoryCount = Math.max(
        1,
        chart.categories.length,
        ...chart.series.map(series => Math.max(
          series.values.length,
          series.categories?.length ?? 0,
        )),
      );
      const percentScales = percent
        ? planThreeDPercentCategoryScales(chart, categoryCount)
        : undefined;
      const visibleInterval = (base: number, end: number): boolean => {
        if (base === end) return false;
        const lower = Math.min(base, end);
        const upper = Math.max(base, end);
        const min = chart.valMin;
        const max = chart.valMax;
        if (min != null && Number.isFinite(min) && upper <= min) return false;
        if (max != null && Number.isFinite(max) && lower >= max) return false;
        return !logarithmic || upper > 0;
      };
      const positiveBase = new Array(categoryCount).fill(0) as number[];
      const negativeBase = new Array(categoryCount).fill(0) as number[];
      for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
        const series = chart.series[seriesIndex];
        const overrides = new Map(
          series.dataPointOverrides?.map(point => [point.idx, point]) ?? [],
        );
        for (let index = 0; index < categoryCount; index++) {
          const value = series.values[index];
          const visible = value != null && Number.isFinite(value)
            && (!logarithmic || value > 0)
            || (value == null && zeroBlanks);
          if (!visible) continue;
          const plottedValue = threeDPlottedValue(chart, seriesIndex, index, percentScales);
          const base = stacked
            ? (plottedValue >= 0 ? positiveBase[index] : negativeBase[index])
            : 0;
          const end = threeDSafeAdd(base, plottedValue);
          if (stacked) {
            if (plottedValue >= 0) positiveBase[index] = end;
            else negativeBase[index] = end;
          }
          if (!visibleInterval(base, end)) continue;
          const paint = threeDDatumPaint(
            chart, series, overrides.get(index), index, seriesIndex,
          );
          paints.push({ fill: paint.fill, line: paint.lineFill });
        }
      }
    } else {
      const area = chart.chartType.toLowerCase().includes('area');
      const stacked = chart.chartType.startsWith('stacked');
      const percent = chart.chartType.endsWith('Pct');
      const zeroBlanks = stacked || chart.dispBlanksAs === 'zero';
      const logarithmic = chart.valAxisLogBase != null
        && Number.isFinite(chart.valAxisLogBase) && chart.valAxisLogBase >= 2;
      const categoryCount = Math.max(
        chart.categories.length,
        ...chart.series.map(series => Math.max(
          series.categories?.length ?? 0,
          series.values.length,
        )),
      );
      const percentScales = percent
        ? planThreeDPercentCategoryScales(chart, categoryCount)
        : undefined;
      const stackedUpper = chart.series.map(() => new Array<number>(categoryCount).fill(0));
      if (stacked) {
        const positive = new Array<number>(categoryCount).fill(0);
        const negative = new Array<number>(categoryCount).fill(0);
        for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
          for (let index = 0; index < categoryCount; index++) {
            const value = threeDPlottedValue(chart, seriesIndex, index, percentScales);
            const base = value >= 0 ? positive[index] : negative[index];
            const end = threeDSafeAdd(base, value);
            stackedUpper[seriesIndex][index] = end;
            if (value >= 0) positive[index] = end;
            else negative[index] = end;
          }
        }
      }
      for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
        const series = chart.series[seriesIndex];
        const overrides = new Map(
          series.dataPointOverrides?.map(point => [point.idx, point]) ?? [],
        );
        const pointCount = categoryCount;
        let runLength = 0;
        let previousValue: number | null = null;
        let hasGeometry = false;
        let defaultLinePaintUsed = false;
        for (let index = 0; index < pointCount; index++) {
          const value = series.values[index];
          const visible = value != null && Number.isFinite(value)
            && (!logarithmic || value > 0)
            || (value == null && zeroBlanks);
          if (visible) {
            runLength++;
            const currentValue = stacked
              ? stackedUpper[seriesIndex][index]
              : threeDPlottedValue(chart, seriesIndex, index, percentScales);
            const lower = previousValue == null ? currentValue : Math.min(previousValue, currentValue);
            const upper = previousValue == null ? currentValue : Math.max(previousValue, currentValue);
            const outsideAuthoredAxis = chart.valMin != null && Number.isFinite(chart.valMin)
              && upper <= chart.valMin
              || chart.valMax != null && Number.isFinite(chart.valMax) && lower >= chart.valMax;
            if (runLength >= 2 && (area || series.smooth === true || !outsideAuthoredAxis)) {
              hasGeometry = true;
              if (area) break;
              const point = overrides.get(index);
              if (threeDPointOwnsLineSegment(point)) {
                const paint = threeDDatumPaint(chart, series, point, index, seriesIndex);
                paints.push({ fill: undefined, line: paint.lineFill });
              } else {
                defaultLinePaintUsed = true;
              }
            }
            previousValue = currentValue;
          } else if ((chart.dispBlanksAs ?? 'gap') === 'gap') {
            runLength = 0;
            previousValue = null;
          }
        }
        if (!hasGeometry) continue;
        const paint = threeDDatumPaint(chart, series, undefined, seriesIndex, seriesIndex);
        if (area) paints.push({ fill: paint.fill, line: paint.lineFill });
        else if (defaultLinePaintUsed) paints.push({ fill: undefined, line: paint.lineFill });
      }
    }
  }
  return threeDPaintsWithinBudget(paints, paint => paint.fill, paint => paint.line);
}

function solidOrStructured(
  decision: Fill | null | undefined,
  fallbackColor: string,
): { color: string; fill: Fill | null | undefined } {
  if (decision?.fillType === 'solid') {
    return { color: `#${decision.color}`, fill: undefined };
  }
  return { color: fallbackColor, fill: decision };
}

/** Resolve one classic 3-D datum without materializing per-face style state.
 * Direct point > direct series > linked dataPoint3D > automatic. Solid colors
 * continue through the bounded automatic material; structured recipes are
 * resolved once for the complete datum after projection. */
export function threeDDatumPaint(
  chart: ChartModel,
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  pointIndex: number,
  seriesIndex: number,
  includePointColor = true,
): ThreeDDatumPaint {
  // `ChartSeries.color` and `dataPointColors` are effective legacy colors:
  // the parser may materialize the automatic accent palette there. Keep them
  // as the final automatic fallback so an authored/linked 3-D paint recipe can
  // still own the mark. Direct spPr provenance is carried by chartexStyle.
  const automaticColor = includePointColor && series.dataPointColors?.[pointIndex]
    ? `#${series.dataPointColors[pointIndex]}`
    : colorFor(seriesIndex, series);
  // Chart Style palettes normally advance by series. Pie slices, and the
  // single-series bar case explicitly carrying varyColors, advance by point.
  // Keep this independent from `pointIndex`: that index still selects direct
  // dPt formatting, while linked/series roles describe the owning series.
  const styleIndex = chartSeriesVariesByPoint(chart, seriesIndex)
    ? pointIndex
    : series.chartexFormatIdx ?? seriesIndex;
  const family = series.seriesType ?? chart.chartType;
  const lineRibbonFamily = family === 'line'
    || family === 'stackedLine'
    || family === 'stackedLinePct';
  const linkedDataPointStyle = chartDataPointStyleRole(chart, 'dataPoint3D', seriesIndex);
  const rawLinkedDataPointStyle = rawLinkedChartStyleRole(chart, 'dataPoint3D');
  let fill: Fill | null | undefined;
  let color = automaticColor;
  const fillDecision = chartThreeDDatumFillDecision(
    chart, series, point, pointIndex, seriesIndex,
  );
  if (fillDecision !== undefined) ({ color, fill } = solidOrStructured(fillDecision, color));

  let lineFill: Fill | null | undefined;
  let lineColor: string | null = null;
  let pointLine = chartStyleDirectLineDecision(
    point?.chartexStyle, rawLinkedDataPointStyle, pointIndex,
  );
  if (pointLine === undefined && point?.lineHidden === true) {
    pointLine = chartStyleDirectNoLineDecision(rawLinkedDataPointStyle);
  }
  if (pointLine !== undefined) {
    if (pointLine?.fillType === 'solid') lineColor = `#${pointLine.color}`;
    else lineFill = pointLine;
  }
  else if (point?.lineColor) lineColor = `#${point.lineColor}`;
  else {
    let seriesLine = chartStyleDirectLineDecision(
      series.chartexStyle, rawLinkedDataPointStyle, styleIndex,
    );
    if (seriesLine === undefined && series.lineHidden === true) {
      seriesLine = chartStyleDirectNoLineDecision(rawLinkedDataPointStyle);
    }
    if (seriesLine !== undefined) {
      if (seriesLine?.fillType === 'solid') lineColor = `#${seriesLine.color}`;
      else lineFill = seriesLine;
    }
    else if (series.lineColor) lineColor = `#${series.lineColor}`;
    else if (lineRibbonFamily) {
      // ECMA-376 §21.2.3.46 Table 1 classifies every 3-D chart mark under
      // "Fills for Data Points (3-D)". A Line3D ribbon is therefore painted
      // by dataPoint3D's fill component; its Table 5 No Line suppresses only
      // the solid's outline, not the ribbon itself. Direct c:ser/c:dPt line
      // formatting above still owns the segment and can explicitly hide it.
      lineFill = fill === null ? null : fill;
    } else {
      const linkedLine = chartStyleLineCascade(
        linkedDataPointStyle, rawLinkedDataPointStyle, styleIndex,
        point?.chartexStyle, series.chartexStyle,
      );
      if (linkedLine !== undefined) {
        if (linkedLine?.fillType === 'solid') lineColor = `#${linkedLine.color}`;
        else lineFill = linkedLine;
      }
    }
  }
  const linkedGeometry = lineRibbonFamily ? undefined : linkedDataPointStyle;
  const lineCapValue = point?.chartexStyle?.lineCap
    ?? series.chartexStyle?.lineCap ?? linkedGeometry?.lineCap;
  const lineJoinValue = point?.chartexStyle?.lineJoin
    ?? series.chartexStyle?.lineJoin ?? linkedGeometry?.lineJoin;
  const pointDash = point?.lineDash != null
    ? { lineDash: point.lineDash, lineDashAuthored: true }
    : point?.chartexStyle;
  const dash = chartStyleDashChoice(pointDash, series.chartexStyle, linkedGeometry);
  return {
    color,
    fill,
    lineColor,
    lineFill,
    lineWidthEmu: point?.lineWidthEmu ?? point?.chartexStyle?.lineWidthEmu
      ?? series.lineWidthEmu ?? series.chartexStyle?.lineWidthEmu
      ?? linkedGeometry?.lineWidthEmu,
    lineDash: dash?.lineDash,
    lineCustomDash: dash?.lineCustomDash,
    lineCap: lineCapValue === 'rnd' ? 'round' : lineCapValue === 'sq' ? 'square' : 'butt',
    lineJoin: lineJoinValue === 'round' || lineJoinValue === 'bevel'
      ? lineJoinValue : 'miter',
  };
}

/** A 3-D line dPt owns only the incoming segment's line component. Shape fill
 * remains irrelevant to Excel's line ribbon, and area dPt paint is not applied
 * to the series body. Keep this predicate narrow so color-only fill overrides
 * do not split or recolor line geometry. */
export function threeDPointOwnsLineSegment(
  point: ChartDataPointOverride | undefined,
): boolean {
  if (!point) return false;
  const style = point.chartexStyle;
  return point.lineHidden != null
    || point.lineColor != null
    || point.lineWidthEmu != null
    || point.lineDash != null
    || style?.linePaintAuthored === true
    || style?.lineHidden != null
    || style?.lineNoStyle === true
    || style?.lineColors?.some(color => color != null) === true
    || style?.linePaints?.some(paint => paint != null) === true
    || style?.lineWidthEmu != null
    || style?.lineDash != null
    || style?.lineCustomDash != null
    || style?.lineCap != null
    || style?.lineJoin != null;
}
