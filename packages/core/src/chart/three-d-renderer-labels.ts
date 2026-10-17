// 3-D chart labels: category labels, markers, data labels, axis titles, and
// the series axis. Split from three-d-renderer.ts (#1857); reached only
// through the optional three-d renderer chunk.

import type {
  ChartDataLabelOverride,
  ChartDisplayUnits,
  ChartExElementStyle,
  ChartModel,
  ChartRect,
  ChartSeries,
} from '../types/chart.js';
import type { Fill } from '../types/common.js';
import { paintChartImageFill } from './image-fill.js';
import { dataLabelIsDeleted } from './data-label-style.js';
import { EMU_PER_PT, PT_TO_PX } from '../units.js';
import { pptxPresetDashArray } from '../draw/dash.js';
import type { NumericValueAxisPlan } from './axis-scale.js';
import {
  axisTitleFontPx,
  axisTitleMargin,
  axisTitleRotationRad,
  chartTextFontSizePx,
  resolveManualLayoutRect,
  type ChartAxisTitleSide,
} from './layout.js';
import { elideToWidth } from './text-elide.js';
import { resolveThemeFontRef } from './shared/fonts.js';
import { resolveDataLabelPlacement } from './data-label-layout.js';
import { paintRichDataLabelBlock, resolveRichDataLabelBlock } from './rich-data-label.js';
import { effectiveDataLabelText } from './data-label-content.js';
import { mergeChartLabelBoxes, paintChartLabelBox } from './label-box.js';
import {
  anchoredDataLabelPoint,
  dataLabelCanvasTextAlign,
  dataLabelInsets,
  effectiveDataLabelTextStyle,
  fitStyledDataLabelLines,
  rotatedDataLabelSize,
  transformDataLabelText,
} from './data-label-style.js';
import type { ChartThreeDProjection } from './three-d.js';
import { resolveNativeFill as resolveFill } from '../shape/paint.js';
import { drawingmlLineDashArray } from '../draw/dash.js';
import { paintChartStyleEffects } from './style-effects.js';
import type { Point } from './three-d-renderer-paint.js';
import {
  threeDAxisGeometry,
  threeDWallGeometry,
  applyThreeDStroke,
  threeDAxisStroke,
  threeDAxisTickStroke,
} from './three-d-renderer-walls.js';

// Library policy: when neither the element nor the theme supplies a face, the
// 3-D renderer falls back to Arial. ECMA-376 defines no last-resort face; it is
// left to the consumer. An unresolved theme reference (`+mn-*` / `+mj-*` with no
// theme threaded into the chart) takes the same fallback.
const fontFamily = (face: string | null | undefined): string => {
  const safe = face && !face.startsWith('+') ? face.replace(/["\\]/g, '') : 'Arial';
  return `"${safe}"`;
};

/** Resolve an element face, including a DrawingML theme font reference
 *  (`+mj-*` / `+mn-*`, ECMA-376 §20.1.4.1.16 against the theme font scheme),
 *  before applying {@link fontFamily}. An absent face keeps the fallback. */
export const elementFontFamily = (chart: ChartModel, face: string | null | undefined): string =>
  fontFamily(resolveThemeFontRef(chart, face));

export const chartFontFamily = (
  chart: ChartModel,
  face: string | null | undefined,
  role: 'major' | 'minor' = 'minor',
): string => {
  const reference = face?.startsWith('+mj-')
    ? chart.themeMajorFontLatin
    : face?.startsWith('+mn-') ? chart.themeMinorFontLatin : face;
  const resolved = reference ?? (role === 'major' ? chart.themeMajorFontLatin : chart.themeMinorFontLatin);
  return fontFamily(resolved);
};

export function authoredCategoryRotation(chart: ChartModel): number | null {
  const raw = chart.catAxisLabelRotation;
  if (raw == null) return null;
  if (!Number.isFinite(raw) || Math.abs(raw) > 5_400_000) return 0;
  return raw / 60_000 * Math.PI / 180;
}

export function drawAngledCategoryLabel(
  ctx: CanvasRenderingContext2D,
  label: string,
  point: Point,
  rotation: number,
  horizontal: boolean,
  outwardY = 1,
  offset = 6,
): void {
  if (horizontal || rotation === 0) {
    ctx.textAlign = horizontal ? 'right' : 'center';
    ctx.textBaseline = horizontal ? 'middle' : outwardY < 0 ? 'bottom' : 'top';
    ctx.fillText(
      label,
      point.x + (horizontal ? -offset : 0),
      point.y + (horizontal ? 0 : outwardY * offset),
    );
    return;
  }
  ctx.save();
  ctx.translate(point.x, point.y + outwardY * offset);
  ctx.rotate(rotation);
  ctx.textAlign = outwardY < 0 ? 'left' : 'right';
  ctx.textBaseline = 'middle';
  ctx.fillText(label, 0, 0);
  ctx.restore();
}

export function paintThreeDMarker(
  ctx: CanvasRenderingContext2D,
  point: Point,
  symbol: string,
  size: number,
  fill: string,
  line: string,
  lineWidth: number,
  fillPaint: Fill | null | undefined = undefined,
  shapeRotationDeg = 0,
  ptToPx = PT_TO_PX,
  effectDirect: ChartExElementStyle | null | undefined = undefined,
  effectFallback: ChartExElementStyle | null | undefined = undefined,
  effectIndex = 0,
  effectFallbackIndex = effectIndex,
  linePaint: Fill | null | undefined = undefined,
  lineDash: string | null | undefined = undefined,
  lineCustomDash: ChartModel['plotAreaLineCustomDash'] = undefined,
  lineCap: string | null | undefined = undefined,
  lineJoin: string | null | undefined = undefined,
): void {
  if (!(size > 0) || !Number.isFinite(point.x) || !Number.isFinite(point.y)) return;
  const radius = size / 2;
  if (effectDirect !== undefined || effectFallback !== undefined) {
    paintChartStyleEffects(
      ctx,
      effectDirect,
      effectFallback,
      effectIndex,
      { x: point.x - radius, y: point.y - radius, w: size, h: size },
      ptToPx,
      target => paintThreeDMarker(
        target, point, symbol, size, fill, line, lineWidth,
        fillPaint, shapeRotationDeg, ptToPx,
        undefined, undefined, 0, 0,
        linePaint, lineDash, lineCustomDash, lineCap, lineJoin,
      ),
      effectFallbackIndex,
    );
    return;
  }
  ctx.beginPath();
  const resolvedLinePaint = linePaint === undefined
    ? line
    : linePaint == null
      ? null
      : linePaint.fillType === 'solid'
        ? `#${linePaint.color}`
        : resolveFill(
            linePaint, ctx, point.x - radius, point.y - radius, size, size, shapeRotationDeg,
          );
  if (resolvedLinePaint) {
    ctx.strokeStyle = resolvedLinePaint;
    ctx.lineWidth = lineWidth;
    ctx.setLineDash(drawingmlLineDashArray(lineCustomDash, lineDash, lineWidth));
    ctx.lineCap = lineCap === 'rnd' ? 'round' : lineCap === 'sq' ? 'square' : 'butt';
    ctx.lineJoin = lineJoin === 'round' || lineJoin === 'bevel' ? lineJoin : 'miter';
  }
  switch (symbol) {
    case 'square':
      ctx.rect(point.x - radius, point.y - radius, size, size);
      break;
    case 'diamond':
      ctx.moveTo(point.x, point.y - radius);
      ctx.lineTo(point.x + radius, point.y);
      ctx.lineTo(point.x, point.y + radius);
      ctx.lineTo(point.x - radius, point.y);
      ctx.closePath();
      break;
    case 'triangle':
      ctx.moveTo(point.x, point.y - radius);
      ctx.lineTo(point.x + radius, point.y + radius);
      ctx.lineTo(point.x - radius, point.y + radius);
      ctx.closePath();
      break;
    case 'x':
    case 'plus': {
      const diagonal = symbol === 'x';
      ctx.moveTo(point.x - radius, point.y + (diagonal ? -radius : 0));
      ctx.lineTo(point.x + radius, point.y + (diagonal ? radius : 0));
      ctx.moveTo(point.x + (diagonal ? -radius : 0), point.y + radius);
      ctx.lineTo(point.x + (diagonal ? radius : 0), point.y - radius);
      if (resolvedLinePaint) ctx.stroke();
      return;
    }
    case 'dash':
      // ECMA-376 ST_MarkerStyle `dash` is the same filled short rectangle in
      // the 2-D and optional 3-D paths; unlike x/plus it consumes fill paint.
      ctx.rect(point.x - radius, point.y - size * 0.1, size, size * 0.2);
      break;
    case 'star': {
      for (let index = 0; index < 10; index++) {
        const angle = -Math.PI / 2 + index * Math.PI / 5;
        const r = index % 2 === 0 ? radius : radius * 0.4;
        const x = point.x + Math.cos(angle) * r;
        const y = point.y + Math.sin(angle) * r;
        index ? ctx.lineTo(x, y) : ctx.moveTo(x, y);
      }
      ctx.closePath();
      break;
    }
    case 'dot':
      // ECMA-376 §21.2.3.27: width=1/2 and height=1/5 of marker size.
      ctx.ellipse(point.x, point.y, size * 0.25, size * 0.1, 0, 0, Math.PI * 2);
      break;
    case 'picture': {
      if (fillPaint?.fillType === 'image') {
        paintChartImageFill(
          ctx, fillPaint, point.x - radius, point.y - radius, size, size, ptToPx,
          shapeRotationDeg,
        );
      }
      if (resolvedLinePaint) ctx.strokeRect(point.x - radius, point.y - radius, size, size);
      return;
    }
    default:
      ctx.arc(point.x, point.y, radius, 0, Math.PI * 2);
      break;
  }
  const imageFill = fillPaint?.fillType === 'image' ? fillPaint : undefined;
  const resolved = imageFill ? null : fillPaint === undefined
    ? (fill === 'transparent' ? null : fill)
    : fillPaint == null
      ? null
      : resolveFill(
          fillPaint, ctx, point.x - radius, point.y - radius, size, size, shapeRotationDeg,
        );
  if (resolved != null) {
    ctx.fillStyle = resolved;
    ctx.fill();
  } else if (imageFill) {
    ctx.save();
    ctx.clip();
    paintChartImageFill(
      ctx, imageFill, point.x - radius, point.y - radius, size, size, ptToPx,
      shapeRotationDeg,
    );
    ctx.restore();
  }
  if (resolvedLinePaint) ctx.stroke();
}

export function drawThreeDDataLabel(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  series: ChartSeries,
  seriesIndex: number,
  categoryIndex: number,
  value: number,
  anchor: Point,
  bounds: ChartRect,
  ptToPx: number,
  markerGap = 0,
  percentValue?: number,
  resolvedOverride?: ChartDataLabelOverride,
  defaultPosition = 't',
  leaderAnchor: Point = anchor,
  leaderLineEligible = true,
  valueDisplayUnits?: ChartDisplayUnits | null,
  axisMaximum?: number,
  plottedValueForAxis = value,
  shapeRotationDeg = 0,
): void {
  // Callers resolve indexed overrides through their per-series Map before the
  // paint loop. Falling back to Array.find here would quietly reintroduce
  // quadratic work for a fully-authored maximum-size public model.
  const override = resolvedOverride;
  const defaults = series.seriesDataLabels;
  if (dataLabelIsDeleted(defaults, override)) return;
  if (chart.showDataLabelsOverMax !== true
    && axisMaximum != null
    && Number.isFinite(axisMaximum)
    && plottedValueForAxis > axisMaximum) return;
  const showVal = override?.showVal ?? defaults?.showVal ?? chart.showDataLabels;
  const showCat = override?.showCatName ?? defaults?.showCatName ?? false;
  const showSeries = override?.showSerName ?? defaults?.showSerName ?? false;
  const showPercent = override?.showPercent ?? defaults?.showPercent ?? false;
  const explicitText = override?.text;
  const text = effectiveDataLabelText({
    customText: explicitText,
    showCategory: showCat,
    showSeries,
    showValue: showVal,
    showPercent,
    category: series.categories?.[categoryIndex]
      ?? chart.categories[categoryIndex] ?? `${categoryIndex + 1}`,
    seriesName: series.name || `Series ${seriesIndex + 1}`,
    sourceValue: value,
    valueDivisor: valueDisplayUnits?.divisor,
    percentRatio: percentValue != null && Number.isFinite(percentValue) ? percentValue : undefined,
    formatCode: override?.formatCode
      ?? defaults?.formatCode ?? chart.dataLabelFormatCode ?? series.valFormatCode,
    percentFormatCode: override?.formatCode
      ?? defaults?.formatCode ?? chart.dataLabelFormatCode ?? '0%',
    date1904: chart.date1904,
    separator: override?.separator ?? defaults?.separator,
  });
  if (!text) return;
  const fontPx = chartTextFontSizePx(
    override?.fontSizeHpt ?? defaults?.fontSizeHpt ?? chart.dataLabelFontSizeHpt,
    ptToPx,
  ) ?? 9 * ptToPx;
  const bold = override?.fontBold ?? defaults?.fontBold ?? chart.dataLabelFontBold ?? false;
  const textStyle = effectiveDataLabelTextStyle(override, defaults);
  const fallbackFamily = chartFontFamily(
    chart, override?.fontFace ?? defaults?.fontFace ?? chart.dataLabelFontFace,
  );
  ctx.font = `${textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fontPx}px ${fallbackFamily}`;
  const fallbackColor = `#${override?.fontColor ?? defaults?.fontColor
    ?? series.labelColor ?? chart.dataLabelFontColor ?? '111111'}`;
  const rich = explicitText && override?.richRuns?.length
    ? resolveRichDataLabelBlock(ctx, {
      runs: override.richRuns,
      ptToPx,
      fontFamily: fallbackFamily,
      fallbackBold: bold,
      fallbackItalic: textStyle.fontItalic,
      fallbackBaseline: textStyle.fontBaseline,
      fallbackColorHidden: textStyle.fontPaintAuthored === true
        && (textStyle.fontHidden === true || textStyle.fontColor == null),
      fontFamilyForFace: face => chartFontFamily(chart, face),
    }, fontPx, fallbackColor)
    : null;
  const lines = rich ? [] : fitStyledDataLabelLines(
    text,
    Math.max(fontPx, bounds.w * 0.45),
    Math.max(fontPx * 1.2, bounds.h * 0.35),
    fontPx * 1.2,
    value => ctx.measureText(value).width,
    textStyle,
  );
  if (!rich && !lines.length) return;
  const textWidth = rich?.width ?? Math.max(...lines.map(line => ctx.measureText(line).width));
  const textHeight = rich?.height ?? lines.length * fontPx * 1.2;
  const insets = dataLabelInsets(textStyle, ptToPx);
  const rotated = rotatedDataLabelSize(
    textWidth + insets.left + insets.right,
    textHeight + insets.top + insets.bottom,
    textStyle.textRotation,
    textStyle.textVerticalMode,
  );
  const placement = resolveDataLabelPlacement(
    {
      kind: 'point', x: anchor.x, y: anchor.y,
      position: override?.position ?? defaults?.position ?? chart.dataLabelPosition ?? defaultPosition,
      markerGap,
    },
    bounds,
    { w: rotated.w, h: rotated.h },
    fontPx,
    override?.manualLayout,
    bounds,
  );
  if (!placement) return;
  const labelBox = mergeChartLabelBoxes(override?.labelBox, defaults?.labelBox);
  if (defaults?.showLeaderLines && defaults.leaderLineHidden !== true && leaderLineEligible) {
    ctx.beginPath();
    ctx.moveTo(leaderAnchor.x, leaderAnchor.y);
    ctx.lineTo(
      Math.max(placement.rect.x, Math.min(leaderAnchor.x, placement.rect.x + placement.rect.w)),
      Math.max(placement.rect.y, Math.min(leaderAnchor.y, placement.rect.y + placement.rect.h)),
    );
    ctx.strokeStyle = `#${defaults.leaderLineColor ?? '808080'}`;
    ctx.lineWidth = defaults.leaderLineWidthEmu != null
      ? Math.max(0.25, defaults.leaderLineWidthEmu / EMU_PER_PT * ptToPx)
      : 0.75 * ptToPx;
    ctx.setLineDash(pptxPresetDashArray(defaults.leaderLineDash ?? 'solid', ctx.lineWidth));
    ctx.stroke();
  }
  paintChartLabelBox(ctx, labelBox, placement.rect, ptToPx, shapeRotationDeg);
  ctx.save();
  ctx.beginPath();
  ctx.rect(placement.clip.x, placement.clip.y, placement.clip.w, placement.clip.h);
  ctx.clip();
  const paintAlign = dataLabelCanvasTextAlign(textStyle, placement.textAlign);
  const anchored = anchoredDataLabelPoint(
    placement.x, placement.y, placement.rect,
    textHeight + insets.top + insets.bottom, textStyle, override?.manualLayout != null,
    paintAlign, placement.textAlign,
    textWidth + insets.left + insets.right, rotated.radians,
  );
  const transformed = transformDataLabelText(
    ctx, anchored.x, anchored.y, rotated.radians, paintAlign,
    placement.textBaseline, insets,
  );
  if (rich) {
    paintRichDataLabelBlock(
      ctx, rich, transformed.x, transformed.y, paintAlign, placement.textBaseline,
      override?.manualLayout
        ? Math.max(0, placement.rect.w - insets.left - insets.right) : rich.width,
    );
    ctx.restore();
    return;
  }
  if (!(textStyle.fontPaintAuthored === true
    && (textStyle.fontHidden === true || textStyle.fontColor == null))) {
    ctx.fillStyle = fallbackColor;
    ctx.textAlign = paintAlign;
    ctx.textBaseline = 'middle';
    const lineHeight = fontPx * 1.2;
    const baselineShift = (textStyle.fontBaseline ?? 0) * fontPx;
    const firstY = transformed.y - (lines.length - 1) * lineHeight / 2 - baselineShift;
    lines.forEach((line, index) => ctx.fillText(line, transformed.x, firstY + index * lineHeight));
  }
  ctx.restore();
}

export function wantsThreeDDataLabel(
  chart: ChartModel,
  series: ChartSeries,
  override: ChartDataLabelOverride | undefined,
): boolean {
  const defaults = series.seriesDataLabels;
  if (dataLabelIsDeleted(defaults, override)) return false;
  if (override?.text) return true;
  return (override?.showVal ?? defaults?.showVal ?? chart.showDataLabels)
    || (override?.showCatName ?? defaults?.showCatName ?? false)
    || (override?.showSerName ?? defaults?.showSerName ?? false)
    || (override?.showPercent ?? defaults?.showPercent ?? false);
}

function drawThreeDAxisTitle(
  ctx: CanvasRenderingContext2D,
  text: string,
  chartRect: ChartRect,
  anchor: Point,
  side: ChartAxisTitleSide,
  fontSizePx: number,
  fontCssFamily: string,
  color: string | null | undefined,
  bold: boolean,
  italic: boolean,
  authoredRotation: number | null | undefined,
  authoredVerticalMode: ChartModel['catAxisTitleVerticalMode'],
  manualLayout: ChartModel['catAxisTitleManualLayout'],
  maxPx: number,
): void {
  ctx.save();
  ctx.font = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fontSizePx}px ${fontCssFamily}`;
  ctx.fillStyle = color ? `#${color}` : '#555';
  const label = manualLayout ? text : elideToWidth(ctx, text, maxPx);
  const rotation = axisTitleRotationRad(side, authoredRotation, authoredVerticalMode);
  let center = anchor;
  if (manualLayout) {
    const textWidth = ctx.measureText(label).width;
    const cos = Math.abs(Math.cos(rotation));
    const sin = Math.abs(Math.sin(rotation));
    const automatic = {
      x: anchor.x - (textWidth * cos + fontSizePx * sin) / 2,
      y: anchor.y - (textWidth * sin + fontSizePx * cos) / 2,
      w: textWidth * cos + fontSizePx * sin,
      h: textWidth * sin + fontSizePx * cos,
    };
    const resolved = resolveManualLayoutRect(
      { ...manualLayout, w: undefined, h: undefined }, chartRect, automatic,
    );
    if (resolved) center = { x: resolved.x + resolved.w / 2, y: resolved.y + resolved.h / 2 };
  }
  ctx.translate(center.x, center.y);
  if (rotation) ctx.rotate(rotation);
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(label, 0, 0);
  ctx.restore();
}

/** Paint 3-D cartesian axis titles through the same authored font/orientation
 * contract as the 2-D families. Projection selects the axis rules; titles are
 * screen annotations in the already-reserved outer gutter, as in Office. */
export function drawThreeDAxisTitles(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  chartRect: ChartRect,
  plot: ChartRect,
  horizontal: boolean,
  ptToPx: number,
): void {
  if (chart.valAxisTitle) {
    const fontPx = axisTitleFontPx(chart.valAxisTitleFontSizeHpt, ptToPx);
    const side: ChartAxisTitleSide = horizontal ? 'horizontal' : 'left';
    drawThreeDAxisTitle(
      ctx, chart.valAxisTitle, chartRect,
      horizontal
        ? { x: plot.x + plot.w / 2, y: plot.y + plot.h + axisTitleMargin(chartRect.h) + fontPx / 2 }
        : { x: plot.x - axisTitleMargin(chartRect.w) - fontPx / 2, y: plot.y + plot.h / 2 },
      side, fontPx, chartFontFamily(chart, chart.valAxisTitleFontFace, 'major'), chart.valAxisTitleFontColor,
      chart.valAxisTitleFontBold ?? true, chart.valAxisTitleFontItalic ?? false,
      chart.valAxisTitleRotation, chart.valAxisTitleVerticalMode,
      chart.valAxisTitleManualLayout, horizontal ? plot.w : plot.h,
    );
  }
  if (chart.catAxisTitle) {
    const fontPx = axisTitleFontPx(chart.catAxisTitleFontSizeHpt, ptToPx);
    const side: ChartAxisTitleSide = horizontal ? 'left' : 'horizontal';
    drawThreeDAxisTitle(
      ctx, chart.catAxisTitle, chartRect,
      horizontal
        ? { x: plot.x - axisTitleMargin(chartRect.w) - fontPx / 2, y: plot.y + plot.h / 2 }
        : { x: plot.x + plot.w / 2, y: plot.y + plot.h + axisTitleMargin(chartRect.h) + fontPx / 2 },
      side, fontPx, chartFontFamily(chart, chart.catAxisTitleFontFace, 'major'), chart.catAxisTitleFontColor,
      chart.catAxisTitleFontBold ?? true, chart.catAxisTitleFontItalic ?? false,
      chart.catAxisTitleRotation, chart.catAxisTitleVerticalMode,
      chart.catAxisTitleManualLayout, horizontal ? plot.h : plot.w,
    );
  }
}

/** Paint `<c:serAx>` for the `standard` 3-D bar arrangement. Unlike a legend,
 * this is a projected coordinate axis: its rule and tick anchors use the same
 * model-space depth coordinate as the bar solids. */
export function drawThreeDSeriesAxis(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  chartRect: ChartRect,
  projection: ChartThreeDProjection,
  axis: NumericValueAxisPlan,
  categoryCount: number,
  categoryBetween: boolean,
  categoryReversed: boolean,
  orientation: 'vertical' | 'horizontal',
  barFamily: boolean,
  ptToPx: number,
): void {
  const spec = chart.threeD?.seriesAxis;
  if (!spec || spec.hidden || chart.threeD?.barGrouping !== 'standard' || chart.series.length === 0) return;
  const geometry = threeDAxisGeometry(
    chart, projection, axis, categoryCount, categoryBetween, categoryReversed, orientation,
  );
  const wallGeometry = threeDWallGeometry(projection);
  // A series axis is the depth edge opposite the visible side wall. In the
  // default Office view the side wall is on the left and `<c:serAx>` is the
  // right-hand floor edge; anchoring it to the value-axis corner incorrectly
  // puts depth ticks and labels on the left.
  const seriesAxisX = orientation === 'vertical'
    ? wallGeometry.seriesAxisX : geometry.axisX;
  const seriesAxisY = orientation === 'horizontal'
    ? (wallGeometry.floorY === projection.front.y
      ? projection.front.y + projection.front.h : projection.front.y)
    : wallGeometry.floorY;
  const start = projection.project(
    seriesAxisX, seriesAxisY, projection.topology.nearDepth,
  );
  const end = projection.project(
    seriesAxisX, seriesAxisY, projection.topology.farDepth,
  );
  const length = Math.hypot(end.x - start.x, end.y - start.y);
  if (!(length > 1e-6)) return;
  const tangent = { x: (end.x - start.x) / length, y: (end.y - start.y) / length };
  let normal = { x: -tangent.y, y: tangent.x };
  const sceneCenter = projection.project(
    projection.front.x + projection.front.w / 2,
    projection.front.y + projection.front.h / 2,
    0.5,
  );
  const midpoint = { x: (start.x + end.x) / 2, y: (start.y + end.y) / 2 };
  if ((midpoint.x - sceneCenter.x) * normal.x + (midpoint.y - sceneCenter.y) * normal.y < 0) {
    normal = { x: -normal.x, y: -normal.y };
  }
  if ((spec.tickLabelPos === 'low' && normal.y < 0)
    || (spec.tickLabelPos === 'high' && normal.y > 0)) {
    normal = { x: -normal.x, y: -normal.y };
  }

  if (!spec.lineHidden) {
    applyThreeDStroke(ctx, threeDAxisStroke(
      spec.lineColor, spec.lineWidthEmu, spec.lineDash, ptToPx,
    ));
    ctx.beginPath();
    ctx.moveTo(start.x, start.y);
    ctx.lineTo(end.x, end.y);
    ctx.stroke();
  }

  const markSkip = Math.max(1, Math.floor(spec.tickMarkSkip ?? 1));
  const labelSkip = Math.max(1, Math.floor(spec.tickLabelSkip ?? 1));
  const tickMode = spec.majorTickMark ?? 'out';
  const fontPx = chartTextFontSizePx(spec.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  ctx.font = `${spec.fontItalic ? 'italic ' : ''}${spec.fontBold ? 'bold ' : ''}${fontPx}px ${chartFontFamily(chart, spec.fontFace)}`;
  ctx.fillStyle = spec.fontColor ? `#${spec.fontColor}` : '#595959';
  ctx.textAlign = Math.abs(normal.x) < 0.2 ? 'center' : normal.x < 0 ? 'right' : 'left';
  ctx.textBaseline = Math.abs(normal.y) < 0.2 ? 'middle' : normal.y < 0 ? 'bottom' : 'top';
  if (barFamily && !spec.lineHidden && tickMode !== 'none') {
    applyThreeDStroke(ctx, threeDAxisTickStroke(
      spec.lineColor, spec.lineWidthEmu, spec.lineDash, ptToPx,
    ));
    const boundaryCount = chart.series.length;
    for (let boundaryIndex = 0; boundaryIndex <= boundaryCount; boundaryIndex += markSkip) {
      const authoredDepth = boundaryIndex / boundaryCount;
      const depth = spec.orientation === 'maxMin' ? 1 - authoredDepth : authoredDepth;
      const point = projection.project(seriesAxisX, seriesAxisY, depth);
      const total = 6 * ptToPx;
      const outer = tickMode === 'cross' ? total / 2 : tickMode === 'out' ? total : 0;
      const inner = tickMode === 'cross' ? total / 2 : tickMode === 'in' ? total : 0;
      ctx.beginPath();
      ctx.moveTo(point.x + normal.x * outer, point.y + normal.y * outer);
      ctx.lineTo(point.x - normal.x * inner, point.y - normal.y * inner);
      ctx.stroke();
    }
  }
  for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
    const authoredDepth = projection.seriesDepth(seriesIndex, chart.series.length, false);
    const depth = spec.orientation === 'maxMin' ? 1 - authoredDepth : authoredDepth;
    const point = projection.project(seriesAxisX, seriesAxisY, depth);
    if (!barFamily && !spec.lineHidden && seriesIndex % markSkip === 0 && tickMode !== 'none') {
      const total = 6 * ptToPx;
      const outer = tickMode === 'cross' ? total / 2 : tickMode === 'out' ? total : 0;
      const inner = tickMode === 'cross' ? total / 2 : tickMode === 'in' ? total : 0;
      ctx.beginPath();
      ctx.moveTo(point.x + normal.x * outer, point.y + normal.y * outer);
      ctx.lineTo(point.x - normal.x * inner, point.y - normal.y * inner);
      ctx.stroke();
    }
    if (spec.tickLabelPos !== 'none' && seriesIndex % labelSkip === 0) {
      ctx.fillText(
        chart.series[seriesIndex].name || `Series ${seriesIndex + 1}`,
        point.x + normal.x * (6 * ptToPx + 3),
        point.y + normal.y * (6 * ptToPx + 3),
      );
    }
  }
  ctx.setLineDash([]);

  if (spec.title) {
    const titleFontPx = axisTitleFontPx(spec.titleFontSizeHpt, ptToPx);
    drawThreeDAxisTitle(
      ctx,
      spec.title,
      chartRect,
      {
        x: midpoint.x + normal.x * (fontPx + titleFontPx + 12),
        y: midpoint.y + normal.y * (fontPx + titleFontPx + 12),
      },
      'horizontal',
      titleFontPx,
      chartFontFamily(chart, spec.titleFontFace, 'major'),
      spec.titleFontColor,
      spec.titleFontBold ?? true,
      spec.titleFontItalic ?? false,
      spec.titleRotation,
      spec.titleVerticalMode,
      spec.titleManualLayout,
      Math.max(projection.front.w, projection.front.h),
    );
  }
}
