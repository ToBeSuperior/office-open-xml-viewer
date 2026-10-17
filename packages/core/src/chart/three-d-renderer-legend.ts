// 3-D chart title / plot-area layout and legend painting. Split from
// three-d-renderer.ts (#1857); reached only through the optional three-d
// renderer chunk.

import type {
  ChartDataPointOverride,
  ChartExElementStyle,
  ChartLegendEntryOverride,
  ChartModel,
  ChartRect,
  ChartSeries,
} from '../types/chart.js';
import type { Fill } from '../types/common.js';
import { threeDDatumStyleIndex } from './three-d-datum-style.js';
import { markerFillColorFor } from './marker-style.js';
import { EMU_PER_PT } from '../units.js';
import {
  axisTitleMargin,
  axisTitleVerticalInsetPx,
  cartesianTitleBand,
  chartLegendBands,
  chartLegendReserve,
  chartAxisTitleBands,
  chartTextFontSizePx,
  computeChartFrame,
  packLegendRows,
  resolveManualLayoutRect,
} from './layout.js';
import { elideToWidth } from './text-elide.js';
import { fitDataLabelLines } from './data-label-layout.js';
import { paintLegendFrame } from './legend-frame.js';
import { paintPlotAreaFrame } from './plot-area-frame.js';
import { resolveNativeFill as resolveFill } from '../shape/paint.js';
import {
  chartStyleDirectFillDecision,
  chartStyleDirectLineDecision,
  chartStyleFillCascade,
  chartStyleLineDecision,
} from './style-paint.js';
import { drawingmlLineDashArray } from '../draw/dash.js';
import { chartStyleEffectOwner, paintChartStyleEffects } from './style-effects.js';
import {
  chartDataPointStyleRole,
  chartStyleDashChoice,
  rawLinkedChartStyleRole,
} from './effective-style.js';
import { threeDDatumPaint, colorFor } from './three-d-renderer-paint.js';
import { elementFontFamily, paintThreeDMarker } from './three-d-renderer-labels.js';

interface ThreeDLegendTextStyle {
  fontPx: number;
  font: string;
  color: string;
}

interface ThreeDLegendMeasure {
  labels: string[];
  styles: ThreeDLegendTextStyle[];
  itemWidths: number[];
}

function threeDLegendOverrideMap(chart: ChartModel): Map<number, ChartLegendEntryOverride> {
  const overrides = new Map<number, ChartLegendEntryOverride>();
  for (const override of chart.legendEntries ?? []) overrides.set(override.idx, override);
  return overrides;
}

function threeDLegendTextStyle(
  chart: ChartModel,
  override: ChartLegendEntryOverride | undefined,
  ptToPx: number,
): ThreeDLegendTextStyle {
  const fontPx = chartTextFontSizePx(
    override?.fontSizeHpt ?? chart.legendFontSizeHpt,
    ptToPx,
  ) ?? 9 * ptToPx;
  const face = override?.fontFace ?? chart.legendFontFace;
  const bold = override?.fontBold ?? chart.legendFontBold ?? false;
  const italic = override?.fontItalic ?? chart.legendFontItalic ?? false;
  return {
    fontPx,
    font: `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fontPx}px ${elementFontFamily(chart, face)}`,
    color: override?.fontColor
      ? `#${override.fontColor}`
      : chart.legendFontColor ? `#${chart.legendFontColor}` : '#595959',
  };
}

export function titleAndPlot(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
  orientation: 'vertical' | 'horizontal' | 'radial',
  shapeRotationDeg: number,
): { plot: ChartRect; legend: ChartRect | null; legendMeasure: ThreeDLegendMeasure } {
  const band = cartesianTitleBand(chart, rect.h, ptToPx);
  if (chart.title) {
    ctx.font = `${chart.titleFontBold === false ? '' : 'bold '}${band.fontPx}px ${elementFontFamily(chart, chart.titleFontFace)}`;
    ctx.fillStyle = chart.titleFontColor ? `#${chart.titleFontColor}` : '#111111';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    const measuredTitleWidth = ctx.measureText(chart.title).width;
    const automatic = {
      x: rect.x + (rect.w - measuredTitleWidth) / 2,
      y: rect.y + band.topPad,
      w: measuredTitleWidth,
      h: Math.max(1, band.fontPx),
    };
    const authored = chart.titleManualLayout
      ? resolveManualLayoutRect(
        { ...chart.titleManualLayout, w: undefined, h: undefined },
        rect,
        automatic,
      )
      : null;
    ctx.fillText(
      chart.title,
      authored ? authored.x + authored.w / 2 : rect.x + rect.w / 2,
      authored?.y ?? automatic.y,
    );
  }
  const legendOverrides = threeDLegendOverrideMap(chart);
  ctx.save();
  const rawLegendLabels = chart.chartType === 'pie'
    ? (chart.series[0]?.categories?.length ? chart.series[0].categories : chart.categories)
    : chart.series.map((series, index) => series.name || `Series ${index + 1}`);
  const legendLabels = rawLegendLabels.flatMap((label, index) =>
    legendOverrides.get(index)?.deleted === true ? [] : [{ label, index }]
  );
  const legendStyles = legendLabels.map(entry =>
    threeDLegendTextStyle(chart, legendOverrides.get(entry.index), ptToPx)
  );
  const legendFontPx = Math.max(0, ...legendStyles.map(style => style.fontPx));
  const itemWidths = legendLabels.map((entry, index) => {
      ctx.font = legendStyles[index].font;
      return 7 * ptToPx + 4 + ctx.measureText(entry.label).width;
    });
  const legendReserve = chartLegendReserve(chart, rect.w, rect.h, {
    itemWidths,
    rowHeight: Math.max(legendFontPx * 1.45, 12),
    itemGap: 12,
    horizontalPadding: 8,
    verticalPadding: 4,
  });
  ctx.restore();
  const legendBands = chartLegendBands(legendReserve, chart.legendOverlay === true);
  const axisBands = chartAxisTitleBands(chart, rect.w, rect.h, ptToPx);
  const catAxisTitleInset = axisTitleVerticalInsetPx(
    chart.catAxisTitleTextVerticalInsetEmu, ptToPx,
  );
  const valAxisTitleInset = axisTitleVerticalInsetPx(
    chart.valAxisTitleTextVerticalInsetEmu, ptToPx,
  );
  const leftTitleBand = orientation === 'horizontal'
    ? chart.catAxisTitle
      ? axisBands.catFontPx + catAxisTitleInset + axisTitleMargin(rect.w) + 4
      : 0
    : orientation === 'vertical' ? axisBands.valBandW : 0;
  const bottomTitleBand = orientation === 'horizontal'
    ? chart.valAxisTitle
      ? axisBands.valFontPx + valAxisTitleInset + axisTitleMargin(rect.h) + 4
      : 0
    : orientation === 'vertical' ? axisBands.catBandH : 0;
  const frame = computeChartFrame(chart, rect.x, rect.y, rect.w, rect.h, ptToPx, {
    titleBand: band,
    legendReserve,
    // Outer pads (4/5/19/13 % of the frame): library policy, see ChartPad in layout.ts.
    pad: {
      t: band.bandH + legendBands.legTopH + rect.h * 0.04,
      r: legendBands.legRightW + rect.w * 0.05,
      b: legendBands.legBottomH + rect.h * 0.19 + bottomTitleBand,
      l: legendBands.legLeftW + rect.w * 0.13 + leftTitleBand,
    },
    honorPlotAreaManualLayout: true,
    manualOuterInsets: {
      t: band.bandH,
      r: legendBands.legRightW,
      b: legendBands.legBottomH + bottomTitleBand,
      l: legendBands.legLeftW + leftTitleBand,
    },
  });
  const plot: ChartRect = {
    x: frame.plotRect.px0,
    y: frame.plotRect.py0,
    w: Math.max(1, frame.plotRect.pw),
    h: Math.max(1, frame.plotRect.ph),
  };
  paintPlotAreaFrame(
    ctx, chart, plot.x, plot.y, plot.w, plot.h, ptToPx, shapeRotationDeg,
  );
  const defaultLegend = !legendReserve ? null
    : legendReserve.side === 'r'
      ? { x: rect.x + rect.w - legendReserve.reserveW, y: plot.y, w: legendReserve.reserveW, h: plot.h }
      : legendReserve.side === 'l'
        ? { x: rect.x, y: plot.y, w: legendReserve.reserveW, h: plot.h }
        : legendReserve.side === 't'
          ? { x: rect.x + 4, y: rect.y + band.bandH, w: Math.max(1, rect.w - 8), h: legendReserve.reserveH }
          : { x: rect.x + 4, y: rect.y + rect.h - legendReserve.reserveH, w: Math.max(1, rect.w - 8), h: legendReserve.reserveH };
  const manualLegend = defaultLegend && chart.legendManualLayout
    ? resolveManualLayoutRect(chart.legendManualLayout, rect, defaultLegend)
    : null;
  return {
    plot,
    legend: manualLegend ?? defaultLegend,
    legendMeasure: {
      labels: legendLabels.map(entry => entry.label),
      styles: legendStyles,
      itemWidths,
    },
  };
}

interface ThreeDMarkerPaint {
  fill: string;
  fillPaint: Fill | null | undefined;
  line: string;
  linePaint: Fill | null | undefined;
  lineWidth: number;
  lineDash: string | null | undefined;
  lineCustomDash: ChartModel['plotAreaLineCustomDash'];
  lineCap: string | null | undefined;
  lineJoin: string | null | undefined;
  effectDirect: ChartExElementStyle | null | undefined;
  effectFallback: ChartExElementStyle | null | undefined;
  effectIndex: number;
  effectFallbackIndex: number;
}

/** Resolve a marker once for the 3-D plot and its legend key. Direct point and
 * series marker formatting remain authoritative; the effective linked-over-
 * numeric marker role supplies only omitted components. */
export function threeDMarkerPaint(
  chart: ChartModel,
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  pointIndex: number,
  seriesIndex: number,
  fallbackColor: string,
  ptToPx: number,
): ThreeDMarkerPaint {
  const linked = chartDataPointStyleRole(chart, 'dataPointMarker', seriesIndex);
  const rawLinked = rawLinkedChartStyleRole(chart, 'dataPointMarker');
  const seriesStyleIndex = series.chartexFormatIdx ?? seriesIndex;
  const pointStyleFill = chartStyleDirectFillDecision(
    point?.markerStyle, rawLinked, pointIndex,
  );
  const seriesStyleFill = chartStyleDirectFillDecision(
    series.markerStyle, rawLinked, seriesStyleIndex,
  );
  const pointFillAuthored = pointStyleFill !== undefined
    || point?.markerFillPaint !== undefined || point?.markerFill != null
    || point?.color != null
    || point?.markerFillPaintAuthored === true && point.markerStyle?.fillHidden !== true;
  const seriesFillAuthored = seriesStyleFill !== undefined
    || series.markerFillPaint !== undefined || series.markerFill != null
    || series.markerFillPaintAuthored === true && series.markerStyle?.fillHidden !== true;
  const linkedFill = chartStyleFillCascade(linked, rawLinked, seriesStyleIndex);
  const fillDecision = pointFillAuthored
    ? pointStyleFill ?? point?.markerFillPaint
    : seriesFillAuthored
      ? seriesStyleFill ?? series.markerFillPaint
      : linkedFill;
  let fill = markerFillColorFor(series, point, pointIndex, fallbackColor);
  let fillPaint = fillDecision;
  if (fillDecision?.fillType === 'solid') {
    fill = fillDecision.color;
    fillPaint = undefined;
  } else if (fillDecision === null) fill = '00000000';

  const pointLine = chartStyleDirectLineDecision(
    point?.markerStyle, rawLinked, pointIndex,
  );
  const seriesLine = chartStyleDirectLineDecision(
    series.markerStyle, rawLinked, seriesStyleIndex,
  );
  const pointLineAuthored = pointLine !== undefined || point?.markerLine != null
    || point?.markerLinePaintAuthored === true && point.markerStyle?.lineHidden !== true;
  const seriesLineAuthored = seriesLine !== undefined || series.markerLine != null
    || series.markerLinePaintAuthored === true && series.markerStyle?.lineHidden !== true;
  const linkedLine = chartStyleLineDecision(linked, seriesStyleIndex);
  const linePaint = pointLineAuthored
    ? pointLine ?? (point?.markerLinePaintAuthored === true ? null : undefined)
    : seriesLineAuthored
      ? seriesLine ?? (series.markerLinePaintAuthored === true ? null : undefined)
      : linkedLine;
  let line = point?.markerLine ?? series.markerLine ?? series.lineColor ?? fallbackColor;
  let effectiveLinePaint = linePaint;
  if (linePaint?.fillType === 'solid') {
    line = linePaint.color;
    effectiveLinePaint = undefined;
  } else if (linePaint === null) line = '00000000';
  const geometry = point?.markerStyle ?? series.markerStyle ?? linked;
  const dash = chartStyleDashChoice(point?.markerStyle, series.markerStyle, linked);
  return {
    fill,
    fillPaint,
    line,
    linePaint: effectiveLinePaint,
    lineWidth: (point?.markerLineWidthEmu ?? series.markerLineWidthEmu
      ?? geometry?.lineWidthEmu) != null
      ? Math.max(0.25, (point?.markerLineWidthEmu ?? series.markerLineWidthEmu
        ?? geometry?.lineWidthEmu ?? 0) / EMU_PER_PT * ptToPx)
      : Math.max(0.75, ptToPx),
    lineDash: dash?.lineDash,
    lineCustomDash: dash?.lineCustomDash,
    lineCap: geometry?.lineCap,
    lineJoin: geometry?.lineJoin,
    effectDirect: chartStyleEffectOwner(
      point?.markerStyle, point?.chartexStyle, series.markerStyle,
    ),
    effectFallback: linked,
    effectIndex: point?.markerStyle || point?.chartexStyle ? pointIndex : seriesStyleIndex,
    effectFallbackIndex: seriesStyleIndex,
  };
}

function paintThreeDDatumLegendSwatch(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  pointIndex: number,
  seriesIndex: number,
  x: number,
  y: number,
  key: number,
  ptToPx: number,
  shapeRotationDeg: number,
  pointOwned: boolean,
): void {
  const paint = threeDDatumPaint(
    chart, series, point, pointIndex, seriesIndex, pointOwned,
  );
  const styleIndex = threeDDatumStyleIndex(chart, series, pointIndex, seriesIndex);
  const bounds = { x, y: y - key / 2, w: key, h: key };
  const body = (target: CanvasRenderingContext2D): void => {
    if (paint.fill !== null) {
      const resolved = paint.fill === undefined
        ? paint.color
        : paint.fill.fillType === 'image'
          ? null // Legend picture swatches are intentionally fail-closed.
          : resolveFill(paint.fill, target, bounds.x, bounds.y, bounds.w, bounds.h,
              shapeRotationDeg);
      if (resolved) {
        target.fillStyle = resolved;
        target.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
      }
    }
    const linePaint = paint.lineFill === undefined
      ? paint.lineColor
      : paint.lineFill == null || paint.lineFill.fillType === 'image'
        ? null
        : resolveFill(paint.lineFill, target, bounds.x, bounds.y, bounds.w, bounds.h,
            shapeRotationDeg);
    if (!linePaint) return;
    target.strokeStyle = linePaint;
    target.lineWidth = paint.lineWidthEmu != null
      ? Math.max(0.25, paint.lineWidthEmu / EMU_PER_PT * ptToPx) : 0.75 * ptToPx;
    target.setLineDash(drawingmlLineDashArray(
      paint.lineCustomDash, paint.lineDash, target.lineWidth,
    ));
    target.lineCap = paint.lineCap;
    target.lineJoin = paint.lineJoin;
    target.strokeRect(bounds.x, bounds.y, bounds.w, bounds.h);
    target.setLineDash([]);
  };
  const directEffect = chartStyleEffectOwner(point?.chartexStyle, series.chartexStyle);
  paintChartStyleEffects(
    ctx,
    directEffect,
    chartDataPointStyleRole(chart, 'dataPoint3D', seriesIndex),
    directEffect === point?.chartexStyle ? pointIndex : styleIndex,
    bounds,
    ptToPx,
    body,
    styleIndex,
  );
}

function paintThreeDLineLegendKey(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  series: ChartSeries,
  seriesIndex: number,
  color: string,
  x: number,
  y: number,
  key: number,
  ptToPx: number,
  shapeRotationDeg: number,
): void {
  const datum = threeDDatumPaint(
    chart, series, undefined, seriesIndex, seriesIndex, false,
  );
  const datumFill = datum.lineFill === undefined
    ? datum.lineColor != null
      ? ({ fillType: 'solid', color: datum.lineColor.replace(/^#/, '') } as const)
      : datum.fill
    : datum.lineFill;
  if (datumFill !== null) {
    const lineWidth = datum.lineWidthEmu != null
      ? Math.max(0.5, datum.lineWidthEmu / EMU_PER_PT * ptToPx)
      : Math.max(1, 2 * ptToPx);
    const keyBounds = { x, y: y - lineWidth / 2, w: key, h: lineWidth };
    const resolved = datumFill === undefined
      ? datum.color
      : datumFill.fillType === 'image'
        ? null
        : datumFill.fillType === 'solid'
          ? `#${datumFill.color}`
        : resolveFill(datumFill, ctx, x, y - lineWidth / 2, key, lineWidth,
            shapeRotationDeg);
    if (resolved) {
      const paintLine = (target: CanvasRenderingContext2D): void => {
        target.beginPath();
        target.moveTo(x, y);
        target.lineTo(x + key, y);
        target.strokeStyle = target === ctx ? resolved : datumFill === undefined
          ? datum.color
          : datumFill.fillType === 'image'
            ? 'rgba(0,0,0,0)'
            : datumFill.fillType === 'solid'
              ? `#${datumFill.color}`
            : resolveFill(
                datumFill, target, x, y - lineWidth / 2, key, lineWidth,
                shapeRotationDeg,
              ) ?? 'rgba(0,0,0,0)';
        target.lineWidth = lineWidth;
        target.setLineDash(drawingmlLineDashArray(
          datum.lineCustomDash, datum.lineDash, lineWidth,
        ));
        target.stroke();
        target.setLineDash([]);
      };
      const styleIndex = threeDDatumStyleIndex(chart, series, seriesIndex, seriesIndex);
      paintChartStyleEffects(
        ctx,
        chartStyleEffectOwner(series.chartexStyle),
        chartDataPointStyleRole(chart, 'dataPoint3D', seriesIndex),
        styleIndex,
        keyBounds,
        ptToPx,
        paintLine,
        styleIndex,
      );
    }
  }
  if (series.showMarker !== true || series.markerSymbol === 'none') return;
  const symbol = series.markerSymbol ?? 'circle';
  const marker = threeDMarkerPaint(
    chart, series, undefined, seriesIndex, seriesIndex, color.replace(/^#/, ''), ptToPx,
  );
  paintThreeDMarker(
    ctx,
    { x: x + key / 2, y },
    symbol,
    Math.min(key, Math.max(2, (series.markerSize ?? 5) * ptToPx)),
    marker.fill === '00000000' ? 'transparent' : `#${marker.fill.replace(/^#/, '')}`,
    marker.line === '00000000' ? 'rgba(0,0,0,0)' : `#${marker.line.replace(/^#/, '')}`,
    marker.lineWidth,
    marker.fillPaint,
    shapeRotationDeg,
    ptToPx,
    marker.effectDirect,
    marker.effectFallback,
    marker.effectIndex,
    marker.effectFallbackIndex,
    marker.linePaint,
    marker.lineDash,
    marker.lineCustomDash,
    marker.lineCap,
    marker.lineJoin,
  );
}

export function simpleLegend(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  bounds: ChartRect | null,
  ptToPx: number,
  categoryDriven = false,
  measured?: ThreeDLegendMeasure,
  shapeRotationDeg = 0,
): void {
  if (!bounds) return;
  paintLegendFrame(ctx, chart, bounds, ptToPx, shapeRotationDeg);
  const indexedPoints = new Map<number, ChartDataPointOverride>(
    chart.series[0]?.dataPointOverrides?.map(override => [override.idx, override]) ?? [],
  );
  const rawEntries = categoryDriven
    ? (chart.series[0]?.categories?.length ? chart.series[0].categories : chart.categories)
      .map((label, index) => {
        const authored = chart.series[0]?.dataPointColors?.[index];
        return {
          label,
          color: authored === '00000000'
            ? 'transparent'
            : authored ? `#${authored}` : colorFor(index),
          series: chart.series[0],
          point: indexedPoints.get(index),
          sourceIndex: index,
          seriesIndex: 0,
        };
      })
    : chart.series.map((series, index) => ({
      label: series.name || `Series ${index + 1}`,
      color: colorFor(index, series),
      series,
      point: undefined,
      sourceIndex: index,
      seriesIndex: index,
    }));
  const legendOverrides = threeDLegendOverrideMap(chart);
  const entries = rawEntries.filter(entry =>
    legendOverrides.get(entry.sourceIndex)?.deleted !== true
  );
  const canReuseMeasure = measured != null
    && measured.labels.length === entries.length
    && measured.labels.every((label, index) => label === entries[index].label);
  const entryStyles = canReuseMeasure
    ? measured.styles
    : entries.map(entry =>
        threeDLegendTextStyle(chart, legendOverrides.get(entry.sourceIndex), ptToPx)
      );
  ctx.textAlign = 'left';
  ctx.textBaseline = 'middle';
  const row = Math.max(Math.max(0, ...entryStyles.map(style => style.fontPx)) * 1.45, 12);
  const key = Math.min(7 * ptToPx, row * 0.7);
  const horizontal = chart.legendPos === 't' || chart.legendPos === 'b'
    || (chart.legendManualLayout != null && bounds.w >= bounds.h);
  if (horizontal) {
    const itemWidths = canReuseMeasure
      ? measured.itemWidths
      : entries.map((entry, index) => {
          ctx.font = entryStyles[index].font;
          return key + 4 + ctx.measureText(entry.label).width;
        });
    const rows = packLegendRows(itemWidths, Math.max(1, bounds.w - 8), 12);
    const visibleRows = rows.slice(
      0,
      Math.max(0, Math.floor((bounds.h - 4 + 1e-6) / row)),
    );
    let rowY = bounds.y + 2 + row / 2;
    for (const indices of visibleRows) {
      const widths = indices.map(index => Math.min(bounds.w, itemWidths[index]));
      const total = widths.reduce((sum, width) => sum + width, 0)
        + Math.max(0, indices.length - 1) * 12;
      let itemX = bounds.x + Math.max(4, (bounds.w - total) / 2);
      for (let position = 0; position < indices.length; position++) {
        const index = indices[position];
        const entry = entries[index];
        const textStyle = entryStyles[index];
        ctx.font = textStyle.font;
        const available = Math.max(0, widths[position] - key - 4);
        const lineKey = !categoryDriven && chart.chartType.toLowerCase().includes('line');
        if (lineKey && entry.series) {
          paintThreeDLineLegendKey(
            ctx, chart, entry.series, entry.seriesIndex, entry.color,
            itemX, rowY, key, ptToPx, shapeRotationDeg,
          );
        } else if (entry.series) paintThreeDDatumLegendSwatch(
          ctx, chart, entry.series, entry.point, entry.sourceIndex, entry.seriesIndex,
          itemX, rowY, key, ptToPx, shapeRotationDeg, categoryDriven,
        );
        ctx.fillStyle = textStyle.color;
        ctx.fillText(elideToWidth(ctx, entry.label, available), itemX + key + 4, rowY);
        itemX += widths[position] + 12;
      }
      rowY += row;
    }
    return;
  }
  let rowTop = bounds.y;
  for (let index = 0; index < entries.length; index++) {
    const entry = entries[index];
    const textStyle = entryStyles[index];
    ctx.font = textStyle.font;
    const textX = bounds.x + 8 + key;
    const availableTextWidth = Math.max(0, bounds.x + bounds.w - 4 - textX);
    const lineHeight = Math.max(textStyle.fontPx * 1.2, 10);
    const lines = fitDataLabelLines(
      entry.label,
      availableTextWidth,
      lineHeight * (categoryDriven ? 1 : 2),
      lineHeight,
      value => ctx.measureText(value).width,
    );
    if (lines.length === 0) continue;
    const itemHeight = Math.max(row, lines.length * lineHeight + 2);
    if (rowTop + itemHeight > bounds.y + bounds.h + 1e-6) break;
    const cy = rowTop + itemHeight / 2;
    const lineKey = !categoryDriven && chart.chartType.toLowerCase().includes('line');
    if (lineKey && entry.series) {
      paintThreeDLineLegendKey(
        ctx, chart, entry.series, entry.seriesIndex, entry.color,
        bounds.x + 4, cy, key, ptToPx, shapeRotationDeg,
      );
    } else if (entry.series) paintThreeDDatumLegendSwatch(
      ctx, chart, entry.series, entry.point, entry.sourceIndex, entry.seriesIndex,
      bounds.x + 4, cy, key, ptToPx, shapeRotationDeg, categoryDriven,
    );
    ctx.fillStyle = textStyle.color;
    const firstY = cy - (lines.length - 1) * lineHeight / 2;
    lines.forEach((line, lineIndex) => ctx.fillText(line, textX, firstY + lineIndex * lineHeight));
    rowTop += itemHeight;
  }
}
