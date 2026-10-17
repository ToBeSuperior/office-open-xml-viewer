import type {
  ChartDataLabelOverride,
  ChartDataPointOverride,
  ChartModel,
  ChartRect,
} from '../types/chart.js';
import type { Fill } from '../types/common.js';
import {
  chartThreeDDatumFillDecision,
  threeDDatumStyleIndex,
} from './three-d-datum-style.js';
import {
  effectiveMarkerSymbol,
  hasVisiblePointMarkerOverride,
  seriesHasMarkerDetail,
} from './marker-style.js';

import { EMU_PER_PT, PT_TO_PX } from '../units.js';
import { finiteDataExtent } from './axis-scale.js';
import { formatCategoryLabel, formatChartValWithCode } from './chart-number-format.js';
import { chartTextFontSizePx } from './layout.js';
import { resolveRichDataLabelBlock } from './rich-data-label.js';
import { effectiveDataLabelText } from './data-label-content.js';
import {
  fitChartThreeDProjectionToPoints,
  fitChartThreeDProjectionToWallThickness,
  pieThreeDThicknessMultiplier,
  planChartThreeDProjection,
  planThreeDBarClusterSlot,
  threeDToMaxScale,
} from './three-d.js';
import { scaleHexColor } from './material-color.js';
import { categoryLabelOffsetPx, categoryPositionFraction } from './category-spacing.js';
import {
  buildThreeDLineRibbonMesh,
  buildThreeDPieSectorMesh,
  type ThreeDMesh,
} from './three-d-mesh.js';
import {
  clipAxisFractionBand,
  clipAxisFractionSegment,
  sortProjectedSceneFaces,
  type ProjectedSceneFace,
} from './three-d-scene.js';
import {
  buildProjectedStrokePrimitives,
  MAX_PROJECTED_STROKE_PRIMITIVES,
  type ProjectedStrokePoint,
} from './three-d-stroke.js';
import { drawingmlLineDashArray } from '../draw/dash.js';
import { chartStyleEffectOwner, paintChartStyleOuterEffectBehind } from './style-effects.js';
import { chartDataPointStyleRole } from './effective-style.js';
import {
  threeDDatumPaint,
  colorFor,
  planThreeDPercentCategoryScales,
  threeDPlottedValue,
  threeDSafeAdd,
  threeDMeshOutlineWidthPx,
  bucketThreeDStackItems,
  threeDPointOwnsLineSegment,
  PALETTE,
  SUPPORTED_CARTESIAN_THREE_D_TYPES,
  threeDChartPaintsWithinBudget,
  type Point,
  type SceneFace,
  type MeshOutlineStyle,
  type ScenePrimitiveBudget,
  type ThreeDDatumPaint,
} from './three-d-renderer-paint.js';
import {
  isTransparentPaint,
  shapeMeshFaces,
  applyScenePaint,
  paintThreeDTooManyDataPoints,
  paintSceneFace,
  paintSceneEffectSilhouette,
  areaStripFaces,
  projectThreeDMesh,
  projectThreeDPieOutline,
} from './three-d-renderer-scene.js';
import {
  threeDAxisGeometry,
  axisPlan,
  walls,
  frontAxes,
  cartesianAxisTicks,
  screenAnnotationOutward,
  projectedAxisTickLabelOffsetPx,
  threeDPieSliceAngles,
  type ThreeDAxisGeometry,
} from './three-d-renderer-walls.js';
import {
  elementFontFamily,
  paintThreeDMarker,
  wantsThreeDDataLabel,
  drawThreeDDataLabel,
  drawThreeDSeriesAxis,
  authoredCategoryRotation,
  drawAngledCategoryLabel,
  drawThreeDAxisTitles,
  chartFontFamily,
} from './three-d-renderer-labels.js';
import { titleAndPlot, threeDMarkerPaint, simpleLegend } from './three-d-renderer-legend.js';

// Historical three-d-renderer.ts exports (tests and the internal
// three-d-renderer entry import them from here); the implementations live in
// the three-d-renderer-* modules split out for #1857.
export { elementFontFamily } from './three-d-renderer-labels.js';
export { threeDPieSliceAngles, threeDWallGeometry } from './three-d-renderer-walls.js';
export type { ThreeDWallGeometry, ThreeDPieSliceAngles } from './three-d-renderer-walls.js';
export { threeDMeshOutlineWidthPx, bucketThreeDStackItems } from './three-d-renderer-paint.js';

function renderCartesian(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
  shapeRotationDeg: number,
): boolean {
  if (!chart.threeD) return false;
  const bars = chart.chartType === 'clusteredBar'
    || chart.chartType === 'clusteredBarH'
    || chart.chartType.startsWith('stackedBar');
  const horizontal = chart.chartType.endsWith('H') || chart.chartType.includes('BarH');
  const stacked = chart.chartType.startsWith('stacked');
  const depthArranged = bars && !stacked && chart.threeD.barGrouping === 'standard';
  const { plot, legend, legendMeasure } = titleAndPlot(
    ctx, chart, rect, ptToPx, horizontal ? 'horizontal' : 'vertical', shapeRotationDeg,
  );
  let projection = planChartThreeDProjection(chart.threeD, plot, {
    // Office places bar/column prisms in a compact depth box, while line and
    // area series planes span a substantially deeper Z scene. Projection and
    // view angles remain identical; only the authored family geometry differs.
    // A clustered/stacked bar group is one compact Z slab because its series
    // are separated on the category axis (or accumulated on the value axis).
    // A `standard` bar group instead uses `<c:serAx>` as a real third axis.
    // ECMA-376 §21.2.2.41 defines the authored model percentage but does not
    // prescribe Office's final projected auto-fit. For the default
    // h=100/depth=100/elevation=15/rotation=20/perspective=30 view, measured
    // Office axes are width:height:depth = 8.1:8.1:2.6. Standard Bar uses the
    // authored ECMA FOV directly; the stronger line/area compatibility gain
    // made its convergence visibly too severe.
    sceneDepthScale: bars ? (depthArranged ? 0.65 : 0.10) : 0.40,
    perspectiveTangentGain: depthArranged ? 1 : 2,
    // MS-OE376 §2.1.1501(a) assigns omitted hPercent to Office automatic
    // scaling. Repeated Excel/PDF line/area probes resolve that automatic
    // model-space height to one third of the scene width; an authored
    // hPercent continues to own the ratio in planChartThreeDProjection.
    sceneHeightScale: !bars
      && !(chart.threeD.heightPercentAuthored ?? chart.threeD.heightPercent != null)
      ? 1 / 3
      : undefined,
  });
  if (!projection) return true;
  projection = fitChartThreeDProjectionToWallThickness(projection, chart.threeD, plot);
  const percent = chart.chartType.endsWith('Pct');
  const categories = chart.series.find(series => (series.categories?.length ?? 0) > 0)?.categories
    ?? chart.categories;
  const categoryCount = Math.max(1, categories.length, ...chart.series.map(series => series.values.length));
  const percentScales = percent
    ? planThreeDPercentCategoryScales(chart, categoryCount)
    : undefined;
  const categoryReversed = chart.catAxisOrientation === 'maxMin';
  const categoryBetween = chart.catAxisCrossBetween === 'between';
  const dispBlanks = chart.dispBlanksAs ?? 'gap';
  const logarithmic = chart.valAxisLogBase != null
    && Number.isFinite(chart.valAxisLogBase)
    && chart.valAxisLogBase >= 2;
  const hasFiniteValue = (seriesIndex: number, categoryIndex: number): boolean => {
    const raw = chart.series[seriesIndex]?.values[categoryIndex];
    return raw != null && Number.isFinite(raw) && (!logarithmic || raw > 0)
      || (raw == null && (stacked || dispBlanks === 'zero'));
  };
  const valueAt = (seriesIndex: number, categoryIndex: number): number => {
    return threeDPlottedValue(chart, seriesIndex, categoryIndex, percentScales);
  };
  let dataMin = 0;
  // Percent-stacked extents follow the data signs. Starting every percent
  // chart at +100 makes an all-negative stack reserve an empty positive half
  // of the scene and moves the zero plane away from its authored edge.
  let dataMax = 0;
  if (stacked) {
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
      let positive = 0;
      let negative = 0;
      for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
        const value = valueAt(seriesIndex, categoryIndex);
        if (value >= 0) positive = threeDSafeAdd(positive, value);
        else negative = threeDSafeAdd(negative, value);
      }
      dataMin = Math.min(dataMin, negative);
      dataMax = Math.max(dataMax, positive);
    }
    if (percent) {
      // Match the shared 2-D percent-stacked contract: each occupied sign gets
      // a complete 100% side of the axis, not merely the largest category's
      // signed share. Authored min/max still win in axisPlan().
      dataMin = dataMin < 0 ? -100 : 0;
      dataMax = dataMax > 0 ? 100 : 0;
      if (dataMin === 0 && dataMax === 0) dataMax = 1;
    }
  } else {
    const finiteValues = chart.series.flatMap(series => series.values).filter(
      (value): value is number => value != null
        && Number.isFinite(value)
        && (!logarithmic || value > 0),
    );
    const extent = finiteDataExtent(
      finiteValues,
      logarithmic ? { min: 1, max: 10 } : { min: 0, max: 1 },
    );
    dataMin = logarithmic ? extent.min : Math.min(0, extent.min);
    dataMax = logarithmic ? extent.max : Math.max(0, extent.max);
  }
  const projectedValueAxisStart = horizontal
    ? projection.project(
        projection.front.x,
        projection.topology.axisY === 'min'
          ? projection.front.y
          : projection.front.y + projection.front.h,
        projection.topology.nearDepth,
      )
    : projection.project(
        projection.topology.axisX === 'min'
          ? projection.front.x
          : projection.front.x + projection.front.w,
        projection.front.y,
        projection.topology.nearDepth,
      );
  const projectedValueAxisEnd = horizontal
    ? projection.project(
        projection.front.x + projection.front.w,
        projection.topology.axisY === 'min'
          ? projection.front.y
          : projection.front.y + projection.front.h,
        projection.topology.nearDepth,
      )
    : projection.project(
        projection.topology.axisX === 'min'
          ? projection.front.x
          : projection.front.x + projection.front.w,
        projection.front.y + projection.front.h,
        projection.topology.nearDepth,
      );
  const projectedValueAxisLenPt = Math.hypot(
    projectedValueAxisEnd.x - projectedValueAxisStart.x,
    projectedValueAxisEnd.y - projectedValueAxisStart.y,
  ) / ptToPx;
  const axis = axisPlan(
    chart, dataMin, dataMax, projectedValueAxisLenPt,
    percent, horizontal ? 'horizontal' : 'vertical',
  );
  // A column/bar datum is zero only while zero belongs to the value-axis
  // domain. Logarithmic axes and explicitly positive/negative-only bounds
  // place the visible datum at the nearest authored edge; projecting zero in
  // log space otherwise creates a many-decades-off-screen cuboid.
  const visibleAxisValue = (value: number): number => {
    if (!Number.isFinite(value)) return axis.min;
    return Math.max(axis.min, Math.min(axis.max, value));
  };
  if (!bars) {
    walls(
      ctx, chart, projection, axis, horizontal ? 'horizontal' : 'vertical',
      categoryCount, categoryBetween, categoryReversed, ptToPx,
    );
  }
  const { front } = projection;
  const seriesCount = Math.max(1, chart.series.length);
  const deferredDataLabels: Array<() => void> = [];
  // Classic parser caches may carry thousands of indexed overrides. Resolve
  // them once per series; point-by-point Array.find would turn a bounded 10k
  // Canvas paint into quadratic work.
  const pointOverrides = chart.series.map(series => new Map<number, ChartDataPointOverride>(
    series.dataPointOverrides?.map(override => [override.idx, override]) ?? [],
  ));
  const labelOverrides = chart.series.map(series => new Map<number, ChartDataLabelOverride>(
    series.dataLabelOverrides?.map(override => [override.idx, override]) ?? [],
  ));
  if (bars) {
    // ECMA-376 §21.2.2.77 keeps two non-stacked 3-D arrangements distinct:
    // `clustered` series are adjacent on the category axis and share one depth
    // interval, while `standard` series reuse one category footprint and
    // occupy separate slots on `<c:serAx>`. Public models that omit the new
    // carrier retain the pre-existing clustered fallback.
    const clusterDepth = projection.prismInterval(0, 1, true);
    const primitives: Array<{
      x: number; y: number; width: number; height: number;
      nearDepth: number; farDepth: number;
      categoryIndex: number; seriesIndex: number; color: string;
      fillPaint: Fill | null | undefined;
      lineFill: Fill | null | undefined;
      shape: string; baseCoord: number; endCoord: number; endScale: number;
      baseScale: number; omitBaseCap: boolean; omitEndCap: boolean;
      outline: boolean; outlineColor: string; outlineWidth: number;
      outlineDash: string; outlineCustomDash: ChartModel['plotAreaLineCustomDash'];
      outlineCap: CanvasLineCap; outlineJoin: CanvasLineJoin;
      labelValue: number;
      plottedLabelValue: number;
    }> = [];
    const positiveBase = new Array(categoryCount).fill(0) as number[];
    const negativeBase = new Array(categoryCount).fill(0) as number[];
    const interval = (horizontal ? front.h : front.w) / categoryCount;
    const gapWidth = chart.barGapWidth != null
      && Number.isFinite(chart.barGapWidth)
      && chart.barGapWidth >= 0
      ? chart.barGapWidth
      : 150;
    for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
      const series = chart.series[seriesIndex];
      const depthInterval = depthArranged
        ? projection.prismInterval(seriesIndex, seriesCount, false)
        : clusterDepth;
      const clusterSlot = planThreeDBarClusterSlot(
        interval,
        gapWidth,
        depthArranged ? 0 : seriesIndex,
        depthArranged ? 1 : seriesCount,
        stacked || depthArranged,
      );
      for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
        if (!hasFiniteValue(seriesIndex, categoryIndex)) continue;
        const value = valueAt(seriesIndex, categoryIndex);
        const pointOverride = pointOverrides[seriesIndex].get(categoryIndex);
        const paint = threeDDatumPaint(
          chart, series, pointOverride, categoryIndex, seriesIndex,
        );
        const color = paint.fill === null ? 'transparent' : paint.color;
        const outline = paint.lineFill !== null
          && (paint.lineColor != null || paint.lineFill !== undefined);
        const base = stacked ? (value >= 0 ? positiveBase[categoryIndex] : negativeBase[categoryIndex]) : 0;
        const end = threeDSafeAdd(base, value);
        if (stacked) {
          if (value >= 0) positiveBase[categoryIndex] = end;
          else negativeBase[categoryIndex] = end;
        }
        const visibleBase = visibleAxisValue(base);
        const visibleEnd = visibleAxisValue(end);
        const shape = series.threeDShape ?? chart.threeD.shape ?? 'box';
        const tapered = shape === 'cone' || shape === 'pyramid';
        const toMax = shape === 'coneToMax' || shape === 'pyramidToMax';
        const ordinaryTaperScale = (coordinate: number): number => {
          if (!tapered) return 1;
          // A taper is linear in the same model coordinate used to paint the
          // value axis. Raw-value interpolation is only equivalent on a linear
          // axis; on a logarithmic axis it turns a clipped cone into the wrong
          // frustum. Zero is outside a log domain, so the authored datum begins
          // with a full ring at the first visible positive axis value.
          const modelBaseValue = logarithmic && !(base > 0) ? axis.min : base;
          const modelBase = axis.fraction(modelBaseValue);
          const modelEnd = axis.fraction(end);
          const modelCoordinate = axis.fraction(coordinate);
          const span = modelEnd - modelBase;
          if (span === 0 || !Number.isFinite(span)
            || !Number.isFinite(modelCoordinate - modelBase)) {
            return coordinate === modelBaseValue ? 1 : 0;
          }
          return Math.max(0, Math.min(1, 1 - (modelCoordinate - modelBase) / span));
        };
        const toMaxScale = (coordinate: number): number => {
          if (!toMax) return 1;
          const bound = end >= base ? axis.max : axis.min;
          const baseline = logarithmic ? axis.min : 0;
          const modelBound = axis.fraction(bound);
          const modelBaseline = axis.fraction(baseline);
          const modelCoordinate = axis.fraction(coordinate);
          const span = Math.abs(modelBound - modelBaseline);
          if (!(span > 0) || ![modelBound, modelBaseline, modelCoordinate].every(Number.isFinite)) {
            return threeDToMaxScale(coordinate, axis.min, axis.max);
          }
          return Math.max(0, Math.min(1, Math.abs(modelBound - modelCoordinate) / span));
        };
        // An authored axis clips the original solid. Preserve its parametric
        // cross-section at each clip plane instead of regenerating a full cone
        // or pyramid inside the visible interval.
        const baseScale = toMax
          ? toMaxScale(visibleBase)
          : ordinaryTaperScale(visibleBase);
        const endScale = toMax
          ? toMaxScale(visibleEnd)
          : ordinaryTaperScale(visibleEnd);
        if (horizontal) {
          const x0 = front.x + axis.fraction(visibleBase) * front.w;
          const x1 = front.x + axis.fraction(visibleEnd) * front.w;
          const categorySlot = categoryReversed ? categoryCount - 1 - categoryIndex : categoryIndex;
          const top = front.y + categorySlot * interval + clusterSlot.offset;
          primitives.push({
            x: Math.min(x0, x1), y: top,
            width: Math.abs(x1 - x0), height: clusterSlot.size,
            nearDepth: depthInterval.near, farDepth: depthInterval.far,
            categoryIndex, seriesIndex, color, fillPaint: paint.fill,
            lineFill: paint.lineFill,
            shape,
            baseCoord: x0, endCoord: x1,
            baseScale, endScale, omitBaseCap: false, omitEndCap: false,
            outline,
            outlineColor: paint.lineColor ?? 'rgba(0,0,0,0.42)',
            outlineWidth: paint.lineWidthEmu != null
              ? threeDMeshOutlineWidthPx(paint.lineWidthEmu, ptToPx)
              : 0.75 * ptToPx / PT_TO_PX,
            outlineDash: paint.lineDash ?? 'solid',
            outlineCustomDash: paint.lineCustomDash,
            outlineCap: paint.lineCap,
            outlineJoin: paint.lineJoin,
            labelValue: percent ? value / 100 : value,
            plottedLabelValue: end,
          });
        } else {
          const y0 = front.y + front.h - axis.fraction(visibleBase) * front.h;
          const y1 = front.y + front.h - axis.fraction(visibleEnd) * front.h;
          const categorySlot = categoryReversed ? categoryCount - 1 - categoryIndex : categoryIndex;
          const left = front.x + categorySlot * interval + clusterSlot.offset;
          primitives.push({
            x: left, y: Math.min(y0, y1),
            width: clusterSlot.size, height: Math.abs(y1 - y0),
            nearDepth: depthInterval.near, farDepth: depthInterval.far,
            categoryIndex, seriesIndex, color, fillPaint: paint.fill,
            lineFill: paint.lineFill,
            shape,
            baseCoord: y0, endCoord: y1,
            baseScale, endScale, omitBaseCap: false, omitEndCap: false,
            outline,
            outlineColor: paint.lineColor ?? 'rgba(0,0,0,0.42)',
            outlineWidth: paint.lineWidthEmu != null
              ? threeDMeshOutlineWidthPx(paint.lineWidthEmu, ptToPx)
              : 0.75 * ptToPx / PT_TO_PX,
            outlineDash: paint.lineDash ?? 'solid',
            outlineCustomDash: paint.lineCustomDash,
            outlineCap: paint.lineCap,
            outlineJoin: paint.lineJoin,
            labelValue: percent ? value / 100 : value,
            plottedLabelValue: end,
          });
        }
      }
    }
    if (stacked) {
      // Remove only genuinely shared internal cross-sections. This converts a
      // compatible stack into one continuous mesh while preserving caps when
      // shapes/scales differ or an authored axis clips away the neighbour.
      const primitivesByCategory = bucketThreeDStackItems(primitives, categoryCount);
      for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
        const categoryPrimitives = primitivesByCategory[categoryIndex];
        for (const sign of [-1, 1] as const) {
          const segments = categoryPrimitives
            .filter(item => Math.sign(item.labelValue) === sign
              && !isTransparentPaint(item.color)
              && Math.abs(item.endCoord - item.baseCoord) > 1e-9)
            .sort((a, b) => a.seriesIndex - b.seriesIndex);
          for (let index = 0; index + 1 < segments.length; index++) {
            const lower = segments[index];
            const upper = segments[index + 1];
            const tolerance = 1e-8 * Math.max(
              1, Math.abs(lower.endCoord), Math.abs(upper.baseCoord),
            );
            if (lower.shape !== upper.shape
              || Math.abs(lower.endCoord - upper.baseCoord) > tolerance
              || Math.abs(lower.endScale - upper.baseScale) > 1e-9
              || lower.nearDepth !== upper.nearDepth
              || lower.farDepth !== upper.farDepth) continue;
            lower.omitEndCap = true;
            upper.omitBaseCap = true;
          }
        }
        // The first positive and negative slabs meet at the zero plane. When
        // both are opaque and have the same cross-section this plane is inside
        // one continuous signed stack, so neither solid may paint a cap there.
        // A noFill neighbour is deliberately excluded: it cannot occlude the
        // opaque slab's exposed cap.
        const positive = categoryPrimitives.find(item =>
          item.labelValue > 0 && !isTransparentPaint(item.color));
        const negative = categoryPrimitives.find(item =>
          item.labelValue < 0 && !isTransparentPaint(item.color));
        if (positive && negative) {
          const tolerance = 1e-8 * Math.max(
            1, Math.abs(positive.baseCoord), Math.abs(negative.baseCoord),
          );
          if (positive.shape === negative.shape
            && Math.abs(positive.baseCoord - negative.baseCoord) <= tolerance
            && Math.abs(positive.baseScale - negative.baseScale) <= 1e-9
            && positive.nearDepth === negative.nearDepth
            && positive.farDepth === negative.farDepth) {
            positive.omitBaseCap = true;
            negative.omitBaseCap = true;
          }
        }
      }
    }
    walls(
      ctx, chart, projection, axis, horizontal ? 'horizontal' : 'vertical',
      categoryCount, categoryBetween, categoryReversed, ptToPx,
    );
    // Build one scene list and sort every visible face by camera-space depth.
    // A logical series index is not a view depth after rotation, and painting
    // a complete cuboid at a time can still put one of its far faces over a
    // nearer neighbour. The compact face sort is the shared painter contract.
    const meshBudget: ScenePrimitiveBudget = {
      remaining: MAX_PROJECTED_STROKE_PRIMITIVES,
      exceeded: false,
    };
    const datumFaces = primitives.map(item => {
      const faces = shapeMeshFaces(
        projection,
        item.shape,
        horizontal,
        item.x, item.y, item.width, item.height,
        item.baseCoord, item.endCoord,
        item.nearDepth, item.farDepth, item.color,
        item.baseScale,
        item.endScale,
        item.omitBaseCap,
        item.omitEndCap,
        item.outline && item.outlineColor ? {
          color: item.outlineColor,
          width: item.outlineWidth,
          dash: drawingmlLineDashArray(
            item.outlineCustomDash, item.outlineDash, item.outlineWidth,
          ),
          cap: item.outlineCap,
          join: item.outlineJoin,
        } : undefined,
        meshBudget,
      );
      applyScenePaint(ctx, faces, 'fill', item.fillPaint, shapeRotationDeg);
      applyScenePaint(ctx, faces, 'outline', item.lineFill, shapeRotationDeg);
      return { item, faces };
    });
    const sceneFaces = datumFaces.flatMap(datum => datum.faces);
    if (meshBudget.exceeded) {
      paintThreeDTooManyDataPoints(ctx, rect);
      return true;
    }
    for (const item of sortProjectedSceneFaces(sceneFaces)) paintSceneFace(ctx, item);
    // 3-D faces from different datums interleave after camera-depth sorting, so
    // applying a normal per-datum effect wrapper would repaint nearer faces in
    // the wrong order. Emit only the native outer effect behind the completed
    // scene; raster effects stay resolved but intentionally suppressed by the
    // helper's documented 3-D boundary.
    for (const { item, faces } of datumFaces) {
      const series = chart.series[item.seriesIndex];
      const point = pointOverrides[item.seriesIndex].get(item.categoryIndex);
      const styleIndex = threeDDatumStyleIndex(
        chart, series, item.categoryIndex, item.seriesIndex,
      );
      const pointEffect = chartStyleEffectOwner(point?.chartexStyle);
      const seriesEffect = chartStyleEffectOwner(series.chartexStyle);
      paintChartStyleOuterEffectBehind(
        ctx,
        pointEffect ?? seriesEffect,
        chartDataPointStyleRole(chart, 'dataPoint3D', item.seriesIndex),
        pointEffect ? item.categoryIndex : styleIndex,
        ptToPx,
        target => paintSceneEffectSilhouette(target, faces),
        styleIndex,
      );
    }
    for (const item of primitives) {
      const series = chart.series[item.seriesIndex];
      const anchor = horizontal
        ? projection.project(
          item.endCoord,
          item.y + item.height / 2,
          (item.nearDepth + item.farDepth) / 2,
        )
        : projection.project(
          item.x + item.width / 2,
          item.endCoord,
          (item.nearDepth + item.farDepth) / 2,
        );
      const labelOverride = labelOverrides[item.seriesIndex].get(item.categoryIndex);
      if (wantsThreeDDataLabel(chart, series, labelOverride)) {
        deferredDataLabels.push(() => drawThreeDDataLabel(
          ctx, chart, series, item.seriesIndex, item.categoryIndex,
          item.labelValue, anchor, rect, ptToPx, 0, undefined, labelOverride,
          't', anchor, true, chart.valAxisDisplayUnits, axis.max, item.plottedLabelValue,
          shapeRotationDeg,
        ));
      }
    }
  } else {
    // Stacked line/area geometry is derived once in source-series order. The
    // lower and upper boundaries then live in the same x/y/z scene as every
    // other primitive; painting a raw series from zero would not be a stacked
    // surface even if its projection were geometrically coherent.
    const stackedLower = chart.series.map(() => new Array<number>(categoryCount).fill(0));
    const stackedUpper = chart.series.map(() => new Array<number>(categoryCount).fill(0));
    if (stacked) {
      const positive = new Array<number>(categoryCount).fill(0);
      const negative = new Array<number>(categoryCount).fill(0);
      for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
        for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
          const value = valueAt(seriesIndex, categoryIndex);
          const base = value >= 0 ? positive[categoryIndex] : negative[categoryIndex];
          stackedLower[seriesIndex][categoryIndex] = base;
          stackedUpper[seriesIndex][categoryIndex] = threeDSafeAdd(base, value);
          if (value >= 0) positive[categoryIndex] = threeDSafeAdd(positive[categoryIndex], value);
          else negative[categoryIndex] = threeDSafeAdd(negative[categoryIndex], value);
        }
      }
    }
    const seriesAverageDepth = (seriesIndex: number): number => {
      const z = projection.seriesDepth(seriesIndex, seriesCount, stacked);
      let sum = 0;
      let count = 0;
      for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
        if (!hasFiniteValue(seriesIndex, categoryIndex)) continue;
        const upperValue = stacked
          ? stackedUpper[seriesIndex][categoryIndex]
          : valueAt(seriesIndex, categoryIndex);
        const x = front.x + categoryPositionFraction(
          categoryIndex, categoryCount, categoryBetween, categoryReversed,
        ) * front.w;
        const y = front.y + front.h - axis.fraction(visibleAxisValue(upperValue)) * front.h;
        sum += projection.cameraDepth(x, y, z);
        count++;
      }
      return count > 0 ? sum / count : Number.NEGATIVE_INFINITY;
    };
    const seriesOrder = chart.series.map((_, seriesIndex) => seriesIndex).sort((a, b) =>
      seriesAverageDepth(a) - seriesAverageDepth(b) || b - a
    );
    const seriesPaints = chart.series.map((series, seriesIndex) =>
      threeDDatumPaint(chart, series, undefined, seriesIndex, seriesIndex)
    );
    const sceneCommands: Array<ProjectedSceneFace & {
      layer: 0 | 1;
      paint: () => void;
    }> = [];
    const foregroundMarkers: Array<() => void> = [];
    let strokeBudgetExceeded = false;
    for (const seriesIndex of seriesOrder) {
      if (strokeBudgetExceeded) break;
      const series = chart.series[seriesIndex];
      const seriesPaint = seriesPaints[seriesIndex];
      const color = seriesPaint.fill === null ? 'transparent' : seriesPaint.color;
      const depthInterval = stacked
        ? projection.prismInterval(0, 1, true)
        : projection.prismInterval(seriesIndex, seriesCount, false);
      const depthProbeX = front.x + front.w / 2;
      const depthProbeY = front.y + front.h / 2;
      // Lines and annotations sit on the camera-facing ridge of the solid
      // area slab.  Ordinary line charts use the same plane without painting
      // the slab faces.
      const z = projection.cameraDepth(depthProbeX, depthProbeY, depthInterval.near)
        >= projection.cameraDepth(depthProbeX, depthProbeY, depthInterval.far)
        ? depthInterval.near : depthInterval.far;
      const points: Array<Point | null> = [];
      const lowerPoints: Array<Point | null> = [];
      const pointDepths: Array<number | null> = [];
      const lowerDepths: Array<number | null> = [];
      const sourceUpperValues: Array<number | null> = [];
      const sceneXs: Array<number | null> = [];
      const upperYs: Array<number | null> = [];
      const lowerYs: Array<number | null> = [];
      const upperFractions: Array<number | null> = [];
      const lowerFractions: Array<number | null> = [];
      for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
        if (!hasFiniteValue(seriesIndex, categoryIndex)) {
          points.push(null);
          lowerPoints.push(null);
          pointDepths.push(null);
          lowerDepths.push(null);
          sourceUpperValues.push(null);
          sceneXs.push(null);
          upperYs.push(null);
          lowerYs.push(null);
          upperFractions.push(null);
          lowerFractions.push(null);
          continue;
        }
        const upperValue = stacked
          ? stackedUpper[seriesIndex][categoryIndex]
          : valueAt(seriesIndex, categoryIndex);
        const lowerValue = stacked
          ? stackedLower[seriesIndex][categoryIndex]
          : 0;
        const sceneX = front.x + categoryPositionFraction(
          categoryIndex, categoryCount, categoryBetween, categoryReversed,
        ) * front.w;
        const upperFraction = axis.fraction(upperValue);
        // Zero is not part of a logarithmic domain.  An ordinary area datum
        // still has a visible base at the lower plot edge, so represent that
        // clipped model boundary as fraction zero rather than asking log(0).
        const rawLowerFraction = axis.fraction(lowerValue);
        const lowerFraction = Number.isFinite(rawLowerFraction)
          ? rawLowerFraction
          : lowerValue <= axis.min ? 0 : 1;
        const upperInside = Number.isFinite(upperFraction)
          && upperFraction >= 0 && upperFraction <= 1;
        const visibleUpperFraction = Number.isFinite(upperFraction)
          ? Math.max(0, Math.min(1, upperFraction))
          : upperValue <= axis.min ? 0 : 1;
        const upperY = front.y + front.h
          - visibleUpperFraction * front.h;
        const lowerY = front.y + front.h
          - Math.max(0, Math.min(1, lowerFraction)) * front.h;
        points.push(
          chart.chartType.toLowerCase().includes('area') || upperInside
            ? projection.project(sceneX, upperY, z)
            : null,
        );
        lowerPoints.push(projection.project(sceneX, lowerY, z));
        pointDepths.push(upperInside
          ? projection.cameraDepth(sceneX, upperY, z)
          : null);
        lowerDepths.push(projection.cameraDepth(sceneX, lowerY, z));
        sourceUpperValues.push(upperValue);
        sceneXs.push(sceneX);
        upperYs.push(upperY);
        lowerYs.push(lowerY);
        upperFractions.push(upperFraction);
        lowerFractions.push(lowerFraction);
      }
      const runs: Array<{
        upper: Point[];
        lower: Point[];
        upperDepths: number[];
        lowerDepths: number[];
        indices: number[];
        sceneXs: number[];
        upperYs: number[];
        lowerYs: number[];
        upperFractions: number[];
        lowerFractions: number[];
      }> = [];
      let run: (typeof runs)[number] | null = null;
      for (let categoryIndex = 0; categoryIndex < points.length; categoryIndex++) {
        const point = points[categoryIndex];
        const lower = lowerPoints[categoryIndex];
        if (!point || !lower) {
          if (dispBlanks === 'gap') {
            if (run) runs.push(run);
            run = null;
          }
          continue;
        }
        run ??= {
          upper: [], lower: [], upperDepths: [], lowerDepths: [], indices: [],
          sceneXs: [], upperYs: [], lowerYs: [], upperFractions: [], lowerFractions: [],
        };
        run.upper.push(point);
        run.lower.push(lower);
        run.upperDepths.push(pointDepths[categoryIndex] ?? 0);
        run.lowerDepths.push(lowerDepths[categoryIndex] ?? 0);
        run.indices.push(categoryIndex);
        run.sceneXs.push(sceneXs[categoryIndex] ?? 0);
        run.upperYs.push(upperYs[categoryIndex] ?? 0);
        run.lowerYs.push(lowerYs[categoryIndex] ?? 0);
        run.upperFractions.push(upperFractions[categoryIndex] ?? 0);
        run.lowerFractions.push(lowerFractions[categoryIndex] ?? 0);
      }
      if (run) runs.push(run);
      const areaStrokeRuns: ProjectedStrokePoint[][] = [];
      const areaPaintFaces: SceneFace[] = [];
      if (chart.chartType.toLowerCase().includes('area')) {
        for (const item of runs) {
          let strokeRun: ProjectedStrokePoint[] | null = null;
          for (let index = 0; index + 1 < item.upper.length; index++) {
            const pieces = clipAxisFractionBand(
              item.lowerFractions[index], item.lowerFractions[index + 1],
              item.upperFractions[index], item.upperFractions[index + 1],
            );
            for (let pieceIndex = 0; pieceIndex < pieces.length; pieceIndex++) {
              const piece = pieces[pieceIndex];
              const sceneX0 = item.sceneXs[index]
                + (item.sceneXs[index + 1] - item.sceneXs[index]) * piece.startT;
              const sceneX1 = item.sceneXs[index]
                + (item.sceneXs[index + 1] - item.sceneXs[index]) * piece.endT;
              const lowerY0 = front.y + front.h - piece.lowerStart * front.h;
              const lowerY1 = front.y + front.h - piece.lowerEnd * front.h;
              const upperY0 = front.y + front.h - piece.upperStart * front.h;
              const upperY1 = front.y + front.h - piece.upperEnd * front.h;
              const ridgeStart = {
                ...projection.project(sceneX0, upperY0, z),
                cameraDepth: projection.cameraDepth(sceneX0, upperY0, z),
                cameraWeight: projection.cameraProjectionWeight(sceneX0, upperY0, z),
              };
              const ridgeEnd = {
                ...projection.project(sceneX1, upperY1, z),
                cameraDepth: projection.cameraDepth(sceneX1, upperY1, z),
                cameraWeight: projection.cameraProjectionWeight(sceneX1, upperY1, z),
              };
              const continues = strokeRun != null && Math.hypot(
                strokeRun.at(-1)!.x - ridgeStart.x,
                strokeRun.at(-1)!.y - ridgeStart.y,
              ) <= 1e-8;
              if (!continues) {
                if (strokeRun && strokeRun.length >= 2) areaStrokeRuns.push(strokeRun);
                strokeRun = [ridgeStart, ridgeEnd];
              } else {
                strokeRun!.push(ridgeEnd);
              }
              const faces = areaStripFaces(
                projection,
                sceneX0, sceneX1,
                lowerY0, lowerY1,
                upperY0, upperY1,
                depthInterval.near, depthInterval.far, color,
                index === 0 && pieceIndex === 0 && piece.startT === 0,
                index + 2 === item.upper.length
                  && pieceIndex + 1 === pieces.length && piece.endT === 1,
              );
              for (const areaFace of faces) {
                if (sceneCommands.length >= MAX_PROJECTED_STROKE_PRIMITIVES) {
                  strokeBudgetExceeded = true;
                  break;
                }
                sceneCommands.push({
                  points: areaFace.points,
                  cameraDepth: areaFace.cameraDepth,
                  cameraDepths: areaFace.cameraDepths,
                  cameraWeights: areaFace.cameraWeights,
                  layer: 0,
                  paint: () => paintSceneFace(ctx, areaFace),
                });
                areaPaintFaces.push(areaFace);
              }
              if (strokeBudgetExceeded) break;
            }
            if (strokeBudgetExceeded) break;
          }
          if (strokeRun && strokeRun.length >= 2) areaStrokeRuns.push(strokeRun);
          if (strokeBudgetExceeded) break;
        }
        applyScenePaint(ctx, areaPaintFaces, 'fill', seriesPaint.fill, shapeRotationDeg);
      }
      const lineStrokeRuns: Array<{
        path: ProjectedStrokePoint[];
        modelPath: Point[];
        ownerIndex: number;
        startClipped: boolean;
        endClipped: boolean;
        dashOffset: number;
      }> = [];
      if (!chart.chartType.toLowerCase().includes('area')) {
        type ModelLinePoint = { x: number; fraction: number; ownerIndex: number };
        const modelRuns: ModelLinePoint[][] = [];
        let modelRun: ModelLinePoint[] | null = null;
        for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
          const value = sourceUpperValues[categoryIndex];
          const fraction = value == null ? Number.NaN : axis.fraction(value);
          if (value == null || !Number.isFinite(fraction)) {
            if (dispBlanks === 'gap' && modelRun) {
              modelRuns.push(modelRun);
              modelRun = null;
            }
            continue;
          }
          modelRun ??= [];
          modelRun.push({
            x: front.x + categoryPositionFraction(
              categoryIndex, categoryCount, categoryBetween, categoryReversed,
            ) * front.w,
            fraction,
            ownerIndex: categoryIndex,
          });
        }
        if (modelRun) modelRuns.push(modelRun);
        const projectModelPoint = (point: ModelLinePoint): ProjectedStrokePoint => {
          const visibleFraction = Math.max(0, Math.min(1, point.fraction));
          const sceneY = front.y + front.h - visibleFraction * front.h;
          return {
            ...projection.project(point.x, sceneY, z),
            cameraDepth: projection.cameraDepth(point.x, sceneY, z),
            cameraWeight: projection.cameraProjectionWeight(point.x, sceneY, z),
          };
        };
        const finishVisibleRun = (current: (typeof lineStrokeRuns)[number] | null) => {
          if (current && current.path.length >= 2) lineStrokeRuns.push(current);
        };
        for (const source of modelRuns) {
          if (source.length < 2) continue;
          const sampled: ModelLinePoint[] = [source[0]];
          for (let index = 0; index + 1 < source.length; index++) {
            const p0 = source[index - 1] ?? source[index];
            const p1 = source[index];
            const p2 = source[index + 1];
            const p3 = source[index + 2] ?? p2;
            if (series.smooth !== true || source.length < 3) {
              sampled.push(p2);
              continue;
            }
            const c1 = {
              x: p1.x + (p2.x - p0.x) / 6,
              fraction: p1.fraction + (p2.fraction - p0.fraction) / 6,
            };
            const c2 = {
              x: p2.x - (p3.x - p1.x) / 6,
              fraction: p2.fraction - (p3.fraction - p1.fraction) / 6,
            };
            for (let step = 1; step <= 12; step++) {
              const t = step / 12;
              const u = 1 - t;
              sampled.push({
                x: u * u * u * p1.x + 3 * u * u * t * c1.x
                  + 3 * u * t * t * c2.x + t * t * t * p2.x,
                fraction: u * u * u * p1.fraction
                  + 3 * u * u * t * c1.fraction
                  + 3 * u * t * t * c2.fraction + t * t * t * p2.fraction,
                ownerIndex: p2.ownerIndex,
              });
            }
          }
          let visible: (typeof lineStrokeRuns)[number] | null = null;
          let traversedLength = 0;
          for (let index = 0; index + 1 < sampled.length; index++) {
            const startModel = sampled[index];
            const endModel = sampled[index + 1];
            const projectUnclipped = (point: ModelLinePoint): Point => projection.project(
              point.x,
              front.y + front.h - point.fraction * front.h,
              z,
            );
            const projectedStart = projectUnclipped(startModel);
            const projectedEnd = projectUnclipped(endModel);
            const projectedLength = Math.hypot(
              projectedEnd.x - projectedStart.x,
              projectedEnd.y - projectedStart.y,
            );
            const modelLength = Number.isFinite(projectedLength)
              ? projectedLength
              : Math.hypot(
                endModel.x - startModel.x,
                (endModel.fraction - startModel.fraction) * front.h,
              );
            const clipped = clipAxisFractionSegment(
              startModel.fraction, endModel.fraction,
            );
            if (!clipped || clipped.endT - clipped.startT <= 1e-12) {
              finishVisibleRun(visible);
              visible = null;
              traversedLength += modelLength;
              continue;
            }
            const at = (t: number): ModelLinePoint => ({
              x: startModel.x + (endModel.x - startModel.x) * t,
              fraction: startModel.fraction
                + (endModel.fraction - startModel.fraction) * t,
              ownerIndex: endModel.ownerIndex,
            });
            const projectedClipStart = projectUnclipped(at(clipped.startT));
            const clippedPrefixLength = Math.hypot(
              projectedClipStart.x - projectedStart.x,
              projectedClipStart.y - projectedStart.y,
            );
            const start = projectModelPoint(at(clipped.startT));
            const end = projectModelPoint(at(clipped.endT));
            const modelStart = at(clipped.startT);
            const modelEnd = at(clipped.endT);
            const toModelPoint = (point: ModelLinePoint): Point => ({
              x: point.x,
              y: front.y + front.h
                - Math.max(0, Math.min(1, point.fraction)) * front.h,
            });
            const continues = visible != null && Math.hypot(
              visible.path.at(-1)!.x - start.x,
              visible.path.at(-1)!.y - start.y,
            ) <= 1e-8 && visible.ownerIndex === endModel.ownerIndex;
            if (!continues) {
              finishVisibleRun(visible);
              visible = {
                path: [start, end],
                modelPath: [toModelPoint(modelStart), toModelPoint(modelEnd)],
                ownerIndex: endModel.ownerIndex,
                startClipped: index > 0 || clipped.startT > 0,
                endClipped: index + 1 < sampled.length - 1 || clipped.endT < 1,
                dashOffset: traversedLength + (Number.isFinite(clippedPrefixLength)
                  ? clippedPrefixLength : modelLength * clipped.startT),
              };
            } else {
              visible!.path.push(end);
              visible!.modelPath.push(toModelPoint(modelEnd));
              visible!.endClipped = index + 1 < sampled.length - 1 || clipped.endT < 1;
            }
            traversedLength += modelLength;
          }
          finishVisibleRun(visible);
        }
      }
      const areaFamily = chart.chartType.toLowerCase().includes('area');
      const authoredAreaLine = series.lineHidden != null
        || series.lineColor != null
        || series.lineWidthEmu != null
        || series.chartexStyle?.lineHidden != null
        || series.chartexStyle?.lineColors?.some(Boolean)
        || series.chartexStyle?.lineWidthEmu != null
        || series.chartexStyle?.lineDash != null
        || series.chartexStyle?.lineCap != null
        || series.chartexStyle?.lineJoin != null
        || seriesPaint.lineColor != null
        || seriesPaint.lineFill !== undefined;
      if (areaFamily ? seriesPaint.lineFill !== null && authoredAreaLine : true) {
        const pointLinePaints = new Map<number, ThreeDDatumPaint>();
        const paintForLineRun = (ownerIndex: number): ThreeDDatumPaint => {
          const point = pointOverrides[seriesIndex].get(ownerIndex);
          if (!threeDPointOwnsLineSegment(point)) return seriesPaint;
          let paint = pointLinePaints.get(ownerIndex);
          if (!paint) {
            paint = threeDDatumPaint(chart, series, point, ownerIndex, seriesIndex);
            pointLinePaints.set(ownerIndex, paint);
          }
          return paint;
        };
        const strokeFaceGroups = new Map<ThreeDDatumPaint, SceneFace[]>();
        const lineRibbonBudget: ScenePrimitiveBudget = {
          remaining: MAX_PROJECTED_STROKE_PRIMITIVES,
          exceeded: false,
        };
        const pushStrokeGeometry = (
          path: readonly ProjectedStrokePoint[],
          paint: ThreeDDatumPaint,
          startCap: CanvasLineCap,
          endCap: CanvasLineCap = startCap,
          dashOffset = 0,
          modelPath?: readonly Point[],
        ) => {
          if (paint.lineFill === null) return;
          // Automatic 3-D line/area edges use the darker Chart Style line
          // role. Direct point line paint owns only its incoming segment.
          const automaticStrokeBase = paint === seriesPaint ? color : paint.color;
          const strokeStyle = paint.lineColor ?? scaleHexColor(automaticStrokeBase, 0.70);
          const lineWidth = paint.lineWidthEmu
            ? Math.max(0.5, paint.lineWidthEmu / EMU_PER_PT) * ptToPx
            : areaFamily
              ? 0.75 * ptToPx
              : Math.max(1, 2 * ptToPx);
          const lineDash = drawingmlLineDashArray(
            paint.lineCustomDash, paint.lineDash, lineWidth,
          );
          const strokeOptions = {
            width: lineWidth,
            dash: lineDash,
            dashOffset,
            lineCap: paint.lineCap,
            startCap,
            endCap,
            lineJoin: paint.lineJoin,
          };
          // ECMA Line3D is a depth-bearing ribbon. Tessellate the same bounded
          // dash/cap/join polygon in the chart's model X/Y plane, extrude it
          // through the series interval, then hand every face to the existing
          // camera/depth/material pipeline. Area ridges remain ordinary edge
          // strokes because their slab already supplies the depth faces.
          if (!areaFamily && modelPath && modelPath.length >= 2) {
            const modelStrokes = buildProjectedStrokePrimitives(
              modelPath.map(point => ({ ...point, cameraDepth: 0, cameraWeight: 1 })),
              strokeOptions,
            );
            if (modelStrokes == null) {
              strokeBudgetExceeded = true;
              return;
            }
            const group = strokeFaceGroups.get(paint) ?? [];
            for (const stroke of modelStrokes) {
              const mesh = buildThreeDLineRibbonMesh({
                outline: stroke.points,
                nearDepth: depthInterval.near,
                farDepth: depthInterval.far,
              });
              if (!mesh) continue;
              const faces = projectThreeDMesh(
                projection, mesh, strokeStyle, undefined, lineRibbonBudget,
              ).map(face => ({ ...face, paintRole: 'outline' as const }));
              if (lineRibbonBudget.exceeded) {
                strokeBudgetExceeded = true;
                return;
              }
              group.push(...faces);
              for (const face of faces) {
                sceneCommands.push({
                  points: face.points,
                  cameraDepth: face.cameraDepth,
                  cameraDepths: face.cameraDepths,
                  cameraWeights: face.cameraWeights,
                  layer: 1,
                  paint: () => paintSceneFace(ctx, face),
                });
              }
            }
            strokeFaceGroups.set(paint, group);
            return;
          }
          const strokes = buildProjectedStrokePrimitives(path, strokeOptions);
          if (strokes == null) {
            strokeBudgetExceeded = true;
            return;
          }
          if (sceneCommands.length + strokes.length > MAX_PROJECTED_STROKE_PRIMITIVES) {
            strokeBudgetExceeded = true;
            return;
          }
          for (const stroke of strokes) {
            const strokeFace: SceneFace = {
              points: stroke.points,
              color: strokeStyle,
              paintRole: 'outline',
              shade: 0,
              cameraDepth: stroke.cameraDepth,
              cameraDepths: stroke.cameraDepths,
              cameraWeights: stroke.cameraWeights,
              outline: false,
            };
            const group = strokeFaceGroups.get(paint) ?? [];
            group.push(strokeFace);
            strokeFaceGroups.set(paint, group);
            sceneCommands.push({
              points: stroke.points,
              cameraDepth: stroke.cameraDepth,
              cameraDepths: stroke.cameraDepths,
              cameraWeights: stroke.cameraWeights,
              layer: 1,
              paint: () => paintSceneFace(ctx, strokeFace),
            });
          }
        };
        if (areaFamily) {
          for (const path of areaStrokeRuns) {
            pushStrokeGeometry(path, seriesPaint, seriesPaint.lineCap);
          }
        } else {
          for (const item of lineStrokeRuns) {
            const paint = paintForLineRun(item.ownerIndex);
            pushStrokeGeometry(
              item.path,
              paint,
              item.startClipped ? 'butt' : paint.lineCap,
              item.endClipped ? 'butt' : paint.lineCap,
              item.dashOffset,
              item.modelPath,
            );
          }
        }
        for (const [paint, faces] of strokeFaceGroups) {
          applyScenePaint(ctx, faces, 'outline', paint.lineFill, shapeRotationDeg);
        }
      }
      const seriesMarkersVisible = (areaFamily
        ? series.showMarker === true || seriesHasMarkerDetail(series)
        : series.showMarker === true) && series.markerSymbol !== 'none';
      if ((chart.chartType.toLowerCase().includes('line') || areaFamily)
        && (seriesMarkersVisible || hasVisiblePointMarkerOverride(series))) {
        for (let categoryIndex = 0; categoryIndex < points.length; categoryIndex++) {
          const point = points[categoryIndex];
          if (!point) continue;
          const override = pointOverrides[seriesIndex].get(categoryIndex);
          const symbol = effectiveMarkerSymbol(
            series, override, 'circle', seriesMarkersVisible,
          );
          if (symbol === 'none') continue;
          const sizePt = override?.markerSize ?? series.markerSize ?? 5;
          const marker = threeDMarkerPaint(
            chart, series, override, categoryIndex, seriesIndex,
            series.color ?? PALETTE[seriesIndex % PALETTE.length], ptToPx,
          );
          foregroundMarkers.push(() => paintThreeDMarker(
            ctx, point, symbol, Math.max(2, sizePt) * ptToPx,
            marker.fill === '00000000' ? 'transparent' : `#${marker.fill.replace(/^#/, '')}`,
            marker.line === '00000000' ? 'rgba(0,0,0,0)'
              : `#${marker.line.replace(/^#/, '')}`,
            marker.lineWidth,
            marker.fillPaint, shapeRotationDeg, ptToPx,
            marker.effectDirect,
            marker.effectFallback,
            marker.effectIndex,
            marker.effectFallbackIndex,
            marker.linePaint,
            marker.lineDash,
            marker.lineCustomDash,
            marker.lineCap,
            marker.lineJoin,
          ));
        }
      }
      for (let categoryIndex = 0; categoryIndex < points.length; categoryIndex++) {
        const point = points[categoryIndex];
        if (!point) continue;
        const sourceValue = valueAt(seriesIndex, categoryIndex);
        const override = pointOverrides[seriesIndex].get(categoryIndex);
        const markerSize = override?.markerSize ?? series.markerSize ?? 5;
        const labelOverride = labelOverrides[seriesIndex].get(categoryIndex);
        if (wantsThreeDDataLabel(chart, series, labelOverride)) {
          deferredDataLabels.push(() => drawThreeDDataLabel(
            ctx, chart, series, seriesIndex, categoryIndex,
            percent ? sourceValue / 100 : sourceValue,
            point, rect, ptToPx,
            series.showMarker === true || override?.markerSymbol != null
              ? markerSize * ptToPx / 2
              : 0,
            undefined,
            labelOverride,
            't', point, true, chart.valAxisDisplayUnits, axis.max,
            stacked ? stackedUpper[seriesIndex][categoryIndex] : sourceValue,
            shapeRotationDeg,
          ));
          }
        }
      }
    // Area faces and explicit stroke polygons participate in one local-overlap
    // depth graph. Average series depth is insufficient when two projected
    // surfaces cross at only one side of the chart.
    if (strokeBudgetExceeded) {
      // The source/mesh preflight rejects ordinary oversized charts before
      // any paint. This late guard handles adversarial sub-pixel dash
      // expansion. Do not cover a noFill chart with an invented white panel;
      // preserve the host background and report the bounded rejection.
      paintThreeDTooManyDataPoints(ctx, rect);
      return true;
    }
    for (const command of sortProjectedSceneFaces(sceneCommands)) command.paint();
    // Markers are annotations on the completed series geometry. Keeping them
    // out of the face/segment depth list prevents a later adjacent segment
    // from cutting through its own endpoint marker.
    for (const paintMarker of foregroundMarkers) paintMarker();
  }
  frontAxes(
    ctx, chart, projection, axis, categoryCount, categoryBetween, categoryReversed,
    ptToPx, horizontal ? 'horizontal' : 'vertical',
  );
  cartesianAxisTicks(
    ctx, chart, projection, axis, categoryCount, categoryBetween, categoryReversed,
    horizontal ? 'horizontal' : 'vertical', bars, ptToPx,
  );
  drawThreeDSeriesAxis(
    ctx,
    chart,
    rect,
    projection,
    axis,
    categoryCount,
    categoryBetween,
    categoryReversed,
    horizontal ? 'horizontal' : 'vertical',
    bars,
    ptToPx,
  );
  const resolvedAxisGeometry = threeDAxisGeometry(
    chart,
    projection,
    axis,
    categoryCount,
    categoryBetween,
    categoryReversed,
    horizontal ? 'horizontal' : 'vertical',
  );
  const edgeGeometry = (
    kind: 'value' | 'category',
    position: string | null | undefined,
  ): ThreeDAxisGeometry => {
    if (position !== 'low' && position !== 'high') return resolvedAxisGeometry;
    const depth = projection.topology.nearDepth;
    let axisX = resolvedAxisGeometry.axisX;
    let axisY = resolvedAxisGeometry.axisY;
    if ((kind === 'value') === horizontal) {
      const minY = front.y;
      const maxY = front.y + front.h;
      const minPoint = projection.project(front.x + front.w / 2, minY, depth);
      const maxPoint = projection.project(front.x + front.w / 2, maxY, depth);
      const lowY = minPoint.y >= maxPoint.y ? minY : maxY;
      const highY = lowY === minY ? maxY : minY;
      axisY = position === 'low' ? lowY : highY;
    } else {
      const minX = front.x;
      const maxX = front.x + front.w;
      const minPoint = projection.project(minX, front.y + front.h / 2, depth);
      const maxPoint = projection.project(maxX, front.y + front.h / 2, depth);
      const lowX = minPoint.x <= maxPoint.x ? minX : maxX;
      const highX = lowX === minX ? maxX : minX;
      axisX = position === 'low' ? lowX : highX;
    }
    return {
      axisX, axisY, depth,
      horizontalStart: projection.project(front.x, axisY, depth),
      horizontalEnd: projection.project(front.x + front.w, axisY, depth),
      verticalStart: projection.project(axisX, front.y + front.h, depth),
      verticalEnd: projection.project(axisX, front.y, depth),
    };
  };
  const valueLabelGeometry = edgeGeometry('value', chart.valAxisTickLabelPos);
  const categoryLabelGeometry = edgeGeometry('category', chart.catAxisTickLabelPos);
  const valueFontPx = chartTextFontSizePx(chart.valAxisFontSizeHpt, ptToPx) ?? 9 * ptToPx;
  ctx.font = `${chart.valAxisFontItalic ? 'italic ' : ''}${chart.valAxisFontBold ? 'bold ' : ''}${valueFontPx}px ${elementFontFamily(chart, chart.valAxisFontFace)}`;
  ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#595959';
  ctx.textAlign = horizontal ? 'center' : 'right';
  ctx.textBaseline = horizontal ? 'top' : 'middle';
  if (!chart.valAxisHidden && chart.valAxisTickLabelPos !== 'none') {
    const { axisX, axisY, depth: axisDepth } = valueLabelGeometry;
    const sceneMidpoint = projection.project(front.x + front.w / 2, front.y + front.h / 2, axisDepth);
    const valueAxisStart = horizontal
      ? valueLabelGeometry.horizontalStart : valueLabelGeometry.verticalStart;
    const valueAxisEnd = horizontal
      ? valueLabelGeometry.horizontalEnd : valueLabelGeometry.verticalEnd;
    const labelNormal = screenAnnotationOutward(
      valueAxisStart,
      valueAxisEnd,
      sceneMidpoint,
      horizontal ? 'vertical' : 'horizontal',
    );
    ctx.textAlign = Math.abs(labelNormal.x) < 0.2 ? 'center' : labelNormal.x < 0 ? 'right' : 'left';
    ctx.textBaseline = Math.abs(labelNormal.y) < 0.2 ? 'middle' : labelNormal.y < 0 ? 'bottom' : 'top';
    const labelOffset = projectedAxisTickLabelOffsetPx(
      chart.valAxisMajorTickMark, chart.valAxisLineHidden, ptToPx, 5,
    );
    const displayUnitDivisor = chart.valAxisDisplayUnits?.divisor;
    for (const value of axis.majorTicks) {
      const point = horizontal
        ? projection.project(front.x + axis.fraction(value) * front.w, axisY, axisDepth)
        : projection.project(axisX, front.y + front.h - axis.fraction(value) * front.h, axisDepth);
      ctx.fillText(formatChartValWithCode(
        percent
          ? value / 100
          : displayUnitDivisor != null
              && Number.isFinite(displayUnitDivisor)
              && displayUnitDivisor > 0
            ? value / displayUnitDivisor
            : value,
        percent ? chart.valAxisFormatCode ?? '0%' : chart.valAxisFormatCode,
        chart.date1904,
      ), point.x + labelNormal.x * labelOffset, point.y + labelNormal.y * labelOffset);
    }
  }
  const categoryFontPx = chartTextFontSizePx(chart.catAxisFontSizeHpt, ptToPx) ?? 9 * ptToPx;
  ctx.font = `${chart.catAxisFontItalic ? 'italic ' : ''}${chart.catAxisFontBold ? 'bold ' : ''}${categoryFontPx}px ${elementFontFamily(chart, chart.catAxisFontFace)}`;
  ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#595959';
  if (!chart.catAxisHidden && chart.catAxisTickLabelPos !== 'none') {
    const labelOffset = categoryLabelOffsetPx(
      projectedAxisTickLabelOffsetPx(
        chart.catAxisMajorTickMark, chart.catAxisLineHidden, ptToPx, 6,
      ),
      chart.catAxisLabelOffsetPercent,
    );
    const formattedCategories = Array.from({ length: categoryCount }, (_, index) =>
      formatCategoryLabel(
        String(categories[index] ?? index + 1),
        chart.catAxisFormatCode,
        chart.date1904,
      )
    );
    let rotation = authoredCategoryRotation(chart);
    if (rotation == null) {
      rotation = 0;
      if (!horizontal && categoryCount > 1) {
        let minimumSpacing = Number.POSITIVE_INFINITY;
        let previous: Point | null = null;
        const categoryAxisY = categoryLabelGeometry.axisY;
        for (let index = 0; index < categoryCount; index++) {
          const fraction = categoryPositionFraction(index, categoryCount, categoryBetween, categoryReversed);
          const current = projection.project(
            front.x + fraction * front.w,
            categoryAxisY,
            projection.topology.nearDepth,
          );
          if (previous) minimumSpacing = Math.min(
            minimumSpacing,
            Math.hypot(current.x - previous.x, current.y - previous.y),
          );
          previous = current;
        }
        const widest = Math.max(0, ...formattedCategories.map(label => ctx.measureText(label).width));
        // Office automatically angles dense cartesian 3-D category labels.
        // The boundary corpus repeats a -45° choice once a horizontal label
        // no longer fits between adjacent projected tick positions. Explicit
        // bodyPr@rot, including an authored zero, remains authoritative.
        if (widest > minimumSpacing * 0.9) rotation = -Math.PI / 4;
      }
    }
    const labelInterval = Math.max(1, Math.floor(chart.catAxisTickLabelSkip ?? 1));
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += labelInterval) {
      const fraction = categoryPositionFraction(
        categoryIndex, categoryCount, categoryBetween, categoryReversed,
      );
      const { axisX, axisY, depth: axisDepth } = categoryLabelGeometry;
      const point = horizontal
        ? projection.project(
          axisX,
          front.y + categoryPositionFraction(
            categoryIndex, categoryCount, categoryBetween, categoryReversed,
          ) * front.h,
          axisDepth,
        )
        : projection.project(front.x + fraction * front.w, axisY, axisDepth);
      if (horizontal) {
        const axisMidpoint = projection.project(axisX, front.y + front.h / 2, axisDepth);
        const sceneMidpoint = projection.project(front.x + front.w / 2, front.y + front.h / 2, axisDepth);
        const onScreenLeft = axisMidpoint.x <= sceneMidpoint.x;
        ctx.textAlign = onScreenLeft ? 'right' : 'left';
        ctx.textBaseline = 'middle';
        ctx.fillText(
          formattedCategories[categoryIndex],
          point.x + (onScreenLeft ? -labelOffset : labelOffset),
          point.y,
        );
      } else {
        const sceneMidpoint = projection.project(
          front.x + front.w / 2,
          front.y + front.h / 2,
          axisDepth,
        );
        const outward = screenAnnotationOutward(
          categoryLabelGeometry.horizontalStart,
          categoryLabelGeometry.horizontalEnd,
          sceneMidpoint,
          'vertical',
        );
        drawAngledCategoryLabel(
          ctx,
          formattedCategories[categoryIndex],
          point,
          rotation,
          horizontal,
          outward.y < 0 ? -1 : 1,
          labelOffset,
        );
      }
    }
  }
  drawThreeDAxisTitles(ctx, chart, rect, plot, horizontal, ptToPx);
  // Data labels are chart foreground annotations. Flush them only after all
  // walls, faces, lines, markers, axes, ticks and tick labels, so later scene
  // primitives cannot cross a label belonging to an earlier series.
  for (const drawLabel of deferredDataLabels) drawLabel();
  simpleLegend(ctx, chart, legend, ptToPx, false, legendMeasure, shapeRotationDeg);
  return true;
}

function renderPie(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
  shapeRotationDeg: number,
): boolean {
  if (!chart.threeD || chart.chartType !== 'pie') return false;
  const series = chart.series[0];
  if (!series) return true;
  const values = series.values.flatMap((value, index) =>
    value != null && Number.isFinite(value) ? [{ index, value: Math.abs(value) }] : []
  );
  let maxMagnitude = 0;
  for (const item of values) maxMagnitude = Math.max(maxMagnitude, item.value);
  if (!(maxMagnitude > 0)) return true;
  const scaledTotal = values.reduce((sum, item) => sum + item.value / maxMagnitude, 0);
  if (!(scaledTotal > 0) || !Number.isFinite(scaledTotal)) return true;
  const { plot, legend, legendMeasure } = titleAndPlot(
    ctx, chart, rect, ptToPx, 'radial', shapeRotationDeg,
  );
  // Pie is a cylindrical sector in X/Z with thickness on Y. Its authored
  // depth/gap fields do not control the radial solid (the depth=100/2000
  // Office references are pixel-identical), but rotX/rotY/perspective must use
  // the same homogeneous camera as every cartesian mesh.
  let projection = planChartThreeDProjection({
    ...chart.threeD,
    // Unlike the cartesian compatibility view, an omitted radial yaw starts
    // at zero: the first-slice ray remains at twelve o'clock. An authored
    // rotY still rotates the complete solid and its slice seams.
    rotationY: chart.threeD.rotationY ?? 0,
    // MS-OE376 §2.1.1501(b): pie3D hPercent scales the default pie
    // thickness. It is not the cartesian scene-height ratio used by the
    // other 3-D families.
    heightPercent: undefined,
    depthPercent: 100,
  }, plot, {
    sceneDepthScale: 1,
    // radius=.45W and thickness=.30r occupy .135W on model Y. A .15W scene
    // box tightly fits the actual cylinder when hPercent is omitted, instead
    // of shrinking it to fit an unused full-height cartesian cuboid.
    sceneHeightScale: 0.15,
  });
  if (!projection) return true;
  const pointOverrides = new Map<number, ChartDataPointOverride>(
    series.dataPointOverrides?.map(override => [override.idx, override]) ?? [],
  );
  let maxExplosion = 0;
  for (const override of pointOverrides.values()) {
    if (override.explosion != null && Number.isFinite(override.explosion)) {
      maxExplosion = Math.max(maxExplosion, Math.max(0, Math.min(100, override.explosion)) / 100);
    }
  }
  const { scene } = projection;
  const radius = Math.min(
    scene.w * 0.45 / (1 + maxExplosion),
    projection.modelDepth * 0.45 / (1 + maxExplosion),
    scene.h / 0.45,
  );
  if (!(radius > 0)) return true;
  const centerX = scene.x + scene.w / 2;
  const centerY = scene.y + scene.h / 2;
  const centerDepth = 0.5;
  // Office's default pie wall is about .29r. The camera now decides which
  // side is visible; no screen-down wall or abs(rotX) special case remains.
  const thickness = radius * 0.30 * pieThreeDThicknessMultiplier(chart.threeD);
  // rotY is already part of the camera and must not be added to the seam a
  // second time. Slice angles are resolved in the X/Z model plane below.
  let cumulativeFraction = 0;
  const slices: Array<{
    index: number;
    start: number;
    end: number;
    color: string;
    fillPaint: Fill | null | undefined;
    value: number;
    percentValue: number;
    centerX: number;
    centerDepth: number;
    segments: number;
    mesh: ThreeDMesh;
    lineColor: string | null;
    lineFill: Fill | null | undefined;
    lineWidthEmu: number | null;
    lineDash: string;
    lineCustomDash: ChartModel['plotAreaLineCustomDash'];
    lineCap: CanvasLineCap;
    lineJoin: CanvasLineJoin;
  }> = [];
  const labelOverrides = new Map<number, ChartDataLabelOverride>(
    series.dataLabelOverrides?.map(override => [override.idx, override]) ?? [],
  );
  // Keep the projected circumference below roughly four pixels per facet,
  // bounded to 128 segments for synchronous Canvas work. A fixed 32-facet
  // ring is visibly polygonal on large/zoomed charts and makes an otherwise
  // constant-width outline appear stepped at every facet boundary.
  const pieRoundSegments = Math.max(48, Math.min(
    128, Math.ceil(Math.PI * 2 * radius / 4),
  ));
  for (const item of values) {
    const percentValue = (item.value / maxMagnitude) / scaledTotal;
    const angles = threeDPieSliceAngles(
      chart.firstSliceAngle, cumulativeFraction, percentValue,
    );
    const pointOverride = pointOverrides.get(item.index);
    const paint = threeDDatumPaint(chart, series, pointOverride, item.index, 0);
    const middle = angles.middle;
    const explosion = pointOverride?.explosion != null && Number.isFinite(pointOverride.explosion)
      ? Math.max(0, Math.min(100, pointOverride.explosion)) / 100
      : 0;
    const sliceCenterX = centerX + Math.cos(middle) * radius * explosion;
    const sliceCenterDepth = centerDepth
      + Math.sin(middle) * radius * explosion / projection.modelDepth;
    const segments = Math.max(2, Math.ceil(pieRoundSegments * percentValue));
    const mesh = buildThreeDPieSectorMesh({
      centerX: sliceCenterX,
      centerY,
      centerDepth: sliceCenterDepth,
      radius,
      modelDepth: projection.modelDepth,
      thickness,
      startAngle: angles.start,
      endAngle: angles.end,
      segments,
    });
    if (!mesh) {
      cumulativeFraction += percentValue;
      continue;
    }
    slices.push({
      index: item.index,
      start: angles.start,
      end: angles.end,
      color: paint.fill === null ? 'transparent' : paint.color,
      fillPaint: paint.fill,
      value: item.value,
      percentValue,
      centerX: sliceCenterX,
      centerDepth: sliceCenterDepth,
      segments,
      mesh,
      lineColor: paint.lineColor,
      lineFill: paint.lineFill,
      lineWidthEmu: paint.lineWidthEmu ?? null,
      lineDash: paint.lineDash ?? 'solid',
      lineCustomDash: paint.lineCustomDash,
      lineCap: paint.lineCap,
      lineJoin: paint.lineJoin,
    });
    cumulativeFraction += percentValue;
  }
  projection = fitChartThreeDProjectionToPoints(
    projection,
    slices.flatMap(slice => slice.mesh.vertices),
    plot,
    0.08,
  );
  const deferredLabels: Array<() => void> = [];
  const pieBudget: ScenePrimitiveBudget = {
    remaining: MAX_PROJECTED_STROKE_PRIMITIVES,
    exceeded: false,
  };
  const outlineStyleForSlice = (slice: (typeof slices)[number]): MeshOutlineStyle | undefined => {
    const width = slice.lineWidthEmu != null
      ? Math.max(0.25, slice.lineWidthEmu / EMU_PER_PT * ptToPx)
      : 0.75 * ptToPx;
    return slice.lineFill !== null
      && (slice.lineColor != null || slice.lineFill !== undefined) ? {
      color: slice.lineColor ?? 'rgba(0,0,0,0.42)',
      width,
      dash: drawingmlLineDashArray(slice.lineCustomDash, slice.lineDash, width),
      cap: slice.lineCap,
      join: slice.lineJoin,
    } : undefined;
  };
  // Fill solids remain independently colored and depth-sorted. A uniform,
  // non-exploded pie uses semantic continuous outline paths; differently
  // styled or exploded points retain independent authored solid outlines.
  const pieFillFaceGroups = slices.map(slice => {
    const faces = projectThreeDMesh(
      projection, slice.mesh, slice.color, undefined, pieBudget,
    );
    applyScenePaint(ctx, faces, 'fill', slice.fillPaint, shapeRotationDeg);
    return { slice, faces };
  });
  const pieFillFaces = pieFillFaceGroups.flatMap(group => group.faces);
  const pieOutlineFaces: SceneFace[] = [];
  const outlineStyles = slices.map(outlineStyleForSlice);
  const firstOutline = outlineStyles[0];
  const outlineKey = (style: MeshOutlineStyle | undefined) => style == null ? null : [
    style.color, style.width, style.dash.join(','), style.cap, style.join,
  ].join('|');
  const uniformOutline = firstOutline != null
    && slices.every(slice => slice.lineFill === undefined)
    && outlineStyles.every(style => outlineKey(style) === outlineKey(firstOutline))
    && slices.every(slice => Math.abs(slice.centerX - centerX) < 1e-9
      && Math.abs(slice.centerDepth - centerDepth) < 1e-9);
  if (uniformOutline) {
    pieOutlineFaces.push(...projectThreeDPieOutline(
      projection, slices, centerY, radius, thickness, firstOutline, pieBudget,
    ));
  } else {
    for (let index = 0; index < slices.length; index++) {
      const style = outlineStyles[index];
      if (!style) continue;
      const faces = projectThreeDMesh(
        projection, slices[index].mesh, 'transparent', style, pieBudget, true,
      );
      applyScenePaint(ctx, faces, 'outline', slices[index].lineFill, shapeRotationDeg);
      pieOutlineFaces.push(...faces);
    }
  }
  if (pieBudget.exceeded) {
    paintThreeDTooManyDataPoints(ctx, rect);
    return true;
  }
  if (uniformOutline) {
    for (const item of sortProjectedSceneFaces(pieFillFaces)) paintSceneFace(ctx, item);
    // Semantic paths already contain only visible cap/wall boundaries. Paint
    // them after all colored faces so neither adjacent coplanar slice can hide
    // half the authored screen-space width.
    for (const item of sortProjectedSceneFaces(pieOutlineFaces)) paintSceneFace(ctx, item);
  } else {
    // Exploded or independently styled slices still require inter-solid depth
    // ordering because their authored outlines can genuinely occlude peers.
    for (const item of sortProjectedSceneFaces([
      ...pieFillFaces, ...pieOutlineFaces,
    ])) paintSceneFace(ctx, item);
  }
  for (const { slice, faces } of pieFillFaceGroups) {
    const point = pointOverrides.get(slice.index);
    const styleIndex = threeDDatumStyleIndex(chart, series, slice.index, 0);
    const pointEffect = chartStyleEffectOwner(point?.chartexStyle);
    paintChartStyleOuterEffectBehind(
      ctx,
      pointEffect ?? chartStyleEffectOwner(series.chartexStyle),
      chartDataPointStyleRole(chart, 'dataPoint3D', 0),
      pointEffect ? slice.index : styleIndex,
      ptToPx,
      target => paintSceneEffectSilhouette(target, faces),
      styleIndex,
    );
  }
  for (const slice of slices) {
    const middle = (slice.start + slice.end) / 2;
    const labelOverride = labelOverrides.get(slice.index);
    if (wantsThreeDDataLabel(chart, series, labelOverride)) {
      const defaults = series.seriesDataLabels;
      const fontPx = chartTextFontSizePx(
        labelOverride?.fontSizeHpt ?? defaults?.fontSizeHpt ?? chart.dataLabelFontSizeHpt,
        ptToPx,
      ) ?? 9 * ptToPx;
      ctx.font = `${labelOverride?.fontBold ?? defaults?.fontBold
        ?? chart.dataLabelFontBold ? 'bold ' : ''}${fontPx}px ${chartFontFamily(
        chart, labelOverride?.fontFace ?? defaults?.fontFace ?? chart.dataLabelFontFace,
      )}`;
      const labelText = effectiveDataLabelText({
        customText: labelOverride?.text,
        showCategory: labelOverride?.showCatName ?? defaults?.showCatName ?? false,
        showSeries: labelOverride?.showSerName ?? defaults?.showSerName ?? false,
        showValue: labelOverride?.showVal ?? defaults?.showVal ?? chart.showDataLabels,
        showPercent: labelOverride?.showPercent ?? defaults?.showPercent ?? false,
        category: series.categories?.[slice.index]
          ?? chart.categories[slice.index] ?? `${slice.index + 1}`,
        seriesName: series.name || 'Series 1',
        sourceValue: slice.value,
        percentRatio: slice.percentValue,
        formatCode: labelOverride?.formatCode
          ?? defaults?.formatCode ?? chart.dataLabelFormatCode ?? series.valFormatCode,
        separator: labelOverride?.separator ?? defaults?.separator,
        date1904: chart.date1904,
      });
      const fallbackFamily = chartFontFamily(
        chart, labelOverride?.fontFace ?? defaults?.fontFace ?? chart.dataLabelFontFace,
      );
      const richMeasure = labelOverride?.text && labelOverride.richRuns?.length
        ? resolveRichDataLabelBlock(ctx, {
          runs: labelOverride.richRuns,
          ptToPx,
          fontFamily: fallbackFamily,
          fallbackBold: labelOverride.fontBold ?? defaults?.fontBold
            ?? chart.dataLabelFontBold ?? false,
          fontFamilyForFace: face => chartFontFamily(chart, face),
        }, fontPx, `#${labelOverride.fontColor ?? defaults?.fontColor
          ?? series.labelColor ?? chart.dataLabelFontColor ?? '111111'}`)
        : null;
      const authoredPosition = labelOverride?.position ?? defaults?.position ?? chart.dataLabelPosition;
      const surfaceTop = centerY - thickness / 2;
      const surfaceBottom = centerY + thickness / 2;
      const surfaceY = projection.cameraDepth(
        slice.centerX, surfaceTop, slice.centerDepth,
      ) >= projection.cameraDepth(slice.centerX, surfaceBottom, slice.centerDepth)
        ? surfaceTop : surfaceBottom;
      let arcCapacity = 0;
      let previousArcPoint: Point | null = null;
      for (let step = 0; step <= 12; step++) {
        const arcAngle = slice.start + (slice.end - slice.start) * step / 12;
        const current = projection.project(
          slice.centerX + Math.cos(arcAngle) * radius * 0.64,
          surfaceY,
          slice.centerDepth + Math.sin(arcAngle) * radius * 0.64 / projection.modelDepth,
        );
        if (previousArcPoint) arcCapacity += Math.hypot(
          current.x - previousArcPoint.x, current.y - previousArcPoint.y,
        );
        previousArcPoint = current;
      }
      const outside = (authoredPosition == null || authoredPosition === 'bestFit')
        && (slice.percentValue === 0 || arcCapacity < (richMeasure?.width
          ?? ctx.measureText(labelText).width));
      const labelOutside = outside || authoredPosition === 'outEnd';
      const labelRadius = radius * (labelOutside ? 1.12 : 0.64);
      const label = projection.project(
        slice.centerX + Math.cos(middle) * labelRadius,
        surfaceY,
        slice.centerDepth + Math.sin(middle) * labelRadius / projection.modelDepth,
      );
      const leader = projection.project(
        slice.centerX + Math.cos(middle) * radius,
        surfaceY,
        slice.centerDepth + Math.sin(middle) * radius / projection.modelDepth,
      );
      deferredLabels.push(() => drawThreeDDataLabel(
        ctx, chart, series, 0, slice.index, slice.value, label, plot, ptToPx,
        0, slice.percentValue, labelOverride, 'ctr', leader,
        labelOutside,
        undefined, undefined, undefined, shapeRotationDeg,
      ));
    }
  }
  for (const drawLabel of deferredLabels) drawLabel();
  const categories = series.categories?.length ? series.categories : chart.categories;
  const legendPointColors = Array.from({ length: categories.length }, (_, index) => {
    const pointOverride = pointOverrides.get(index);
    const decision = chartThreeDDatumFillDecision(
      chart, series, pointOverride, index, 0,
    );
    if (decision === null) return '00000000';
    const authored = decision?.fillType === 'solid'
      ? decision.color
      : pointOverride?.color ?? series.dataPointColors?.[index] ?? series.color;
    if (authored === '00000000') return '00000000';
    return scaleHexColor(authored ? `#${authored}` : colorFor(index), 0.80).replace(/^#/, '');
  });
  simpleLegend(ctx, {
    ...chart,
    categories,
    series: [{ ...series, categories, dataPointColors: legendPointColors }],
  }, legend, ptToPx, true, legendMeasure, shapeRotationDeg);
  return true;
}

/** Render a bounded compatibility projection for classic 3-D chart groups.
 * Returns false for ordinary 2-D charts so existing family paths remain exact. */
export function renderSimpleThreeDChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): boolean {
  if (!chart.threeD) return false;
  const supported = chart.chartType === 'pie'
    || SUPPORTED_CARTESIAN_THREE_D_TYPES.has(chart.chartType);
  if (!supported) return false;
  if (!threeDChartPaintsWithinBudget(chart)) {
    paintThreeDTooManyDataPoints(ctx, rect);
    return true;
  }
  if (renderPie(ctx, chart, rect, ptToPx, shapeRotationDeg)) return true;
  return renderCartesian(ctx, chart, rect, ptToPx, shapeRotationDeg);
}
