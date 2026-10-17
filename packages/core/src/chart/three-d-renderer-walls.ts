// 3-D chart walls, floor, and axes: wall / pie-slice geometry, axis strokes,
// projected ticks, and axis geometry. Split from three-d-renderer.ts (#1857);
// reached only through the optional three-d renderer chunk.

import type { ChartModel } from '../types/chart.js';
import { chartImageFillSource } from './image-fill.js';
import { paintChartThreeDSurfacePicture } from './three-d-surface-picture.js';
import { EMU_PER_PT } from '../units.js';
import { pptxPresetDashArray } from '../draw/dash.js';
import { planNumericValueAxis, type NumericValueAxisPlan } from './axis-scale.js';
import {
  planChartThreeDSurfaceGridSegments,
  planChartThreeDSurfaceGeometry,
  type ChartThreeDProjection,
  type ThreeDScenePoint,
} from './three-d.js';
import {
  categoryGridlineFractions,
  categoryMinorGridlineFractions,
  categoryPositionFraction,
} from './category-spacing.js';
import { resolveNativeFill as resolveFill } from '../shape/paint.js';
import { chartThreeDSurfacePaint } from './style-paint.js';
import { drawingmlLineDashArray } from '../draw/dash.js';
import type { Point } from './three-d-renderer-paint.js';
import { polygon } from './three-d-renderer-scene.js';

export function axisPlan(
  chart: ChartModel,
  dataMin: number,
  dataMax: number,
  axisLenPt: number,
  percent: boolean,
  orientation: 'vertical' | 'horizontal',
): NumericValueAxisPlan {
  const factor = percent ? 100 : 1;
  const minorTickMark = chart.valAxisMinorTickMark ?? 'none';
  const options = {
    dataMin,
    dataMax,
    explicitMin: chart.valMin == null ? (percent ? dataMin : null) : chart.valMin * factor,
    explicitMax: chart.valMax == null ? (percent ? dataMax : null) : chart.valMax * factor,
    majorUnit: chart.valAxisMajorUnit == null ? null : chart.valAxisMajorUnit * factor,
    minorUnit: chart.valAxisMinorUnit == null ? null : chart.valAxisMinorUnit * factor,
    axisLenPt,
    axisOrientation: orientation,
    logBase: chart.valAxisLogBase,
    reversed: chart.valAxisOrientation === 'maxMin',
    needMinor: chart.valAxisMinorGridlines === true || minorTickMark !== 'none',
  } as const;
  return planNumericValueAxis(options);
}

interface ThreeDStroke {
  color: string;
  width: number;
  dash: number[];
}

export interface ThreeDWallGeometry {
  floor: readonly Point[];
  sideWall: readonly Point[];
  backWall: readonly Point[];
  /** Model-space category coordinate used by the visible side wall. */
  sideX: number;
  /** Opposite category coordinate used by the series/depth axis. */
  seriesAxisX: number;
  floorY: number;
  oppositeFloorY: number;
  nearDepth: 0 | 1;
  farDepth: 0 | 1;
}

export interface ThreeDPieSliceAngles {
  /** Ascending model-angle interval consumed by the sector mesh. */
  start: number;
  end: number;
  middle: number;
  /** Authored leading ray, where the clockwise slice begins. */
  leading: number;
}

/** Convert OOXML's clockwise first-slice angle from screen twelve o'clock into
 * the X/Z model plane used by the shared 3-D camera. Positive model Z projects
 * upward for the ordinary positive elevation, so twelve o'clock is +π/2 (not
 * Canvas' -π/2). The mesh interval stays ascending while retaining the
 * authored clockwise sector. */
export function threeDPieSliceAngles(
  firstSliceAngle: number | null | undefined,
  cumulativeFraction: number,
  sliceFraction: number,
): ThreeDPieSliceAngles {
  const normalizedFirst = firstSliceAngle != null && Number.isFinite(firstSliceAngle)
    ? ((firstSliceAngle % 360) + 360) % 360 : 0;
  const cumulative = Number.isFinite(cumulativeFraction)
    ? Math.max(0, Math.min(1, cumulativeFraction)) : 0;
  const fraction = Number.isFinite(sliceFraction)
    ? Math.max(0, Math.min(1 - cumulative, sliceFraction)) : 0;
  const leading = Math.PI / 2 - (normalizedFirst * Math.PI / 180 + cumulative * Math.PI * 2);
  const trailing = leading - fraction * Math.PI * 2;
  return {
    start: Math.min(leading, trailing),
    end: Math.max(leading, trailing),
    middle: (leading + trailing) / 2,
    leading,
  };
}

/** Resolve the three authored CT_Surface faces from the same projected chart
 * cuboid used by bars, grids and axes. Keeping the shared edges identical is
 * what makes the side wall, floor and back-wall separator close as one drawing
 * instead of a collection of screen-space approximations. */
export function threeDWallGeometry(
  projection: ChartThreeDProjection,
): ThreeDWallGeometry {
  const { front } = projection;
  const xMin = front.x;
  const xMax = front.x + front.w;
  const sideX = projection.topology.farX === 'min' ? xMin : xMax;
  const seriesAxisX = sideX === xMin ? xMax : xMin;
  const floorY = projection.topology.axisY === 'min' ? front.y : front.y + front.h;
  const oppositeFloorY = floorY === front.y ? front.y + front.h : front.y;
  const { nearDepth, farDepth } = projection.topology;
  const floorNearMin = projection.project(xMin, floorY, nearDepth);
  const floorNearMax = projection.project(xMax, floorY, nearDepth);
  const floorFarMax = projection.project(xMax, floorY, farDepth);
  const floorFarMin = projection.project(xMin, floorY, farDepth);
  const backTopMin = projection.project(xMin, oppositeFloorY, farDepth);
  const backTopMax = projection.project(xMax, oppositeFloorY, farDepth);
  const sideNearFloor = projection.project(sideX, floorY, nearDepth);
  const sideFarFloor = projection.project(sideX, floorY, farDepth);
  const sideFarTop = projection.project(sideX, oppositeFloorY, farDepth);
  const sideNearTop = projection.project(sideX, oppositeFloorY, nearDepth);
  return {
    floor: [floorNearMin, floorNearMax, floorFarMax, floorFarMin],
    sideWall: [sideNearFloor, sideFarFloor, sideFarTop, sideNearTop],
    backWall: [floorFarMin, floorFarMax, backTopMax, backTopMin],
    sideX,
    seriesAxisX,
    floorY,
    oppositeFloorY,
    nearDepth,
    farDepth,
  };
}

function threeDStroke(
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  dash: string | null | undefined,
  ptToPx: number,
  fallbackColor = 'A6A6A6',
  fallbackWidth = 0.75,
): ThreeDStroke {
  const width = widthEmu != null && Number.isFinite(widthEmu) && widthEmu >= 0
    ? Math.max(0.25, widthEmu / EMU_PER_PT * ptToPx)
    : fallbackWidth * ptToPx;
  return {
    color: `#${color ?? fallbackColor}`,
    width,
    dash: pptxPresetDashArray(dash ?? 'solid', width),
  };
}

export function threeDAxisStroke(
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  dash: string | null | undefined,
  ptToPx: number,
): ThreeDStroke {
  const stroke = threeDStroke(color, widthEmu, dash, ptToPx, '898989', 1);
  if (widthEmu == null || !Number.isFinite(widthEmu) || widthEmu < 0) return stroke;
  // `<c:*Ax><c:spPr><a:ln@w>` owns the coordinate-rule width. Keep it in the
  // same projected point scale as the chart instead of applying a device-pixel
  // floor; otherwise zoom changes and a thin authored axis can become thicker
  // than the surrounding CT_Surface rules.
  const authoredWidth = widthEmu / EMU_PER_PT * ptToPx;
  if (!(authoredWidth > 0)) return stroke;
  const width = authoredWidth;
  return { ...stroke, width, dash: pptxPresetDashArray(dash ?? 'solid', width) };
}

export function threeDAxisTickStroke(
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  dash: string | null | undefined,
  ptToPx: number,
): ThreeDStroke {
  // Tick rules use exactly the same authored DrawingML point scale as their
  // coordinate axis. A separate device-pixel floor would make ticks thicker
  // than the owning rule at low zoom.
  return threeDAxisStroke(color, widthEmu, dash, ptToPx);
}

export function applyThreeDStroke(ctx: CanvasRenderingContext2D, stroke: ThreeDStroke): void {
  ctx.strokeStyle = stroke.color;
  ctx.lineWidth = stroke.width;
  ctx.setLineDash(stroke.dash);
}

export function walls(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  projection: ChartThreeDProjection,
  axis: NumericValueAxisPlan,
  orientation: 'vertical' | 'horizontal',
  categoryCount: number,
  categoryBetween: boolean,
  categoryReversed: boolean,
  ptToPx: number,
): void {
  const { front } = projection;
  const xMin = front.x;
  const xMax = front.x + front.w;
  const geometry = threeDWallGeometry(projection);
  const {
    sideX: farX, floorY, oppositeFloorY, nearDepth, farDepth,
  } = geometry;
  const axisGeometry = threeDAxisGeometry(
    chart, projection, axis, categoryCount, categoryBetween, categoryReversed, orientation,
  );
  const axisOwnedEdges: Array<readonly [Point, Point]> = [];
  // Linked chart styles may resolve the same color/width fields without a
  // direct axis line. Only a direct `<c:*Ax><c:spPr><a:ln>` owns a coincident
  // CT_Surface perimeter; style-derived axes retain the complete wall frame.
  const categoryAxisPaintAuthored = chart.catAxisLinePaintAuthored === true;
  const valueAxisPaintAuthored = chart.valAxisLinePaintAuthored === true;
  const horizontalAxisVisible = orientation === 'vertical'
    ? categoryAxisPaintAuthored && !chart.catAxisHidden && !chart.catAxisLineHidden
    : valueAxisPaintAuthored && !chart.valAxisHidden && !chart.valAxisLineHidden;
  const verticalAxisVisible = orientation === 'vertical'
    ? valueAxisPaintAuthored && !chart.valAxisHidden && !chart.valAxisLineHidden
    : categoryAxisPaintAuthored && !chart.catAxisHidden && !chart.catAxisLineHidden;
  if (horizontalAxisVisible) {
    axisOwnedEdges.push([axisGeometry.horizontalStart, axisGeometry.horizontalEnd]);
  }
  if (verticalAxisVisible) {
    axisOwnedEdges.push([axisGeometry.verticalStart, axisGeometry.verticalEnd]);
  }
  const seriesAxis = chart.threeD?.seriesAxis;
  const seriesAxisPaintAuthored = seriesAxis?.linePaintAuthored === true;
  if (seriesAxis && !seriesAxis.hidden && !seriesAxis.lineHidden
    && seriesAxisPaintAuthored
    && chart.threeD?.barGrouping === 'standard' && chart.series.length > 0) {
    const seriesAxisX = orientation === 'vertical' ? geometry.seriesAxisX : axisGeometry.axisX;
    const seriesAxisY = orientation === 'horizontal'
      ? (floorY === projection.front.y
        ? projection.front.y + projection.front.h : projection.front.y)
      : floorY;
    axisOwnedEdges.push([
      projection.project(seriesAxisX, seriesAxisY, nearDepth),
      projection.project(seriesAxisX, seriesAxisY, farDepth),
    ]);
  }
  const samePoint = (left: Point, right: Point): boolean =>
    Math.hypot(left.x - right.x, left.y - right.y) <= 1e-6;
  const axisOwnsEdge = (start: Point, end: Point): boolean =>
    axisOwnedEdges.some(([axisStart, axisEnd]) =>
      (samePoint(start, axisStart) && samePoint(end, axisEnd))
      || (samePoint(start, axisEnd) && samePoint(end, axisStart))
    );
  const projectedSurfaceFaces = (
    kind: 'floor' | 'sideWall' | 'backWall',
    surface: NonNullable<ChartModel['threeD']>['floor'],
  ) => {
    const slab = planChartThreeDSurfaceGeometry(projection, kind, surface?.thicknessPercent);
    const faces = slab.faces
      .map((scenePoints, faceIndex) => ({ scenePoints, faceIndex }))
      .filter(({ scenePoints }) => slab.thickness === 0 || projection.cameraFacing(scenePoints))
      .map(({ scenePoints, faceIndex }) => ({
        faceIndex,
        scenePoints,
        points: scenePoints.map(point =>
          projection.projectUnbounded(point.x, point.y, point.depth)
        ),
        depth: scenePoints.reduce(
          (sum, point) => sum + projection.cameraDepth(point.x, point.y, point.depth),
          0,
        ) / scenePoints.length,
      }))
      .sort((left, right) => left.depth - right.depth);
    return {
      slab,
      faces,
      visibleFaceIndices: new Set(faces.map(face => face.faceIndex)),
    };
  };
  const floorSurface = projectedSurfaceFaces('floor', chart.threeD?.floor);
  const sideWallSurface = projectedSurfaceFaces('sideWall', chart.threeD?.sideWall);
  const backWallSurface = projectedSurfaceFaces('backWall', chart.threeD?.backWall);
  const floorFaces = floorSurface.faces;
  const sideWallFaces = sideWallSurface.faces;
  const backWallFaces = backWallSurface.faces;
  const drawSurfaceFill = (
    group: ReturnType<typeof projectedSurfaceFaces>,
    surface: NonNullable<ChartModel['threeD']>['floor'],
    role: 'floor' | 'wall',
    kind: 'floor' | 'sideWall' | 'backWall',
  ) => {
    // CT_Surface does not imply a paint when c:spPr/fill is omitted. Excel's
    // classic 3-D default leaves these faces unfilled; only an authored fill
    // creates an opaque wall. The wall/grid rules are painted independently.
    const effective = chartThreeDSurfacePaint(chart, surface, role);
    const faces = group.faces;
    if (!effective.fill || !faces.length) return;
    if (effective.fill.fillType === 'image') {
      const image = chartImageFillSource(effective.fill);
      if (!image) return;
      const project = (point: ThreeDScenePoint): Point =>
        projection.projectUnbounded(point.x, point.y, point.depth);
      paintChartThreeDSurfacePicture(
        ctx, effective.fill, image, surface, kind,
        group.slab, group.faces.map(face => face.faceIndex),
        project, axis.max - axis.min,
      );
      return;
    }
    const points = faces.flatMap(face => face.points);
    const minX = Math.min(...points.map(point => point.x));
    const maxX = Math.max(...points.map(point => point.x));
    const minY = Math.min(...points.map(point => point.y));
    const maxY = Math.max(...points.map(point => point.y));
    const fill = effective.fill.fillType === 'solid'
      ? `#${effective.fill.color}`
      : resolveFill(effective.fill, ctx, minX, minY, maxX - minX, maxY - minY);
    if (!fill) return;
    for (const face of faces) {
      if (face.points.length < 3) continue;
      polygon(ctx, face.points);
      ctx.fillStyle = fill;
      ctx.fill();
    }
  };
  drawSurfaceFill(floorSurface, chart.threeD?.floor, 'floor', 'floor');
  drawSurfaceFill(sideWallSurface, chart.threeD?.sideWall, 'wall', 'sideWall');
  drawSurfaceFill(backWallSurface, chart.threeD?.backWall, 'wall', 'backWall');
  const strokeSurfaceGridRule = (
    group: ReturnType<typeof projectedSurfaceFaces>,
    kind: 'floor' | 'sideWall' | 'backWall',
    coordinate: 'x' | 'y',
    fraction: number,
    includeExterior = true,
  ): void => {
    for (const segment of planChartThreeDSurfaceGridSegments(
      group.slab,
      kind,
      coordinate,
      fraction,
    )) {
      if (!includeExterior && segment.faceIndex !== 0) continue;
      if (!group.visibleFaceIndices.has(segment.faceIndex)) continue;
      const [start, end] = segment.scenePoints.map(point =>
        projection.projectUnbounded(point.x, point.y, point.depth)
      );
      ctx.beginPath();
      ctx.moveTo(start.x, start.y);
      ctx.lineTo(end.x, end.y);
      ctx.stroke();
    }
  };
  const drawValueGrid = (
    values: readonly number[],
    stroke: ThreeDStroke,
    authoredSurfaceRule: boolean,
  ) => {
    applyThreeDStroke(ctx, stroke);
    for (const value of values) {
      const fraction = axis.fraction(value);
      if (authoredSurfaceRule) {
        if (orientation === 'horizontal') {
          strokeSurfaceGridRule(floorSurface, 'floor', 'x', fraction);
          strokeSurfaceGridRule(backWallSurface, 'backWall', 'x', fraction);
        } else {
          strokeSurfaceGridRule(sideWallSurface, 'sideWall', 'y', fraction);
          strokeSurfaceGridRule(backWallSurface, 'backWall', 'y', fraction);
        }
        continue;
      }
      if (orientation === 'horizontal') {
        const x = front.x + fraction * front.w;
        const near = projection.project(x, floorY, farDepth);
        const far = projection.project(x, oppositeFloorY, farDepth);
        ctx.beginPath(); ctx.moveTo(near.x, near.y); ctx.lineTo(far.x, far.y); ctx.stroke();
      } else {
        // The same value plane crosses both the visible side wall and the back
        // wall. Leaving out the side segment makes the grid appear to change
        // direction at the wall boundary even though the scene is projective.
        const y = front.y + front.h - fraction * front.h;
        const sideNear = projection.project(farX, y, nearDepth);
        const sideFar = projection.project(farX, y, farDepth);
        const backEnd = projection.project(farX === xMin ? xMax : xMin, y, farDepth);
        ctx.beginPath(); ctx.moveTo(sideNear.x, sideNear.y); ctx.lineTo(sideFar.x, sideFar.y); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(sideFar.x, sideFar.y); ctx.lineTo(backEnd.x, backEnd.y); ctx.stroke();
      }
    }
  };
  if (chart.valAxisMinorGridlines === true) {
    drawValueGrid(axis.minorTicks, threeDStroke(
      chart.valAxisMinorGridlineColor,
      chart.valAxisMinorGridlineWidthEmu,
      chart.valAxisMinorGridlineDash,
      ptToPx,
      'D9D9D9',
      0.5,
    ), true);
  }
  // CT_ValAx/majorGridlines is an optional element, not an on-by-default
  // property. Its absence therefore contributes no value-grid geometry. This
  // mirrors the 2-D axis path and prevents synthetic back-wall rules from
  // appearing in 3-D charts that author only the axis line.
  if (chart.valAxisMajorGridlines === true) {
    drawValueGrid(axis.majorTicks, threeDStroke(
      chart.valAxisGridlineColor,
      chart.valAxisGridlineWidthEmu,
      chart.valAxisGridlineDash,
      ptToPx,
      '898989',
      1,
    ), true);
  }
  // CT_CatAx/majorGridlines is optional. Only an authored element contributes
  // category-grid geometry; omission must not synthesize depth rays. Authored
  // rules continue across visible thick surface faces.
  const strokeCategorySurfaceGrid = (
    fractions: readonly number[],
    stroke: ThreeDStroke,
  ): void => {
    applyThreeDStroke(ctx, stroke);
    for (const fraction of fractions) {
      if (orientation === 'vertical') {
        strokeSurfaceGridRule(floorSurface, 'floor', 'x', fraction);
        strokeSurfaceGridRule(backWallSurface, 'backWall', 'x', fraction);
      } else {
        strokeSurfaceGridRule(sideWallSurface, 'sideWall', 'y', fraction);
        strokeSurfaceGridRule(backWallSurface, 'backWall', 'y', fraction);
      }
    }
  };
  if (chart.catAxisMinorGridlines === true) {
    strokeCategorySurfaceGrid(
      categoryMinorGridlineFractions(categoryCount, categoryBetween),
      threeDStroke(
        chart.catAxisMinorGridlineColor,
        chart.catAxisMinorGridlineWidthEmu,
        chart.catAxisMinorGridlineDash,
        ptToPx,
        'E0E0E0',
        0.5,
      ),
    );
  }
  if (chart.catAxisMajorGridlines === true) {
    strokeCategorySurfaceGrid(
      categoryGridlineFractions(categoryCount, categoryBetween),
      threeDStroke(
        chart.catAxisGridlineColor,
        chart.catAxisGridlineWidthEmu,
        chart.catAxisGridlineDash,
        ptToPx,
        'E0E0E0',
        0.5,
      ),
    );
  }
  const strokeSurface = (
    faces: readonly { points: readonly Point[] }[],
    surface: NonNullable<ChartModel['threeD']>['floor'],
    role: 'floor' | 'wall',
  ) => {
    if (!faces.length) return;
    const effective = chartThreeDSurfacePaint(chart, surface, role);
    if (effective.line === null) return;
    const linePaint = effective.line ?? { fillType: 'solid' as const, color: '898989' };
    const points = faces.flatMap(face => face.points);
    const minX = Math.min(...points.map(point => point.x));
    const maxX = Math.max(...points.map(point => point.x));
    const minY = Math.min(...points.map(point => point.y));
    const maxY = Math.max(...points.map(point => point.y));
    const line = linePaint.fillType === 'solid'
      ? `#${linePaint.color}`
      : resolveFill(linePaint, ctx, minX, minY, maxX - minX, maxY - minY);
    if (!line) return;
    const width = effective.lineWidthEmu != null
      ? Math.max(0.25, effective.lineWidthEmu / EMU_PER_PT * ptToPx)
      : ptToPx;
    ctx.strokeStyle = line;
    ctx.lineWidth = width;
    ctx.setLineDash(drawingmlLineDashArray(
      effective.lineCustomDash,
      effective.lineDash,
      width,
    ));
    ctx.lineCap = effective.lineCap === 'rnd'
      ? 'round' : effective.lineCap === 'sq' ? 'square' : 'butt';
    ctx.lineJoin = effective.lineJoin === 'round' || effective.lineJoin === 'bevel'
      ? effective.lineJoin : 'miter';
    for (const face of faces) {
      if (face.points.length < 2) continue;
      const edges = face.points.map((point, index) => ({
        start: point,
        end: face.points[(index + 1) % face.points.length],
      }));
      if (edges.some(edge => axisOwnsEdge(edge.start, edge.end))) {
        // A visible axis and a CT_Surface perimeter can describe the same
        // projected cuboid edge with different authored paints. Office gives
        // the coordinate axis visual ownership; painting both produces the
        // gray halo seen around a thinner black axis. Retain every independent
        // surface edge and omit only the exactly coincident segment.
        for (const edge of edges) {
          if (axisOwnsEdge(edge.start, edge.end)) continue;
          ctx.beginPath();
          ctx.moveTo(edge.start.x, edge.start.y);
          ctx.lineTo(edge.end.x, edge.end.y);
          ctx.stroke();
        }
        continue;
      }
      ctx.beginPath();
      ctx.moveTo(face.points[0].x, face.points[0].y);
      for (let index = 1; index < face.points.length; index++) {
        ctx.lineTo(face.points[index].x, face.points[index].y);
      }
      ctx.lineTo(face.points[0].x, face.points[0].y);
      ctx.closePath();
      ctx.stroke();
    }
  };
  // Paint each CT_Surface perimeter in chart order. The back wall is last, so
  // its authored rule owns the shared floor/back separator without drawing an
  // extra synthetic line. This also restores the complete side-wall outline.
  strokeSurface(floorFaces, chart.threeD?.floor, 'floor');
  strokeSurface(sideWallFaces, chart.threeD?.sideWall, 'wall');
  strokeSurface(backWallFaces, chart.threeD?.backWall, 'wall');
  ctx.setLineDash([]);
}

/** Front scene axes are an overlay, not part of the back-wall paint. Repaint
 * them after data faces so a near cuboid cannot erase the coordinate frame. */
export function frontAxes(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  projection: ChartThreeDProjection,
  axis: NumericValueAxisPlan,
  categoryCount: number,
  categoryBetween: boolean,
  categoryReversed: boolean,
  ptToPx: number,
  orientation: 'vertical' | 'horizontal',
): void {
  const geometry = threeDAxisGeometry(
    chart, projection, axis, categoryCount, categoryBetween, categoryReversed, orientation,
  );
  const drawHorizontal = orientation === 'vertical'
    ? !chart.catAxisHidden && !chart.catAxisLineHidden
    : !chart.valAxisHidden && !chart.valAxisLineHidden;
  const drawVertical = orientation === 'vertical'
    ? !chart.valAxisHidden && !chart.valAxisLineHidden
    : !chart.catAxisHidden && !chart.catAxisLineHidden;
  if (drawHorizontal) {
    applyThreeDStroke(ctx, threeDAxisStroke(
      orientation === 'vertical' ? chart.catAxisLineColor : chart.valAxisLineColor,
      orientation === 'vertical' ? chart.catAxisLineWidthEmu : chart.valAxisLineWidthEmu,
      orientation === 'vertical' ? chart.catAxisLineDash : chart.valAxisLineDash,
      ptToPx,
    ));
    ctx.beginPath();
    ctx.moveTo(geometry.horizontalStart.x, geometry.horizontalStart.y);
    ctx.lineTo(geometry.horizontalEnd.x, geometry.horizontalEnd.y);
    ctx.stroke();
  }
  if (drawVertical) {
    applyThreeDStroke(ctx, threeDAxisStroke(
      orientation === 'vertical' ? chart.valAxisLineColor : chart.catAxisLineColor,
      orientation === 'vertical' ? chart.valAxisLineWidthEmu : chart.catAxisLineWidthEmu,
      orientation === 'vertical' ? chart.valAxisLineDash : chart.catAxisLineDash,
      ptToPx,
    ));
    ctx.beginPath();
    ctx.moveTo(geometry.verticalStart.x, geometry.verticalStart.y);
    ctx.lineTo(geometry.verticalEnd.x, geometry.verticalEnd.y);
    ctx.stroke();
  }
  ctx.setLineDash([]);
}

/** Draw a tick in screen space from the normal of its already-projected axis.
 * The anchor comes from the same 3-D point as the axis/grid/data geometry; only
 * the physical tick length is a CSS/print dimension. This keeps 6pt/4pt tick
 * lengths readable without reintroducing an unprojected horizontal/vertical
 * segment that would disagree with a rotated axis. */
export function screenAnnotationOutward(
  axisStart: Point,
  axisEnd: Point,
  projectedSceneCenter: Point,
  screenDirection: 'horizontal' | 'vertical',
): Point {
  const axisMidpoint = {
    x: (axisStart.x + axisEnd.x) / 2,
    y: (axisStart.y + axisEnd.y) / 2,
  };
  return screenDirection === 'horizontal'
    ? { x: axisMidpoint.x <= projectedSceneCenter.x ? -1 : 1, y: 0 }
    : { x: 0, y: axisMidpoint.y <= projectedSceneCenter.y ? -1 : 1 };
}

function projectedAxisTick(
  ctx: CanvasRenderingContext2D,
  mode: string | null | undefined,
  anchor: Point,
  axisStart: Point,
  axisEnd: Point,
  projectedSceneCenter: Point,
  screenDirection: 'horizontal' | 'vertical',
  level: 'major' | 'minor',
  ptToPx: number,
): void {
  if (!mode || mode === 'none') return;
  // Office paints 3-D ticks as screen annotations: value-axis ticks remain
  // horizontal and category-axis ticks vertical even when their projected
  // axis rule is sloped. Only the outward sign follows the selected front edge.
  const normal = screenAnnotationOutward(
    axisStart, axisEnd, projectedSceneCenter, screenDirection,
  );
  const length = projectedAxisTickLengthPx(level, ptToPx);
  const sideLength = mode === 'cross' ? length / 2 : length;
  const outer = mode === 'out' || mode === 'cross' ? sideLength : 0;
  const inner = mode === 'in' || mode === 'cross' ? sideLength : 0;
  ctx.beginPath();
  ctx.moveTo(anchor.x + normal.x * outer, anchor.y + normal.y * outer);
  ctx.lineTo(anchor.x - normal.x * inner, anchor.y - normal.y * inner);
  ctx.stroke();
}

function projectedAxisTickLengthPx(level: 'major' | 'minor', ptToPx: number): number {
  return (level === 'minor' ? 4 : 6) * ptToPx;
}

function projectedAxisTickOutwardExtentPx(
  mode: string | null | undefined,
  level: 'major' | 'minor',
  ptToPx: number,
): number {
  if (mode !== 'out' && mode !== 'cross') return 0;
  const length = projectedAxisTickLengthPx(level, ptToPx);
  return mode === 'cross' ? length / 2 : length;
}

export function projectedAxisTickLabelOffsetPx(
  mode: string | null | undefined,
  lineHidden: boolean | null | undefined,
  ptToPx: number,
  previousMinimum: number,
): number {
  if (lineHidden) return previousMinimum;
  const tickOutset = projectedAxisTickOutwardExtentPx(mode, 'major', ptToPx);
  return Math.max(previousMinimum, tickOutset + 3 * ptToPx);
}

export interface ThreeDAxisGeometry {
  axisX: number;
  axisY: number;
  depth: number;
  horizontalStart: Point;
  horizontalEnd: Point;
  verticalStart: Point;
  verticalEnd: Point;
}

/** Resolve authored axis crossing in model space before projection.
 *
 * `catAxisCrosses*` locates the category-axis rule on the numeric value axis;
 * `valAxisCrosses*` locates the value-axis rule on the category axis. Keeping
 * this choice in model space makes axis rules, ticks and labels share the same
 * homogeneous camera instead of moving a projected line after the fact.
 */
export function threeDAxisGeometry(
  chart: ChartModel,
  projection: ChartThreeDProjection,
  axis: NumericValueAxisPlan,
  categoryCount: number,
  categoryBetween: boolean,
  categoryReversed: boolean,
  orientation: 'vertical' | 'horizontal',
): ThreeDAxisGeometry {
  const { front } = projection;
  const clampFraction = (value: number): number => Math.max(0, Math.min(1, value));
  const topologyX = projection.topology.axisX === 'min' ? front.x : front.x + front.w;
  const topologyY = projection.topology.axisY === 'min' ? front.y : front.y + front.h;

  const numericCrossFraction = (): number => {
    if (chart.catAxisCrossesAt != null && Number.isFinite(chart.catAxisCrossesAt)) {
      // Percent-stacked render geometry uses the internal -100..100 scale,
      // while OOXML authors crossesAt in the underlying 0..1 value domain.
      const authored = chart.chartType.endsWith('Pct')
        ? chart.catAxisCrossesAt * 100
        : chart.catAxisCrossesAt;
      return clampFraction(axis.fraction(authored));
    }
    const crosses = chart.catAxisCrosses ?? 'autoZero';
    if (crosses === 'min') return clampFraction(axis.fraction(axis.min));
    if (crosses === 'max') return clampFraction(axis.fraction(axis.max));
    return clampFraction(axis.fraction(0));
  };
  const categoryCrossFraction = (): number | null => {
    const explicit = chart.valAxisCrossesAt;
    if (explicit != null && Number.isFinite(explicit)) {
      // Category-axis crossesAt is one-based in classic chart markup. A value
      // of 1 selects the first category position; clamp malformed/public-model
      // values before they enter camera geometry.
      return categoryPositionFraction(
        explicit - 1, categoryCount, categoryBetween, categoryReversed,
      );
    }
    const crosses = chart.valAxisCrosses;
    if (crosses !== 'min' && crosses !== 'max') return null;
    const fraction = crosses === 'max' ? 1 : 0;
    return categoryReversed ? 1 - fraction : fraction;
  };

  const numericFraction = numericCrossFraction();
  const categoryFraction = categoryCrossFraction();
  const axisX = orientation === 'horizontal'
    ? front.x + numericFraction * front.w
    : categoryFraction == null ? topologyX : front.x + categoryFraction * front.w;
  const axisY = orientation === 'vertical'
    ? front.y + front.h - numericFraction * front.h
    : categoryFraction == null ? topologyY : front.y + categoryFraction * front.h;
  const depth = projection.topology.nearDepth;
  return {
    axisX,
    axisY,
    depth,
    horizontalStart: projection.project(front.x, axisY, depth),
    horizontalEnd: projection.project(front.x + front.w, axisY, depth),
    verticalStart: projection.project(axisX, front.y + front.h, depth),
    verticalEnd: projection.project(axisX, front.y, depth),
  };
}

export function cartesianAxisTicks(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  projection: ChartThreeDProjection,
  axis: NumericValueAxisPlan,
  categoryCount: number,
  categoryBetween: boolean,
  categoryReversed: boolean,
  orientation: 'vertical' | 'horizontal',
  barFamily: boolean,
  ptToPx: number,
): void {
  const { front } = projection;
  const geometry = threeDAxisGeometry(
    chart, projection, axis, categoryCount, categoryBetween, categoryReversed, orientation,
  );
  const { axisX, axisY, depth } = geometry;
  const projectedCenter = projection.project(
    front.x + front.w / 2,
    front.y + front.h / 2,
    depth,
  );
  const valueMinorTickMark = chart.valAxisMinorTickMark ?? 'none';
  const axisTickStroke = barFamily ? threeDAxisTickStroke : threeDAxisStroke;
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    applyThreeDStroke(ctx, axisTickStroke(
      chart.valAxisLineColor, chart.valAxisLineWidthEmu, chart.valAxisLineDash,
      ptToPx,
    ));
    const valueAnchor = (value: number) => orientation === 'horizontal'
      ? projection.project(front.x + axis.fraction(value) * front.w, axisY, depth)
      : projection.project(axisX, front.y + front.h - axis.fraction(value) * front.h, depth);
    const valueStart = orientation === 'horizontal'
      ? geometry.horizontalStart : geometry.verticalStart;
    const valueEnd = orientation === 'horizontal'
      ? geometry.horizontalEnd : geometry.verticalEnd;
    for (const value of axis.majorTicks) {
      projectedAxisTick(
        ctx, chart.valAxisMajorTickMark, valueAnchor(value), valueStart, valueEnd,
        projectedCenter, orientation === 'vertical' ? 'horizontal' : 'vertical', 'major', ptToPx,
      );
    }
    for (const value of axis.minorTicks) {
      projectedAxisTick(
        ctx, valueMinorTickMark, valueAnchor(value), valueStart, valueEnd,
        projectedCenter, orientation === 'vertical' ? 'horizontal' : 'vertical', 'minor', ptToPx,
      );
    }
  }
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    applyThreeDStroke(ctx, axisTickStroke(
      chart.catAxisLineColor, chart.catAxisLineWidthEmu, chart.catAxisLineDash,
      ptToPx,
    ));
    const categoryStart = orientation === 'vertical'
      ? geometry.horizontalStart : geometry.verticalStart;
    const categoryEnd = orientation === 'vertical'
      ? geometry.horizontalEnd : geometry.verticalEnd;
    const tickSkip = Math.max(1, Math.floor(chart.catAxisTickMarkSkip ?? 1));
    const drawCategoryTick = (fraction: number): void => {
      const anchor = orientation === 'vertical'
        ? projection.project(front.x + fraction * front.w, axisY, depth)
        : projection.project(axisX, front.y + fraction * front.h, depth);
      projectedAxisTick(
        ctx, chart.catAxisMajorTickMark, anchor, categoryStart, categoryEnd,
        projectedCenter, orientation === 'vertical' ? 'vertical' : 'horizontal', 'major', ptToPx,
      );
    };
    if (barFamily) {
      const boundaries = categoryGridlineFractions(categoryCount, categoryBetween);
      for (let index = 0; index < boundaries.length; index += tickSkip) {
        drawCategoryTick(categoryReversed ? 1 - boundaries[index] : boundaries[index]);
      }
    } else {
      for (let index = 0; index < categoryCount; index += tickSkip) {
        drawCategoryTick(categoryPositionFraction(
          index, categoryCount, categoryBetween, categoryReversed,
        ));
      }
    }
    const minorUnit = chart.catAxisMinorUnit;
    if (chart.catAxisMinorTickMark && chart.catAxisMinorTickMark !== 'none'
      && minorUnit != null && Number.isFinite(minorUnit) && minorUnit > 0) {
      const majorUnit = chart.catAxisMajorUnit != null
        && Number.isFinite(chart.catAxisMajorUnit) && chart.catAxisMajorUnit > 0
        ? chart.catAxisMajorUnit
        : tickSkip;
      const boundedMinorCount = Math.min(512, Math.ceil(categoryCount / minorUnit));
      for (let index = 1; index < boundedMinorCount; index++) {
        const categoryPosition = index * minorUnit;
        if (!(categoryPosition < categoryCount)) break;
        if (Math.abs(categoryPosition / majorUnit - Math.round(categoryPosition / majorUnit)) < 1e-9) continue;
        const fraction = categoryPositionFraction(
          categoryPosition, categoryCount, categoryBetween, categoryReversed,
        );
        const anchor = orientation === 'vertical'
          ? projection.project(front.x + fraction * front.w, axisY, depth)
          : projection.project(axisX, front.y + fraction * front.h, depth);
        projectedAxisTick(
          ctx, chart.catAxisMinorTickMark, anchor, categoryStart, categoryEnd,
          projectedCenter, orientation === 'vertical' ? 'vertical' : 'horizontal', 'minor', ptToPx,
        );
      }
    }
  }
  ctx.setLineDash([]);
}
