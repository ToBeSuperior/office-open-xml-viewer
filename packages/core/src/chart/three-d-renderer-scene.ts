// 3-D chart scene geometry and face painting: mesh and pie-outline
// projection into sorted scene faces, and face / effect-silhouette paint.
// Split from three-d-renderer.ts (#1857); reached only through the optional
// three-d renderer chunk.

import type { ChartRect } from '../types/chart.js';
import type { Fill } from '../types/common.js';
import { pptxPresetDashArray } from '../draw/dash.js';
import type { ChartThreeDProjection, ThreeDScenePoint } from './three-d.js';
import { scaleHexColor } from './material-color.js';
import {
  buildThreeDAreaStripMeshes,
  buildThreeDShapeMesh,
  type ThreeDMesh,
} from './three-d-mesh.js';
import {
  buildProjectedStrokeJunction,
  buildProjectedStrokePrimitives,
  type ProjectedStrokePoint,
} from './three-d-stroke.js';
import { buildThreeDOutlineTopology } from './three-d-outline.js';
import { resolveNativeFill as resolveFill } from '../shape/paint.js';
import type {
  Point,
  SceneFace,
  MeshOutlineStyle,
  ScenePrimitiveBudget,
} from './three-d-renderer-paint.js';

/** Automatic 3-D chart material evaluated from a real camera-space face
 * normal. This stage changes color only; it never moves or invents geometry. */
function meshMaterialFactor(
  projection: ChartThreeDProjection,
  points: readonly ThreeDScenePoint[],
): number {
  const normal = projection.cameraNormal(points);
  if (!normal) return 1;
  // Office's automatic material behaves primarily like a camera-side soft
  // light: the centre of a revolved face remains close to the authored color,
  // while the silhouette shoulder darkens. A small upper-left component keeps
  // planar caps distinguishable without turning side faces muddy.
  const light = { x: -0.20, y: 0.25, z: 1 };
  const lightLength = Math.hypot(light.x, light.y, light.z);
  const lambert = Math.max(0,
    (normal.x * light.x + normal.y * light.y + normal.z * light.z) / lightLength,
  );
  // Office's automatic 3-D material keeps the illuminated band near the
  // authored accent and lowers visible shoulders to roughly 78%. The former
  // 62% ambient term over-darkened every real mesh even though its normals and
  // occlusion were correct. Keep one directional Lambert response, but raise
  // the ambient contribution instead of applying family-specific fill hacks.
  return Math.max(0.78, Math.min(1, 0.78 + 0.24 * lambert));
}

export function polygon(ctx: CanvasRenderingContext2D, points: readonly Point[]): void {
  if (!points.length) return;
  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  for (let index = 1; index < points.length; index++) ctx.lineTo(points[index].x, points[index].y);
  ctx.closePath();
}

export const isTransparentPaint = (color: string): boolean =>
  color === 'transparent' || color === '#00000000' || color === 'rgba(0,0,0,0)';

function face(
  ctx: CanvasRenderingContext2D,
  points: readonly Point[],
  color: string,
  shade: number,
): void {
  if (isTransparentPaint(color)) return;
  polygon(ctx, points);
  ctx.fillStyle = color;
  ctx.fill();
  if (shade > 0) {
    polygon(ctx, points);
    ctx.fillStyle = `rgba(0,0,0,${shade})`;
    ctx.fill();
  }
}

/** Build an extruded area strip from one adjacent category interval.
 *
 * Excel's 3-D area is a solid ribbon, not a zero-thickness polygon.  Both
 * broad faces, the value ridge and the datum edge are projected from the same
 * model-space slab as bars, walls and axes.  Internal category seams are not
 * emitted here; callers add only the two run-end caps so adjacent strips form
 * one continuous solid.
 */
export function areaStripFaces(
  projection: ChartThreeDProjection,
  x0: number,
  x1: number,
  lower0: number,
  lower1: number,
  upper0: number,
  upper1: number,
  nearDepth: number,
  farDepth: number,
  color: string,
  capStart: boolean,
  capEnd: boolean,
): SceneFace[] {
  return buildThreeDAreaStripMeshes({
    x0, x1, lower0, lower1, upper0, upper1,
    nearDepth, farDepth, capStart, capEnd,
  }).flatMap(mesh => projectThreeDMesh(projection, mesh, color)
    .map(projected => ({ ...projected, outline: false, outlineSegments: undefined })));
}

export function shapeMeshFaces(
  projection: ChartThreeDProjection,
  shape: string,
  horizontal: boolean,
  x: number,
  y: number,
  width: number,
  height: number,
  baseCoord: number,
  endCoord: number,
  nearDepth: number,
  farDepth: number,
  color: string,
  baseScale: number,
  endScale: number,
  omitBaseCap = false,
  omitEndCap = false,
  outlineStyle?: MeshOutlineStyle,
  budget?: ScenePrimitiveBudget,
): SceneFace[] {
  if (![x, y, width, height, baseCoord, endCoord, nearDepth, farDepth].every(Number.isFinite)
    || width <= 0 || height <= 0 || baseCoord === endCoord) return [];
  const mesh = buildThreeDShapeMesh({
    shape,
    horizontal,
    crossStart: horizontal ? y : x,
    crossSize: horizontal ? height : width,
    baseCoord,
    endCoord,
    nearDepth,
    farDepth,
    baseScale,
    endScale,
    omitBaseCap,
    omitEndCap,
  });
  if (!mesh) return [];
  return projectThreeDMesh(projection, mesh, color, outlineStyle, budget);
}

/** Project one model-space solid without changing its topology.
 *
 * Mesh construction owns vertices, faces and outward winding. This stage owns
 * only camera visibility, depth and material response; Canvas paint receives
 * already projected faces. Keeping these stages separate prevents a shape
 * from being approximated by screen-space rectangles or ad-hoc fill bands.
 */
export function projectThreeDMesh(
  projection: ChartThreeDProjection,
  mesh: ThreeDMesh,
  color: string,
  outlineStyle?: MeshOutlineStyle,
  budget?: ScenePrimitiveBudget,
  outlineOnly = false,
): SceneFace[] {
  if (budget?.exceeded) return [];
  const candidates = mesh.faces.map(meshFace => {
    const scenePoints = meshFace.indices.map(index => mesh.vertices[index]);
    const facing = projection.cameraFacing(scenePoints);
    if (!facing) return { meshFace, facing, face: null as SceneFace | null };
    const points = scenePoints.map(point => projection.project(point.x, point.y, point.depth));
    const twiceArea = points.reduce((sum, point, index) => {
      const next = points[(index + 1) % points.length];
      return sum + point.x * next.y - point.y * next.x;
    }, 0);
    if (!Number.isFinite(twiceArea) || Math.abs(twiceArea) < 1e-7) {
      return { meshFace, facing, face: null as SceneFace | null };
    }
    const face: SceneFace = {
      points,
      color: scaleHexColor(color, meshMaterialFactor(projection, scenePoints)),
      paintRole: 'fill',
      shade: 0,
      cameraDepth: scenePoints.reduce(
        (sum, point) => sum + projection.cameraDepth(point.x, point.y, point.depth),
        0,
      ) / scenePoints.length,
      cameraDepths: scenePoints.map(point =>
        projection.cameraDepth(point.x, point.y, point.depth)),
      cameraWeights: scenePoints.map(point =>
        projection.cameraProjectionWeight(point.x, point.y, point.depth)),
      // Authored outlines are derived once from mesh topology below. Painting
      // each face perimeter would double shared edges and restart dash phase.
      outline: false,
    };
    return { meshFace, facing, face };
  });
  const visibleFaces = outlineOnly ? [] : candidates
    .map(candidate => candidate.face)
    .filter((face): face is SceneFace => face != null);
  if (budget) {
    if (visibleFaces.length > budget.remaining) {
      budget.exceeded = true;
      return [];
    }
    budget.remaining -= visibleFaces.length;
  }
  const outlineEdges = new Map<string, readonly [number, number]>();
  const addOutlineEdge = (startIndex: number, endIndex: number) => {
    if (startIndex === endIndex) return;
    const key = startIndex < endIndex
      ? `${startIndex}:${endIndex}` : `${endIndex}:${startIndex}`;
    if (!outlineEdges.has(key)) outlineEdges.set(key, [startIndex, endIndex]);
  };
  // Polygonal faces and smooth caps have real boundary edges. Add them once
  // for the complete visible mesh, never once per face.
  for (const candidate of candidates) {
    if (!candidate.facing || !candidate.face) continue;
    const { indices, smoothSurface, role } = candidate.meshFace;
    if (smoothSurface && role === 'side') continue;
    for (let index = 0; index < indices.length; index++) {
      addOutlineEdge(indices[index], indices[(index + 1) % indices.length]);
    }
  }

  // A smooth mesh has no authored facet seams. If an outline is authored,
  // derive only its longitudinal silhouette from visibility transitions in
  // the actual side-face topology.
  const sideCandidates = candidates
    .filter(candidate => candidate.meshFace.role === 'side')
    .sort((a, b) => (a.meshFace.segmentIndex ?? 0) - (b.meshFace.segmentIndex ?? 0));
  if (mesh.silhouetteEdges.length === sideCandidates.length && sideCandidates.length > 0) {
    for (let index = 0; index < sideCandidates.length; index++) {
      const previous = sideCandidates[(index + sideCandidates.length - 1) % sideCandidates.length];
      const current = sideCandidates[index];
      if (previous.facing === current.facing) continue;
      const visible = current.facing ? current.face : previous.face;
      if (!visible) continue;
      const [startIndex, endIndex] = mesh.silhouetteEdges[index];
      addOutlineEdge(startIndex, endIndex);
    }
  }
  const baseCapFacing = candidates.some(candidate =>
    candidate.meshFace.role === 'baseCap' && candidate.facing && candidate.face != null);
  const endCapFacing = candidates.some(candidate =>
    candidate.meshFace.role === 'endCap' && candidate.facing && candidate.face != null);
  for (const candidate of sideCandidates) {
    if (!candidate.facing || !candidate.face) continue;
    const rimEdges = [
      !baseCapFacing ? candidate.meshFace.baseRimEdge : undefined,
      !endCapFacing ? candidate.meshFace.endRimEdge : undefined,
    ].filter((edge): edge is readonly [number, number] => edge != null);
    for (const [startIndex, endIndex] of rimEdges) {
      addOutlineEdge(startIndex, endIndex);
    }
  }
  const topology = buildThreeDOutlineTopology([...outlineEdges.values()].map(([startIndex, endIndex]) => [
    mesh.vertices[startIndex], mesh.vertices[endIndex],
  ]));
  if (!outlineStyle) return visibleFaces;
  const sameScenePoint = (left: ThreeDScenePoint, right: ThreeDScenePoint) =>
    Math.hypot(left.x - right.x, left.y - right.y, left.depth - right.depth) <= 1e-9;
  const isJunction = (point: ThreeDScenePoint) =>
    topology.junctions.some(junction => sameScenePoint(junction.point, point));
  for (const scenePath of topology.paths) {
    const projectedPath = scenePath.map(point => projection.project(point.x, point.y, point.depth));
    const cameraDepths = scenePath.map(point =>
      projection.cameraDepth(point.x, point.y, point.depth));
    const cameraWeights = scenePath.map(point =>
      projection.cameraProjectionWeight(point.x, point.y, point.depth));
    const stroke = buildProjectedStrokePrimitives(projectedPath.map((point, index) => ({
      ...point,
      cameraDepth: cameraDepths[index],
      cameraWeight: cameraWeights[index],
    })), {
      width: outlineStyle.width,
      dash: outlineStyle.dash,
      lineCap: outlineStyle.cap,
      startCap: isJunction(scenePath[0]) ? 'butt' : outlineStyle.cap,
      endCap: isJunction(scenePath.at(-1)!) ? 'butt' : outlineStyle.cap,
      overlapStart: isJunction(scenePath[0]),
      overlapEnd: isJunction(scenePath.at(-1)!),
      lineJoin: outlineStyle.join,
    });
    if (!stroke || (budget && stroke.length > budget.remaining)) {
      if (budget) budget.exceeded = true;
      return [];
    }
    if (budget) budget.remaining -= stroke.length;
    visibleFaces.push(...stroke.map(primitive => {
      // Mesh edges are coplanar with their adjacent fill. A minute camera-space
      // bias prevents a later coplanar slice from erasing one side of an
      // authored circumference; genuinely nearer solids still occlude it.
      const bias = 1e-6 * Math.max(1, Math.abs(primitive.cameraDepth));
      return {
        points: primitive.points,
        color: outlineStyle.color,
        paintRole: 'outline' as const,
        shade: 0,
        cameraDepth: primitive.cameraDepth + bias,
        cameraDepths: primitive.cameraDepths?.map(depth => depth + bias),
        cameraWeights: primitive.cameraWeights,
        outline: false,
      };
    }));
  }
  for (const junction of topology.junctions) {
    const centerPoint = projection.project(
      junction.point.x, junction.point.y, junction.point.depth,
    );
    const center: ProjectedStrokePoint = {
      ...centerPoint,
      cameraDepth: projection.cameraDepth(
        junction.point.x, junction.point.y, junction.point.depth,
      ),
      cameraWeight: projection.cameraProjectionWeight(
        junction.point.x, junction.point.y, junction.point.depth,
      ),
    };
    const neighbours: ProjectedStrokePoint[] = junction.neighbours.map(point => ({
      ...projection.project(point.x, point.y, point.depth),
      cameraDepth: projection.cameraDepth(point.x, point.y, point.depth),
      cameraWeight: projection.cameraProjectionWeight(point.x, point.y, point.depth),
    }));
    const primitive = buildProjectedStrokeJunction(center, neighbours, {
      width: outlineStyle.width,
      lineJoin: outlineStyle.join,
    });
    if (!primitive) continue;
    if (budget && budget.remaining < 1) {
      budget.exceeded = true;
      return [];
    }
    if (budget) budget.remaining -= 1;
    const bias = 1e-6 * Math.max(1, Math.abs(primitive.cameraDepth));
    visibleFaces.push({
      points: primitive.points,
      color: outlineStyle.color,
      paintRole: 'outline',
      shade: 0,
      cameraDepth: primitive.cameraDepth + bias,
      cameraDepths: primitive.cameraDepths?.map(depth => depth + bias),
      cameraWeights: primitive.cameraWeights,
      outline: false,
    });
  }
  return visibleFaces;
}

interface ThreeDPieOutlineSlice {
  readonly start: number;
  readonly end: number;
  readonly segments: number;
  readonly centerX: number;
  readonly centerDepth: number;
}

/** Build a pie outline from semantic curves instead of the generic mesh-edge
 * graph. A pie has four meaningful stroke families: the visible cap rim,
 * radial separators, the visible lower wall rim, and exposed outer vertical
 * separators. Keeping them as independent continuous paths avoids high-degree
 * graph junctions that create spikes and locally doubled widths. */
export function projectThreeDPieOutline(
  projection: ChartThreeDProjection,
  slices: readonly ThreeDPieOutlineSlice[],
  centerY: number,
  radius: number,
  thickness: number,
  style: MeshOutlineStyle,
  budget: ScenePrimitiveBudget,
): SceneFace[] {
  if (!slices.length || budget.exceeded) return [];
  // `threeDPieSliceAngles` stores each clockwise slice as an ascending model
  // interval, so source order runs from high angles toward low angles. Curved
  // paths must be traversed in geometric angle order; concatenating source
  // order inserts long chords between slices and paints those chords again as
  // radial separators—the severe doubled lines seen in enlarged pies.
  const orderedSlices = [...slices].sort((left, right) => left.start - right.start);
  const topY = centerY - thickness / 2;
  const bottomY = centerY + thickness / 2;
  const reference = orderedSlices[0];
  const capY = projection.cameraDepth(reference.centerX, topY, reference.centerDepth)
    >= projection.cameraDepth(reference.centerX, bottomY, reference.centerDepth)
    ? topY : bottomY;
  const lowerY = capY === topY ? bottomY : topY;
  const pointAt = (slice: ThreeDPieOutlineSlice, angle: number, y: number): ThreeDScenePoint => ({
    x: slice.centerX + Math.cos(angle) * radius,
    y,
    depth: slice.centerDepth + Math.sin(angle) * radius / projection.modelDepth,
  });
  const centerAt = (slice: ThreeDPieOutlineSlice): ThreeDScenePoint => ({
    x: slice.centerX, y: capY, depth: slice.centerDepth,
  });
  const result: SceneFace[] = [];
  const addPath = (scenePath: readonly ThreeDScenePoint[], closed = false): void => {
    if (scenePath.length < 2 || budget.exceeded) return;
    const first = scenePath[0];
    const last = scenePath.at(-1)!;
    const alreadyClosed = Math.hypot(
      first.x - last.x, first.y - last.y, first.depth - last.depth,
    ) <= 1e-9;
    const points = closed && !alreadyClosed ? [...scenePath, first] : [...scenePath];
    const projected: ProjectedStrokePoint[] = points.map(point => ({
      ...projection.project(point.x, point.y, point.depth),
      cameraDepth: projection.cameraDepth(point.x, point.y, point.depth),
      cameraWeight: projection.cameraProjectionWeight(point.x, point.y, point.depth),
    }));
    const primitives = buildProjectedStrokePrimitives(projected, {
      width: style.width,
      dash: style.dash,
      lineCap: style.cap,
      startCap: 'butt',
      endCap: 'butt',
      lineJoin: style.join,
    });
    if (!primitives || primitives.length > budget.remaining) {
      budget.exceeded = true;
      return;
    }
    budget.remaining -= primitives.length;
    for (const primitive of primitives) {
      const bias = 1e-6 * Math.max(1, Math.abs(primitive.cameraDepth));
      result.push({
        points: primitive.points,
        color: style.color,
        paintRole: 'outline',
        shade: 0,
        cameraDepth: primitive.cameraDepth + bias,
        cameraDepths: primitive.cameraDepths?.map(depth => depth + bias),
        cameraWeights: primitive.cameraWeights,
        outline: false,
      });
    }
  };

  const capRim: ThreeDScenePoint[] = [];
  const lowerRuns: ThreeDScenePoint[][] = [];
  let lowerRun: ThreeDScenePoint[] = [];
  const visibleBoundaryAngles = new Set<string>();
  const angleKey = (angle: number) => {
    const turn = Math.PI * 2;
    const normalized = ((angle % turn) + turn) % turn;
    return (normalized < 1e-8 || turn - normalized < 1e-8 ? 0 : normalized).toFixed(8);
  };
  const boundarySlices = new Map<string, { slice: ThreeDPieOutlineSlice; angle: number }>();
  for (const slice of orderedSlices) {
    boundarySlices.set(angleKey(slice.start), { slice, angle: slice.start });
    boundarySlices.set(angleKey(slice.end), { slice, angle: slice.end });
    for (let segment = 0; segment <= slice.segments; segment++) {
      const angle = slice.start + (slice.end - slice.start) * segment / slice.segments;
      const capPoint = pointAt(slice, angle, capY);
      if (!capRim.length || Math.hypot(
        capRim.at(-1)!.x - capPoint.x,
        capRim.at(-1)!.y - capPoint.y,
        capRim.at(-1)!.depth - capPoint.depth,
      ) > 1e-9) capRim.push(capPoint);
      if (segment === slice.segments) continue;
      const nextAngle = slice.start + (slice.end - slice.start) * (segment + 1)
        / slice.segments;
      const topStart = pointAt(slice, angle, topY);
      const bottomStart = pointAt(slice, angle, bottomY);
      const bottomEnd = pointAt(slice, nextAngle, bottomY);
      const topEnd = pointAt(slice, nextAngle, topY);
      const sideVisible = projection.cameraFacing([
        topStart, bottomStart, bottomEnd, topEnd,
      ]);
      if (sideVisible) {
        const lowerStart = pointAt(slice, angle, lowerY);
        const lowerEnd = pointAt(slice, nextAngle, lowerY);
        if (!lowerRun.length) lowerRun.push(lowerStart);
        lowerRun.push(lowerEnd);
        visibleBoundaryAngles.add(angleKey(angle));
        visibleBoundaryAngles.add(angleKey(nextAngle));
      } else if (lowerRun.length) {
        lowerRuns.push(lowerRun);
        lowerRun = [];
      }
    }
  }
  if (lowerRun.length) lowerRuns.push(lowerRun);
  if (lowerRuns.length > 1) {
    const first = lowerRuns[0];
    const last = lowerRuns.at(-1)!;
    const firstPoint = first[0];
    const lastPoint = last.at(-1)!;
    if (Math.hypot(
      firstPoint.x - lastPoint.x,
      firstPoint.y - lastPoint.y,
      firstPoint.depth - lastPoint.depth,
    ) <= 1e-9) {
      lowerRuns[0] = [...last, ...first.slice(1)];
      lowerRuns.pop();
    }
  }

  addPath(capRim, true);
  const verticalAngles = new Map<string, { slice: ThreeDPieOutlineSlice; angle: number }>();
  for (const { slice, angle } of boundarySlices.values()) {
    addPath([centerAt(slice), pointAt(slice, angle, capY)]);
    if (visibleBoundaryAngles.has(angleKey(angle))) {
      verticalAngles.set(angleKey(angle), { slice, angle });
    }
  }
  for (const run of lowerRuns) {
    addPath(run);
    // The visible cylindrical wall ends at two silhouette generators. They
    // are not necessarily slice boundaries, but they must connect the upper
    // and lower rims; omitting them leaves the two short black stubs visible
    // at the left and right extremes of the pie.
    for (const endpoint of [run[0], run.at(-1)!]) {
      const angle = Math.atan2(
        (endpoint.depth - reference.centerDepth) * projection.modelDepth,
        endpoint.x - reference.centerX,
      );
      verticalAngles.set(angleKey(angle), { slice: reference, angle });
    }
  }
  for (const { slice, angle } of verticalAngles.values()) {
    addPath([pointAt(slice, angle, capY), pointAt(slice, angle, lowerY)]);
  }
  return result;
}

export function paintSceneFace(ctx: CanvasRenderingContext2D, item: SceneFace): void {
  if (item.paint === null) return;
  if (item.paint !== undefined) {
    polygon(ctx, item.points);
    ctx.fillStyle = item.paint;
    ctx.fill();
  } else {
    face(ctx, item.points, item.color, item.shade);
  }
  if (!item.outline) return;
  ctx.strokeStyle = item.outlineColor ?? 'rgba(0,0,0,0.42)';
  ctx.lineWidth = item.outlineWidth ?? 0.75;
  ctx.setLineDash(pptxPresetDashArray(item.outlineDash ?? 'solid', ctx.lineWidth));
  ctx.lineCap = item.outlineCap ?? 'butt';
  ctx.lineJoin = item.outlineJoin ?? 'miter';
  if (item.outline) {
    polygon(ctx, item.points);
    ctx.stroke();
  }
  ctx.setLineDash([]);
}

/** Paint one combined projected silhouette so a native Canvas effect is cast
 * once per datum, not once per visible mesh face. Null paint stays absent, but
 * outline polygons still participate when a datum is authored noFill. */
export function paintSceneEffectSilhouette(
  ctx: CanvasRenderingContext2D,
  faces: readonly SceneFace[],
): void {
  const visible = faces.filter(face => face.paint !== null && face.points.length >= 3);
  if (!visible.length) return;
  ctx.beginPath();
  for (const face of visible) {
    ctx.moveTo(face.points[0].x, face.points[0].y);
    for (let index = 1; index < face.points.length; index++) {
      ctx.lineTo(face.points[index].x, face.points[index].y);
    }
    ctx.closePath();
  }
  ctx.fillStyle = '#000000';
  ctx.fill();
}

/** Resolve one authored fill recipe once for a complete projected datum, then
 * share the Canvas paint across its faces. This keeps gradient-stop work
 * O(datums × stops), never O(faces × stops). */
export function applyScenePaint(
  ctx: CanvasRenderingContext2D,
  faces: SceneFace[],
  role: SceneFace['paintRole'],
  fill: Fill | null | undefined,
  shapeRotationDeg: number,
): void {
  if (fill === undefined) return;
  const targets = faces.filter(face => face.paintRole === role);
  if (!targets.length) return;
  if (fill === null) {
    for (const face of targets) face.paint = null;
    return;
  }
  const points = targets.flatMap(face => face.points);
  const minX = Math.min(...points.map(point => point.x));
  const maxX = Math.max(...points.map(point => point.x));
  const minY = Math.min(...points.map(point => point.y));
  const maxY = Math.max(...points.map(point => point.y));
  const paint = resolveFill(
    fill, ctx, minX, minY, maxX - minX, maxY - minY, shapeRotationDeg,
  );
  for (const face of targets) face.paint = paint;
}

export function paintThreeDTooManyDataPoints(
  ctx: CanvasRenderingContext2D,
  rect: ChartRect,
): void {
  ctx.fillStyle = '#888';
  ctx.font = '12px sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('(too many data points)', rect.x + rect.w / 2, rect.y + rect.h / 2);
}
