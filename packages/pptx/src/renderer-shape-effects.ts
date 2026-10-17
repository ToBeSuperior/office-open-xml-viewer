// PPTX shape effect painting: shadow/glow context state, effect transforms,
// raster effect compositing (effect mask), and scene3d bevel / extrusion
// projection. Split from renderer.ts (#1857).

import { ownValue } from '@silurus/ooxml-core/internal/own-key';
import type { ShapeElement, Shadow, Glow } from './types';
import {
  measurePaintBounds,
  withInheritedPatternScope,
  applyOuterShadow,
  applyInnerShadow,
  createRasterCoveragePaint,
  rasterEffectCoverageBounds,
  applySoftEdge,
  applyReflection,
  computeScene3dQuad,
  drawProjected,
  expandProjectedQuad,
  projectQuadPoint,
  createAuxCanvas,
  applyBevelShading,
  applyExtrusion,
  computeDepthOffset,
  lightDirFromRig,
  materialClass,
  clampCanvasSize,
  MAX_CANVAS_DIMENSION,
  MAX_CANVAS_AREA,
} from '@silurus/ooxml-core';
import type {
  CameraInput,
  Vec2,
  BevelInput,
  ExtrusionInput,
  BevelRegion,
} from '@silurus/ooxml-core';
import { emuToPx, hexToRgba } from './renderer-context';

function applyShadow(ctx: CanvasRenderingContext2D, shadow: Shadow | null, scale: number) {
  if (!shadow) return;
  const dirRad = (shadow.dir * Math.PI) / 180;
  const dist = emuToPx(shadow.dist, scale);
  ctx.shadowColor = hexToRgba(shadow.color, shadow.alpha);
  ctx.shadowBlur = 0;
  ctx.shadowOffsetX = Math.cos(dirRad) * dist;
  ctx.shadowOffsetY = Math.sin(dirRad) * dist;
}

/**
 * ECMA-376 §20.1.8.17 — apply glow as a coloured halo. The Canvas shadow
 * primitive is a fine fit: zero offset + the glow's blur radius produces a
 * symmetric coloured blur centred on every drawn pixel.
 */
export function applyGlow(ctx: CanvasRenderingContext2D, glow: Glow | null | undefined, scale: number) {
  if (!glow) return;
  ctx.shadowColor = hexToRgba(glow.color, glow.alpha);
  ctx.shadowBlur = emuToPx(glow.radius, scale);
  ctx.shadowOffsetX = 0;
  ctx.shadowOffsetY = 0;
}

export function clearShadow(ctx: CanvasRenderingContext2D) {
  ctx.shadowColor = 'transparent';
  ctx.shadowBlur = 0;
  ctx.shadowOffsetX = 0;
  ctx.shadowOffsetY = 0;
}

/**
 * DrawingML camera projection consumes the shape-local face after xfrm flips;
 * the ordinary 2-D rotation remains outside the projected face. Keeping this
 * split explicit prevents a post-projection flip from mirroring the camera
 * footprint itself.
 */
export function splitScene3dShapeTransform(
  rotation: number,
  flipH: boolean,
  flipV: boolean,
): { outerRotation: number; localFlipH: boolean; localFlipV: boolean } {
  return {
    outerRotation: rotation,
    localFlipH: flipH,
    localFlipV: flipV,
  };
}

interface EffectTransform {
  a: number;
  b: number;
  c: number;
  d: number;
  e: number;
  f: number;
}

/**
 * Axis-aligned device-pixel bounds of a shape after the current Canvas
 * transform. DrawingML group transforms are already folded into that matrix;
 * effect surfaces must therefore be cropped from these transformed bounds,
 * not from the shape's group-local coordinates.
 */
export function transformedEffectBounds(
  transform: EffectTransform,
  x: number,
  y: number,
  w: number,
  h: number,
): { x: number; y: number; w: number; h: number } {
  const points = [
    { x, y },
    { x: x + w, y },
    { x, y: y + h },
    { x: x + w, y: y + h },
  ].map((point) => ({
    x: transform.a * point.x + transform.c * point.y + transform.e,
    y: transform.b * point.x + transform.d * point.y + transform.f,
  }));
  const minX = Math.min(...points.map((point) => point.x));
  const minY = Math.min(...points.map((point) => point.y));
  const maxX = Math.max(...points.map((point) => point.x));
  const maxY = Math.max(...points.map((point) => point.y));
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
}

export type EffectPaint = (target: CanvasRenderingContext2D) => void;

interface RasterEffects {
  shadow?: ShapeElement['shadow'];
  innerShadow?: ShapeElement['innerShadow'];
  glow?: ShapeElement['glow'];
  softEdge?: ShapeElement['softEdge'];
  reflection?: ShapeElement['reflection'];
}

const ALIGNMENT_UV: Record<NonNullable<Shadow['algn']>, readonly [number, number]> = {
  tl: [0, 0], t: [0.5, 0], tr: [1, 0],
  l: [0, 0.5], ctr: [0.5, 0.5], r: [1, 0.5],
  bl: [0, 1], b: [0.5, 1], br: [1, 1],
};

function transformPoint(transform: EffectTransform, x: number, y: number): [number, number] {
  return [
    transform.a * x + transform.c * y + transform.e,
    transform.b * x + transform.d * y + transform.f,
  ];
}

export function authoredAlignmentAnchor(
  transform: EffectTransform,
  x: number,
  y: number,
  w: number,
  h: number,
  algn: Shadow['algn'],
): [number, number] {
  const [u, v] = ownValue(ALIGNMENT_UV, algn ?? 'b') ?? ALIGNMENT_UV.b;
  return transformPoint(transform, x + u * w, y + v * h);
}

/** One surface plane for body projection, effect bounds/anchors and masks.
 * MS-OI29500 §2.1.1268 places bevelB at the extrusion base. Keep the
 * existing planar compositor's face selection, including its winding test;
 * effect dispatch must not silently move that surface back to z=0. Depth comes
 * from a prepared bottom bevel or wireframe input; this planar painter still
 * does not construct an unbeveled solid's back mesh. Quantize padding here too,
 * so cache/effect bounds include the same device-pixel margin as the raster.
 */
function projectedSurface(
  camera: CameraInput, w: number, h: number, backDepthCss = 0,
  edgePadCss = 0, devScale = 1,
) {
  // The selected face and its effect footprint belong to one solid, including
  // when a caller supplies a rear plane without pre-binding sceneDepth.
  if (backDepthCss > (camera.sceneDepth ?? 0)) camera = { ...camera, sceneDepth: backDepthCss };
  const front = computeScene3dQuad(camera, w, h);
  const [p0, p1, , p3] = front.corners;
  const backFacing = (p1.x - p0.x) * (p3.y - p0.y) - (p1.y - p0.y) * (p3.x - p0.x) < 0;
  const quad = backFacing && backDepthCss > 0
    ? computeScene3dQuad(camera, w, h, backDepthCss) : front;
  let padDev = Math.max(0, Math.ceil(edgePadCss * devScale));
  const expanded = padDev > 0
    ? expandProjectedQuad(quad.corners, padDev / devScale / w, padDev / devScale / h) : null;
  if (!expanded) padDev = 0;
  return { face: quad.corners, corners: expanded ?? quad.corners,
    backFacing, padDev, padCss: padDev / devScale };
}

export function projectedEffectGeometry(
  camera: CameraInput,
  transform: EffectTransform,
  x: number,
  y: number,
  w: number,
  h: number,
  edgePadCss: number,
  algn: Shadow['algn'],
  backDepthCss = 0,
): { bbox: { x: number; y: number; w: number; h: number }; anchor: [number, number] } {
  const det = Math.abs(transform.a * transform.d - transform.b * transform.c);
  const devScale = det > 0 ? Math.sqrt(det) : 1;
  const { face, corners: painted } = projectedSurface(camera, w, h, backDepthCss, edgePadCss, devScale);
  const devicePoints = painted.map((point) => transformPoint(transform, x + point.x, y + point.y));
  const xs = devicePoints.map(([px]) => px);
  const ys = devicePoints.map(([, py]) => py);
  const minX = Math.min(...xs);
  const minY = Math.min(...ys);
  const maxX = Math.max(...xs);
  const maxY = Math.max(...ys);
  const [u, v] = ownValue(ALIGNMENT_UV, algn ?? 'b') ?? ALIGNMENT_UV.b;
  const projectedAnchor = projectQuadPoint(face, u, v);
  return {
    bbox: { x: minX, y: minY, w: maxX - minX, h: maxY - minY },
    anchor: projectedAnchor
      ? transformPoint(transform, x + projectedAnchor.x, y + projectedAnchor.y)
      : authoredAlignmentAnchor(transform, x, y, w, h, algn),
  };
}

/** Cache an expensive projected body for this one render call only. */
export function cacheDevicePaint(
  paint: EffectPaint,
  liveTransform: EffectTransform,
  bbox: { x: number; y: number; w: number; h: number },
  viewport?: { w: number; h: number },
  source?: CanvasRenderingContext2D,
): EffectPaint {
  const x = Math.floor(bbox.x) - 1;
  const y = Math.floor(bbox.y) - 1;
  const w = Math.max(1, Math.ceil(bbox.x + bbox.w) - x + 1);
  const h = Math.max(1, Math.ceil(bbox.y + bbox.h) - y + 1);
  // This cache is an optimization, not a rendering dependency. Do not allocate
  // a fully viewport-external or oversized projected surface: replaying the source
  // projection is slower but preserves both effect reach and the shared canvas
  // safety limits without risking an OOM/DOMException.
  if (viewport && (x + w <= 0 || y + h <= 0 || x >= viewport.w || y >= viewport.h)) return paint;
  if (clampCanvasSize(w, h).clamped) return paint;
  let canvas: ReturnType<typeof createAuxCanvas> = null;
  try {
    canvas = createAuxCanvas(w, h);
  } catch {
    return paint;
  }
  const cache = canvas?.getContext('2d') as CanvasRenderingContext2D | null;
  if (!canvas || !cache) return paint;
  cache.setTransform(
    liveTransform.a,
    liveTransform.b,
    liveTransform.c,
    liveTransform.d,
    liveTransform.e - x,
    liveTransform.f - y,
  );
  if (source) withInheritedPatternScope(source, cache, () => paint(cache), { x, y });
  else paint(cache);
  const identity = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } as DOMMatrix;
  return (target) => {
    target.save();
    target.setTransform(identity);
    target.drawImage(canvas as CanvasImageSource, x, y);
    target.restore();
  };
}

export function paintWithRasterEffects(
  ctx: CanvasRenderingContext2D,
  effects: RasterEffects,
  paintBody: EffectPaint,
  paintMask: EffectPaint,
  bbox: { x: number; y: number; w: number; h: number },
  alignmentAnchor: readonly [number, number],
  scale: number,
  effectScale: number,
  liveTransform: EffectTransform,
  hasInterior = true,
  rasterMask = false,
): void {
  const deviceW = (ctx.canvas as { width: number }).width || 0;
  const deviceH = (ctx.canvas as { height: number }).height || 0;
  const haveAux = deviceW > 0 && deviceH > 0;
  const identity = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } as DOMMatrix;
  const applyLiveTransform = (target: CanvasRenderingContext2D) => target.setTransform(liveTransform);
  const body = (target: CanvasRenderingContext2D) => {
    applyLiveTransform(target);
    withInheritedPatternScope(ctx, target, () => paintBody(target));
  };
  const mask = (target: CanvasRenderingContext2D) => {
    applyLiveTransform(target);
    withInheritedPatternScope(ctx, target, () => paintMask(target));
  };
  // Projected bodies supply the complete geometry, but their RGBA opacity
  // and color must never become an effect mask. Convert once and share the
  // dedicated coverage raster across softEdge and innerShdw.
  const coverage = rasterMask && haveAux && (effects.softEdge || effects.innerShadow)
    ? createRasterCoveragePaint(mask as never, rasterEffectCoverageBounds(
        bbox, effects, effectScale, deviceW, deviceH,
      ))
    : null;
  const haveMask = !rasterMask || Boolean(coverage);
  let nativeShadowFallback = false;
  if (effects.shadow && haveAux) {
    ctx.save();
    ctx.setTransform(identity);
    nativeShadowFallback = !applyOuterShadow(
      ctx,
      body as never,
      bbox,
      effects.shadow,
      effectScale,
      deviceW,
      deviceH,
      Math.atan2(liveTransform.b, liveTransform.a) * 180 / Math.PI,
      alignmentAnchor,
    );
    ctx.restore();
  } else if (effects.shadow) {
    nativeShadowFallback = true;
  }
  if (effects.reflection && haveAux) {
    ctx.save();
    ctx.setTransform(identity);
    applyReflection(ctx, body as never, bbox, effects.reflection, effectScale, deviceW, deviceH);
    ctx.restore();
  }
  if (nativeShadowFallback) applyShadow(ctx, effects.shadow ?? null, scale);
  else if (effects.glow) applyGlow(ctx, effects.glow, scale);

  if (effects.softEdge && haveAux && haveMask) {
    ctx.save();
    ctx.setTransform(identity);
    applySoftEdge(
      ctx,
      body as never,
      bbox,
      effects.softEdge,
      effectScale,
      deviceW,
      deviceH,
      (coverage ?? mask) as never,
    );
    ctx.restore();
  } else {
    paintBody(ctx);
  }
  if (nativeShadowFallback || effects.glow) clearShadow(ctx);

  if (effects.innerShadow && hasInterior && haveAux && haveMask) {
    ctx.save();
    ctx.setTransform(identity);
    applyInnerShadow(
      ctx,
      (coverage ?? mask) as never,
      bbox,
      effects.innerShadow,
      effectScale,
      deviceW,
      deviceH,
    );
    ctx.restore();
  }
}


/** Local view of the parsed `<a:sp3d>` (1:1 with the Sp3d TS type). */
export interface Sp3dLike {
  extrusionH?: number;
  extrusionClr?: string;
  bevelT?: { w: number; h: number; prst: string };
  bevelB?: { w: number; h: number; prst: string };
}
/** Local view of `<a:lightRig>` (1:1 with the LightRig TS type). */
export interface LightRigLike {
  rig: string;
  dir: string;
  /** Optional `<a:rot lat lon rev>` (§20.1.5.11) — rev rotates the key azimuth. */
  rot?: { lat: number; lon: number; rev: number };
}

/**
 * Build the bevel-shading inputs for an sp3d, in DEVICE px. ECMA-376 §20.1.5.12
 * `bevelT`/`bevelB`. Returns an empty list when there is no bevel to shade.
 *
 * The bevel band/height are EMU (§20.1.5.3); convert to device px via the EMU→
 * CSS-px `scale` and the canvas `devScale`, because the shading reads/writes the
 * device-resolution offscreen. The light vector comes from the scene's lightRig
 * (§20.1.5.9); when absent we default to the OOXML default rig "threePt"/"t"
 * (PowerPoint's default 3-D scene light) so a bevel with no explicit rig still
 * picks up a top key light rather than rendering flat.
 */
export function buildBevelInputs(
  sp3d: Sp3dLike | undefined,
  lightRig: LightRigLike | undefined,
  prstMaterial: string | undefined,
  scale: number,
  devScale: number,
  camera?: CameraInput,
): BevelInput[] {
  if (!sp3d) return [];
  const rig = lightRig?.rig ?? 'threePt';
  const dir = lightRig?.dir ?? 't';
  // ECMA §§20.1.5.9/.11 and MS-OI29500 §2.1.1274(b): preserve explicit
  // zero rot as an override of dir. Core rotates the entire rig (including
  // flat-face illumination), with revolution about the tilted rig axis.
  const light = lightDirFromRig(rig, dir, lightRig?.rot);
  const material = materialClass(prstMaterial);
  const emuToDev = scale * devScale;
  const out: BevelInput[] = [];
  if (sp3d.bevelT && sp3d.bevelT.w > 0 && sp3d.bevelT.h > 0) {
    out.push({
      widthPx: sp3d.bevelT.w * emuToDev,
      heightPx: sp3d.bevelT.h * emuToDev,
      prst: sp3d.bevelT.prst || 'circle',
      material,
      light,
      camera,
      wireframeLinePx: devScale,
    });
  }
  if (sp3d.bevelB && sp3d.bevelB.w > 0 && sp3d.bevelB.h > 0) {
    out.push({
      widthPx: sp3d.bevelB.w * emuToDev,
      heightPx: sp3d.bevelB.h * emuToDev,
      prst: sp3d.bevelB.prst || 'circle',
      material,
      light,
      bottom: true,
      camera,
      depthPx: (sp3d.extrusionH ?? 0) * emuToDev,
      wireframeLinePx: devScale,
    });
  }
  // Wireframe has geometry even without a bevel; it must not retain a solid
  // face just because there is no bevel band to send to the compositor.
  if (material === 'legacyWireframe') {
    for (const bottom of [false, true]) {
      if (!out.some(bevel => Boolean(bevel.bottom) === bottom))
        out.push({ widthPx: 0, heightPx: 0, prst: 'circle', material, light, camera,
          bottom, depthPx: bottom ? (sp3d.extrusionH ?? 0) * emuToDev : 0,
          wireframeLinePx: devScale });
    }
  }
  return out;
}

/**
 * Build the extrusion side-wall input (device px) for an sp3d, or null when
 * there is no extrusion or the camera is face-on (the wall would be invisible).
 * ECMA-376 §20.1.5.12 `extrusionH` / `extrusionClr`. The side-wall colour is the
 * explicit `extrusionClr` when present, else a darkened mid-grey stand-in (the
 * spec leaves the default to the renderer; PowerPoint uses a shaded copy of the
 * face — a neutral dark wall is the conservative Phase B choice).
 */
export function buildExtrusion(
  sp3d: Sp3dLike | undefined,
  camera: CameraInput,
  w: number,
  h: number,
  scale: number,
  devScale: number,
): ExtrusionInput | null {
  if (!sp3d || !sp3d.extrusionH || sp3d.extrusionH <= 0) return null;
  const depthDev = sp3d.extrusionH * scale * devScale;
  // Depth offset is computed in the device-px box (w·devScale × h·devScale) so
  // it matches the device-resolution offscreen the side wall is baked into.
  const off = computeDepthOffset(
    { ...camera, unitsPerInch: (camera.unitsPerInch ?? 96) * devScale,
      sceneDepth: (camera.sceneDepth ?? 0) * devScale },
    w * devScale, h * devScale, depthDev,
  );
  if (Math.hypot(off.x, off.y) < 0.75) return null;
  let rgb: [number, number, number] = [64, 64, 64];
  if (sp3d.extrusionClr) {
    const hex = sp3d.extrusionClr.replace('#', '');
    if (hex.length >= 6) {
      rgb = [
        parseInt(hex.slice(0, 2), 16),
        parseInt(hex.slice(2, 4), 16),
        parseInt(hex.slice(4, 6), 16),
      ];
    }
  }
  return { offsetX: off.x, offsetY: off.y, rgb };
}

/**
 * Project a shape/picture body through a DrawingML 3D camera (scene3d, Phase A),
 * with optional sp3d bevel shading baked in BEFORE the projection (Phase B).
 *
 * ECMA-376 §20.1.5.5: the camera acts in the shape's *local* space, ahead of
 * the shape's 2D placement. We therefore render the body — normally drawn at
 * (x,y,w,h) in `target`'s coordinate space — into a separate device-resolution
 * offscreen canvas sized to the w×h box, then warp that bitmap onto `target`
 * with `drawProjected` (a piecewise-affine homography). Because `target` already
 * carries the element's 2D rotation/flip transform when this runs, the
 * projection composes *inside* that transform, exactly matching "scene3d first,
 * then the 2D xfrm" ordering.
 *
 * Phase B: when `bevels` is non-empty, the bevel lip shading is baked into the
 * offscreen bitmap (`applyBevelShading`) right after `paintBody` and before the
 * warp, so the lit/shadowed rim rides the same camera homography as the body
 * (matching PowerPoint, which shades the 3-D solid then projects it).
 *
 * `paintBody(octx, ox, oy, ow, oh)` paints the body (clip + fill/stroke/image)
 * into the offscreen context at the given local rect. Returns true if it
 * projected, false when no offscreen canvas is available (headless) — callers
 * then fall back to painting directly.
 */
interface Project3dOpts {
  /** Preserve overflow of the separate, unshaded text pass. Body material
   * and effect rasters keep their established domains. */
  measureOverflow?: boolean;
  /** Extrusion-base depth of the prepared bottom surface, in CSS px. */
  backDepthCss?: number;
  /** Text front faces retain the slide pattern frame before camera projection. */
  preservePatternFrame?: boolean;
  /** Bevel lips to bake into the body before the warp (§20.1.5.12 bevelT/B). */
  bevels?: BevelInput[];
  /** Extrusion side-wall to bake in before the bevel (§20.1.5.12 extrusionH). */
  extrusion?: ExtrusionInput;
  /**
   * Paint edges that sit OUTSIDE the beveled front face (the sp3d contour, an
   * outside-aligned rim of the extruded side edge) into the offscreen AFTER the
   * bevel shading, in the same local rect. The a:ln border belongs to the front
   * face itself — PowerPoint's bevel lights the framed picture as one surface —
   * so the border is painted inside `paintBody`, BEFORE the bevel, and only the
   * contour goes here (it must not contaminate the lip's distance transform).
   */
  paintEdges?: (octx: CanvasRenderingContext2D, ox: number, oy: number, ow: number, oh: number) => void;
  /**
   * Margin (CSS px) added on every side of the offscreen around the shape's
   * w×h box. The box alone CLIPS everything a shape legitimately paints past
   * it: the outer half of the centre-aligned a:ln border, the outside-aligned
   * contour, and the extrusion sweep. For a silhouette that touches its box
   * (an inscribed ellipse touches it at all four apices) that clip cuts the
   * rim along the straight box edge — the "sliced ellipse" slide-6 bug. The
   * margin also keeps the silhouette away from the offscreen border so the
   * bevel's distance transform (which treats out-of-bounds as background)
   * measures the true silhouette distance instead of the canvas edge.
   */
  edgePadCss?: number;
}

/** Map the source slide/device frame into a padded, shape-local effect raster.
 * Bevels blit that raster; scene3d warps it through the camera after painting. */
function patternFrameToLocalRaster(
  tf: DOMMatrix,
  devScale: number,
  padDev: number,
  x: number,
  y: number,
  padYDev = padDev,
) {
  const det = tf.a * tf.d - tf.b * tf.c;
  if (Math.abs(det) <= 1e-12) return undefined;
  const ia = tf.d / det;
  const ib = -tf.b / det;
  const ic = -tf.c / det;
  const id = tf.a / det;
  const ie = -(ia * tf.e + ic * tf.f);
  const iff = -(ib * tf.e + id * tf.f);
  return {
    a: devScale * ia, b: devScale * ib,
    c: devScale * ic, d: devScale * id,
    e: padDev + devScale * (ie - x),
    f: padYDev + devScale * (iff - y),
  };
}

export function projectScene3dPaint(
  target: CanvasRenderingContext2D,
  camera: CameraInput,
  x: number,
  y: number,
  w: number,
  h: number,
  paintBody: (octx: CanvasRenderingContext2D, ox: number, oy: number, ow: number, oh: number) => void,
  opts: Project3dOpts = {},
): boolean {
  if (w <= 0 || h <= 0) return false;
  // Device scale folded into target's current transform, so the offscreen is
  // rasterised at the same pixel density as the live canvas (no blur on HiDPI).
  const tf = target.getTransform();
  const det = Math.abs(tf.a * tf.d - tf.b * tf.c);
  const devScale = det > 0 ? Math.sqrt(det) : 1;

  // Edge margin (see Project3dOpts.edgePadCss). Quantised to whole device px so
  // the body lands on the same pixel grid as the unpadded layout, then mapped
  // onto the destination by extrapolating the camera quad through its own
  // homography. The camera is defined for the authored w×h face, so it must not
  // be recomputed from the padded raster size. Degenerate extrapolation → pad 0.
  const backDepth = opts.backDepthCss ?? 0;
  const { face, corners: paddedCorners, backFacing, padDev } =
    projectedSurface(camera, w, h, backDepth, opts.edgePadCss, devScale);
  let quadCorners = paddedCorners;
  let originX = padDev, originY = padDev;
  let ow = Math.max(1, Math.ceil(w * devScale) + 2 * padDev);
  let oh = Math.max(1, Math.ceil(h * devScale) + 2 * padDev);
  if (opts.measureOverflow) {
    // ECMA-376 §21.1.2.1.1 bodyPr defaults horzOverflow/vertOverflow to
    // overflow; §21.1.2.1.2 noAutofit does not scale text into its box.
    // Measure actual paint in the local authored frame, including insets,
    // vertical/stacked text and glyph overhangs. The camera still belongs to
    // w×h: growing that face would change its centre, lens and text placement.
    const ink = measurePaintBounds(target, octx => paintBody(octx, 0, 0, w, h));
    if (ink) {
      // Independent axis bounds keep a long, thin text run O(ink area), rather
      // than allocating a square with its longest overflow as four-side padding.
      // One device pixel encloses raster AA; integer origins preserve the grid.
      const x0 = Math.min(-padDev, Math.floor(ink.x * devScale) - 1);
      const y0 = Math.min(-padDev, Math.floor(ink.y * devScale) - 1);
      const x1 = Math.max(ow - padDev, Math.ceil((ink.x + ink.w) * devScale) + 1);
      const y1 = Math.max(oh - padDev, Math.ceil((ink.y + ink.h) * devScale) + 1);
      const expanded = [[x0, y0], [x1, y0], [x1, y1], [x0, y1]]
        .map(([px, py]) => projectQuadPoint(face, px / devScale / w, py / devScale / h));
      // An authored degenerate face has no drawable homography, as in
      // drawProjected itself. Otherwise a null extrapolation crosses the
      // projection horizon: retain the flat fallback, not an unbounded warp.
      if (expanded.some(point => !point)) return projectQuadPoint(face, .5, .5) === null;
      quadCorners = expanded as [Vec2, Vec2, Vec2, Vec2];
      originX = -x0; originY = -y0;
      ow = x1 - x0; oh = y1 - y0;
    }
  }
  if (opts.measureOverflow && (!Number.isFinite(ow * oh)
    || ow > MAX_CANVAS_DIMENSION || oh > MAX_CANVAS_DIMENSION
    || ow * oh > MAX_CANVAS_AREA)) return false;
  const aux = createAuxCanvas(ow, oh);
  if (!aux) return false;
  const octx = aux.getContext('2d') as CanvasRenderingContext2D | null;
  if (!octx) return false;

  // Paint in authored local coordinates; only the raster origin is displaced.
  octx.save();
  octx.scale(devScale, devScale);
  octx.translate(originX / devScale, originY / devScale);
  // PowerPoint PDF rasterises projected pattern text as a front-face image
  // before the camera warp. Its 8 pt tile stays in the slide frame: moving an
  // otherwise identical dnDiag text shape by 4 pt changes the embedded image's
  // colour phase, while camera rotation and extrusion retain that phase. Map
  // the live slide root into this padded local raster before painting.
  if (opts.preservePatternFrame) {
    const sourceToAux = patternFrameToLocalRaster(tf, devScale, originX, x, y, originY);
    withInheritedPatternScope(target, octx, () => paintBody(octx, 0, 0, w, h), undefined, sourceToAux);
  } else {
    paintBody(octx, 0, 0, w, h);
  }
  octx.restore();

  // The body silhouette occupies the offscreen's inner box (device px):
  // originX/Y + [0, w/h·devScale] (the a:ln border, painted inside
  // paintBody, extends ±half-line-width past it; the pad reserves that room).
  // Restricting the bevel/extrusion distance-transform + sweep to this box grown
  // by the effect's reach (perf: A3) skips the transparent pad border. Equivalent
  // to the whole offscreen because a band pixel is within `bandPx` of the
  // silhouette (lip inside it) and a wall pixel within `|offset|` of it. When the
  // offscreen is already shape-tight (the usual case) the region ≈ the whole
  // canvas and this is a harmless no-op. See BevelRegion.
  const bodyDevW = Math.ceil(w * devScale);
  const bodyDevH = Math.ceil(h * devScale);
  const regionFor = (reachPx: number): BevelRegion => ({
    x: originX - reachPx,
    y: originY - reachPx,
    w: bodyDevW + 2 * reachPx,
    h: bodyDevH + 2 * reachPx,
  });

  // Extrusion side wall first (it sits behind/around the front face), then the
  // bevel lip on the front-face silhouette (§20.1.5.12). Both are baked into the
  // offscreen so they ride the camera warp.
  const extrusion = opts.extrusion && backFacing && backDepth > 0
    ? { ...opts.extrusion, offsetX: -opts.extrusion.offsetX, offsetY: -opts.extrusion.offsetY }
    : opts.extrusion;
  const wireframe = opts.bevels?.some(bevel => bevel.material === 'legacyWireframe');
  const shadeBackBeforeWall = backFacing && opts.bevels?.some(bevel => bevel.bottom);
  const paintWall = () => {
    if (!extrusion || wireframe) return;
    const reach = Math.ceil(Math.hypot(extrusion.offsetX, extrusion.offsetY)) + 2;
    applyExtrusion(octx, extrusion, regionFor(reach));
  };
  if (!shadeBackBeforeWall) paintWall();
  // Bake bevel lip shading into the body bitmap (device-px band) before the
  // warp so the lit/shadowed rim rides the camera homography (§20.1.5.12).
  if (opts.bevels && opts.bevels.length > 0) {
    for (const bevel of opts.bevels) {
      const wireframeDepth = wireframe && extrusion
        ? { x: extrusion.offsetX, y: extrusion.offsetY } : undefined;
      const reach = bevel.widthPx + (wireframeDepth ? Math.hypot(wireframeDepth.x, wireframeDepth.y) : 0);
      applyBevelShading(octx, wireframeDepth ? { ...bevel, wireframeDepth } : bevel,
        regionFor(Math.ceil(reach) + 2));
    }
  }

  // The exposed back lip is measured from the back-face silhouette, not from
  // the outside of a wall swept toward the front. Shade it before that sweep;
  // keep the established front-face compositing order above.
  if (shadeBackBeforeWall) paintWall();

  // Post-bevel edges (the contour rim) on top of the beveled body, still
  // pre-projection. The a:ln border is painted inside paintBody (see
  // Project3dOpts.paintEdges).
  if (opts.paintEdges) {
    octx.save();
    octx.scale(devScale, devScale);
    octx.translate(originX / devScale, originY / devScale);
    opts.paintEdges(octx, 0, 0, w, h);
    octx.restore();
  }

  // Offset the (possibly pad-expanded) quad to the element's (x,y).
  const corners = quadCorners.map((c) => ({ x: x + c.x, y: y + c.y })) as [
    Vec2,
    Vec2,
    Vec2,
    Vec2,
  ];
  // drawProjected runs in target's CSS-px space (the device scale lives in the
  // target transform), warping the ow×oh offscreen onto the CSS-space quad.
  drawProjected(aux as unknown as CanvasImageSource, target, ow, oh, corners);
  return true;
}

/**
 * Bake bevel shading + edges into a body that is NOT camera-projected (identity
 * camera, e.g. orthographicFront). Renders the body to a device-px offscreen,
 * shades the bevel, paints the post-bevel edges (contour), then blits the
 * offscreen back at (x,y). Falls back to false (caller paints flat) when no
 * offscreen is available.
 *
 * `edgePadCss` grows the offscreen by a margin on every side (same rationale
 * as Project3dOpts.edgePadCss): without it the box-sized offscreen clips the
 * outer half of the centre-aligned border wherever the silhouette touches its
 * bounding box, and the bevel's distance transform reads the offscreen border
 * as the silhouette edge there — both showed up as the straight-cut rim on the
 * slide-6 ellipse apices.
 */
export function paintBeveledFlat(
  target: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  bevels: BevelInput[],
  paintBody: (octx: CanvasRenderingContext2D, ox: number, oy: number, ow: number, oh: number) => void,
  paintEdges?: (octx: CanvasRenderingContext2D, ox: number, oy: number, ow: number, oh: number) => void,
  edgePadCss = 0,
): boolean {
  if (w <= 0 || h <= 0 || bevels.length === 0) return false;
  const tf = target.getTransform();
  const det = Math.abs(tf.a * tf.d - tf.b * tf.c);
  const devScale = det > 0 ? Math.sqrt(det) : 1;
  // Whole-device-px margin so the body stays on the unpadded pixel grid.
  const padDev = Math.max(0, Math.ceil(edgePadCss * devScale));
  const padCss = padDev / devScale;
  const ow = Math.max(1, Math.ceil(w * devScale) + 2 * padDev);
  const oh = Math.max(1, Math.ceil(h * devScale) + 2 * padDev);
  const aux = createAuxCanvas(ow, oh);
  if (!aux) return false;
  const octx = aux.getContext('2d') as CanvasRenderingContext2D | null;
  if (!octx) return false;

  octx.save();
  octx.scale(devScale, devScale);
  octx.translate(padCss, padCss);
  // PowerPoint keeps the pattern's slide axes inside an orthographic bevel.
  // The body is repainted in this shape-local device canvas, then blitted
  // through the shape CTM. Transfer the complete slide frame, including the
  // shape's rotation/flip, before resolving its pattern fill.
  const sourceToAux = patternFrameToLocalRaster(tf, devScale, padDev, x, y);
  withInheritedPatternScope(target, octx, () => paintBody(octx, 0, 0, w, h), undefined, sourceToAux);
  octx.restore();
  // Restrict the bevel distance-transform to the body's inner box grown by the
  // band width (perf: A3 — skips the transparent pad border). Equivalent because
  // the lip sits within `bandPx` of the silhouette, which fills this box. No-op
  // when the offscreen is already shape-tight. See BevelRegion / projectScene3dPaint.
  const bodyDevW = Math.ceil(w * devScale);
  const bodyDevH = Math.ceil(h * devScale);
  for (const bevel of bevels) {
    const reach = Math.ceil(bevel.widthPx) + 2;
    applyBevelShading(octx, bevel, {
      x: padDev - reach,
      y: padDev - reach,
      w: bodyDevW + 2 * reach,
      h: bodyDevH + 2 * reach,
    });
  }
  if (paintEdges) {
    octx.save();
    octx.scale(devScale, devScale);
    octx.translate(padCss, padCss);
    paintEdges(octx, 0, 0, w, h);
    octx.restore();
  }
  // Blit the device-px offscreen back into target's CSS-px space, 1:1 in
  // device pixels, with the margin hanging past (x,y) on every side.
  target.drawImage(
    aux as unknown as CanvasImageSource,
    x - padCss,
    y - padCss,
    ow / devScale,
    oh / devScale,
  );
  return true;
}
