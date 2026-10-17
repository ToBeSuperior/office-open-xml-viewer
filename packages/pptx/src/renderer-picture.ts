// PPTX picture and media element preparation and painting, including the
// media poster bitmap. Split from renderer.ts (#1857).

import type { PictureElement, MediaElement } from './types';
import {
  usesPathShade,
  trackPaintPath,
  EMU_PER_INCH,
  EMU_PER_PT as PT_TO_EMU,
  hasPreset,
  buildPresetGeometryFillPath,
  getPresetGeometryBounds,
  getCustomGeometryBounds,
  isScene3dNonIdentity,
  getCachedSvgImageByPath,
  getCachedBitmapByPath,
  getCachedDuotoneBitmapByPath,
  preferVectorBlip,
  drawImageCropped,
  metafileRasterSize,
  isOoxmlDecodedImageLimitError,
  OoxmlResourceLimitError,
  isTiffDecodeError,
  isOptionalImageCodecUnavailableError,
  paintOptionalImagePlaceholder,
} from '@silurus/ooxml-core';
import type { TiffRenderer, SvgBlobDecoder, DecodedImageTargetPlan } from '@silurus/ooxml-core';
import { drawPlayBadge } from './media-chrome';
import {
  rasterTargetOptions,
  pixelTransform,
  imagePlanKey,
  emuToPx,
  hexToRgba,
  plannedRasterOptions,
  buildCustomPath,
  pictureResourcePath,
  applyFrameTransform,
  type PlannedRasterOptions,
} from './renderer-context';
import { applyStroke, resolveShapeFill, NO_PAINT, type PreparedPaint } from './renderer-shape-fill';
import {
  buildBevelInputs,
  buildExtrusion,
  projectScene3dPaint,
  projectedEffectGeometry,
  cacheDevicePaint,
  paintWithRasterEffects,
  transformedEffectBounds,
  paintBeveledFlat,
  authoredAlignmentAnchor,
  type Sp3dLike,
  type LightRigLike,
  type EffectPaint,
} from './renderer-shape-effects';

type FetchMedia = (path: string) => Promise<Blob>;
export type PosterFetchImage = (path: string, mimeType: string) => Promise<Blob>;
const posterFetchImageByMedia = new WeakMap<FetchMedia, PosterFetchImage>();

export function posterFetchImage(fetchMedia: FetchMedia): PosterFetchImage {
  let fetchImage = posterFetchImageByMedia.get(fetchMedia);
  if (!fetchImage) {
    fetchImage = async (path, mimeType) => {
      const blob = await fetchMedia(path);
      return blob.type === mimeType ? blob : new Blob([blob], { type: mimeType });
    };
    posterFetchImageByMedia.set(fetchMedia, fetchImage);
  }
  return fetchImage;
}

// Exported for the RB1 poster decode-bomb neutralization test (asserts an
// over-budget poster is rejected before `createImageBitmap`). Not part of the
// public package surface.
export function getPosterBitmap(
  el: MediaElement,
  fetchMedia: FetchMedia,
  bitmapOwner: PosterFetchImage = posterFetchImage(fetchMedia),
  tiff?: TiffRenderer,
  target?: PlannedRasterOptions | { targetWidthPx: number; targetHeightPx: number },
  svgDecoder?: SvgBlobDecoder,
): Promise<ImageBitmap> {
  const fetchPoster = posterFetchImage(fetchMedia);
  return getCachedBitmapByPath(
    el.posterPath,
    el.posterMimeType || 'application/octet-stream',
    fetchPoster,
    { tiff, svgDecoder, ...(target ?? {}) },
    bitmapOwner,
  ).then((bitmap) => {
    if (!bitmap) throw new Error('Media poster could not be decoded');
    return bitmap;
  });
}

function paintUnavailablePicture(
  ctx: CanvasRenderingContext2D,
  el: PictureElement,
  scale: number,
  format: 'tiff' | 'pict' = 'tiff',
): void {
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);
  ctx.save();
  try {
    if (el.alpha != null) ctx.globalAlpha *= el.alpha;
    if (el.rotation !== 0 || el.flipH || el.flipV) {
      ctx.translate(x + w / 2, y + h / 2);
      ctx.rotate((el.rotation * Math.PI) / 180);
      ctx.scale(el.flipH ? -1 : 1, el.flipV ? -1 : 1);
      ctx.translate(-(x + w / 2), -(y + h / 2));
    }
    paintOptionalImagePlaceholder(ctx, format, { x, y, width: w, height: h });
  } finally {
    ctx.restore();
  }
}

/** Resolve a picture's drawable source; the returned step paints it without
 *  yielding (see {@link PreparedPaint}). */
export async function preparePicture(
  el: PictureElement,
  scale: number,
  fetchImage?: (path: string, mime: string) => Promise<Blob>,
  tiff?: TiffRenderer,
  dpr = 1,
  svgDecoder?: SvgBlobDecoder,
  imagePlan?: DecodedImageTargetPlan,
): Promise<PreparedPaint> {
  // A PICT picture is an opaque unavailable-image resource, including OOXML
  // parts labeled with this MIME. QuickDraw is unsupported: keep the authored
  // bounds with a visible diagnostic instead of invoking a platform decoder.
  // Each source validates its own container format; this renderer gate
  // validates no PICT bytes and catches no errors from supported image codecs.
  if (el.mimeType === 'image/x-pict') {
    return { paint: (ctx) => paintUnavailablePicture(ctx, el, scale, 'pict') };
  }
  // No byte source → nothing to draw (the lazy pipeline always supplies one in
  // both render modes; this guards the rare misconfiguration).
  if (!fetchImage) return NO_PAINT;
  try {
    // Prefer the vector original (Microsoft svgBlip extension); fall back to the
    // raster on any SVG decode failure. `bitmap` widens to the union of the two
    // drawable sources — both expose numeric .width/.height (used for the
    // srcRect crop below) and both are valid `ctx.drawImage` sources, so every
    // downstream path stays unchanged.
    // `imagePath` is normally a raster (PNG/JPEG), but for a pure-SVG picture
    // with no raster blip it is the SVG part itself — and `createImageBitmap`
    // (getCachedBitmapByPath) cannot rasterize SVG in every browser, so such a
    // picture must also go through the <img>-based SVG decoder (keyed by path).
    const dataIsSvg = el.mimeType === 'image/svg+xml';
    // The picture's intended draw size in points sizes any metafile raster
    // (el.width/height are EMU; /PT_TO_EMU = pt). Unused by the raster/SVG paths.
    // A cropped metafile rasterizes at its FULL picture frame (scaled up by
    // 1/(1−crop)) so the fractional crop below lands correctly; raster blips and
    // uncropped metafiles pass through unchanged. NB: getCachedBitmapByPath is keyed by
    // imagePath plus the required retained resolution. A previously decoded
    // larger variant can satisfy a smaller request; a larger/crop-magnified
    // request receives its own variant, so traversal order cannot reduce quality.
    const rasterSize = metafileRasterSize(
      el.mimeType,
      el.srcRect,
      el.width / PT_TO_EMU,
      el.height / PT_TO_EMU,
    );
    if (!rasterSize) return NO_PAINT;
    const { widthPt, heightPt } = rasterSize;
    const rawTarget = rasterTargetOptions(
      emuToPx(el.width, scale),
      emuToPx(el.height, scale),
      dpr,
      el.srcRect,
    );
    const vector = preferVectorBlip(el) || dataIsSvg;
    const target = vector
      ? rawTarget
      : imagePlan && !pixelTransform(el)
        ? plannedRasterOptions(
          imagePlan,
          imagePlanKey(pictureResourcePath(el), vector ? undefined : pixelTransform(el)),
        )
        : undefined;
    const svgPixelLimit = target && 'maxRetainedPixels' in target
      ? target.maxRetainedPixels as number
      : undefined;
    const svgOptions = {
      ...(target ? {
        targetWidthPx: target.targetWidthPx,
        targetHeightPx: target.targetHeightPx,
        ...(svgPixelLimit === undefined ? {} : { maxRetainedPixels: svgPixelLimit }),
      } : {}),
      workerDecoder: svgDecoder,
    };
    // `null` is reachable when the raster path resolves to an unsupported
    // metafile (a true EMF, or a WMF with no geometry); guarded below.
    let bitmap: ImageBitmap | HTMLImageElement | null;
    if (preferVectorBlip(el)) {
      // No crop: prefer the vector original (shared gate; see preferVectorBlip).
      // With an a:srcRect crop we skip this branch — the crop math below
      // multiplies fractional srcRect edges by the source's pixel dims, and an
      // SVG HTMLImageElement that declares only a viewBox (no intrinsic
      // width/height) reports the 300×150 default rather than its logical size,
      // so a 9-arg drawImage with a source rect samples the wrong basis. When a
      // real raster exists it has exact pixel dims that make the ECMA-376
      // §20.1.8.55 fractional crop well-defined (handled in the final branch);
      // when only the SVG exists the `dataIsSvg` branch below still draws it
      // (uncropped is the overwhelmingly common case for icons).
      try {
        bitmap = await getCachedSvgImageByPath(el.svgImagePath, fetchImage, svgOptions);
      } catch (error) {
        if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error)) throw error;
        // §20.1.8.23 duotone recolour applies only to the raster fallback — an
        // SVG vector original has no readable pixel grid (matches xlsx).
        bitmap = dataIsSvg
          ? await getCachedSvgImageByPath(el.imagePath, fetchImage, svgOptions)
          : await getCachedDuotoneBitmapByPath(el.imagePath, el.mimeType, pixelTransform(el), fetchImage, {
              widthPt,
              heightPt,
              ...(target ?? {}),
              tiff,
            });
      }
    } else if (dataIsSvg) {
      // SVG-only picture (here either because it has a crop, or — defensively —
      // because no svgImagePath was surfaced): decode through the SVG path since
      // createImageBitmap can't. A duotone on a vector picture is a rare edge
      // case left un-recoloured (no readable pixel grid), matching xlsx.
      bitmap = await getCachedSvgImageByPath(el.imagePath, fetchImage, svgOptions);
    } else {
      // §20.1.8.23 duotone recolour on the raster blip (once, at decode time,
      // cached under a colour-suffixed key). No duotone ⇒ this is exactly the
      // former `getCachedBitmapByPath` decode.
      bitmap = await getCachedDuotoneBitmapByPath(
        el.imagePath,
        el.mimeType,
        pixelTransform(el),
        fetchImage,
        { widthPt, heightPt, ...(target ?? {}), tiff },
      );
    }
    // Skip a picture whose blip is an unsupported metafile (null bitmap), the
    // same way an SVG-decode failure that also fails its raster fallback would
    // throw out of this try — here we simply return without painting.
    if (!bitmap) return NO_PAINT;
    const source = bitmap;
    return { paint: (ctx) => paintResolvedPicture(ctx, el, source, scale) };
  } catch (error) {
    return pictureFailure(el, scale, error);
  }
}

/** A picture whose source or paint failed: a missing optional TIFF codec
 *  paints the unavailable-picture placeholder, a decoded-image budget or TIFF
 *  decode error propagates, and any other broken image is skipped. */
function pictureFailure(el: PictureElement, scale: number, error: unknown): PreparedPaint {
  // Resource failures (including lazy package extraction) are fatal even when
  // ordinary unsupported/corrupt image content may produce a fallback.
  if (isOptionalImageCodecUnavailableError(error, 'tiff')) {
    return { paint: (ctx) => paintUnavailablePicture(ctx, el, scale) };
  }
  if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error) || isTiffDecodeError(error)) {
    return { ...NO_PAINT, failure: { error } };
  }
  return NO_PAINT;
}

function paintResolvedPicture(
  ctx: CanvasRenderingContext2D,
  el: PictureElement,
  bitmap: ImageBitmap | HTMLImageElement,
  scale: number,
): void {
  try {
    ctx.save();
    if (el.alpha != null) ctx.globalAlpha *= el.alpha;
    const x = emuToPx(el.x, scale);
    const y = emuToPx(el.y, scale);
    const w = emuToPx(el.width, scale);
    const h = emuToPx(el.height, scale);
    if (el.rotation !== 0 || el.flipH || el.flipV) {
      ctx.translate(x + w / 2, y + h / 2);
      ctx.rotate((el.rotation * Math.PI) / 180);
      if (el.flipH) ctx.scale(-1, 1);
      if (el.flipV) ctx.scale(1, -1);
      ctx.translate(-(x + w / 2), -(y + h / 2));
    }

    // ── srcRect sub-rectangle (ECMA-376 §20.1.8.55 a:srcRect). Resolve once via
    // the shared core helper so both the live paint and the effect aux paints
    // share identical crop coordinates. Applies to raster blips AND metafiles
    // alike: a cropped metafile was rasterized at its full picture frame above
    // (`metafileRasterSize`), so its bitmap maps to the same fractional source.
    // Trace the picture's clip silhouette (roundRect / custGeom / plain rect).
    // Shared by the clip and the border / contour strokes so the outline always
    // hugs the exact silhouette the bitmap is trimmed to. ECMA-376 §20.1.9.8:
    // a `<p:pic>` may carry `<a:custGeom>` (e.g. a laptop frame) or a roundRect
    // preset clip.
    //
    // `...Subpath` appends the silhouette as a fresh subpath of the CURRENT path
    // (no `beginPath`). Used when combining the silhouette with an enclosing
    // rect for an even-odd "outside" clip region.
    const tracePictureSilhouetteSubpath = (
      target: CanvasRenderingContext2D,
      cx: number,
      cy: number,
      cw: number,
      ch: number,
    ): void => {
      if (el.custGeom && el.custGeom.length > 0) {
        // custGeom takes priority over prstGeom (ECMA-376 §20.1.9.8).
        buildCustomPath(target, el.custGeom, cx, cy, cw, ch);
      } else if (el.prstGeom) {
        // §20.1.9.18 — the picture's preset geometry is its clip silhouette.
        // Driven by the shared preset-geometry engine (roundRect, ellipse, and
        // the other 184 presets), with the avLst adjust handles; the engine
        // substitutes each preset's declared default for any omitted guide.
        const ok = buildPresetGeometryFillPath(
          target,
          el.prstGeom,
          cx,
          cy,
          cw,
          ch,
          el.prstAdjust ?? [],
        );
        // Unknown preset name → fall back to a plain rectangle so the bitmap
        // still draws (matches the pre-generalisation rect fallback).
        if (!ok) target.rect(cx, cy, cw, ch);
      } else {
        target.rect(cx, cy, cw, ch);
      }
    };

    const tracePictureSilhouette = (
      target: CanvasRenderingContext2D,
      cx: number,
      cy: number,
      cw: number,
      ch: number,
    ): void => {
      target.beginPath();
      tracePictureSilhouetteSubpath(target, cx, cy, cw, ch);
    };

    // Apply the picture clip (roundRect / custGeom) at an arbitrary local rect.
    // The rect is parameterised so the same clip applies to the live draw
    // (x,y,w,h) and to the scene3d offscreen (0,0,w,h). A plain rectangle needs
    // no clip (the bitmap already fills the rect), so we only clip when there
    // is an actual non-rectangular / rounded silhouette.
    const applyClipAt = (
      target: CanvasRenderingContext2D,
      cx: number,
      cy: number,
      cw: number,
      ch: number,
    ): void => {
      if (el.prstGeom || (el.custGeom && el.custGeom.length > 0)) {
        tracePictureSilhouette(target, cx, cy, cw, ch);
        target.clip();
      }
    };

    // Stroke the picture's silhouette with the `<a:ln>` border and/or the
    // `<a:sp3d>` contour edge. Drawn *after* the bitmap inside `paintImageAt`,
    // so when scene3d is active the strokes are warped through the same camera
    // homography and effectLst (reflection / soft edge) re-paints them too —
    // matching PowerPoint, which applies the 3D transform and effects to the
    // framed picture as a whole.
    //
    // The two edges are split because they sit on OPPOSITE sides of the bevel
    // shading: the a:ln border is part of the FRONT FACE (PowerPoint's bevel
    // lip lights the framed picture as one surface — sample-11.pdf p6 shows
    // the lit shelf and dark rim crease ON the beige border), so it is painted
    // before applyBevelShading; the contour approximates the extruded side
    // edge OUTSIDE the front face and must stay out of the lip's distance
    // transform, so it is painted after.
    const strokeLnBorder = (
      target: CanvasRenderingContext2D,
      ox: number,
      oy: number,
      ow: number,
      oh: number,
    ): void => {
      // a:ln picture border (ECMA-376 §20.1.2.2.24). Centre-aligned stroke,
      // the Canvas default — PowerPoint draws the picture frame straddling
      // the silhouette edge.
      if (el.stroke) {
        if (usesPathShade(el.stroke?.fill)) target = trackPaintPath(target);
        target.save();
        tracePictureSilhouette(target, ox, oy, ow, oh);
        applyStroke(target, el.stroke, scale, { x: ox, y: oy, w: ow, h: oh }, el.rotation);
        target.stroke();
        target.restore();
      }
    };
    const strokeContourEdge = (
      target: CanvasRenderingContext2D,
      ox: number,
      oy: number,
      ow: number,
      oh: number,
    ): void => {
      // sp3d contour edge (ECMA-376 §20.1.5.12 `contourW` / `<a:contourClr>`).
      //    The spec's contour is the extruded 3D edge surface lit by the scene's
      //    light rig. We draw a FLAT approximation: a uniform-width outline in
      //    the contour colour, with no per-edge light-rig response. (The bevel
      //    lip itself IS shaded — applyBevelShading runs on the body before this;
      //    the contour is the thin rim OUTSIDE that bevel.) Position assumption:
      //    the contour grows OUTWARD from the front face in 3D, so we draw it as
      //    an OUTSIDE-aligned stroke (the framed edge sits just beyond the
      //    picture, not over the image). Canvas has no outside-stroke mode, so we
      //    (a) clip to the region OUTSIDE the silhouette — traced silhouette + an
      //    enclosing rect with the even-odd rule — then (b) stroke the silhouette
      //    at 2× width centred on the edge; only the outer half survives the clip.
      const sp3d = el.sp3d;
      if (sp3d && (sp3d.contourW ?? 0) > 0 && sp3d.contourClr) {
        const wPx = Math.max(0.5, (sp3d.contourW as number) * scale);
        target.save();
        // Clip to everything OUTSIDE the silhouette: trace the silhouette plus a
        // generously enlarged enclosing rect, filled even-odd, so the silhouette
        // interior is excluded from the clip region.
        target.beginPath();
        const pad = wPx * 2 + Math.max(ow, oh);
        target.rect(ox - pad, oy - pad, ow + 2 * pad, oh + 2 * pad);
        tracePictureSilhouetteSubpath(target, ox, oy, ow, oh);
        target.clip('evenodd');
        target.beginPath();
        tracePictureSilhouette(target, ox, oy, ow, oh);
        target.strokeStyle = hexToRgba(sp3d.contourClr);
        target.lineWidth = wPx * 2;
        target.setLineDash([]);
        target.stroke();
        target.restore();
      }
    };

    // scene3d (ECMA-376 §20.1.5.5 camera projection). Active only when the
    // camera actually transforms the shape; identity/front cameras fall through
    // to the normal flat draw (but sp3d bevel/extrusion still apply via the flat
    // path below). sp3d bevel/extrusion shading is wired in just after this; the
    // contour edge is the flat approximation in strokeContourEdge.
    const pictureCamera = el.scene3d && { ...el.scene3d.camera, unitsPerInch: EMU_PER_INCH * scale,
      sceneDepth: (el.sp3d?.extrusionH ?? 0) * scale };
    const scene3d = pictureCamera && isScene3dNonIdentity(
      pictureCamera, w, h, (el.sp3d?.extrusionH ?? 0) * scale,
    ) ? { ...el.scene3d, camera: pictureCamera } : null;

    // Paint JUST the (clipped, optionally cropped) bitmap body at a local rect —
    // no border/contour. Separated from the edges so the bevel shading reads the
    // body silhouette's alpha without the outside-aligned contour stroke.
    const paintBitmapBodyAt = (
      target: CanvasRenderingContext2D,
      ox: number,
      oy: number,
      ow: number,
      oh: number,
    ): void => {
      // spPr fill (§19.3.1.37): painted inside the silhouette BEHIND the blip,
      // visible through transparent pixels. Image fills need their own decode
      // and are not painted here.
      const backing = el.fill && el.fill.fillType !== 'none' && el.fill.fillType !== 'image'
        ? resolveShapeFill(el.fill, target, ox, oy, ow, oh, el.rotation, scale * PT_TO_EMU, tracePictureSilhouetteSubpath)
        : null;
      if (backing) {
        target.save();
        tracePictureSilhouette(target, ox, oy, ow, oh);
        target.fillStyle = backing;
        target.fill();
        target.restore();
      }
      target.save();
      applyClipAt(target, ox, oy, ow, oh);
      drawImageCropped(target, bitmap, el.srcRect, ox, oy, ow, oh);
      target.restore();
    };

    // Paint the full picture body (bitmap + edges) at a local rect. Used by the
    // effect aux re-paints (reflection / soft edge) and the headless flat path,
    // which don't apply bevel. Border (a:ln) and the flat sp3d contour edge are
    // stroked AFTER the bitmap so they sit on top of / around the image.
    const paintImageAt = (
      target: CanvasRenderingContext2D,
      ox: number,
      oy: number,
      ow: number,
      oh: number,
    ): void => {
      paintBitmapBodyAt(target, ox, oy, ow, oh);
      strokeLnBorder(target, ox, oy, ow, oh);
      strokeContourEdge(target, ox, oy, ow, oh);
    };

    // Front face for the bevel paths: the clipped bitmap PLUS its a:ln border.
    // The bevel's distance transform then starts at the border's outer edge and
    // its lit/shadowed lip shades the border itself, matching PowerPoint (and
    // the p:sp path, whose recursive body render also includes the stroke).
    const paintFaceAt = (
      target: CanvasRenderingContext2D,
      ox: number,
      oy: number,
      ow: number,
      oh: number,
    ): void => {
      paintBitmapBodyAt(target, ox, oy, ow, oh);
      strokeLnBorder(target, ox, oy, ow, oh);
    };

    // ── sp3d bevel (ECMA-376 §20.1.5.12 bevelT/bevelB, Phase B). Device-px lip
    //    shading baked into the body bitmap before any camera projection so the
    //    rim rides the warp. The light comes from the scene's lightRig; with no
    //    rig PowerPoint's default 3-D scene light (threePt / "t") applies.
    const tf0 = ctx.getTransform();
    const det0 = Math.abs(tf0.a * tf0.d - tf0.b * tf0.c);
    const ctxDevScale = det0 > 0 ? Math.sqrt(det0) : 1;
    const bevels = buildBevelInputs(
      el.sp3d as Sp3dLike | undefined,
      el.scene3d?.lightRig as LightRigLike | undefined,
      (el.sp3d as Sp3dLike | undefined) ? (el.sp3d as { prstMaterial?: string }).prstMaterial : undefined,
      scale,
      ctxDevScale,
      el.scene3d?.camera,
    );
    const backDepthCss = (bevels.find(bevel => bevel.bottom)?.depthPx ?? 0) / ctxDevScale;
    // Extrusion side wall only resolves under a camera that turns it visible
    // (§20.1.5.12). Computed against the active camera so a face-on view yields
    // null (no wall) and a tilted view reveals it.
    const extrusion = scene3d
      ? buildExtrusion(el.sp3d as Sp3dLike | undefined, scene3d.camera, w, h, scale, ctxDevScale)
      : null;

    // Offscreen edge margin (CSS px) for the bevel/scene3d paths: everything a
    // picture legitimately paints OUTSIDE its w×h box. Half the centre-aligned
    // border, the outside-aligned contour, the extrusion's screen sweep, plus
    // 2 px so the silhouette's antialiasing never touches the offscreen border
    // (the bevel EDT treats that border as background). See edgePadCss docs.
    const strokeHalfCss = el.stroke ? (el.stroke.width * scale) / 2 : 0;
    const contourCss = el.sp3d?.contourW ? el.sp3d.contourW * scale : 0;
    const extrusionCss = extrusion
      ? Math.hypot(extrusion.offsetX, extrusion.offsetY) / ctxDevScale
      : 0;
    const edgePadCss = strokeHalfCss + contourCss + extrusionCss + 2;

    // Draw the (clipped, optionally cropped) bitmap into a target context. This
    // is the picture "body" that the effect helpers re-paint onto aux canvases,
    // so reflections/soft edges mirror the real image rather than a flat shape.
    // When scene3d is active the body is warped through the camera homography
    // first; effects then composite over the projected result (PowerPoint
    // applies effects after the 3D transform). When sp3d carries a bevel the lip
    // shading is baked into the body before the warp (or blit, for an identity
    // camera) so it tracks the silhouette and the projection.
    const paintImage = (target: CanvasRenderingContext2D): void => {
      if (scene3d) {
        const ok = projectScene3dPaint(target, scene3d.camera, x, y, w, h, paintFaceAt, {
          bevels,
          backDepthCss,
          extrusion: extrusion ?? undefined,
          paintEdges: strokeContourEdge,
          edgePadCss,
        });
        if (ok) return;
        // Headless fallback: draw flat (no bevel).
      } else if (bevels.length > 0) {
        const ok = paintBeveledFlat(
          target,
          x,
          y,
          w,
          h,
          bevels,
          paintFaceAt,
          strokeContourEdge,
          edgePadCss,
        );
        if (ok) return;
      }
      paintImageAt(target, x, y, w, h);
    };

    // Flat unbeveled pictures retain their opaque clip mask. Rasterized 3-D
    // bodies supply full mesh geometry to the dedicated coverage pass below.
    const paintMaskAt = (
      target: CanvasRenderingContext2D,
      color: string,
      ox: number,
      oy: number,
      ow: number,
      oh: number,
    ): void => {
      target.save();
      applyClipAt(target, ox, oy, ow, oh);
      target.fillStyle = color;
      target.fillRect(ox, oy, ow, oh);
      target.restore();
    };

    // ── effectLst (§19.3.1.37 routes p:pic's spPr through CT_ShapeProperties,
    // so §20.1.8.16 effects apply to images). Same sequence as the p:sp path.
    const liveTransform = ctx.getTransform();
    const det = Math.abs(
      liveTransform.a * liveTransform.d - liveTransform.b * liveTransform.c,
    );
    const devScale = det > 0 ? Math.sqrt(det) : 1;
    const paintedPad = strokeHalfCss + contourCss;
    const clipGeometryBounds = el.custGeom && el.custGeom.length > 0
      ? getCustomGeometryBounds(el.custGeom, x, y, w, h)
      : el.prstGeom && hasPreset(el.prstGeom.toLowerCase())
        ? getPresetGeometryBounds(
            el.prstGeom.toLowerCase(),
            x,
            y,
            w,
            h,
            el.prstAdjust ?? [],
          )
        : null;
    const flatGeometryBounds = clipGeometryBounds ?? { x, y, w, h };
    const projectedGeometry = scene3d
      ? projectedEffectGeometry(
          scene3d.camera,
          liveTransform,
          x,
          y,
          w,
          h,
          edgePadCss,
          el.shadow?.algn,
          backDepthCss,
        )
      : {
          bbox: transformedEffectBounds(
            liveTransform,
            flatGeometryBounds.x - paintedPad,
            flatGeometryBounds.y - paintedPad,
            flatGeometryBounds.w + paintedPad * 2,
            flatGeometryBounds.h + paintedPad * 2,
          ),
          anchor: authoredAlignmentAnchor(
            liveTransform,
            x,
            y,
            w,
            h,
            el.shadow?.algn,
          ),
        };
    const effScale = scale * devScale; // EMU → device px
    const hasRasterEffects = Boolean(
      el.shadow || el.innerShadow || el.glow || el.softEdge || el.reflection,
    );
    const rawMask: EffectPaint = (target) => paintMaskAt(target, '#000', x, y, w, h);
    const effectBody = scene3d && hasRasterEffects
      ? cacheDevicePaint(paintImage, liveTransform, projectedGeometry.bbox, {
          w: (ctx.canvas as { width: number }).width || 0,
          h: (ctx.canvas as { height: number }).height || 0,
        })
      : paintImage;
    // ECMA-376 §20.1.5.12 includes bevels, extrusion and contour in the 3-D
    // shape. The shared effects dispatcher derives a dedicated opaque coverage
    // mask from the complete projected raster, preserving walls and mesh holes.
    // Its source color/opacity remain exclusively in the body; innerShdw owns
    // its own color, and softEdge must preserve interior fill opacity (§20.1.8).
    // Bounds/padding already come from the body's shared projectedSurface.
    const rasterMask = Boolean(scene3d) || bevels.length > 0;
    const effectMask = rasterMask ? effectBody : rawMask;
    // The caller paints this prepared picture in the slide pattern scope,
    // including transparent blip areas and any authored rotation or flip.
    paintWithRasterEffects(
      ctx,
      el,
      effectBody,
      effectMask,
      projectedGeometry.bbox,
      projectedGeometry.anchor,
      scale,
      effScale,
      liveTransform,
      true,
      rasterMask,
    );

    ctx.restore();
    // bitmap is owned by getCachedBitmapByPath's cache — do not close it here.
  } catch (error) {
    const failure = pictureFailure(el, scale, error);
    failure.paint(ctx);
    if (failure.failure) throw failure.failure.error;
  }
}

/** Resolve a media element's poster frame; the returned step paints the
 *  poster (or the plain media fill) without yielding (see {@link PreparedPaint}). */
export async function prepareMedia(
  el: MediaElement,
  scale: number,
  fetchMedia?: (path: string) => Promise<Blob>,
  skipControls?: boolean,
  bitmapOwner?: PosterFetchImage,
  tiff?: TiffRenderer,
  dpr = 1,
  svgDecoder?: SvgBlobDecoder,
  imagePlan?: DecodedImageTargetPlan,
): Promise<PreparedPaint> {
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);

  let poster: ImageBitmap | undefined;
  let posterCodecUnavailable = false;
  if (el.posterPath && fetchMedia) {
    try {
      // Poster is cached; do not close it here — it is reused across renders
      // of the same slide.
      const target = el.posterMimeType === 'image/svg+xml'
        ? rasterTargetOptions(w, h, dpr)
        : imagePlan
          ? plannedRasterOptions(imagePlan, imagePlanKey(el.posterPath))
          : undefined;
      poster = await getPosterBitmap(el, fetchMedia, bitmapOwner, tiff, target, svgDecoder);
    } catch (error) {
      if (isOptionalImageCodecUnavailableError(error, 'tiff')) {
        posterCodecUnavailable = true;
      } else {
        if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error) || isTiffDecodeError(error)) {
          return { ...NO_PAINT, failure: { error } };
        }
        // fall through to plain fill
      }
    }
  }

  return {
    paint: (ctx) => {
      ctx.save();
      applyFrameTransform(ctx, el, scale);
      if (poster) {
        ctx.drawImage(poster, x, y, w, h);
      } else {
        ctx.fillStyle = el.mediaKind === 'video' ? '#111' : '#f0f0f0';
        ctx.fillRect(x, y, w, h);
        if (posterCodecUnavailable) {
          paintOptionalImagePlaceholder(ctx, 'tiff', { x, y, width: w, height: h });
        }
      }

      if (!skipControls) drawPlayBadge(ctx, x + w / 2, y + h / 2, w, h, 'paused');
      ctx.restore();
    },
  };
}
