// PPTX fill and outline painting: shape fill resolution, slide backgrounds,
// tiled and shape image fills, compound lines, and stroke application.
// Split from renderer.ts (#1857).

import type { PictureElement, Fill, TileInfo, Stroke, ImageFill } from './types';
import {
  usesPathShade,
  currentStrokeBounds,
  resolveFill as resolveFillCore,
  applyStroke as applyStrokeCore,
  EMU_PER_PX,
  EMU_PER_PT as PT_TO_EMU,
  createAuxCanvas,
  MAX_CANVAS_DIMENSION,
  MAX_CANVAS_AREA,
  getCachedDuotoneBitmapByPath,
  inspectCachedRasterSource,
  drawImageCropped,
  isOoxmlDecodedImageLimitError,
  OoxmlResourceLimitError,
  isTiffDecodeError,
  isOptionalImageCodecUnavailableError,
  paintOptionalImagePlaceholder,
} from '@silurus/ooxml-core';
import type {
  TiffRenderer,
  SvgBlobDecoder,
  SvgImageSource,
  DecodedImageTargetPlan,
} from '@silurus/ooxml-core';
import {
  pixelTransform,
  imagePlanKey,
  emuToPx,
  hexToRgba,
  plannedRasterOptions,
  type PreparedShapeFill,
} from './renderer-context';

/** Context-aware fill resolver that creates a CanvasGradient for gradient
 * fills and a CanvasPattern for preset pattern fills. */
export function resolveShapeFill(
  fill: Fill | null,
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  shapeRotationDeg = 0,
  patternPtToUserUnits = 4 / 3,
  outline?: import('@silurus/ooxml-core').FillOutline,
  paintBounds?: { x: number; y: number; w: number; h: number },
): string | CanvasGradient | CanvasPattern | null {
  return resolveFillCore(fill, ctx, x, y, w, h, shapeRotationDeg, patternPtToUserUnits, undefined, outline, paintBounds);
}

// ===== Element renderers =====

/**
 * One slide-paint step whose asynchronous inputs (decoded images, posters) are
 * already resolved, so `paint` issues its canvas calls synchronously.
 *
 * Why the slide is painted in one task: Chrome records Canvas2D calls and
 * rasterizes the pending recording when it flushes the canvas, which happens
 * at the end of any task that drew to a displayed canvas (for example when a
 * frame is produced while the renderer awaits a decode). Rasterization of an
 * antialiased path is not independent of that chunking: in Chrome 153's
 * software canvas, an antialiased concave path covers a few edge pixels
 * differently once the same unflushed recording already holds six such
 * "slow" paths or clips (cc's `kMinNumberOfSlowPathsForMSAA` heuristic). A
 * paint that awaits between draw calls therefore rasterizes differently
 * depending on when decodes settle, so an unchanged renderer and document could
 * produce different pixels on a first paint than on a repaint (issue #1580).
 * Resolving every input first and then painting without yielding keeps the
 * whole slide in one task, which also stops a half-painted slide from being
 * presented while decodes are still pending.
 *
 * `failure` carries an error that must propagate (decoded-image budget, TIFF
 * decode). It is rethrown right after `paint`, at this step's position in the
 * slide's paint order, so the canvas holds exactly what the former
 * interleaved paint left before throwing.
 */
export interface PreparedPaint {
  readonly paint: (ctx: CanvasRenderingContext2D) => void;
  readonly failure?: { readonly error: unknown };
}

export const NO_PAINT: PreparedPaint = { paint: () => {} };

export async function prepareBackground(
  fill: Fill | null,
  canvasW: number,
  canvasH: number,
  scale: number,
  fetchImage?: (path: string, mime: string) => Promise<Blob>,
  tiff?: TiffRenderer,
  svgDecoder?: SvgBlobDecoder,
  imagePlan?: DecodedImageTargetPlan,
): Promise<PreparedPaint> {
  if (fill && fill.fillType === 'image') {
    // ECMA-376 §20.1.8.14 — image (blipFill) background. Paint an opaque white
    // base first so a partially transparent image (alphaModFix) composites over
    // white, and so a decode failure still leaves a defined background.
    const paintBase = (ctx: CanvasRenderingContext2D) => {
      ctx.fillStyle = '#FFFFFF';
      ctx.fillRect(0, 0, canvasW, canvasH);
    };
    // The lazy pipeline always emits imagePath + mimeType for blip fills; bail
    // to the white base if either the path or the byte source is missing.
    if (!fill.imagePath || !fill.mimeType || !fetchImage) return { paint: paintBase };
    const failed = (error: unknown): PreparedPaint => {
      if (isOptionalImageCodecUnavailableError(error, 'tiff')) {
        return {
          paint: (ctx) => {
            paintBase(ctx);
            paintOptionalImagePlaceholder(ctx, 'tiff', {
              x: 0, y: 0, width: canvasW, height: canvasH,
            });
          },
        };
      }
      if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error) || isTiffDecodeError(error)) {
        return { paint: paintBase, failure: { error } };
      }
      // Decode failed — the white base remains as the fallback.
      return { paint: paintBase };
    };
    let bitmap: Awaited<ReturnType<typeof getCachedDuotoneBitmapByPath>>;
    let sourceInspection: Awaited<ReturnType<typeof inspectCachedRasterSource>> | undefined;
    try {
      // Size the metafile raster from the fill box (canvasW/H are CSS px;
      // scale is px-per-EMU, so px/scale = EMU, /PT_TO_EMU = pt).
      // §20.1.8.23 duotone recolour on the raster blip (issue #889): route
      // through the shared duotone cache (keyed by path + colours). No duotone ⇒
      // this is exactly the former `getCachedBitmapByPath` decode, byte-identical.
      const planned = imagePlan && !pixelTransform(fill)
        ? plannedRasterOptions(imagePlan, imagePlanKey(fill.imagePath, pixelTransform(fill)))
        : undefined;
      sourceInspection = fill.tile
        ? await inspectCachedRasterSource(fill.imagePath, fill.mimeType, fetchImage)
        : undefined;
      bitmap = await getCachedDuotoneBitmapByPath(
        fill.imagePath,
        fill.mimeType,
        pixelTransform(fill),
        fetchImage,
        {
          widthPt: canvasW / scale / PT_TO_EMU,
          heightPt: canvasH / scale / PT_TO_EMU,
          ...(planned ?? {}),
          tiff,
          svgDecoder,
        },
      );
    } catch (error) {
      return failed(error);
    }
    // A null bitmap (unsupported metafile, e.g. true EMF) → keep the white
    // base as the fallback, exactly like a decode failure.
    if (!bitmap) return { paint: paintBase };
    const source = bitmap;
    return {
      paint: (ctx) => {
        paintBase(ctx);
        try {
          const fr = fill.fillRect ?? {};
          const l = fr.l ?? 0;
          const t = fr.t ?? 0;
          const r = fr.r ?? 0;
          const b = fr.b ?? 0;
          const dx = l * canvasW;
          const dy = t * canvasH;
          const dw = canvasW * (1 - l - r);
          const dh = canvasH * (1 - t - b);
          ctx.save();
          // Clip to the slide rectangle so overscan (negative insets) or tile
          // bleed is cropped at the slide edge rather than spilling onto
          // neighbouring content.
          ctx.beginPath();
          ctx.rect(0, 0, canvasW, canvasH);
          ctx.clip();
          if (fill.alpha != null) ctx.globalAlpha = fill.alpha;
          if (fill.tile) {
            // §20.1.8.58 — tiled placement: repeat the blip at its native size.
            paintTiledBackground(
              ctx,
              source,
              fill.tile,
              canvasW,
              canvasH,
              scale,
              fill.srcRect,
              sourceInspection?.dimensions ?? undefined,
            );
          } else {
            // §20.1.8.56 stretch into the destination rect from the §20.1.8.30
            // fillRect insets. l/t are left/top insets, r/b are right/bottom
            // insets, so the destination spans [l, 1-r] × [t, 1-b] of the box;
            // negative edges overscan past the box.
            drawImageCropped(ctx, source, fill.srcRect, dx, dy, dw, dh);
          }
          ctx.restore();
        } catch (error) {
          if (isOptionalImageCodecUnavailableError(error, 'tiff')) {
            paintOptionalImagePlaceholder(ctx, 'tiff', {
              x: 0, y: 0, width: canvasW, height: canvasH,
            });
            return;
          }
          if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error) || isTiffDecodeError(error)) throw error;
          // Paint failed — the white base painted above remains as the fallback.
        }
      },
    };
  }
  return {
    paint: (ctx) => {
      const bg = resolveShapeFill(fill, ctx, 0, 0, canvasW, canvasH, 0, scale * PT_TO_EMU);
      ctx.fillStyle = bg ?? '#FFFFFF';
      ctx.fillRect(0, 0, canvasW, canvasH);
    },
  };
}

/**
 * Compute the anchor (origin) of the first tile inside the fill box for a
 * given §20.1.8.41 ST_RectAlignment value. The returned point is where the
 * tile grid is registered; tx/ty then shift it further. The pattern is then
 * phase-shifted so a tile edge passes through this anchor.
 *
 * - `tl` → box origin (0,0); the first tile's top-left sits at the box origin.
 * - `ctr` → box centre; a tile is centred in the box.
 * - `br` → box bottom-right corner; a tile's bottom-right sits there.
 * etc. The anchor is expressed as the position (px) of the tile's top-left
 * for the corresponding alignment, before tx/ty.
 */
export function tileAnchorOffset(
  algn: string,
  boxW: number,
  boxH: number,
  tileW: number,
  tileH: number,
): { ax: number; ay: number } {
  // Horizontal: tl/l/bl = left, t/ctr/b = centre, tr/r/br = right.
  let ax: number;
  if (algn === 't' || algn === 'ctr' || algn === 'b') {
    ax = (boxW - tileW) / 2;
  } else if (algn === 'tr' || algn === 'r' || algn === 'br') {
    ax = boxW - tileW;
  } else {
    ax = 0; // tl, l, bl (and any unknown) anchor left.
  }
  // Vertical: tl/t/tr = top, l/ctr/r = middle, bl/b/br = bottom.
  let ay: number;
  if (algn === 'l' || algn === 'ctr' || algn === 'r') {
    ay = (boxH - tileH) / 2;
  } else if (algn === 'bl' || algn === 'b' || algn === 'br') {
    ay = boxH - tileH;
  } else {
    ay = 0; // tl, t, tr (and any unknown) anchor top.
  }
  return { ax, ay };
}

/** Native pixel extent of the §20.1.8.55 source rectangle duplicated by tile
 * mode. The logical rectangle can be larger than the bitmap for negative
 * outsets; that extra extent remains transparent when painted. */
export function tileSourceExtent(
  width: number,
  height: number,
  srcRect?: PictureElement['srcRect'],
): { width: number; height: number } {
  return {
    width: width * (srcRect ? 1 - srcRect.l - srcRect.r : 1),
    height: height * (srcRect ? 1 - srcRect.t - srcRect.b : 1),
  };
}

/**
 * Paint a tiled blip background (ECMA-376 §20.1.8.58 CT_TileInfoProperties).
 *
 * The blip repeats at its native pixel size (interpreted at its authored DPI)
 * scaled by sx/sy and the slide `scale`. `flip` mirrors alternate tiles, which
 * we pre-compose into a 2×2 "super-tile" so a plain `repeat` pattern reproduces
 * the mirror cadence. `algn` registers the grid against a box corner/edge and
 * tx/ty add a further EMU offset. The whole pattern is phase-shifted via a
 * `DOMMatrix` translate on the pattern transform.
 *
 * The caller has already clipped to the fill box and set globalAlpha. Background
 * callers omit `dpi` and preserve their prior 96-DPI behavior. `coverage` can
 * extend beyond this alignment box when the caller counter-rotates the brush;
 * expanding coverage must not shift the tile registration or phase.
 */
function paintTiledBackground(
  ctx: CanvasRenderingContext2D,
  bitmap: SvgImageSource,
  tile: TileInfo,
  canvasW: number,
  canvasH: number,
  scale: number,
  srcRect?: PictureElement['srcRect'],
  intrinsicSize?: Readonly<{ width: number; height: number }>,
  dpi = 96,
  coverage = { x: 0, y: 0, w: canvasW, h: canvasH },
): boolean {
  // Native tile size in slide px: image px → EMU @96dpi → × sx/sy → × scale.
  // §20.1.8.55 applies before the fill mode: for tile mode the cropped source
  // rectangle is the content duplicated (Annex L.4.8.4.3), so its logical
  // width/height determine the native tile extent. Negative edges retain their
  // transparent outset through drawImageCropped below.
  // Established shape compatibility defaults remain local to the PPTX shape
  // renderer; the shared model preserves absence for spec-first chart paint.
  const source = tileSourceExtent(
    intrinsicSize?.width ?? bitmap.width,
    intrinsicSize?.height ?? bitmap.height,
    srcRect,
  );
  // A blip without an authored DPI uses the 96-DPI EMU-per-px basis.
  const pixelToEmu = EMU_PER_PX * 96 / dpi;
  const tileW = source.width * pixelToEmu * (tile.sx ?? 1) * scale;
  const tileH = source.height * pixelToEmu * (tile.sy ?? 1) * scale;
  if (!(tileW > 0) || !(tileH > 0)) return false;

  const flipX = tile.flip === 'x' || tile.flip === 'xy';
  const flipY = tile.flip === 'y' || tile.flip === 'xy';

  // Build the repeating cell. Without flip it is one tile; with flip it is a
  // 2×2 block whose neighbours are mirrored, so the seam between repeats is a
  // mirror line (PowerPoint's tile-flip behaviour).
  const cellW = tileW * (flipX ? 2 : 1);
  const cellH = tileH * (flipY ? 2 : 1);
  // createAuxCanvas rounds each requested axis up. Charge the actual allocation,
  // not the fractional geometry, so two individually-safe fractions cannot
  // cross the retained scratch-area limit after independent ceil operations.
  if (!Number.isFinite(cellW) || !Number.isFinite(cellH)) return false;
  // Preserve the previous rasterization for cells that fit: draw at the
  // authored size into the ceil-sized canvas and repeat those integer pixels.
  // Normalizing even fractional cells to their exact extent changes sampling
  // and seam placement. Only cells previously omitted use the bounded path.
  let scratchW = Math.max(1, Math.ceil(cellW));
  let scratchH = Math.max(1, Math.ceil(cellH));
  const oversized = scratchW > MAX_CANVAS_DIMENSION || scratchH > MAX_CANVAS_DIMENSION
    || scratchW > Math.floor(MAX_CANVAS_AREA / scratchH);
  if (oversized) {
    // Raster size is resource policy, not a DrawingML geometry limit. Avoid
    // upsampling small sources; the pattern transform restores the authored
    // extent (§20.1.8.58), including mirrored-neighbour cadence.
    scratchW = Math.min(MAX_CANVAS_DIMENSION,
      Math.max(1, Math.ceil(Math.min(cellW, source.width * (flipX ? 2 : 1)))));
    scratchH = Math.min(MAX_CANVAS_DIMENSION,
      Math.max(1, Math.ceil(Math.min(cellH, source.height * (flipY ? 2 : 1)))));
    if (scratchW * scratchH > MAX_CANVAS_AREA) {
      const reduction = Math.sqrt(MAX_CANVAS_AREA / (scratchW * scratchH));
      scratchW = Math.max(1, Math.floor(scratchW * reduction));
      scratchH = Math.max(1, Math.floor(scratchH * reduction));
    }
  }
  const rasterScaleX = oversized ? scratchW / cellW : 1;
  const rasterScaleY = oversized ? scratchH / cellH : 1;
  const aux = createAuxCanvas(scratchW, scratchH);
  if (!aux) return false;
  const actx = aux.getContext('2d') as CanvasRenderingContext2D | null;
  if (!actx) return false;

  actx.scale(rasterScaleX, rasterScaleY);
  const drawCell = (cx: number, cy: number, mx: boolean, my: boolean) => {
    actx.save();
    actx.translate(cx + (mx ? tileW : 0), cy + (my ? tileH : 0));
    actx.scale(mx ? -1 : 1, my ? -1 : 1);
    drawImageCropped(actx, bitmap, srcRect, 0, 0, tileW, tileH);
    actx.restore();
  };
  // Top-left tile is always un-mirrored; mirror the X/Y neighbours.
  drawCell(0, 0, false, false);
  if (flipX) drawCell(tileW, 0, true, false);
  if (flipY) drawCell(0, tileH, false, true);
  if (flipX && flipY) drawCell(tileW, tileH, true, true);

  const pattern = ctx.createPattern(aux as unknown as CanvasImageSource, 'repeat');
  if (!pattern) return false;

  // Phase: register the grid against the alignment anchor, then add tx/ty.
  // PowerPoint registers an omitted algn at the top-left. CT_TileInfoProperties
  // itself declares no schema default, so keep this host observation here rather
  // than normalizing it in the shared OOXML parser.
  const { ax, ay } = tileAnchorOffset(tile.algn ?? 'tl', canvasW, canvasH, tileW, tileH);
  const px = ax + emuToPx(tile.tx ?? 0, scale);
  const py = ay + emuToPx(tile.ty ?? 0, scale);
  // `setTransform` exists where DOMMatrix is available (browser + skia-canvas).
  // The translate aligns a cell origin with (px, py); `repeat` covers the box.
  if (typeof pattern.setTransform === 'function' && typeof DOMMatrix !== 'undefined') {
    pattern.setTransform(new DOMMatrix().translateSelf(px, py).scaleSelf(1 / rasterScaleX, 1 / rasterScaleY));
    ctx.fillStyle = pattern;
    ctx.fillRect(coverage.x, coverage.y, coverage.w, coverage.h);
  } else {
    // Fallback: bake the phase into the fill origin by translating the context.
    ctx.save();
    ctx.translate(px, py);
    ctx.scale(1 / rasterScaleX, 1 / rasterScaleY);
    ctx.fillStyle = pattern;
    ctx.fillRect((coverage.x - px) * rasterScaleX, (coverage.y - py) * rasterScaleY,
      coverage.w * rasterScaleX, coverage.h * rasterScaleY);
    ctx.restore();
  }
  return true;
}

/** Key of one prepared (decoded + pixel-transformed) shape blip. The decode
 * applies the CT_Blip effects (§20.1.8.13) in document order, so the key must
 * carry the complete ordered transform, not only the source path: two shapes
 * that share a PNG but differ in grayscl/biLevel/clrChange/lum/duotone need
 * distinct prepared images. `imagePlanKey` is the same path + ordered-effect
 * key the decode cache uses. */
export function shapeFillKey(fill: ImageFill): string {
  return JSON.stringify([
    fill.svgImagePath ?? null,
    imagePlanKey(fill.imagePath, pixelTransform(fill)),
  ]);
}

/** EG_FillModeProperties (§20.1.8.14) is an optional choice of `tile` or
 * `stretch` with no schema default, and the two are mutually exclusive. An
 * ordinary shape paints a blip only when exactly one mode is authored: an
 * omitted mode has no specified placement and no Office evidence here, and a
 * tile+stretch pair is schema-invalid, so both fail closed (nothing is
 * painted, as before shape image fills were supported). This is the same rule
 * as the core chart picture-fill path. */
export function shapeImageFillModeIsPaintable(fill: ImageFill): boolean {
  return (fill.tile != null) !== (fill.stretch === true);
}

/** Paint a predecoded shape blip into the current geometry path. */
export function paintPreparedShapeImageFill(
  ctx: CanvasRenderingContext2D,
  fill: ImageFill,
  prepared: PreparedShapeFill | null | undefined,
  bounds: Readonly<{ x: number; y: number; w: number; h: number }>,
  scale: number,
  evenOdd = false,
  shapeTransform: Readonly<{ rotation: number; flipH: boolean; flipV: boolean }> = { rotation: 0, flipH: false, flipV: false },
): boolean {
  if (!prepared || !(bounds.w > 0) || !(bounds.h > 0)) return false;
  if (!shapeImageFillModeIsPaintable(fill)) return false;
  const { x, y, w, h } = bounds;
  ctx.save();
  try {
    ctx.clip(evenOdd ? 'evenodd' : 'nonzero');
    if (fill.alpha != null) ctx.globalAlpha *= Math.max(0, Math.min(1, fill.alpha));
    // CT_BlipFillProperties (§20.1.8.14): keep the clipped shape geometry,
    // but cancel its rotation about the fill-box centre. A reflected local
    // frame reverses the rotation sign (R F R = F); the reflection itself is
    // retained because rotWithShape controls rotation, not flips.
    let fillRotation = 0;
    if (fill.rotWithShape === false && shapeTransform.rotation !== 0) {
      const sign = shapeTransform.flipH !== shapeTransform.flipV ? 1 : -1;
      ctx.translate(x + w / 2, y + h / 2);
      fillRotation = sign * shapeTransform.rotation * Math.PI / 180;
      ctx.rotate(fillRotation);
      ctx.translate(-(x + w / 2), -(y + h / 2));
    }
    if (fill.tile) {
      // §20.1.8.58 repeats across the entire clipped shape. Counter-rotation
      // changes the clip's extent in the brush frame, not the tile's phase or
      // its alignment box. Cover the inverse-rotated box about the same centre.
      const cos = Math.abs(Math.cos(fillRotation));
      const sin = Math.abs(Math.sin(fillRotation));
      const coverW = w * cos + h * sin;
      const coverH = w * sin + h * cos;
      ctx.translate(x, y);
      return paintTiledBackground(
        ctx, prepared.image, fill.tile, w, h, scale, fill.srcRect, prepared.intrinsicSize,
        fill.dpi && fill.dpi > 0 ? fill.dpi : 96,
        { x: (w - coverW) / 2, y: (h - coverH) / 2, w: coverW, h: coverH },
      );
    }
    const rect = fill.fillRect ?? {};
    const dx = x + (rect.l ?? 0) * w;
    const dy = y + (rect.t ?? 0) * h;
    const dw = (1 - (rect.l ?? 0) - (rect.r ?? 0)) * w;
    const dh = (1 - (rect.t ?? 0) - (rect.b ?? 0)) * h;
    if (!(dw > 0) || !(dh > 0)) return false;
    drawImageCropped(ctx, prepared.image, fill.srcRect, dx, dy, dw, dh);
    return true;
  } finally {
    ctx.restore();
  }
}

// ===== Table renderer =====

/**
 * Re-stroke a straight connector under one of the ECMA-376 §20.1.8.42
 * compound-line styles. Each sub-line gets its own offset along the
 * perpendicular and its own thickness; widths are taken from Office's
 * usual interpretation:
 *
 *   dbl       → two equal-thickness lines, each w/3, gap w/3
 *   thinThick → thin (w/4) + gap (w/4) + thick (w/2)
 *   thickThin → thick (w/2) + gap (w/4) + thin (w/4)
 *   tri       → three lines: thin(w/5), thick(w/5*3 ≈ 3w/5), thin(w/5)
 *
 * The starting position of each sub-line is computed so the *outer envelope*
 * stays w wide — i.e. centred on the original geometry. The base stroke
 * already drew at the centre line; we erase it (destination-out) and then
 * paint the compound sub-lines on top so dash/headEnd remain consistent.
 */
export function drawCompoundLine(
  ctx: CanvasRenderingContext2D,
  start: { x: number; y: number },
  end: { x: number; y: number },
  stroke: Stroke,
  cmpd: string,
  scale: number,
  shapeRotationDeg: number,
): void {
  const totalW = Math.max(0.5, emuToPx(stroke.width, scale));
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const len = Math.hypot(dx, dy);
  if (len === 0) return;
  const px = -dy / len;
  const py = dx / len;

  // sub: relative position along the perpendicular axis (-1..1, with 0 being
  // the centre line) and width as a fraction of totalW.
  type Sub = { offset: number; widthFrac: number };
  let subs: Sub[];
  switch (cmpd) {
    case 'dbl':
      subs = [
        { offset: -1 / 3, widthFrac: 1 / 3 },
        { offset:  1 / 3, widthFrac: 1 / 3 },
      ];
      break;
    case 'thinThick':
      subs = [
        { offset: -3 / 8, widthFrac: 1 / 4 },
        { offset:  1 / 4, widthFrac: 1 / 2 },
      ];
      break;
    case 'thickThin':
      subs = [
        { offset: -1 / 4, widthFrac: 1 / 2 },
        { offset:  3 / 8, widthFrac: 1 / 4 },
      ];
      break;
    case 'tri':
      subs = [
        { offset: -2 / 5, widthFrac: 1 / 5 },
        { offset:  0,     widthFrac: 3 / 5 },
        { offset:  2 / 5, widthFrac: 1 / 5 },
      ];
      break;
    default:
      return;
  }

  ctx.save();
  // 1. Erase the centre line that was already drawn at full width.
  ctx.globalCompositeOperation = 'destination-out';
  ctx.strokeStyle = '#000';
  ctx.lineWidth = totalW + 0.5; // small overshoot to fully erase antialiasing fringe
  ctx.setLineDash([]);
  ctx.beginPath();
  ctx.moveTo(start.x, start.y); ctx.lineTo(end.x, end.y);
  ctx.stroke();

  // 2. Paint each sub-line.
  ctx.globalCompositeOperation = 'source-over';
  const strokePaint = stroke.fill
    ? resolveShapeFill(
        stroke.fill,
        ctx,
        Math.min(start.x, end.x),
        Math.min(start.y, end.y),
        Math.max(1, Math.abs(end.x - start.x)),
        Math.max(1, Math.abs(end.y - start.y)),
        shapeRotationDeg,
        scale * PT_TO_EMU,
      )
    : null;
  ctx.strokeStyle = strokePaint ?? hexToRgba(stroke.color);
  for (const sub of subs) {
    const ox = px * (totalW * sub.offset);
    const oy = py * (totalW * sub.offset);
    ctx.lineWidth = Math.max(0.5, totalW * sub.widthFrac);
    ctx.beginPath();
    ctx.moveTo(start.x + ox, start.y + oy);
    ctx.lineTo(end.x + ox, end.y + oy);
    if (stroke.fill && usesPathShade(stroke.fill)) {
      // Parallel compound segments have different offsets and widths. Resolve
      // after building each real segment, with the common authored shade frame.
      const paint = resolveShapeFill(stroke.fill, ctx,
        Math.min(start.x, end.x), Math.min(start.y, end.y),
        Math.max(1, Math.abs(end.x - start.x)), Math.max(1, Math.abs(end.y - start.y)),
        shapeRotationDeg, scale * PT_TO_EMU, undefined, currentStrokeBounds(ctx));
      if (paint) ctx.strokeStyle = paint;
    }
    ctx.stroke();
  }
  ctx.restore();
}

export function applyStroke(
  ctx: CanvasRenderingContext2D,
  stroke: Stroke | null,
  scale: number,
  bounds?: { x: number; y: number; w: number; h: number },
  shapeRotationDeg = 0,
) {
  // `scale` is EMU → px factor (canvasWidthPx / slideWidthEMU).
  applyStrokeCore(ctx, stroke, scale);
  if (stroke?.fill && bounds) {
    const paint = resolveShapeFill(
      stroke.fill,
      ctx,
      bounds.x,
      bounds.y,
      bounds.w,
      bounds.h,
      shapeRotationDeg,
      scale * PT_TO_EMU,
      undefined, currentStrokeBounds(ctx),
    );
    if (paint) ctx.strokeStyle = paint;
  }
}
