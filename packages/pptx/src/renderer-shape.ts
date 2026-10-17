// PPTX shape element rendering (geometry, fill, outline, effects, text).
// Split from renderer.ts (#1857).

import type { ShapeElement } from './types';
import {
  usesPathShade,
  trackPaintPath,
  resolveArrowPaint,
  buildShapePath,
  EMU_PER_INCH,
  EMU_PER_PT as PT_TO_EMU,
  renderPresetShape,
  pathFillModeOverlay,
  hasPreset,
  buildPresetGeometryFillPath,
  getPresetGeometryBounds,
  getCustomGeometryBounds,
  getConnectorAnchors,
  getCustGeomEndpoints,
  drawArrowHead,
  lineEndPaintExtent,
  lineEndRetract,
  traceTrimmedPath,
  tracePresetLeader,
  trimGeometryPathEnds,
  isScene3dNonIdentity,
} from '@silurus/ooxml-core';
import { isSmartArtFallbackShape } from './smartart-fallback-contrast';
import {
  emuToPx,
  hexToRgba,
  buildCustomPath,
  type RenderContext,
  type TextRunCallback,
  type FetchImage,
} from './renderer-context';
import {
  shapeImageFillModeIsPaintable,
  applyStroke,
  resolveShapeFill,
  shapeFillKey,
  paintPreparedShapeImageFill,
  drawCompoundLine,
} from './renderer-shape-fill';
import {
  buildBevelInputs,
  buildExtrusion,
  splitScene3dShapeTransform,
  projectScene3dPaint,
  projectedEffectGeometry,
  cacheDevicePaint,
  paintWithRasterEffects,
  applyGlow,
  clearShadow,
  transformedEffectBounds,
  paintBeveledFlat,
  authoredAlignmentAnchor,
  type Sp3dLike,
  type LightRigLike,
  type EffectPaint,
} from './renderer-shape-effects';
import { renderTextBody } from './renderer-text-body';

/**
 * Preset-geometry text rectangle (ECMA-376 §20.1.9.21 `<a:rect>` /
 * presetTextRectangle), as an absolute sub-rect of the shape bbox. PowerPoint
 * lays text out inside this rect (then applies the `lIns/tIns/rIns/bIns`
 * insets), NOT in the full bounding box. For arrows the text rect is the shaft
 * (left of the arrowhead), so centered text sits in the body — without this a
 * `rightArrow`'s label drifts right into the arrowhead. Returns null when the
 * geometry's text rect is the whole bbox (the common case).
 *
 * Arrowhead length uses `ss = min(w, h)` per the spec's `dx1 = ss * adj2`.
 */
function presetTextRect(
  geom: string,
  x: number, y: number, w: number, h: number,
  adj1?: number | null, adj2?: number | null,
): { tx: number; ty: number; tw: number; th: number } | null {
  const ss = Math.min(w, h);
  switch (geom) {
    case 'rightarrow':
    case 'leftarrow': {
      const a1 = Math.min(Math.max(adj1 ?? 50000, 0), 100000); // shaft height fraction
      const a2 = Math.min(Math.max(adj2 ?? 50000, 0), 100000); // arrowhead length fraction
      const dx = (ss * a2) / 100000;        // arrowhead length (ss-based, ECMA dx1)
      const dy = (h * a1) / 200000;          // shaft half-height about the center
      const ty = y + h / 2 - dy;
      const th = 2 * dy;
      const tw = Math.max(0, w - dx);
      return geom === 'rightarrow'
        ? { tx: x, ty, tw, th }              // shaft on the left, arrowhead on the right
        : { tx: x + dx, ty, tw, th };        // leftarrow: arrowhead on the left
    }
    case 'roundrect': {
      // Text rect inset from the rounded corners: il = radius * (1 - 1/√2),
      // radius = ss * adj / 100000 (ECMA roundRect adj default 16667).
      const a = Math.min(Math.max(adj1 ?? 16667, 0), 100000) as number;
      const il = (ss * a) / 100000 * (1 - 1 / Math.SQRT2);
      return { tx: x + il, ty: y + il, tw: Math.max(0, w - 2 * il), th: Math.max(0, h - 2 * il) };
    }
    default:
      return null;
  }
}

/**
 * The shape-level default text colour passed to {@link renderTextBody} (used
 * by runs whose colour resolves to null). A style-derived `defaultTextColor`
 * (p:style > fontRef) always wins; the synthetic SmartArt fallback shape —
 * which never carries one — takes the slide's contrast-aware default from
 * {@link RenderContext.smartArtFallbackTextColor} so its data-model text is
 * legible on a dark background (issue #805). All other shapes: null (theme
 * default applies downstream).
 */
function shapeDefaultTextColor(el: ShapeElement, rc: RenderContext): string | null {
  if (el.defaultTextColor) return hexToRgba(el.defaultTextColor);
  if (rc.smartArtFallbackTextColor != null && isSmartArtFallbackShape(el)) {
    return rc.smartArtFallbackTextColor;
  }
  return null;
}

/**
 * Resolve the rotation PowerPoint uses for text in a singly-reflected shape.
 *
 * ECMA-376 §20.1.7.6 permits rotation and axis flips in the same transform.
 * For a one-axis reflection, two decompositions describe the same geometry:
 * rotating by another 180 degrees while exchanging flipH and flipV leaves the
 * shape matrix unchanged. PowerPoint chooses the decomposition whose text
 * rotation is in the readable -90..90 degree half-plane. Geometry continues
 * to use the authored transform verbatim; only its unmirrored text frame uses
 * this equivalent rotation.
 */
export function reflectedShapeTextRotation(
  rotation: number,
  flipH: boolean,
  flipV: boolean,
): number {
  if (flipH === flipV || !Number.isFinite(rotation)) return rotation;

  let readable = ((rotation + 180) % 360 + 360) % 360 - 180;
  if (readable > 90) readable -= 180;
  else if (readable < -90) readable += 180;
  return Object.is(readable, -0) ? 0 : readable;
}

/**
 * Text rotation of a shape's unmirrored text frame. Stacked vertical text
 * (`wordArtVert` / `wordArtVertRtl`) does not take the readable decomposition
 * of {@link reflectedShapeTextRotation}: in a flipV shape PowerPoint turns it
 * 180° and flipH leaves it unchanged (#1626 wordartvert slide 13: rotation 0
 * with flipH and with flipV, both directions).
 */
export function shapeTextRotation(vert: string, rotation: number, flipH: boolean, flipV: boolean): number {
  if (vert === 'wordArtVert' || vert === 'wordArtVertRtl') return rotation + (flipV ? 180 : 0);
  return reflectedShapeTextRotation(rotation, flipH, flipV);
}

export function renderShape(ctx: CanvasRenderingContext2D, el: ShapeElement, scale: number, themeDefaultColor = '#000000', slideNumber?: number, rc: RenderContext = { themeMajorFont: null, themeMinorFont: null, dpr: 1 }, onTextRun?: TextRunCallback, fetchImage?: FetchImage) {
  if (usesPathShade(el.stroke?.fill)) ctx = trackPaintPath(ctx);
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);
  const shapeTextRunCallback: TextRunCallback | undefined = onTextRun && el.id !== undefined
    ? (run) => onTextRun({ ...run, shapeId: el.id })
    : onTextRun;

  // anchor="b" + h=0: shape grows upward from y; render stroke as bottom border,
  // then let renderTextBody handle positioning.
  if (h === 0 && el.textBody?.verticalAnchor === 'b') {
    if (el.stroke) {
      ctx.save();
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      applyStroke(ctx, el.stroke, scale, { x, y, w, h: 1 }, el.rotation);
      ctx.stroke();
      ctx.restore();
    }
    if (el.textBody) {
      const defaultTextColor = shapeDefaultTextColor(el, rc);
      renderTextBody(ctx, el.textBody, x, y, w, h, scale, defaultTextColor, el.rotation, el.flipH, el.flipV, themeDefaultColor, slideNumber, rc, shapeTextRunCallback, false, fetchImage);
    }
    return;
  }

  // ── scene3d camera projection for p:sp (ECMA-376 §20.1.5.5, Phase B) ──────
  // A text-bearing shape with a non-identity 3-D camera uses local offscreens:
  // the body is bevel-shaded and text is painted unshaded onto
  // the live ctx through the camera homography — the same pipeline the picture
  // path uses. Authored flips are part of the local face before projection;
  // the element's 2-D rotation stays on the live ctx outside the warp.
  //
  // LIMITATION (documented, not silent): the HTML text-selection overlay tracks
  // glyphs by their un-projected layout rect and CSS transforms; it cannot
  // follow the per-pixel perspective warp applied here. So a projected shape's
  // text is drawn into the bitmap but EXCLUDED from the selection overlay
  // (onTextRun is suppressed below). See README "Bevel / 3D" note.
  // Camera framing uses document inches, even when the slide renders at an
  // arbitrary target width. Group sizing is already folded into w/h by Annex L.
  const spCamera = el.scene3d && { ...el.scene3d.camera, unitsPerInch: EMU_PER_INCH * scale,
    sceneDepth: (el.sp3d?.extrusionH ?? 0) * scale };
  const spScene3d = spCamera && isScene3dNonIdentity(
    spCamera, w, h, (el.sp3d?.extrusionH ?? 0) * scale,
  ) ? { ...el.scene3d, camera: spCamera } : null;
  if (spScene3d && w > 0 && h > 0) {
    const tf = ctx.getTransform();
    const det = Math.abs(tf.a * tf.d - tf.b * tf.c);
    const ctxDevScale = det > 0 ? Math.sqrt(det) : 1;
    const bevels = buildBevelInputs(
      el.sp3d as Sp3dLike | undefined,
      el.scene3d?.lightRig as LightRigLike | undefined,
      (el.sp3d as { prstMaterial?: string } | undefined)?.prstMaterial,
      scale,
      ctxDevScale,
      spScene3d.camera,
    );
    const backDepthCss = (bevels.find(bevel => bevel.bottom)?.depthPx ?? 0) / ctxDevScale;
    const extrusion = buildExtrusion(
      el.sp3d as Sp3dLike | undefined,
      spScene3d.camera,
      w,
      h,
      scale,
      ctxDevScale,
    );
    const sceneTransform =
      splitScene3dShapeTransform(el.rotation, el.flipH, el.flipV);
    // Apply the element's 2-D rotation on the live ctx; the warp composes
    // inside it. Flips stay with the local face below.
    ctx.save();
    if (sceneTransform.outerRotation !== 0) {
      ctx.translate(x + w / 2, y + h / 2);
      ctx.rotate((sceneTransform.outerRotation * Math.PI) / 180);
      ctx.translate(-(x + w / 2), -(y + h / 2));
    }
    // Local copy with the camera and 2-D placement neutralised: rendered at the
    // origin with no scene3d so the recursive call paints a flat body+text, then
    // we warp it. Text selection is dropped (onTextRun omitted) per the note.
    const localEl: ShapeElement = {
      ...el,
      x: 0,
      y: 0,
      rotation: 0,
      flipH: sceneTransform.localFlipH,
      flipV: sceneTransform.localFlipV,
      scene3d: undefined,
      // The outer projection owns bevel and extrusion. Retaining sp3d here
      // would shade the recursive flat body and then shade it again outside.
      sp3d: undefined,
      // Raster effects apply to the projected silhouette, not to the flat
      // source face. Keeping them here clips their reach at the source
      // offscreen and then warps the already-clipped result.
      shadow: null,
      innerShadow: undefined,
      glow: undefined,
      softEdge: undefined,
      reflection: undefined,
    };
    const localBodyEl: ShapeElement = { ...localEl, textBody: null };
    const localTextEl: ShapeElement = {
      ...localEl,
      fill: null,
      stroke: null,
    };
    const edgePadCss =
      (el.stroke ? (el.stroke.width * scale) / 2 : 0) +
      (el.sp3d?.contourW ? el.sp3d.contourW * scale : 0) +
      (extrusion ? Math.hypot(extrusion.offsetX, extrusion.offsetY) / ctxDevScale : 0) +
      2;
    const projectedTextHasPattern = el.textBody?.paragraphs.some(p =>
      p.runs.some(run => run.type === 'text' && (
        run.patternFill?.fillType === 'pattern' ||
        run.outline?.fill?.fillType === 'pattern' ||
        run.underlineFill?.fillType === 'pattern'
      )),
    ) ?? false;
    const paintProjectedElement = (
      target: CanvasRenderingContext2D,
      projectedElement: ShapeElement,
      padded: boolean,
      preservePatternFrame = false,
    ): boolean =>
      projectScene3dPaint(
        target,
        spScene3d.camera,
        x,
        y,
        w,
        h,
        (octx) => {
        // localEl is at the origin (x=y=0) with the element's own EMU size, so
        // the recursive render fills the (0,0,w,h) offscreen at the same scale.
        // No onTextRun → the projected text is not selectable (see the note).
          renderShape(octx, projectedElement, scale, themeDefaultColor, slideNumber, rc, undefined);
        },
        padded
          ? { bevels, backDepthCss, extrusion: extrusion ?? undefined, edgePadCss, preservePatternFrame }
          : { preservePatternFrame, measureOverflow: true },
      );
    const paintProjectedBody = (target: CanvasRenderingContext2D): boolean =>
      paintProjectedElement(target, localBodyEl, true);
    const paintProjectedText = (target: CanvasRenderingContext2D): void => {
      if (!el.textBody || paintProjectedElement(target, localTextEl, false, projectedTextHasPattern)) return;
      // Match the body allocation/horizon fallback: retain readable flat text.
      target.save();
      target.translate(x, y);
      renderShape(target, localTextEl, scale, themeDefaultColor, slideNumber, rc);
      target.restore();
    };
    const hasRasterEffects = Boolean(
      el.shadow || el.innerShadow || el.glow || el.softEdge || el.reflection,
    );
    if (hasRasterEffects) {
      const liveTransform = ctx.getTransform();
      const det = Math.abs(liveTransform.a * liveTransform.d - liveTransform.b * liveTransform.c);
      const devScale = det > 0 ? Math.sqrt(det) : 1;
      const projectedGeometry = projectedEffectGeometry(
        spScene3d.camera,
        liveTransform,
        x,
        y,
        w,
        h,
        edgePadCss,
        el.shadow?.algn,
        backDepthCss,
      );
      let projectionSucceeded = false;
      const rawPaint: EffectPaint = (target) => {
        projectionSucceeded = paintProjectedBody(target) || projectionSucceeded;
      };
      const cachedPaint = cacheDevicePaint(
        rawPaint,
        liveTransform,
        projectedGeometry.bbox,
        {
          w: (ctx.canvas as { width: number }).width || 0,
          h: (ctx.canvas as { height: number }).height || 0,
        },
        ctx,
      );
      // A declined cache is still a valid projection path: the compositor will
      // replay rawPaint. Check success only AFTER that first paint so an
      // off-viewport or oversized cache request cannot demote a 3-D shape to
      // the flat fallback merely because the optimization was skipped.
      paintWithRasterEffects(
        ctx,
        el,
        cachedPaint,
        cachedPaint,
        projectedGeometry.bbox,
        projectedGeometry.anchor,
        scale,
        scale * devScale,
        liveTransform,
        Boolean(el.fill),
        true,
      );
      if (projectionSucceeded) {
        paintProjectedText(ctx);
        ctx.restore();
        return;
      }
    } else {
      // sp3d belongs to the shape body (§20.1.5.12); text has its own run
      // properties. Use the same unshaded text pass as the raster-effect path.
      // Without a material pass, retain the combined face's rasterization and
      // single warp, including the established pattern coordinate frame.
      const projected = bevels.length > 0
        ? paintProjectedBody(ctx)
        : paintProjectedElement(ctx, localEl, true,
            projectedTextHasPattern && el.fill?.fillType !== 'pattern');
      if (projected) {
        if (bevels.length > 0) paintProjectedText(ctx);
        ctx.restore();
        return;
      }
    }
    // Headless fallback (no offscreen): draw flat below, but still without the
    // projection. Restore and fall through to the normal path.
    ctx.restore();
  }

  ctx.save();
  if (el.rotation !== 0 || el.flipH || el.flipV) {
    ctx.translate(x + w / 2, y + h / 2);
    ctx.rotate((el.rotation * Math.PI) / 180);
    if (el.flipH) ctx.scale(-1, 1);
    if (el.flipV) ctx.scale(1, -1);
    ctx.translate(-(x + w / 2), -(y + h / 2));
  }

  const geom = el.geometry.toLowerCase();
  const fillOutline: import('@silurus/ooxml-core').FillOutline = (target, bx, by, bw, bh) => {
    if (el.custGeom) {
      const paths = el.custGeomPaint?.length === el.custGeom.length
        ? el.custGeom.filter((_, index) => el.custGeomPaint?.[index].fill !== 'none')
        : el.custGeom;
      buildCustomPath(target, paths, bx, by, bw, bh);
    } else if (!buildPresetGeometryFillPath(target, geom, bx, by, bw, bh,
      [el.adj, el.adj2, el.adj3, el.adj4, el.adj5, el.adj6, el.adj7, el.adj8])) {
      buildShapePath(target, geom, bx, by, bw, bh, el.adj, el.adj2, el.adj3, el.adj4);
    }
  };
  // The slide may render at any requested width. Convert the PDF-measured
  // one-point pattern cell through this render's EMU-to-canvas scale.
  const fillStyle = resolveShapeFill(el.fill, ctx, x, y, w, h, el.rotation, scale * PT_TO_EMU, fillOutline);
  const imageFill = el.fill?.fillType === 'image' && shapeImageFillModeIsPaintable(el.fill)
    ? el.fill
    : null;
  const preparedImageFill = imageFill
    ? rc.shapeFillImages?.get(shapeFillKey(imageFill))
    : undefined;
  const hasPreparedImageFill = Boolean(imageFill && preparedImageFill);

  // The Canvas API exposes a single shadow slot, so when both an outer shadow
  // and a glow are configured we let the outer shadow win (visually dominant)
  // and fall back to the glow only when no outer shadow is present. The outer
  // shadow itself is composited from the complete fill+stroke silhouette below;
  // applying native shadow state here makes each path cast a separate shadow.
  if (!el.shadow) applyGlow(ctx, el.glow ?? null, scale);

  const CONNECTOR_GEOMS = new Set([
    'line', 'straightconnector1',
    'bentconnector2', 'bentconnector3', 'bentconnector4', 'bentconnector5',
    'curvedconnector2', 'curvedconnector3', 'curvedconnector4', 'curvedconnector5',
  ]);

  // callout family — `<a:ln><a:headEnd|tailEnd>` decorate the LEADER line
  // (the geometry's last `<path>` in presets.json), not the text rectangle
  // (path 0) or the accent bar. callout1 = 2-point leader, callout2/3 =
  // 3-/4-point polylines; getConnectorAnchors resolves the leader's two ends
  // for every variant, so all twelve share the connector decoration path.
  const CALLOUT_GEOMS = new Set([
    'callout1', 'callout2', 'callout3',
    'bordercallout1', 'bordercallout2', 'bordercallout3',
    'accentcallout1', 'accentcallout2', 'accentcallout3',
    'accentbordercallout1', 'accentbordercallout2', 'accentbordercallout3',
  ]);

  // A leader line we re-stroke from its vertices, trimmed at its decorated
  // ends: every callout plus the straight / bent (polyline) connectors.
  // Curved connectors are excluded here, because their leader is a Bézier and
  // getConnectorAnchors.vertices only captures segment endpoints, so
  // re-stroking from vertices would straighten the curve. When an end
  // retracts, their preset leader is trimmed along the curve instead
  // (`trimsCurvedLeader` below).
  const isRetractableLeader = (g: string): boolean =>
    CALLOUT_GEOMS.has(g) || g === 'line' || g === 'straightconnector1' || g.startsWith('bentconnector');
  // How far each decorated end pulls its line back (observed Office rule in
  // core arrow-geometry.ts). Paths that cannot be re-stroked from straight
  // vertices — curved connectors, custom geometry, compound lines — are
  // trimmed along the path instead, so every line kind stops at the same place.
  const headTrim = el.stroke?.headEnd ? lineEndRetract(el.stroke.headEnd, el.stroke, scale, PT_TO_EMU) : 0;
  const tailTrim = el.stroke?.tailEnd ? lineEndRetract(el.stroke.tailEnd, el.stroke, scale, PT_TO_EMU) : 0;
  const trimsCurvedLeader = geom.startsWith('curvedconnector') && (headTrim > 0 || tailTrim > 0);
  const custEnds = el.custGeom && el.custGeom.length > 0 && (headTrim > 0 || tailTrim > 0)
    ? getCustGeomEndpoints(el.custGeom) : null;
  const custHeadTrim = custEnds?.start ? headTrim : 0;
  const custTailTrim = custEnds?.end ? tailTrim : 0;

  // ── Dispatch to preset engine when possible ────────────────────────────
  // Preference order: custGeom → generic preset engine → buildShapePath
  // (only `rect` or a schema-invalid name, which the engine does not define).
  // `arc` (ECMA-376 §20.1.10.56 ST_ShapeType "arc") goes through the engine
  // too: its presetShapeDefinitions geometry is a two-<path> shape — path 0
  // (stroke="false") fills the pie wedge (arc + lnTo centre + close) and path 1
  // (fill="none") strokes only the open arc edge.
  const usePresetEngine =
    !el.custGeom && hasPreset(geom);

  /**
   * Paint the shape's body (fill + stroke) into an arbitrary target context,
   * dispatching through the same preset / custGeom / legacy paths as the live
   * render. Extracted into a closure so the edge/blur effect helpers
   * (innerShdw §20.1.8.40, softEdge §20.1.8.53, reflection §20.1.8.50) can
   * re-paint the shape onto auxiliary canvases.
   *
   * `silhouette`: when set, paint a flat opaque silhouette in that colour
   * (filled path only, no stroke, no gradients) — used by innerShdw to build
   * its mask. The silhouette honours the same even-odd / preset path topology.
   */
  const paintShapeBody = (
    target: CanvasRenderingContext2D,
    silhouette?: string,
    bounds: { x: number; y: number; w: number; h: number } = { x, y, w, h },
  ): void => {
    if (usesPathShade(el.stroke?.fill)) target = trackPaintPath(target);
    const { x: bx, y: by, w: bw, h: bh } = bounds;
    const tFill = silhouette ??
      (target === ctx && bx === x && by === y && bw === w && bh === h
        ? fillStyle
        : resolveShapeFill(el.fill, target, bx, by, bw, bh, el.rotation, scale * PT_TO_EMU, fillOutline));
    const tStroke = silhouette
      ? null
      : el.stroke
        ? () => {
            applyStroke(target, el.stroke!, scale, { x: bx, y: by, w: bw, h: bh }, el.rotation);
            target.stroke();
          }
        : null;
    const tClearShadow = () => clearShadow(target);
    const evenOddFill = geom === 'donut' || geom === 'smileyface' || geom === 'frame';
    const paintImageFill = imageFill && !silhouette
      ? (paintCtx: CanvasRenderingContext2D) => paintPreparedShapeImageFill(
          paintCtx,
          imageFill,
          preparedImageFill,
          { x: bx, y: by, w: bw, h: bh },
          scale,
          evenOddFill,
          el,
        )
      : undefined;

    if (usePresetEngine && !silhouette) {
      const skipTrailingStroke = isRetractableLeader(geom) || trimsCurvedLeader;
      renderPresetShape(
        target, geom, bx, by, bw, bh,
        [el.adj, el.adj2, el.adj3, el.adj4, el.adj5, el.adj6, el.adj7, el.adj8],
        tFill, tStroke, tClearShadow,
        // A retractable leader (callout / straight / bent connector) is
        // re-stroked retracted from its decorated ends in the line-end block
        // below, so suppress the preset engine's full-length leader stroke to
        // avoid a double line / a cap poking through the arrow tip.
        skipTrailingStroke || paintImageFill
          ? { ...(skipTrailingStroke ? { skipTrailingStroke: true } : {}), paintFill: paintImageFill }
          : undefined,
      );
      return;
    }

    // ECMA-376 §20.1.9.15: custom geometry paths carry their own fill mode and
    // stroke flag. Paint each path on its own, like the preset engine does, so
    // an unfilled or unstroked path stays unfilled / unstroked. A silhouette
    // uses the fill-bearing paths only.
    const pathPaint = el.custGeom && el.custGeomPaint?.length === el.custGeom.length
      ? el.custGeomPaint
      : null;
    if (el.custGeom && pathPaint) {
      let shadowCleared = false;
      el.custGeom.forEach((cmds, index) => {
        const paint = pathPaint[index];
        const filled = paint.fill !== 'none';
        if (silhouette && !filled) return;
        target.beginPath();
        buildCustomPath(target, [cmds], bx, by, bw, bh);
        if (filled) {
          let painted = false;
          if (paintImageFill) {
            target.save();
            try { painted = paintImageFill(target); } finally { target.restore(); }
          } else if (tFill) {
            target.fillStyle = tFill;
            target.fill();
            painted = true;
          }
          const overlay = painted && !silhouette ? pathFillModeOverlay(paint.fill) : null;
          if (overlay) {
            target.save();
            target.fillStyle = overlay;
            target.fill();
            target.restore();
          }
          if (painted && !silhouette && !shadowCleared) {
            tClearShadow();
            shadowCleared = true;
          }
        }
        if (paint.stroke && tStroke) {
          const head = index === 0 ? custHeadTrim : 0;
          const tail = index === el.custGeom!.length - 1 ? custTailTrim : 0;
          if (head > 0 || tail > 0) {
            target.beginPath();
            traceTrimmedPath(target, (sink) => buildCustomPath(sink, [cmds], bx, by, bw, bh), head, tail);
          }
          tStroke();
        }
      });
      return;
    }
    target.beginPath();
    if (el.custGeom && el.custGeom.length > 0) {
      buildCustomPath(target, el.custGeom, bx, by, bw, bh);
    } else {
      // Silhouette of a preset shape (or any other geometry name).
      // buildShapePath resolves spec-defined names through the ECMA-376
      // preset engine (all of its paths).
      buildShapePath(target, geom, bx, by, bw, bh, el.adj, el.adj2, el.adj3, el.adj4);
    }
    // Arc bodies render through the preset engine above, which owns arc's
    // pie-wedge fill (path 0) and open-arc stroke (path 1); never fill an arc
    // outline here, where both paths would be filled together.
    if (paintImageFill && geom !== 'arc') {
      if (paintImageFill(target)) tClearShadow();
    } else if (tFill && geom !== 'arc') {
      target.fillStyle = tFill;
      if (evenOddFill) {
        target.fill('evenodd');
      } else {
        target.fill();
      }
      if (!silhouette) tClearShadow();
    }
    if (tStroke) {
      if (el.custGeom && el.custGeom.length > 0 && (custHeadTrim > 0 || custTailTrim > 0)) {
        target.beginPath();
        traceTrimmedPath(target, (sink) => buildCustomPath(sink, el.custGeom!, bx, by, bw, bh), custHeadTrim, custTailTrim);
      }
      tStroke();
    }
  };

  // ── effectLst edge/blur effects (independent siblings, ECMA-376 §20.1.8.25)
  // The effect helpers operate in DEVICE pixels: the aux silhouette is painted
  // through the live transform (which already folds in devicePixelRatio +
  // rotation + flip), and the blit happens at identity. So bbox / radii passed
  // to the helpers must be in device px too. The uniform device scale equals
  // sqrt(|det|) of the transform's linear part — for scale(dpr)·rot·flip this
  // is exactly dpr, independent of rotation or mirroring.
  const liveTransform = ctx.getTransform();
  const det = Math.abs(liveTransform.a * liveTransform.d - liveTransform.b * liveTransform.c);
  const devScale = det > 0 ? Math.sqrt(det) : 1;
  // Preset adjustments may place painted vertices outside the nominal xfrm
  // rectangle (for example, a wedge callout's tail). Crop effects to the
  // geometry bounds so those protrusions participate in the shadow/reflection
  // instead of being clipped before compositing.
  const needsRasterEffectBounds = Boolean(
    el.shadow || el.reflection || el.softEdge || el.innerShadow,
  );
  const adjustedGeometryBounds = needsRasterEffectBounds
    ? usePresetEngine
      ? getPresetGeometryBounds(
          geom,
          x,
          y,
          w,
          h,
          [el.adj, el.adj2, el.adj3, el.adj4, el.adj5, el.adj6, el.adj7, el.adj8],
        )
      : el.custGeom && el.custGeom.length > 0
        ? getCustomGeometryBounds(el.custGeom, x, y, w, h)
        : null
    : null;
  const geometryBounds = adjustedGeometryBounds ?? { x, y, w, h };
  const strokeHalf = el.stroke ? el.stroke.width * scale / 2 : 0;
  const lineEndExtent = el.stroke
    ? Math.max(
        el.stroke.headEnd ? lineEndPaintExtent(el.stroke.headEnd, el.stroke, scale, PT_TO_EMU) : 0,
        el.stroke.tailEnd ? lineEndPaintExtent(el.stroke.tailEnd, el.stroke, scale, PT_TO_EMU) : 0,
      )
    : 0;
  const contourExtent = el.sp3d?.contourW ? el.sp3d.contourW * scale : 0;
  const paintedPad = Math.max(strokeHalf, lineEndExtent, contourExtent);
  const effectBounds = paintedPad > 0
    ? {
        x: geometryBounds.x - paintedPad,
        y: geometryBounds.y - paintedPad,
        w: geometryBounds.w + paintedPad * 2,
        h: geometryBounds.h + paintedPad * 2,
      }
    : geometryBounds;
  const effBBox = transformedEffectBounds(
    liveTransform,
    effectBounds.x,
    effectBounds.y,
    effectBounds.w,
    effectBounds.h,
  );
  const effScale = scale * devScale; // EMU → device px
  // A face-on camera does not need a homography, but sp3d bevels still alter
  // the front surface. Pictures already take this flat offscreen path; shapes
  // must do the same instead of treating an identity camera as "no 3-D".
  const flatBevels = !spScene3d
    ? buildBevelInputs(
        el.sp3d as Sp3dLike | undefined,
        el.scene3d?.lightRig as LightRigLike | undefined,
        el.sp3d?.prstMaterial,
        scale,
        devScale,
        el.scene3d?.camera,
      )
    : [];
  const flatBevelEdgePadCss = (el.stroke ? (el.stroke.width * scale) / 2 : 0) + 2;
  const paintLineDecorations = (target: CanvasRenderingContext2D): void => {
    if (usesPathShade(el.stroke?.fill)) target = trackPaintPath(target);
    if (el.stroke && (CONNECTOR_GEOMS.has(geom) || CALLOUT_GEOMS.has(geom))) {
      // The preset body deliberately suppresses retractable leader strokes. Paint
      // the shortened leader and its line ends into the same target as the body
      // so outer shadows, reflections and soft edges see one composed object.
      const anchors = getConnectorAnchors(
        geom,
        x,
        y,
        w,
        h,
        [el.adj, el.adj2, el.adj3, el.adj4, el.adj5, el.adj6, el.adj7, el.adj8],
      );
      if (!anchors) return;
      const cmpd = el.stroke.cmpd;
      const isStraight = geom === 'line' || geom === 'straightconnector1';
      if (isRetractableLeader(geom) && anchors.vertices.length >= 2 && !(cmpd && isStraight)) {
        // Trim across the whole leader, so a short terminal leg cannot stop
        // the retraction early; a leader consumed by both ends paints nothing.
        const pts = anchors.vertices;
        target.beginPath();
        traceTrimmedPath(target, (sink) => {
          sink.moveTo(pts[0].x, pts[0].y);
          for (let i = 1; i < pts.length; i++) sink.lineTo(pts[i].x, pts[i].y);
        }, headTrim, tailTrim);
        applyStroke(target, el.stroke, scale, { x, y, w, h }, el.rotation);
        target.stroke();
      }
      if (trimsCurvedLeader) {
        target.beginPath();
        if (tracePresetLeader(target, geom, x, y, w, h,
          [el.adj, el.adj2, el.adj3, el.adj4, el.adj5, el.adj6, el.adj7, el.adj8], headTrim, tailTrim)) {
          applyStroke(target, el.stroke, scale, { x, y, w, h }, el.rotation);
          target.stroke();
        }
      }
      if (cmpd && isStraight) {
        // Both ends trim the same straight run; when they consume it, no
        // parallel line remains (independent retraction would cross over).
        const run = trimGeometryPathEnds([
          { op: 'moveTo', args: [anchors.start.x, anchors.start.y] },
          { op: 'lineTo', args: [anchors.end.x, anchors.end.y] },
        ], headTrim, tailTrim);
        if (run.length === 2) {
          drawCompoundLine(
            target,
            { x: run[0].args[0], y: run[0].args[1] },
            { x: run[1].args[0], y: run[1].args[1] },
            el.stroke,
            cmpd,
            scale,
            el.rotation,
          );
        }
      }
      if (el.stroke.tailEnd) {
        drawArrowHead(
          target,
          anchors.end.x,
          anchors.end.y,
          anchors.end.angle,
          el.stroke.tailEnd,
          el.stroke,
          scale,
          PT_TO_EMU,
          resolveArrowPaint(target, el.stroke, scale, PT_TO_EMU, anchors.end.x, anchors.end.y, anchors.end.angle,
            el.stroke.tailEnd, { x, y, w, h }, el.rotation, scale * PT_TO_EMU),
        );
      }
      if (el.stroke.headEnd) {
        drawArrowHead(
          target,
          anchors.start.x,
          anchors.start.y,
          anchors.start.angle,
          el.stroke.headEnd,
          el.stroke,
          scale,
          PT_TO_EMU,
          resolveArrowPaint(target, el.stroke, scale, PT_TO_EMU, anchors.start.x, anchors.start.y, anchors.start.angle,
            el.stroke.headEnd, { x, y, w, h }, el.rotation, scale * PT_TO_EMU),
        );
      }
      return;
    }

    if (
      !el.stroke ||
      !el.custGeom ||
      el.custGeom.length === 0 ||
      ((!el.stroke.headEnd || el.stroke.headEnd.type === 'none') &&
        (!el.stroke.tailEnd || el.stroke.tailEnd.type === 'none'))
    ) return;

    const { start, end } = getCustGeomEndpoints(el.custGeom);
    if (start && el.stroke.headEnd && el.stroke.headEnd.type !== 'none') {
      drawArrowHead(
        target,
        x + start.x * w,
        y + start.y * h,
        Math.atan2(start.dy * h, start.dx * w),
        el.stroke.headEnd,
        el.stroke,
        scale,
        PT_TO_EMU,
        resolveArrowPaint(target, el.stroke, scale, PT_TO_EMU, x + start.x * w, y + start.y * h,
          Math.atan2(start.dy * h, start.dx * w), el.stroke.headEnd, { x, y, w, h }, el.rotation, scale * PT_TO_EMU),
      );
    }
    if (end && el.stroke.tailEnd && el.stroke.tailEnd.type !== 'none') {
      drawArrowHead(
        target,
        x + end.x * w,
        y + end.y * h,
        Math.atan2(end.dy * h, end.dx * w),
        el.stroke.tailEnd,
        el.stroke,
        scale,
        PT_TO_EMU,
        resolveArrowPaint(target, el.stroke, scale, PT_TO_EMU, x + end.x * w, y + end.y * h,
          Math.atan2(end.dy * h, end.dx * w), el.stroke.tailEnd, { x, y, w, h }, el.rotation, scale * PT_TO_EMU),
      );
    }
  };
  const paintVisibleShapeBody = (target: CanvasRenderingContext2D): void => {
    if (flatBevels.length > 0) {
      const ok = paintBeveledFlat(
        target,
        x,
        y,
        w,
        h,
        flatBevels,
        (octx, ox, oy, ow, oh) =>
          paintShapeBody(octx, undefined, { x: ox, y: oy, w: ow, h: oh }),
        undefined,
        flatBevelEdgePadCss,
      );
      if (ok) {
        paintLineDecorations(target);
        clearShadow(target);
        return;
      }
    }
    paintShapeBody(target);
    paintLineDecorations(target);
  };
  paintWithRasterEffects(
    ctx,
    el,
    paintVisibleShapeBody,
    paintVisibleShapeBody,
    effBBox,
    authoredAlignmentAnchor(
      liveTransform,
      x,
      y,
      w,
      h,
      el.shadow?.algn,
    ),
    scale,
    effScale,
    liveTransform,
    Boolean(fillStyle) || hasPreparedImageFill,
    true,
  );

  // Render text inside the rotation context so text follows shape rotation
  if (el.textBody) {
    const defaultTextColor = shapeDefaultTextColor(el, rc);
    const textRotation = shapeTextRotation(el.textBody.vert, el.rotation, el.flipH, el.flipV);
    ctx.save();
    if (el.flipH || el.flipV) {
      const cx = x + w / 2;
      const cy = y + h / 2;
      // The shape itself stays mirrored, but text should remain readable.
      // Apply the same flip again around the shape centre to cancel only the text mirror.
      ctx.translate(cx, cy);
      if (el.flipH) ctx.scale(-1, 1);
      if (el.flipV) ctx.scale(1, -1);
      ctx.translate(-cx, -cy);
    }
    const rotationDelta = (textRotation - el.rotation) % 360;
    if (rotationDelta !== 0) {
      const cx = x + w / 2;
      const cy = y + h / 2;
      ctx.translate(cx, cy);
      ctx.rotate((rotationDelta * Math.PI) / 180);
      ctx.translate(-cx, -cy);
    }
    // For ellipses, PowerPoint positions text relative to the inscribed rectangle
    // (the maximum-area rectangle that fits inside the ellipse: sides = a/√2, b/√2).
    // This only affects non-ctr anchors; ctr anchor is invariant to this inset.
    let tx = x, ty = y, tw = w, th = h;
    if (el.textRect) {
      // SmartArt drawings carry an explicit text frame (<dsp:txXfrm>) that
      // PowerPoint pre-computed against the actual layout — e.g. an arrow's
      // label sits past an overlapping circle node, a roundRect's label avoids
      // an overlapping bottom badge. Honour it verbatim (insets apply within).
      tx = emuToPx(el.textRect.x, scale);
      ty = emuToPx(el.textRect.y, scale);
      tw = emuToPx(el.textRect.width, scale);
      th = emuToPx(el.textRect.height, scale);
    } else if (geom === 'ellipse') {
      const insetX = w * (1 - 1 / Math.SQRT2) / 2;
      const insetY = h * (1 - 1 / Math.SQRT2) / 2;
      tx = x + insetX; ty = y + insetY;
      tw = w / Math.SQRT2; th = h / Math.SQRT2;
    } else {
      // Preset text rectangle (ECMA-376 §20.1.9.21): e.g. an arrow's label sits
      // in the shaft, not the full bbox. Insets (lIns/…) apply within this rect.
      const tr = presetTextRect(geom, x, y, w, h, el.adj, el.adj2);
      if (tr) { tx = tr.tx; ty = tr.ty; tw = tr.tw; th = tr.th; }
    }
    // Carry the same readable rotation into the selectable text overlay.
    renderTextBody(ctx, el.textBody, tx, ty, tw, th, scale, defaultTextColor, textRotation, false, false, themeDefaultColor, slideNumber, rc, shapeTextRunCallback, false, fetchImage);
    ctx.restore();
  }

  ctx.restore();
}
