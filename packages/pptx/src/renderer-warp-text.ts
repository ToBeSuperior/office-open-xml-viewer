// PPTX WordArt (prstTxWarp) text rendering through warped glyph strips.
// Split from renderer.ts (#1857).

import type { TextBody, Paragraph } from './types';
import {
  resolveFill as resolveFillCore,
  withInheritedPatternScope,
  EMU_PER_PT as PT_TO_EMU,
  createAuxCanvas,
  DECORATION_LINE_WIDTH_EM,
  buildWarpEnvelope,
  warpGlyphTransform,
  followPathUScale,
} from '@silurus/ooxml-core';
import type { WarpEnvelope, WarpGlyphTransform } from '@silurus/ooxml-core';
import { emuToPx, hexToRgba, type RenderContext } from './renderer-context';
import {
  layoutParagraph,
  codePointCount,
  resolveSegmentTextPaint,
  type TextPaint,
  type LayoutLine,
} from './renderer-text-layout';

// ── WordArt paired-edge warp: piecewise-affine strip subdivision ─────────────
//
// Each glyph is drawn as a stack of vertical STRIPS, every strip re-issuing
// `fillText` under its own affine with a clip rect bounding the strip's share
// of the glyph advance. Drawing VECTOR text per strip (instead of slicing a
// pre-rasterised glyph bitmap, as an earlier revision did) rasterises the
// glyph at the strip's final device transform, which kills three
// bitmap-transfer artifact classes a real browser exposed: blur from
// re-magnifying a raster (vScale stretches up to ~4×), glyph parts cut by the
// offscreen slab rectangle, and detached hairlines from stretched source-edge
// antialiasing. It also needs no auxiliary canvas at all, so the draw works
// identically on the main thread, in a worker, and under headless node.

// BASE strip width in DEVICE px — the coarse first guess only; the adaptive
// deviation loop below is what actually guarantees accuracy. Width alone bounds
// just the chord-vs-curve error of the BASELINE curve (∝ width², sub-pixel at
// 8 px for every preset), but NOT the error away from the baseline: a slab
// corner at distance d from the baseline anchor is additionally misplaced by
// the angle/vScale mismatch across the strip (≈ d·Δangle + |y|·ΔvScale), which
// at 8 px width is measured at ~0.7 px (Inflate) to ~0.2 px (Wave1/Chevron) at
// the baseline but grows to several px at the em-box extremes (Wave1 ≈2.4,
// Button ≈5.9) and tens of px on high-curvature ring/pour families
// (CirclePour ≈11, CanUp ≈44 at em-box edges) — the strips fan apart into
// radial slivers. Hence the width criterion is only the starting k; the
// deviation-driven refinement below subdivides until every painted corner is
// within WARP_STRIP_DEVIATION_BUDGET_DEVICE_PX of the exact envelope map.
const WARP_STRIP_DEVICE_PX = 8;
// Extend each strip's CLIP rect by 1 device px into its neighbour on each
// side. Two abutting clip edges antialiased in separate fill passes leave the
// classic shared-edge coverage dip (each pass covers the boundary pixel
// partially; src-over of two ~50 % coverages of colour C over white composes
// to ~0.75·C, a faint light seam through every glyph) — and the two edges can
// additionally diverge by up to 2·budget at the ink extremes. Overlapping the
// clips makes strip i paint the boundary band at FULL coverage before strip
// i+1's partial-coverage edge lands on top; painting the same opaque colour
// over itself is idempotent, so the seam vanishes. That idempotence only holds
// for an OPAQUE composite: a translucent fill (alpha < 1, from an 8-digit
// RRGGBBAA run colour or an inherited shape opacity) composes the 2-px band
// TWICE and darkens it (issue #879 — a visible pinstripe/moiré). The translucent
// branch of drawWarpedGlyphStrips fixes that by painting the strips OPAQUE into a
// per-glyph device-resolution layer (where the overlap is idempotent again) and
// compositing the layer ONCE with the effective alpha. The overlap is symmetric,
// so it does not bias the mapping.
const WARP_STRIP_OVERLAP_DEVICE_PX = 1;
// Max deviation (device px) of any painted slab corner from the EXACT envelope
// map, enforced by the adaptive loop. Derivation: the true map places the
// glyph's vertical ruling at flat-x b onto Q(b, y) = A(u_b)·(0, y), where A(u)
// is the envelope's local frame at u (anchor T(u)+bandFrac·gap, advance-axis
// angle, gap shear/scale — exactly what warpGlyphTransform returns). A strip
// centred at c paints that same ruling at S(b, y) = A(u_c)·(b−c, y). The
// deviation |S−Q| at the slab's four corners (both strip edges × slab
// top/bottom) is the exact per-strip error of the piecewise-affine map — it
// contains ALL mismatch components at once: rotational fan-out d·Δangle
// (dominant on rings/pours, whose tangent sweeps up to ~345°), vertical-scale
// slide |y|·ΔvScale (dominant on textButton, whose mean-tangent angle is
// constant 0 so a pure angle criterion would miss it), and the baseline chord
// error. Two adjacent strips each deviate ≤ budget from the SAME true ruling at
// their shared boundary, so their mutual divergence is ≤ 2·budget = the
// 2 device px their overlapping clips share — i.e. budget 1 is the largest
// value the clip overlap provably covers. Every quantity is geometric; no
// preset branch.
const WARP_STRIP_DEVIATION_BUDGET_DEVICE_PX = 1;
// Hard cap on strips per glyph, bounding draw cost for extreme geometry (one
// glyph sweeping a large arc of a ring at high em-height). At the cap the
// residual deviation is reported by the loop's last measurement and the seam
// stays proportional (a full-ring glyph at ~250 device px em-height measures a
// few px — degraded but bounded, and localized to that pathological glyph).
const WARP_STRIP_MAX_PER_GLYPH = 256;

/**
 * Draw one glyph across a PAIRED-EDGE warp envelope as a stack of vertical
 * strips (piecewise-affine approximation of the envelope map, ECMA-376
 * §20.1.9.19).
 *
 * The #866 per-glyph transform samples the envelope's local Jacobian at a SINGLE
 * u (the glyph centre) and applies one affine (rotate + shear + non-uniform
 * scale) to the whole glyph. That is exact only where the map is locally affine;
 * on Inflate/Deflate the vertical stretch `vScale` varies WITHIN a glyph (the
 * top/bottom edges are mirror-symmetric so the gap stays vertical — shear 0,
 * angle 0 — but the gap MAGNITUDE peaks at the centre and shrinks toward the
 * ends), which a single affine cannot represent: it leaves the glyph uniformly
 * scaled instead of stretched more where the envelope is taller. PowerPoint maps
 * the glyph OUTLINE continuously through `P(u,v) = (1−v)·T(u) + v·B(u)`, so a
 * wide glyph leans/stretches asymmetrically across its own width.
 *
 * Subdividing the glyph into narrow strips and mapping each strip by the Jacobian
 * at ITS centre-u converges to that continuous map as the strip count grows (the
 * general solution — no preset special-casing). The count is chosen ADAPTIVELY:
 * from a base width criterion, strips double until every painted corner deviates
 * from the exact envelope map by at most 1 device px (see
 * WARP_STRIP_DEVIATION_BUDGET_DEVICE_PX), so high-curvature envelopes (rings,
 * pours — tangent sweep up to ~345°) subdivide as finely as they need while flat
 * ones stay coarse.
 *
 * Each strip re-issues `fillText` under its own affine, clipped to the strip's
 * share of the glyph advance. The clip rect is expressed in GLYPH-LOCAL
 * (flat-advance) coordinates, so adjacent strips cut the glyph at the same
 * flat-x boundary; the text itself is rasterised by the canvas at the strip's
 * final device-space transform — full device resolution, no intermediate
 * bitmap. The clip is effectively unbounded vertically (the subdivision is
 * horizontal only) and the OUTER edges of a glyph's first/last strips are
 * unbounded outward, so side-bearing overhang ink (italic tails, swashes)
 * rides the end strips instead of being cut at the advance box.
 *
 * `chW` is the glyph's flat advance (css px, incl. letter-spacing `ls`);
 * `flatX0` is the glyph's flat left edge along the whole line (css px). The
 * remaining args mirror the per-glyph draw: `hScale` stretches flat ink to the
 * box width, `warpBoxH`/`bandFrac` drive `warpGlyphTransform`, `boxX`/`boxY`
 * are the shape origin, `totalW`/`followScale` map flat-x to the envelope
 * fraction u. `ctx.font` / `ctx.fillStyle` must already be set for this glyph
 * (save/restore around each strip preserves them).
 */
function drawWarpedGlyphStrips(
  ctx: CanvasRenderingContext2D,
  ch: string,
  ls: number,
  chW: number,
  devScale: number,
  env: WarpEnvelope,
  flatX0: number,
  totalW: number,
  followScale: number,
  hScale: number,
  warpBoxH: number,
  bandFrac: number,
  boxX: number,
  boxY: number,
  paint: TextPaint,
  resolveAuxPaint?: (target: CanvasRenderingContext2D) => TextPaint,
  strokeWidth = 0,
): void {
  if (chW <= 0) return;
  // Ink extremes about the baseline (css px), from real metrics — this is where
  // the deviation budget must hold, since that is where ink actually lands.
  // Fallbacks keep degenerate metrics safe (mirrors the flat renderer).
  const m = ctx.measureText(ch);
  const asc = m.actualBoundingBoxAscent > 0 ? m.actualBoundingBoxAscent : chW;
  const desc = m.actualBoundingBoxDescent > 0 ? m.actualBoundingBoxDescent : chW * 0.25;
  // Horizontal ink extremes about the pen origin (css px). Used only to bound the
  // translucent compositing layer (issue #879) tightly around real ink, so
  // side-bearing overhang on the outward-unbounded end strips still fits.
  const inkL = m.actualBoundingBoxLeft > 0 ? m.actualBoundingBoxLeft : 0;
  const inkR = m.actualBoundingBoxRight > 0 ? m.actualBoundingBoxRight : chW;

  // ── Adaptive strip count ──────────────────────────────────────────────────
  // Start from the width criterion (one strip per WARP_STRIP_DEVICE_PX of the
  // glyph's WARPED advance, chW·hScale·devScale), then DOUBLE the count until
  // every painted ink corner is within the deviation budget of the exact
  // envelope map (see WARP_STRIP_DEVIATION_BUDGET_DEVICE_PX for the formula and
  // why the budget ties to the clip overlap). Deviation from the smooth part of
  // the envelope halves with each doubling (it is ∝ the strip's u-span to first
  // order), so the loop converges in a few steps for arcs of any curvature —
  // this is what keeps high-curvature ring/pour envelopes (tangent sweep up to
  // ~345°) contiguous instead of fanning into radial slivers. If a doubling
  // fails to reduce the deviation by ≥ 25 %, the residual sits at a C0 corner of
  // the envelope itself (e.g. the chevron apex): the true map genuinely creases
  // there, refinement can only narrow the affected band — keep the finer
  // partition and stop rather than spinning to the cap.
  const warpedAdvDev = chW * hScale * devScale;
  let k = Math.min(
    WARP_STRIP_MAX_PER_GLYPH,
    Math.max(1, Math.round(warpedAdvDev / WARP_STRIP_DEVICE_PX)),
  );
  const build = (count: number): WarpStrip[] =>
    buildWarpStrips(count, env, chW, flatX0, totalW, followScale, warpBoxH, bandFrac);
  let strips = build(k);
  let dev = maxWarpStripDeviationDev(
    strips, env, flatX0, totalW, followScale, warpBoxH, bandFrac, hScale, devScale, -asc, desc,
  );
  while (dev > WARP_STRIP_DEVIATION_BUDGET_DEVICE_PX && k < WARP_STRIP_MAX_PER_GLYPH) {
    const k2 = Math.min(WARP_STRIP_MAX_PER_GLYPH, k * 2);
    const strips2 = build(k2);
    const dev2 = maxWarpStripDeviationDev(
      strips2, env, flatX0, totalW, followScale, warpBoxH, bandFrac, hScale, devScale, -asc, desc,
    );
    if (dev2 >= dev * 0.75) {
      // Not converging: C0 crease in the envelope (see above). The finer
      // partition still localizes the crease to a narrower strip — keep it.
      strips = strips2;
      break;
    }
    k = k2;
    strips = strips2;
    dev = dev2;
  }

  // Vertical clip extent: the subdivision is horizontal only, so the clip's
  // sole job is to bound x. ±10⁴ css px is four orders of magnitude beyond any
  // glyph's extent (so no ink is ever cut in y), kept finite only so the
  // rasterizer's float precision on the boundary line stays sub-thousandth-px.
  const CLIP_Y = 1e4;
  // Clip overlap in glyph-local flat units: the local x axis is scaled by
  // hScale (then the canvas dpr) on the way to device px, so this makes the
  // painted overlap WARP_STRIP_OVERLAP_DEVICE_PX device px wide.
  const ov = WARP_STRIP_OVERLAP_DEVICE_PX / (hScale * devScale);

  const last = strips.length - 1;
  // Glyph-local clip interval [x0, x1] of strip `i` (flat css px, origin at the
  // strip centre): interior boundaries sit at the shared s0/s1 (± the overlap);
  // the outer edges of the first/last strips are unbounded so side-bearing
  // overhang ink is not cut.
  const stripClipX0 = (i: number, s0: number, centre: number): number =>
    i === 0 ? -CLIP_Y : s0 - centre - ov;
  const stripClipX1 = (i: number, s1: number, centre: number): number =>
    i === last ? CLIP_Y : s1 - centre + ov;

  // Paint the whole strip stack onto `target` in `fillStyle`. The transform
  // chain per strip is the single-affine per-glyph draw anchored at the STRIP
  // centre so vScale/shear/angle track this slice's u.
  const paintStrips = (target: CanvasRenderingContext2D, fillStyle: TextPaint): void => {
    if (strokeWidth > 0) {
      target.strokeStyle = fillStyle;
      target.lineWidth = strokeWidth;
      target.lineJoin = 'round';
    } else {
      target.fillStyle = fillStyle;
    }
    for (let i = 0; i <= last; i++) {
      const { s0, s1, g } = strips[i];
      const centre = (s0 + s1) / 2;
      target.save();
      target.translate(boxX + g.x, boxY + g.y);
      target.rotate(g.angle);
      if (g.shear !== 0) target.transform(1, 0, g.shear, 1, 0, 0);
      if (hScale !== 1 || g.vScale !== 1) target.scale(hScale, g.vScale);
      target.beginPath();
      const x0 = stripClipX0(i, s0, centre);
      const x1 = stripClipX1(i, s1, centre);
      target.rect(x0, -CLIP_Y, x1 - x0, 2 * CLIP_Y);
      target.clip();
      // Pen origin sits at local −centre; the ls/2 shift centres the ink inside
      // its letter-spaced advance, matching the flat draw's origin convention.
      if (strokeWidth > 0) target.strokeText(ch, -centre + ls / 2, 0);
      else target.fillText(ch, -centre + ls / 2, 0);
      target.restore();
    }
  };

  // ── Composite path selection (issue #879) ──────────────────────────────────
  // The overlapping strip clips are idempotent only for a FULLY OPAQUE composite.
  // Effective opacity = the fill's own alpha × the inherited ctx.globalAlpha
  // (shape opacity). When both are 1 the direct draw is exact and byte-identical
  // to the pre-#879 path — the overwhelmingly common WordArt case, zero
  // allocation. Otherwise the 1-device-px overlap band would double-compose and
  // darken, so route through the layer below.
  const fillAlpha = typeof paint === 'string' ? rgbaAlpha(paint) : 1;
  const destAlpha = typeof ctx.globalAlpha === 'number' ? ctx.globalAlpha : 1;
  if (typeof paint === 'string' && fillAlpha >= 1 && destAlpha >= 1) {
    paintStrips(ctx, paint);
    return;
  }
  if (fillAlpha <= 0 || destAlpha <= 0) return; // fully transparent — nothing to draw

  // Translucent: paint the strips OPAQUE into a per-glyph device-resolution layer
  // (so the overlap band is idempotent again) and composite the layer ONCE with
  // the effective alpha. The layer is rasterised under the SAME device transform
  // as the live canvas, so there is no re-magnification of any raster — the fix
  // the issue calls for. Falls back to the direct draw (today's slightly darker
  // band) when no auxiliary canvas or live transform is available (a headless
  // mock ctx), which is never pixel-verified anyway.
  const base = typeof ctx.getTransform === 'function' ? ctx.getTransform() : null;
  if (!base) {
    paintStrips(ctx, paint);
    return;
  }

  // Device-space AABB of the actual ink: per strip, the visible ink is the
  // measured ink rectangle intersected with the strip's clip interval, mapped
  // through the strip affine and the live (css→device) transform. Bounding by
  // measured ink — not the ±CLIP_Y clip — keeps the layer glyph-tight.
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (let i = 0; i <= last; i++) {
    const { s0, s1, g } = strips[i];
    const centre = (s0 + s1) / 2;
    const textX = -centre + ls / 2;
    const x0 = Math.max(stripClipX0(i, s0, centre), textX - inkL);
    const x1 = Math.min(stripClipX1(i, s1, centre), textX + inkR);
    if (x1 <= x0) continue; // no ink in this strip's clip
    for (const [lx, ly] of [
      [x0, -asc],
      [x1, -asc],
      [x0, desc],
      [x1, desc],
    ] as const) {
      const p = mapWarpLocalPoint(g, hScale, lx, ly);
      const cx = boxX + p.x;
      const cy = boxY + p.y;
      const dx = base.a * cx + base.c * cy + base.e;
      const dy = base.b * cx + base.d * cy + base.f;
      if (dx < minX) minX = dx;
      if (dx > maxX) maxX = dx;
      if (dy < minY) minY = dy;
      if (dy > maxY) maxY = dy;
    }
  }
  if (!(maxX > minX && maxY > minY)) return; // no ink mapped — nothing to composite

  const pad = 2 + strokeWidth * devScale; // AA and outline around the ink AABB
  const originX = Math.floor(minX - pad);
  const originY = Math.floor(minY - pad);
  const layerW = Math.ceil(maxX + pad) - originX;
  const layerH = Math.ceil(maxY + pad) - originY;
  const aux = createAuxCanvas(layerW, layerH);
  const auxCtx = aux ? (aux.getContext('2d') as CanvasRenderingContext2D | null) : null;
  if (!aux || !auxCtx) {
    paintStrips(ctx, paint);
    return;
  }
  // The layer's own transform is the live transform translated by −origin (a
  // pure device-space shift, so it just subtracts origin from the base
  // translation — no DOMMatrix multiply needed), placing device pixels on the
  // same grid as the main canvas but offset into the layer.
  auxCtx.font = ctx.font;
  auxCtx.textAlign = 'left';
  auxCtx.textBaseline = 'alphabetic';
  auxCtx.setTransform(base.a, base.b, base.c, base.d, base.e - originX, base.f - originY);
  if (typeof paint === 'string') {
    paintStrips(auxCtx, opaqueRgba(paint));
  } else {
    // A patterned warp can have translucent cells. Build the overlapping strip
    // stack as one opaque glyph mask, then colour that mask once on the slide
    // grid. This avoids double-compositing where adjacent strip clips overlap.
    paintStrips(auxCtx, '#fff');
    auxCtx.save();
    auxCtx.setTransform(1, 0, 0, 1, 0, 0);
    auxCtx.globalCompositeOperation = 'source-in';
    withInheritedPatternScope(ctx, auxCtx, () => {
      auxCtx.fillStyle = resolveAuxPaint?.(auxCtx) ?? paint;
      auxCtx.fillRect(0, 0, layerW, layerH);
    }, { x: originX, y: originY });
    auxCtx.restore();
  }

  // Composite the opaque layer once at the effective alpha, at identity so the
  // (already device-oriented) layer blits 1:1. save/restore preserves the live
  // transform, clip, and globalAlpha for the next glyph.
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.globalAlpha = destAlpha * fillAlpha;
  ctx.drawImage(aux, originX, originY);
  ctx.restore();
}

/** Alpha (0..1) of an `rgb()/rgba()` colour string. The pptx text pipeline emits
 *  colours via {@link hexToRgba} (`rgba(r,g,b,a)`); a 3-component `rgb()` or any
 *  unparseable value is treated as opaque so it never routes into the #879 layer. */
function rgbaAlpha(color: string): number {
  const m = /^rgba?\(\s*[\d.]+\s*,\s*[\d.]+\s*,\s*[\d.]+\s*,\s*([\d.]+)\s*\)$/i.exec(color);
  if (!m) return 1;
  const a = parseFloat(m[1]);
  return Number.isFinite(a) ? Math.min(1, Math.max(0, a)) : 1;
}

/** Drop the alpha channel of an `rgb()/rgba()` colour, yielding an opaque
 *  `rgb(r,g,b)`. Used to paint the #879 compositing layer at full opacity. */
function opaqueRgba(color: string): string {
  const m = /^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)/i.exec(color);
  return m ? `rgb(${m[1]}, ${m[2]}, ${m[3]})` : color;
}

/** One strip of a glyph: flat sub-range `[s0, s1]` (css px from the glyph's flat
 *  left edge / pen origin) and the envelope frame sampled at its centre-u. */
interface WarpStrip {
  s0: number;
  s1: number;
  g: WarpGlyphTransform;
}

/** Uniform-in-u partition of a glyph's advance into `k` strips, each carrying
 *  the envelope frame at its own centre-u. */
function buildWarpStrips(
  k: number,
  env: WarpEnvelope,
  chW: number,
  flatX0: number,
  totalW: number,
  followScale: number,
  warpBoxH: number,
  bandFrac: number,
): WarpStrip[] {
  const strips: WarpStrip[] = new Array(k);
  for (let j = 0; j < k; j++) {
    const s0 = (j / k) * chW;
    const s1 = ((j + 1) / k) * chW;
    const u = ((flatX0 + (s0 + s1) / 2) / totalW) * followScale;
    strips[j] = { s0, s1, g: warpGlyphTransform(env, u, warpBoxH, bandFrac) };
  }
  return strips;
}

/** Apply a strip's affine (scale → shear → rotate → translate, the same chain
 *  the draw path pushes onto the canvas) to a strip-local css point. */
function mapWarpLocalPoint(
  g: WarpGlyphTransform,
  hScale: number,
  x: number,
  y: number,
): { x: number; y: number } {
  const sx = x * hScale;
  const sy = y * g.vScale;
  const hx = sx + g.shear * sy;
  const c = Math.cos(g.angle);
  const s = Math.sin(g.angle);
  return { x: g.x + c * hx - s * sy, y: g.y + s * hx + c * sy };
}

/**
 * Worst-case deviation (device px) of the strips' painted slab corners from the
 * EXACT envelope map — the quantity the adaptive loop drives under the budget.
 *
 * For each strip edge b (a boundary of the flat sub-range), the exact map
 * places the glyph ruling at `Q(b, y) = A(u_b)·(0, y)` — the envelope frame
 * sampled AT the boundary — while the strip paints it at
 * `S(b, y) = A(u_centre)·(b − centre, y)`. `|S − Q|` is evaluated at the slab's
 * top and bottom (`yTop`/`yBot`, css, baseline-relative) for both edges of
 * every strip; the max over all of them bounds every painted pixel's error
 * (corners are the extreme points of an affine image of a rectangle).
 */
function maxWarpStripDeviationDev(
  strips: WarpStrip[],
  env: WarpEnvelope,
  flatX0: number,
  totalW: number,
  followScale: number,
  warpBoxH: number,
  bandFrac: number,
  hScale: number,
  devScale: number,
  yTop: number,
  yBot: number,
): number {
  let max = 0;
  for (const strip of strips) {
    const centre = (strip.s0 + strip.s1) / 2;
    for (const b of [strip.s0, strip.s1]) {
      const uB = ((flatX0 + b) / totalW) * followScale;
      const gB = warpGlyphTransform(env, uB, warpBoxH, bandFrac);
      for (const y of [yTop, yBot]) {
        const q = mapWarpLocalPoint(gB, hScale, 0, y);
        const p = mapWarpLocalPoint(strip.g, hScale, b - centre, y);
        const d = Math.hypot(p.x - q.x, p.y - q.y) * devScale;
        if (d > max) max = d;
      }
    }
  }
  return max;
}

/**
 * Render a WordArt text body through its `<a:prstTxWarp>` envelope (ECMA-376
 * §20.1.9.19). Canvas 2D cannot deform glyph outlines, so this is the standard
 * PER-GLYPH approximation: the text is laid out flat (reusing
 * {@link layoutParagraph} so run font/size/colour resolution is identical to the
 * normal path), then each glyph is placed against the preset's flattened
 * envelope via {@link warpGlyphTransform} and drawn with a local
 * translate/rotate/scale.
 *
 * The envelope spans the shape's box (minus insets). Text is fitted to that
 * width: a glyph whose centre sits at horizontal fraction `u` of the total text
 * width is mapped to the envelope point at `u`. Multiple lines stack vertically
 * inside the box, each occupying its own horizontal band of the envelope's
 * height — enough for the common single-line WordArt title.
 *
 * onTextRun is intentionally NOT invoked here: the transparent text-selection
 * overlay assumes axis-aligned run boxes, which per-glyph warped text violates.
 * Warped WordArt is decorative; skipping the overlay leaves it unselectable
 * rather than mis-boxed. This is called only when a warp preset is present, so
 * all non-WordArt text keeps its overlay.
 */
export function renderWarpedText(
  ctx: CanvasRenderingContext2D,
  body: TextBody,
  preset: string,
  adj: number[],
  bx: number,
  by: number,
  bw: number,
  bh: number,
  scale: number,
  defaultColor: string,
  rc: RenderContext,
): void {
  // The warp envelope spans the FULL shape bounding box, NOT the text rect
  // inset by lIns/tIns/rIns/bIns. The preset guide formulas (ECMA-376
  // §20.1.9.19 / presetTextWarpDefinitions.xml) are written against the
  // shape's `w`/`h` built-ins, and PowerPoint's own PDF output of the warp
  // fixture confirms it: paired-edge warp ink spans 6.04–6.18in of a 6.2in
  // shape (insets would cap it at 6.0) and up to 1.48in of a 1.5in-tall one.
  const boxX = bx;
  const boxY = by;
  const boxW = Math.max(1, bw);
  const boxH = Math.max(1, bh);

  const env = buildWarpEnvelope(preset, adj, boxW, boxH);
  if (!env) return; // unknown preset — nothing to draw (caller already gated)

  const bodyDefaultBold = body.defaultBold ?? false;
  const bodyDefaultItalic = body.defaultItalic ?? false;
  const bodyDefaultFontSizePx = (body.defaultFontSize ?? 18) * PT_TO_EMU * scale;

  // Lay each paragraph out flat at the natural size (no wrap — WordArt fits the
  // shape width itself). Collect every line's segments in order.
  const lines: Array<{ line: LayoutLine; alignment: Paragraph['alignment']; baseline: number }> = [];
  let flatTop = 0;
  for (const para of body.paragraphs) {
    const paraDefaultFontSizePx =
      para.defFontSize != null ? para.defFontSize * PT_TO_EMU * scale : bodyDefaultFontSizePx;
    const paraDefaultColor = para.defColor ? hexToRgba(para.defColor) : defaultColor;
    const laid = layoutParagraph(
      ctx,
      para,
      Infinity, // no wrap: the envelope maps the whole run onto the width
      paraDefaultFontSizePx,
      paraDefaultColor,
      scale,
      0,
      bodyDefaultBold,
      bodyDefaultItalic,
      1.0,
      undefined,
      rc,
      0,
    );
    if (lines.length > 0) flatTop += (para.spaceBefore ?? 0) / 100 * PT_TO_EMU * scale;
    for (const line of laid) {
      const size = line.segments.reduce((max, seg) => Math.max(max, seg.sizePx), 0) || paraDefaultFontSizePx;
      // Same natural line box as ordinary PowerPoint text. The independent
      // flat/Follow Path Office controls cover 50–200%, fixed point spacing,
      // mixed sizes and paragraph gaps; no warp-specific pitch multiplier.
      const height = para.spaceLine?.type === 'pts'
        ? para.spaceLine.val * PT_TO_EMU * scale
        : size * 1.2 * (para.spaceLine ? para.spaceLine.val / 100000 : 1);
      lines.push({ line, alignment: para.alignment, baseline: flatTop + height * 0.8 });
      flatTop += height;
    }
    flatTop += (para.spaceAfter ?? 0) / 100 * PT_TO_EMU * scale;
  }
  if (lines.length === 0) return;

  ctx.save();
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';

  // Device scale folded into the live transform (dpr · slide scale), so glyph
  // bitmaps for the paired-edge strip path are rasterised at the same pixel
  // density as the canvas — no HiDPI blur. sqrt(|det|) of the transform's linear
  // part is dpr regardless of rotation/flip (same idiom as the scene3d/effect
  // offscreens). Computed LAZILY on the first paired-edge glyph only: single-edge
  // (Follow Path) warps never allocate an offscreen, and some unit tests drive a
  // mock ctx without getTransform, so the single-edge path must never touch it.
  let devScaleCache = -1;
  const getDevScale = (): number => {
    if (devScaleCache >= 0) return devScaleCache;
    const tf = typeof ctx.getTransform === 'function' ? ctx.getTransform() : null;
    const det = tf ? Math.abs(tf.a * tf.d - tf.b * tf.c) : 1;
    devScaleCache = det > 0 ? Math.sqrt(det) : 1;
    return devScaleCache;
  };

  const lineCount = lines.length;
  // Follow Path is not the paired-edge vertical-band operation in §20.1.9.19.
  // Office-produced controls show concentric ellipses: clockwise paths anchor
  // the last line, counterclockwise paths the first. Derive winding from the
  // authored path, not screen up/down (adjusted arches can cross either half).
  const p0 = env.top[0];
  const p1 = env.top[1];
  const clockwise = p0 && p1
    ? (p0.x - boxW / 2) * (p1.y - boxH / 2) - (p0.y - boxH / 2) * (p1.x - boxW / 2) > 0
    : true;
  const anchoredBaseline = lines[clockwise ? lineCount - 1 : 0]!.baseline;
  for (let li = 0; li < lineCount; li++) {
    const { line, alignment } = lines[li];
    // Per-line vertical band [v0, v1] of the envelope's height. One line fills
    // the whole band; multiple lines split it evenly top→bottom.
    const v0 = li / lineCount;
    const v1 = (li + 1) / lineCount;

    // Measure the line: total advance width, max font size, and the line's INK
    // extents (max actualBoundingBox ascent/descent over the segments). The ink
    // box — not the em box — is what PowerPoint normalises onto the envelope.
    let totalW = 0;
    let maxSize = 0;
    let maxA = 0; // ink ascent above the baseline
    let maxD = 0; // ink descent below the baseline
    for (const seg of line.segments) {
      if (seg.math) {
        totalW += seg.math.width;
        maxSize = Math.max(maxSize, seg.sizePx);
        maxA = Math.max(maxA, seg.math.ascent);
        maxD = Math.max(maxD, seg.math.descent);
        continue;
      }
      ctx.font = seg.font;
      const ls = seg.letterSpacingPx ?? 0;
      const m = ctx.measureText(seg.text);
      totalW += m.width + ls * codePointCount(seg.text);
      maxSize = Math.max(maxSize, seg.sizePx);
      if (m.actualBoundingBoxAscent > 0) maxA = Math.max(maxA, m.actualBoundingBoxAscent);
      if (m.actualBoundingBoxDescent > 0) maxD = Math.max(maxD, m.actualBoundingBoxDescent);
    }
    if (totalW <= 0) continue;
    // Empty manual lines affect authored baseline spacing, but have no glyphs
    // to map. Avoid rebuilding a sampled curve for them.
    const outset = env.singleEdge
      ? Math.max(0, clockwise ? anchoredBaseline - lines[li]!.baseline : lines[li]!.baseline - anchoredBaseline)
      : 0;
    const lineEnv = outset > 0
      ? buildWarpEnvelope(preset, adj, boxW + 2 * outset, boxH + 2 * outset) ?? env
      : env;

    // PowerPoint's WordArt semantics for the envelope (paired-edge) presets:
    // the FLAT text's INK RECTANGLE (totalW × ink height) is first STRETCHED to
    // fill the shape box, and the warp then bends that stretched rectangle
    // between the two edges. Measured against PowerPoint's own PDF of the warp
    // fixture (calibrated print scale), paired-edge ink spans 6.04–6.18in of
    // the 6.2in shape and up to 1.48in of the 1.5in one — i.e. the ink touches
    // the envelope edges. Passing the body height here instead would collapse
    // vScale = gap/boxH ≈ 1 and leave glyphs at their tiny natural size.
    //
    // - Vertical: the line's flat ink box (maxA+maxD px, measured — no 0.8em
    //   heuristic) maps to this line's SHARE of the local gap.
    //   warpGlyphTransform returns vScale = gap/boxHeight, so passing
    //   `inkH / (v1 - v0)` yields gap·(v1−v0)/inkH, and a baseline fraction of
    //   maxA/inkH pins the ink top to T(u) and the ink bottom to B(u).
    // - Horizontal: the glyph itself is widened by boxW/totalW (not just the
    //   spacing). Glyph centres sit Δu·boxW ≈ chW·hScale apart, so scaled
    //   glyphs tile the envelope without gaps or overlaps.
    //
    // Single-edge (arch/circle) presets keep natural glyph size: PowerPoint's
    // "Follow Path" semantics place the text at its NATURAL width along the
    // arc without stretching the glyphs (vScale stays 1 inside
    // warpGlyphTransform); they keep the flat renderer's 0.8 ascent fallback
    // for the baseline drop below the arc. A glyph's `u` fraction is scaled by
    // naturalWidth/arcLength via `followPathUScale`, and paragraph alignment
    // chooses where that compact segment sits on the path. Paired-edge presets
    // return scale 1 (they stretch the flat ink box to fill the envelope width).
    const inkH = maxA + maxD > 0 ? maxA + maxD : maxSize;
    const baselineFrac = env.singleEdge ? 0.8 : inkH > 0 ? maxA / inkH : 0.8;
    const hScale = env.singleEdge ? 1 : boxW / totalW;
    // `warpGlyphTransform` expects the FLAT line box height. For Follow Path
    // presets this determines only the small baseline-to-curve descent. Passing
    // the entire shape height displaced short text by tens of pixels normal to
    // the arc (and, after shape rotation, outside the intended ring segment).
    const warpBoxH = env.singleEdge ? inkH : inkH / (v1 - v0);
    // Follow Path: fraction of the arc the natural-width text actually spans.
    // 1 for paired-edge (no clamp); ≤1 for arch/circle.
    const followScale = followPathUScale(lineEnv, totalW);
    // A Follow Path warp keeps the text's natural arc length, but its paragraph
    // alignment still selects where that shorter segment sits on the path.
    // PowerPoint-generated circular diagrams commonly use centred paragraphs;
    // pinning every short label to stAng piles the labels onto a wedge edge.
    const followOffset = env.singleEdge
      ? (1 - followScale) * (alignment === 'r' ? 1 : alignment === 'ctr' ? 0.5 : 0)
      : 0;

    // Walk glyphs left→right. `penW` accumulates the flat advance so each glyph's
    // CENTRE maps to its u fraction; the glyph is drawn at a per-glyph transform.
    let penW = 0;
    for (const seg of line.segments) {
      const segmentStartW = penW;
      const segmentBandFrac = env.singleEdge ? baselineFrac : v0 + baselineFrac * (v1 - v0);
      if (seg.math) {
        // Equations inside WordArt are exotic; advance without warping them.
        penW += seg.math.width;
        continue;
      }
      ctx.font = seg.font;
      const glyphPaint = resolveSegmentTextPaint(ctx, seg, boxX, boxY, scale);
      ctx.fillStyle = glyphPaint;
      const outline = seg.outline && seg.outline.width > 0 ? seg.outline : undefined;
      const outlineFill = outline?.fill;
      const outlinePaint = outline
        ? outlineFill
          ? resolveFillCore(outlineFill, ctx, boxX, boxY, boxW, boxH, 0, scale * PT_TO_EMU)
          : outline.color ? `#${outline.color}` : seg.noFill ? undefined : glyphPaint
        : undefined;
      const outlineWidth = outline ? Math.max(0.5, emuToPx(outline.width, scale)) : 0;
      const ls = seg.letterSpacingPx ?? 0;
      const chars = [...seg.text];
      for (const ch of chars) {
        const chW = ctx.measureText(ch).width + ls;
        if (seg.noFill && !outlinePaint) { penW += chW; continue; }
        // Blend the per-line vertical band into the baseline fraction so line 2
        // sits below line 1 within the envelope.
        const bandFrac = env.singleEdge ? baselineFrac : v0 + baselineFrac * (v1 - v0);

        // Paired-edge (envelope) presets: the envelope map is NON-affine within a
        // glyph (on Inflate/Deflate the vertical stretch varies across the glyph's
        // own width; on waves the slope does), so a single per-glyph affine loses
        // that intra-glyph deformation. Draw the glyph via piecewise-affine STRIPS
        // (each strip a clipped vector fillText under its own affine, sampled at
        // its own centre-u) so the mapping converges to PowerPoint's continuous
        // outline warp (§20.1.9.19). Needs no auxiliary canvas, so it runs
        // unconditionally (main thread, worker, headless node alike). The
        // single-edge (Follow Path) branch below keeps rigid per-glyph rotation;
        // its separate placement rules preserve natural size and alignment.
        if (!env.singleEdge && chW > 0) {
          if (!seg.noFill) {
            drawWarpedGlyphStrips(
              ctx, ch, ls, chW, getDevScale(), env, penW, totalW, followScale,
              hScale, warpBoxH, bandFrac, boxX, boxY, glyphPaint,
              seg.patternFill
                ? target => resolveSegmentTextPaint(target, seg, boxX, boxY, scale)
                : undefined,
            );
          }
          if (outlinePaint && outline) {
            drawWarpedGlyphStrips(
              ctx, ch, ls, chW, getDevScale(), env, penW, totalW, followScale,
              hScale, warpBoxH, bandFrac, boxX, boxY, outlinePaint,
              outlineFill
                ? target => resolveFillCore(outlineFill, target, boxX, boxY, boxW, boxH,
                  0, scale * PT_TO_EMU) ?? 'rgba(0,0,0,0)'
                : undefined,
              outlineWidth,
            );
          }
          penW += chW;
          continue;
        }

        // Horizontal fraction of THIS glyph's centre along the whole line,
        // scaled and offset so single-edge (arch/circle) text spans only its
        // naturally sized, paragraph-aligned arc segment rather than the whole
        // path. `followScale` is 1 for paired-edge presets (unchanged).
        const u = followOffset + ((penW + chW / 2) / totalW) * followScale;
        const g = warpGlyphTransform(lineEnv, u, warpBoxH, bandFrac);
        g.x -= outset;
        g.y -= outset;
        ctx.save();
        ctx.translate(boxX + g.x, boxY + g.y);
        ctx.rotate(g.angle);
        // Local envelope shear (§20.1.9.19): on a sloped paired edge the em-height
        // axis (gap vector B−T) is not perpendicular to the advance axis, so the
        // glyph must skew — vertical strokes track the gap while the baseline
        // follows the slope. `g.shear` is the horizontal skew in the rotated
        // frame; 0 for flat points and single-edge (arch/circle) presets, so this
        // is a no-op there and leaves Follow Path glyphs rigidly rotated.
        if (g.shear !== 0) ctx.transform(1, 0, g.shear, 1, 0, 0);
        if (hScale !== 1 || g.vScale !== 1) ctx.scale(hScale, g.vScale);
        // Draw the glyph centred on the mapped point: shift left by half its
        // advance so its own centre lands on `u`.
        if (!seg.noFill) ctx.fillText(ch, -chW / 2 + ls / 2, 0);
        if (outlinePaint) {
          ctx.strokeStyle = outlinePaint;
          ctx.lineWidth = outlineWidth;
          ctx.lineJoin = 'round';
          ctx.strokeText(ch, -chW / 2 + ls / 2, 0);
        }
        ctx.restore();
        penW += chW;
      }
      // A separately authored underline remains visible for noFill WordArt.
      // Trace it on the same warped baseline as the glyphs; uFill and uLn are
      // independent CT_TextCharacterProperties children (§21.1.2.3.9).
      if (seg.underline && !seg.underlineLineNoFill && penW > segmentStartW
        && (!seg.noFill || seg.underlineFill || seg.underlineColor || seg.underlineLine?.fill)) {
        const underlinePaint = seg.underlineFill
          ? resolveFillCore(seg.underlineFill, ctx, boxX, boxY, boxW, boxH,
              0, scale * PT_TO_EMU)
          : seg.underlineLine?.fill
            ? resolveFillCore(seg.underlineLine.fill, ctx, boxX, boxY, boxW, boxH,
                0, scale * PT_TO_EMU)
            : seg.underlineColor ?? glyphPaint;
        if (underlinePaint) {
          ctx.save();
          ctx.strokeStyle = underlinePaint;
          ctx.lineWidth = seg.underlineLine?.width
            ? emuToPx(seg.underlineLine.width, scale)
            : Math.max(1, seg.sizePx * DECORATION_LINE_WIDTH_EM);
          ctx.beginPath();
          const samples = Math.max(2, Math.ceil((penW - segmentStartW) / 8));
          for (let sample = 0; sample <= samples; sample++) {
            const flat = segmentStartW + (penW - segmentStartW) * sample / samples;
            const u = followOffset + flat / totalW * followScale;
            const g = warpGlyphTransform(lineEnv, u, warpBoxH, segmentBandFrac);
            const offset = Math.max(2, ctx.lineWidth) * g.vScale;
            const x = boxX + g.x - outset - Math.sin(g.angle) * offset;
            const y = boxY + g.y + Math.cos(g.angle) * offset;
            if (sample === 0) ctx.moveTo(x, y);
            else ctx.lineTo(x, y);
          }
          ctx.stroke();
          ctx.restore();
        }
      }
    }
  }
  ctx.restore();
}
