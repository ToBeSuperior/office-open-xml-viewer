// XLSX anchored drawings: pictures, shapes and shape text, equations, and
// charts. Split from renderer.ts (#1857).

import { drawShapeStackedText } from './shape-stacked-text.js';
import type { CjkLang } from '@silurus/ooxml-core';
import {
  breakDrawingMlText,
  drawingMlTextRect,
  drawingMlLineHeight,
  drawingMlSpacedLineBox,
  drawingMlParagraphSpacing,
  drawingMlBlockTop,
  drawingMlLineX,
  paintDrawingMlLine,
  type DrawingMlInputRun,
} from '@silurus/ooxml-core/internal/drawingml-text';
import type {
  Worksheet,
  ShapeInfo,
  ShapeAnchor,
  ImageAnchor,
  ChartAnchor,
  Duotone,
  PathInfo,
} from './types.js';
import type {
  Stroke,
  ChartThreeDRenderer,
  ChartRegionMapRenderer,
  ChartExRenderer,
} from '@silurus/ooxml-core';
import {
  usesPathShade,
  trackPaintPath,
  currentStrokeBounds,
  chartImageFillKey,
  paintOptionalImagePlaceholder,
  pathFillModeOverlay,
  withDrawingMLShapeTransform,
  buildPresetGeometryFillPath,
} from '@silurus/ooxml-core';
import {
  renderChart,
  renderPresetShape,
  PT_TO_PX,
  EMU_PER_PX,
  mathToMathML,
  rasterizeMathSvg,
  cachedTintedMathRaster,
  isCjkBreakChar,
  drawImageCropped,
  applyStroke,
  resolveFill,
  type MathNode,
  type MathRenderer,
  type RasterizedMathSvg,
} from '@silurus/ooxml-core';
import {
  canvasShapeFontBoxProbe,
  excelShapeSpacing,
  shapeOfficeRouteKey,
  shapeRunLineRatios,
  type ShapeRunLineRatios,
} from './shape-office-line.js';
import type { GridAxisGeometry } from './internal/grid-axis-geometry.js';
import { resolveWorksheetAnchorRect } from './internal/initial-anchor-sizes.js';
import { isOptionalImageUnavailable } from './internal/optional-image-fallback.js';
import { rotatedImageBounds } from './internal/image-anchor-transform.js';
import {
  fontStackFor,
  officeRoute,
  googleSubstitutesByContext,
  contextRegularAlias,
  DEFAULT_FONT_SIZE,
  sheetAnchoredRectX,
  officeRoutesByContext,
} from './renderer-font-metrics.js';

/** Cache key for a decoded image in the shared `loadedImages` map. A plain
 *  picture is keyed by its zip `imagePath`; a picture carrying a `<a:duotone>`
 *  effect (ECMA-376 §20.1.8.23) is keyed by the path PLUS both resolved endpoint
 *  colours, so a recoloured bitmap is cached and looked up separately from the
 *  raw blip. The render-orchestrator's `prefetchImages` stores each decoded
 *  source under this exact key, and the renderer computes it per anchor for the
 *  synchronous lookup. Kept here (not in the orchestrator) so both the
 *  synchronous renderer and the async orchestrator share one definition without
 *  a circular import. */
export function imageCacheKey(imagePath: string, duotone?: Duotone | null): string {
  return duotone ? `${imagePath}|duo:${duotone.clr1}:${duotone.clr2}` : imagePath;
}

// ────────────────────────────────────────────────────────────────
// Image anchors  (ECMA-376 §20.5, <xdr:twoCellAnchor>)
// ────────────────────────────────────────────────────────────────

type AnchoredDrawing =
  | { kind: 'image'; zOrder: number; fallbackOrder: number; anchor: ImageAnchor }
  | { kind: 'shape'; zOrder: number; fallbackOrder: number; anchor: ShapeAnchor; shape: ShapeInfo }
  | { kind: 'chart'; zOrder: number; fallbackOrder: number; anchor: ChartAnchor };

/** Paint worksheet drawing objects in their DrawingML document order.
 *
 * ECMA-376 §20.5.2.35 defines the worksheet drawing as an ordered sequence;
 * later objects are visually above earlier objects. A grouped chart can contain
 * a chart frame followed by title/footer text boxes, so rendering by object type
 * (all shapes, then all charts) incorrectly covered those text boxes with the
 * chart background. Parser-emitted `zOrder` values are byte positions in the
 * selected drawing-part XML and therefore preserve that authored order. */
export function renderAnchoredDrawings(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  colAxis: GridAxisGeometry,
  rowAxis: GridAxisGeometry,
  loadedImages: Map<string, CanvasImageSource | null> | undefined,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
  rtl: boolean,
  canvasW: number,
  threeD?: ChartThreeDRenderer,
  regionMap?: ChartRegionMapRenderer,
  chartEx?: ChartExRenderer,
  cjkFallback?: CjkLang,
): void {
  const drawings: AnchoredDrawing[] = [];
  let fallbackOrder = 0;
  if (loadedImages) {
    for (const anchor of ws.images ?? []) {
      drawings.push({ kind: 'image', zOrder: anchor.zOrder ?? Number.MAX_SAFE_INTEGER, fallbackOrder: fallbackOrder++, anchor });
    }
  }
  for (const anchor of ws.shapeGroups ?? []) {
    for (const shape of anchor.shapes) {
      drawings.push({ kind: 'shape', zOrder: shape.zOrder ?? Number.MAX_SAFE_INTEGER, fallbackOrder: fallbackOrder++, anchor, shape });
    }
  }
  for (const anchor of ws.charts ?? []) {
    drawings.push({ kind: 'chart', zOrder: anchor.zOrder ?? Number.MAX_SAFE_INTEGER, fallbackOrder: fallbackOrder++, anchor });
  }
  drawings.sort((a, b) => a.zOrder - b.zOrder || a.fallbackOrder - b.fallbackOrder);

  for (const drawing of drawings) {
    if (drawing.kind === 'image') {
      renderImages(
        ctx, ws, colAxis, rowAxis, loadedImages as Map<string, CanvasImageSource | null>, cs,
        startRow, startCol, scrollOffsetX, scrollOffsetY,
        scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH, rtl, canvasW,
        [drawing.anchor],
      );
    } else if (drawing.kind === 'shape') {
      renderShapeGroups(
        ctx, ws, colAxis, rowAxis, cs,
        startRow, startCol, scrollOffsetX, scrollOffsetY,
        scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH,
        loadedImages, rtl, canvasW,
        [{ ...drawing.anchor, shapes: [drawing.shape] }], cjkFallback
      );
    } else {
      renderCharts(
        ctx, ws, colAxis, rowAxis, cs,
        startRow, startCol, scrollOffsetX, scrollOffsetY,
        scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH, rtl, canvasW,
        loadedImages,
        [drawing.anchor],
        threeD,
        regionMap,
        chartEx,
      );
    }
  }
}

/** Sum scaled column widths for cols 1..n-1 (sheet-space X of col n in scaled px). */
export function sheetXForCol(
  axis: GridAxisGeometry,
  col1: number, // 1-indexed column number
): number {
  return axis.offsetOf(col1);
}

/** Sum scaled row heights for rows 1..n-1 (sheet-space Y of row n in scaled px). */
export function sheetYForRow(
  axis: GridAxisGeometry,
  row1: number, // 1-indexed row number
): number {
  return axis.offsetOf(row1);
}


function renderImages(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  colAxis: GridAxisGeometry,
  rowAxis: GridAxisGeometry,
  loadedImages: Map<string, CanvasImageSource | null>,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
  rtl: boolean,
  canvasW: number,
  anchors: readonly ImageAnchor[] = ws.images,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;

  // Sheet-space origin of the current scroll viewport's first visible cell
  const scrollOriginSheetX = sheetXForCol(colAxis, startCol);
  const scrollOriginSheetY = sheetYForRow(rowAxis, startRow);

  ctx.save();
  ctx.beginPath();
  const clipX = sheetAnchoredRectX(scrollAreaX, scrollAreaW, canvasW, rtl);
  ctx.rect(clipX, scrollAreaY, scrollAreaW, scrollAreaH);
  ctx.clip();

  for (const anchor of anchors) {
    // A `<a:duotone>` picture was recoloured at decode time and cached under a
    // colour-suffixed key (§20.1.8.23); look it up with the same key.
    const lookupKey = imageCacheKey(anchor.imagePath, anchor.duotone);
    const img = loadedImages.get(lookupKey);
    const tiffUnavailable = isOptionalImageUnavailable(loadedImages, lookupKey, 'tiff');
    if (!img && !tiffUnavailable) continue;

    // One resolver owns the display rectangle for paint, hit-testing and
    // culling/decode. The normative anchor facts and the library policies
    // (tagged initial rect, untagged compatibility) are documented in
    // internal/cell-anchor-geometry.ts.
    const {
      x: imgSheetX1, y: imgSheetY1, width: imgW, height: imgH,
    } = resolveWorksheetAnchorRect(ws, anchor, colAxis, rowAxis, cs);
    if (imgW <= 0 || imgH <= 0) continue;

    // Translate to canvas coordinates of the scrollable viewport
    const logicalCanvasX = scrollAreaX + (imgSheetX1 - scrollOriginSheetX) - scrollOffsetX;
    const canvasX = sheetAnchoredRectX(logicalCanvasX, imgW, canvasW, rtl);
    const canvasY = scrollAreaY + (imgSheetY1 - scrollOriginSheetY) - scrollOffsetY;

    // A rotated rectangle can intersect the viewport even when its unrotated
    // box does not. Flips preserve the same axis-aligned bounds.
    const bounds = rotatedImageBounds(
      { x: canvasX, y: canvasY, width: imgW, height: imgH },
      anchor.rotation,
    );
    if (bounds.x + bounds.width < clipX || bounds.x > clipX + scrollAreaW) continue;
    if (bounds.y + bounds.height < scrollAreaY || bounds.y > scrollAreaY + scrollAreaH) continue;

    // ECMA-376 §20.1.8.6 `<a:alphaModFix>`: scale the picture's opacity so it
    // composites over the cells beneath it. Saved/restored so it never leaks
    // into a later anchor's draw.
    const paint = () => {
      if (img) {
        drawImageCropped(ctx, img, anchor.srcRect, canvasX, canvasY, imgW, imgH);
      } else {
        paintOptionalImagePlaceholder(ctx, 'tiff', {
          x: canvasX, y: canvasY, width: imgW, height: imgH,
        });
      }
    };
    const paintWithAlpha = () => {
      if (anchor.alpha != null && anchor.alpha < 1) {
        ctx.save();
        try {
          ctx.globalAlpha = anchor.alpha;
          paint();
        } finally {
          ctx.restore();
        }
      } else {
        paint();
      }
    };
    const rotation = anchor.rotation ?? 0;
    const flipH = anchor.flipH ?? false;
    const flipV = anchor.flipV ?? false;
    if (rotation === 0 && !flipH && !flipV) {
      paintWithAlpha();
    } else {
      withDrawingMLShapeTransform(ctx, {
        rect: { x: canvasX, y: canvasY, w: imgW, h: imgH },
        geometry: { kind: 'preset', name: 'rect', adjustments: [] },
        fill: null,
        stroke: null,
        transform: { rotationDeg: rotation, flipH, flipV },
      }, paintWithAlpha);
    }
  }

  ctx.restore();
}

function renderShapeGroups(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  colAxis: GridAxisGeometry,
  rowAxis: GridAxisGeometry,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
  loadedImages: Map<string, CanvasImageSource | null> | undefined,
  rtl: boolean,
  canvasW: number,
  anchors: readonly ShapeAnchor[] = ws.shapeGroups ?? [],
  cjkFallback?: CjkLang,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;
  if (anchors.length === 0) return;

  const scrollOriginSheetX = sheetXForCol(colAxis, startCol);
  const scrollOriginSheetY = sheetYForRow(rowAxis, startRow);

  ctx.save();
  ctx.beginPath();
  const clipX = sheetAnchoredRectX(scrollAreaX, scrollAreaW, canvasW, rtl);
  ctx.rect(clipX, scrollAreaY, scrollAreaW, scrollAreaH);
  ctx.clip();

  for (const anchor of anchors) {
    // Same resolver as renderImages. Normalized child transforms below are
    // raw and only scale into this rectangle.
    const { x: x1, y: y1, width: w, height: h } = resolveWorksheetAnchorRect(
      ws, anchor, colAxis, rowAxis, cs,
    );
    if (w <= 0 || h <= 0) continue;

    const logicalCanvasX = scrollAreaX + (x1 - scrollOriginSheetX) - scrollOffsetX;
    const canvasX = sheetAnchoredRectX(logicalCanvasX, w, canvasW, rtl);
    const canvasY = scrollAreaY + (y1 - scrollOriginSheetY) - scrollOffsetY;

    if (canvasX + w < clipX || canvasX > clipX + scrollAreaW) continue;
    if (canvasY + h < scrollAreaY || canvasY > scrollAreaY + scrollAreaH) continue;

    for (const shape of anchor.shapes) {
      const sx = canvasX + shape.x * w;
      const sy = canvasY + shape.y * h;
      const sw = shape.w * w;
      const sh = shape.h * h;
      if (sw <= 0 || sh <= 0) continue;
      drawShape(ctx, shape, sx, sy, sw, sh, cs, loadedImages, cjkFallback);
    }
  }

  ctx.restore();
}

// XLSX custom coordinates remain path-local until this format boundary.
// Both ordinary painting and gradient outlines use the same arc/Bezier rules.
function appendSpreadsheetCustomPath(
  ctx: CanvasRenderingContext2D, path: PathInfo,
  x: number, y: number, w: number, h: number,
): void {
  if (path.w <= 0 || path.h <= 0) return;
  const kx = w / path.w;
  const ky = h / path.h;
  // Track pen position for arcTo center computation.
  let penX = x, penY = y;
  // Track subpath start for close lineTo.
  let subX = x, subY = y;
  for (const cmd of path.commands) {
    switch (cmd.op) {
      case 'moveTo': {
        const px = x + cmd.x * kx, py = y + cmd.y * ky;
        ctx.moveTo(px, py);
        penX = subX = px; penY = subY = py;
        break;
      }
      case 'lineTo': {
        const px = x + cmd.x * kx, py = y + cmd.y * ky;
        ctx.lineTo(px, py);
        penX = px; penY = py;
        break;
      }
      case 'cubicBezTo': {
        const ex = x + cmd.x3 * kx, ey = y + cmd.y3 * ky;
        ctx.bezierCurveTo(
          x + cmd.x1 * kx, y + cmd.y1 * ky,
          x + cmd.x2 * kx, y + cmd.y2 * ky,
          ex, ey,
        );
        penX = ex; penY = ey;
        break;
      }
      case 'quadBezTo': {
        const ex = x + cmd.x2 * kx, ey = y + cmd.y2 * ky;
        ctx.quadraticCurveTo(x + cmd.x1 * kx, y + cmd.y1 * ky, ex, ey);
        penX = ex; penY = ey;
        break;
      }
      case 'arcTo': {
        // ECMA-376 §20.1.9.3: pen lies on ellipse at stAng;
        // derive center from pen + stAng, then sweep swAng.
        const rx = cmd.wr * kx, ry = cmd.hr * ky;
        if (rx <= 0 || ry <= 0) break;
        const stRad = (cmd.stAng / 60000) * (Math.PI / 180);
        const swRad = (cmd.swAng / 60000) * (Math.PI / 180);
        const cx = penX - Math.cos(stRad) * rx;
        const cy = penY - Math.sin(stRad) * ry;
        const endRad = stRad + swRad;
        ctx.ellipse(cx, cy, rx, ry, 0, stRad, endRad, swRad < 0);
        penX = cx + Math.cos(endRad) * rx;
        penY = cy + Math.sin(endRad) * ry;
        break;
      }
      case 'close':
        ctx.closePath();
        penX = subX; penY = subY;
        break;
    }
  }
}

function appendSpreadsheetShapeOutline(
  ctx: CanvasRenderingContext2D, shape: ShapeInfo,
  x: number, y: number, w: number, h: number,
): void {
  if (shape.geom.type === 'custom') {
    for (const path of shape.geom.paths) {
      if (path.fill !== 'none') appendSpreadsheetCustomPath(ctx, path, x, y, w, h);
    }
  } else if (shape.geom.type !== 'preset'
    || !buildPresetGeometryFillPath(ctx, shape.geom.name, x, y, w, h, shape.geom.adj ?? [])) {
    ctx.rect(x, y, w, h);
  }
}

function drawShape(
  ctx: CanvasRenderingContext2D,
  shape: ShapeInfo,
  sx: number, sy: number, sw: number, sh: number,
  cs: number,
  loadedImages?: Map<string, CanvasImageSource | null>,
  cjkFallback?: CjkLang,
): void {
  if (usesPathShade(shape.strokeFill)) ctx = trackPaintPath(ctx);
  ctx.save();
  if (shape.rot !== 0 || shape.flipH || shape.flipV) {
    ctx.translate(sx + sw / 2, sy + sh / 2);
    ctx.rotate((shape.rot * Math.PI) / 180);
    ctx.scale(shape.flipH ? -1 : 1, shape.flipV ? -1 : 1);
    ctx.translate(-sw / 2, -sh / 2);
  } else {
    ctx.translate(sx, sy);
  }

  if (shape.geom.type === 'custom') {
    const fill = resolveSpreadsheetShapeFill(ctx, shape, sw, sh, cs);
    for (const path of shape.geom.paths) {
      if (path.w <= 0 || path.h <= 0) continue;
      ctx.beginPath();
      appendSpreadsheetCustomPath(ctx, path, 0, 0, sw, sh);
      // ECMA-376 §20.1.9.15: each custom path carries its own fill mode and
      // stroke flag. `fill="none"` leaves the path unfilled and `stroke="0"`
      // unstroked. The lighten/darken modes shade the fill by the amounts
      // measured from PowerPoint's output (shared with the preset engine).
      if (path.fill !== 'none' && fillShape(ctx, shape, sw, sh, cs, fill)) {
        const overlay = pathFillModeOverlay(path.fill);
        if (overlay) {
          ctx.fillStyle = overlay;
          ctx.fill();
        }
      }
      if (path.stroke !== false) strokeShapePath(ctx, shape, sw, sh, cs);
    }
  } else if (shape.geom.type === 'preset') {
    // Drive the shape off the ECMA-376 §20.1.9 spec-driven preset engine
    // (presets.json from presetShapeDefinitions.xml). It honours each path's
    // own fill/stroke attributes, so a parallelogram / rtTriangle / callout
    // renders with its true outline instead of the old rect fallback. The
    // engine returns false for presets it doesn't carry; only then do we fall
    // back to a plain rectangle.
    const baseFill = resolveSpreadsheetShapeFill(ctx, shape, sw, sh, cs);
    const applyAndStroke = shape.strokeColor && shape.strokeWidth > 0
      ? () => strokeShapePath(ctx, shape, sw, sh, cs)
      : null;
    const drawn = renderPresetShape(
      ctx,
      shape.geom.name,
      0,
      0,
      sw,
      sh,
      shape.geom.adj ?? [],
      baseFill,
      applyAndStroke,
      // xlsx shapes carry no drop shadow on the body, so there's nothing to
      // clear between the engine's fill and stroke passes.
      () => {},
    );
    if (!drawn) {
      ctx.beginPath();
      ctx.rect(0, 0, sw, sh);
      fillAndStroke(ctx, shape, sw, sh, cs);
    }
  } else if (shape.geom.type === 'image') {
    // Image leaf inside a group (e.g. a sun-emoji clip-art nested in the
    // calendar header). The caller pre-decodes every image path seen in
    // `ws.shapeGroups[*].shapes[*].geom` via XlsxWorkbook.renderViewport,
    // so we should normally have it in `loadedImages` (keyed by imagePath).
    // Missing ordinary sources remain a silent skip. A recognized TIFF that the
    // optional codec cannot provide carries an explicit frame-local placeholder
    // mark.
    const imageGeom = shape.geom;
    const lookupKey = imageCacheKey(imageGeom.imagePath, imageGeom.duotone);
    const img = loadedImages?.get(lookupKey);
    const tiffUnavailable = isOptionalImageUnavailable(loadedImages, lookupKey, 'tiff');
    if (img || tiffUnavailable) {
      // Honor an `<a:srcRect>` crop on the leaf pic (oneCellAnchor / grpSp leaf),
      // same as the top-level anchor path (ECMA-376 §20.1.8.55). Apply the leaf's
      // `<a:alphaModFix>` opacity (§20.1.8.6) via globalAlpha, saved/restored.
      const leafAlpha = imageGeom.alpha;
      const paint = () => {
        if (img) {
          drawImageCropped(ctx, img, imageGeom.srcRect, 0, 0, sw, sh);
        } else {
          paintOptionalImagePlaceholder(ctx, 'tiff', { x: 0, y: 0, width: sw, height: sh });
        }
      };
      if (leafAlpha != null && leafAlpha < 1) {
        ctx.save();
        ctx.globalAlpha = leafAlpha;
        paint();
        ctx.restore();
      } else {
        paint();
      }
    }
  }
  // Shape text body (ECMA-376 §20.5.2.34 `<xdr:txBody>`). Drawn after
  // fill/stroke so it sits on top of the shape's background.
  if (shape.text) {
    const stacked = shape.text.vert === 'wordArtVert' || shape.text.vert === 'wordArtVertRtl';
    if (stacked && (shape.flipH || shape.flipV)) {
      // Issue #1668 Excel controls: WordArt stays readable under flipH;
      // flipV turns its unmirrored frame 180 degrees. Cancel the shape mirror
      // first, keeping the authored rotation and physical text insets.
      ctx.save();
      ctx.translate(sw / 2, sh / 2);
      ctx.scale(shape.flipH ? -1 : 1, shape.flipV ? -1 : 1);
      if (shape.flipV) ctx.rotate(Math.PI);
      ctx.translate(-sw / 2, -sh / 2);
      drawShapeText(ctx, shape.text, sw, sh, cs, cjkFallback);
      ctx.restore();
    } else {
      drawShapeText(ctx, shape.text, sw, sh, cs, cjkFallback);
    }
  }
  ctx.restore();
}

/**
 * Render a shape's `<xdr:txBody>` content into the local (already-translated)
 * coordinate system spanning [0..sw, 0..sh].
 *
 * Handles only the subset that real-world Excel files exercise heavily —
 * single column of paragraphs, per-run bold/italic/size/color/font, paragraph
 * align (`l`/`ctr`/`r`), body anchor (`t`/`ctr`/`b`). Text wrapping uses
 * canvas measurements when `bodyPr@wrap="square"` (the default), and the
 * inset is taken from `<a:bodyPr>` (the OOXML defaults are 91440 EMU left/right
 * and 45720 EMU top/bottom). Only independently verified single-resource,
 * single-line shapes receive the Office-observed natural line projection.
 */
// ── Math (OMML) rendering in shapes ─────────────────────────────────────────
// Equations are converted to SVG by MathJax once, cached by their MathNode[]
// reference (stable from parse through render), then drawn as images inside
// shape text. Mirrors the pptx/docx pipeline so typesetting is identical; the
// engine is opt-in (injected once at XlsxWorkbook.load via LoadOptions.math,
// stored on the instance) and tree-shakes out otherwise.
interface MathRender {
  /** The equation rasterized as opaque black glyphs on transparent. */
  raster: RasterizedMathSvg;
  /** baseline-relative extents in em (1em = the equation's font size in px). */
  widthEm: number;
  ascentEm: number;
  descentEm: number;
  /** Per-colour tinted copies (the black `img` recoloured via source-in). */
  tinted: Map<string, CanvasImageSource | null>;
}
const mathRenders = new WeakMap<MathNode[], MathRender>();

/** Gather every math run reachable from a worksheet's shapes. Equations live
 *  only in shape text bodies (ECMA-376 §22.1) — never in cell values. */
function collectWorksheetMath(ws: Worksheet): { nodes: MathNode[]; display: boolean }[] {
  const found: { nodes: MathNode[]; display: boolean }[] = [];
  for (const anchor of ws.shapeGroups ?? []) {
    for (const shape of anchor.shapes) {
      for (const p of shape.text?.paragraphs ?? []) {
        for (const run of p.runs) {
          if (run.type === 'math') found.push({ nodes: run.nodes, display: run.display });
        }
      }
    }
  }
  return found;
}

/** True iff the worksheet has at least one equation not yet rasterized. Lets
 *  the caller keep steady-state scroll/zoom frames fully synchronous (no await)
 *  — only the first frame that reveals new equations pays the async cost. */
export function worksheetHasUncachedMath(ws: Worksheet): boolean {
  for (const anchor of ws.shapeGroups ?? []) {
    for (const shape of anchor.shapes) {
      for (const p of shape.text?.paragraphs ?? []) {
        for (const run of p.runs) {
          if (run.type === 'math' && !mathRenders.has(run.nodes)) return true;
        }
      }
    }
  }
  return false;
}

/** Pre-rasterize every equation in the worksheet (async). Idempotent: skips
 *  already-rasterized equations and only loads MathJax when math is present.
 *  MUST be awaited BEFORE the canvas resize (off the synchronous draw path), so
 *  the old frame stays visible until the new one is ready (no white flash). */
export async function prepareWorksheetMath(ws: Worksheet, math: MathRenderer): Promise<void> {
  const uncached = collectWorksheetMath(ws).filter((r) => !mathRenders.has(r.nodes));
  if (uncached.length === 0) return;
  await math.loadMathJax();
  for (const r of uncached) {
    if (mathRenders.has(r.nodes)) continue;
    try {
      const out = await math.mathMLToSvg(mathToMathML(r.nodes, r.display));
      const raster = await rasterizeMathSvg(out, '#000000');
      mathRenders.set(r.nodes, {
        raster,
        widthEm: out.widthEm,
        ascentEm: out.ascentEm,
        descentEm: out.descentEm,
        tinted: new Map(),
      });
    } catch {
      // Conversion failure: leave the equation unrendered rather than throw.
    }
  }
}

export function drawShapeText(
  ctx: CanvasRenderingContext2D,
  txt: import('./types.js').ShapeText,
  sw: number, sh: number,
  cs: number,
  cjkFallback?: CjkLang,
): void {
  if (sw <= 0 || sh <= 0 || txt.paragraphs.length === 0) return;
  // The shape box (sw,sh) is already cellScale-scaled by the caller, but font
  // sizes are authored in points. Multiply every px size (padding, text, math)
  // by `cs` so the contents grow/shrink with the box on Excel's zoom slider —
  // matching how cell text scales via buildFont(font, cs). Without this the box
  // would zoom while the glyphs stayed fixed, drifting out of alignment.
  // Text insets from `<a:bodyPr>` (ECMA-376 §21.1.2.1.1), EMU → px, scaled by
  // `cs` like every other px size here. The parser always emits them (spec
  // default 91440 EMU L/R = 9.6 px, 45720 EMU T/B = 4.8 px). This replaces the
  // former empirical padX=7 / padY=4 constants with the real, per-side insets —
  // the box is inset asymmetrically when the shape authors distinct lIns/rIns or
  // tIns/bIns.
  const rect = drawingMlTextRect(sw, sh, {
    lIns: txt.lIns, rIns: txt.rIns, tIns: txt.tIns, bIns: txt.bIns,
  }, cs / EMU_PER_PX);
  const padLeft = rect.left;
  const innerW = rect.width;
  const innerH = rect.height;
  if (innerW <= 0 || innerH <= 0) return;

  // Font string + px size for a text run (math runs have no run-level font).
  const textFont = (run: Extract<import('./types.js').ShapeTextRun, { type: 'text' }>, fontScale = 1): { font: string; px: number } => {
    const size = run.size > 0 ? run.size : DEFAULT_FONT_SIZE;
    const px = size * PT_TO_PX * cs * fontScale;
    const family = fontStackFor(run.fontFace, cjkFallback, run.text,
      officeRoute(ctx, run.fontFace, run.bold, run.italic),
      googleSubstitutesByContext.get(ctx) === true,
      undefined, contextRegularAlias(ctx, run.fontFace));
    return { font: `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${px}px ${family}`, px };
  };

  if (txt.vert === 'wordArtVert' || txt.vert === 'wordArtVertRtl') {
    // ECMA-376 §21.1.2.1.3: scale each run's original font size before
    // measuring its stacked cell. An omitted scale means 100%.
    const fontScale = txt.autoFit === 'norm' ? txt.fontScale ?? 1 : 1;
    if (fontScale === 0) return; // A zero font scale has no glyph ink.
    drawShapeStackedText(ctx, txt, txt.vert, rect, PT_TO_PX * cs, (run, text) => {
      const face = isCjkBreakChar(text.codePointAt(0) ?? 0)
        ? run.fontFaceEa ?? run.fontFace : run.fontFace;
      return { ...textFont({ ...run, text, fontFace: face }, fontScale), face };
    }, (run, precedingSizePt) => {
      const render = mathRenders.get(run.nodes);
      // Match horizontal input: optional/failed equations contribute neither
      // ink nor a display break; preparation deliberately leaves them uncached.
      if (!render) return undefined;
      const pxSize = (run.fontSize ?? precedingSizePt) * PT_TO_PX * cs * fontScale;
      const width = render.widthEm * pxSize;
      const height = (render.ascentEm + render.descentEm) * pxSize;
      return { width, height, draw: (x, y) => {
        if (width > 0 && height > 0) ctx.drawImage(cachedTintedMathRaster(render, run.color ?? '#000000'), x, y, width, height);
      } };
    });
    return;
  }

  // Excel's natural line box follows each run's font metrics (see
  // excelDrawingMlLineRatios): Arial 1.150 em, Calibri 1.221 em, Meiryo
  // 1.95 em, and so on, instead of a flat 1.2 em. A name alone cannot identify
  // the selected bytes or their geometry. Use the metric box only when every
  // text run's exact local face has loaded with unambiguous reference metrics.
  // Otherwise the whole body keeps the ordinary 1.2 em Canvas line box, so one
  // body never mixes two line-box models.
  const runLineRatios = new Map<object, ShapeRunLineRatios>();
  let metricBody = !txt.paragraphs.some((p) => p.runs.some((run) => run.type === 'math'));
  const routes = officeRoutesByContext.get(ctx);
  const fontBoxProbe = canvasShapeFontBoxProbe(ctx);
  for (const paragraph of txt.paragraphs) for (const run of paragraph.runs) {
    if (!metricBody) break;
    if (run.type !== 'text') continue;
    const ratios = run.fontFace
      ? shapeRunLineRatios(run, routes?.[shapeOfficeRouteKey(run)], fontBoxProbe)
      : undefined;
    if (!ratios) metricBody = false;
    else runLineRatios.set(run, ratios);
  }
  if (!txt.paragraphs.some((p) => p.runs.some((run) => run.type === 'text'))) metricBody = false;

  // A laid-out segment: measured text or a rasterized equation. `w` is the
  // advance width (px); math also carries baseline-relative ascent/descent.
  type Seg =
    | { kind: 'text'; text: string; font: string; color: string; w: number }
    | { kind: 'tab'; w: number }
    | { kind: 'math'; render: MathRender; color: string; w: number; ascent: number; descent: number };
  // `leftInset` = px from padLeft to this line's left edge (paragraph left
  // margin, plus the first-line indent on a paragraph's first line). `availW` = the
  // width of the alignment region for this line (paraW, minus the first-line
  // indent on the first line). ECMA-376 §21.1.2.2.7 (marL/marR/indent).
  // `height`/`ascent` are the spaced line box. `gapBefore` is the paragraph
  // spacing (spcAft of the previous paragraph plus spcBef of this one) that
  // precedes a paragraph's first line.
  type Line = {
    segs: Seg[]; align: string; height: number; ascent: number; hasMath: boolean;
    leftInset: number; availW: number; gapBefore: number;
  };

  // Real font ascent for a line's alphabetic baseline (used on lines that mix
  // text with an inline/display equation, where text and math share a baseline
  // = lineTop + ascent). Mirrors the pptx renderer: measure the rendered glyphs'
  // `actualBoundingBoxAscent` rather than the old flat 0.85×em heuristic, which
  // over-/under-estimated for CJK and tall fonts. `font` must already be a valid
  // CSS font string (measureText keys off `ctx.font`). Falls back to 0.85×px
  // when the metric is unavailable (0), matching the prior constant.
  const measuredAscent = (font: string, px: number): number => {
    const prev = ctx.font;
    ctx.font = font;
    const a = ctx.measureText('M').actualBoundingBoxAscent;
    ctx.font = prev;
    return a > 0 ? a : px * 0.85;
  };

  // Wrap each paragraph into lines (segments preserve run order).
  const wrap = txt.wrap !== 'none';
  const lines: Line[] = [];
  // Natural (unspaced) height of each line, indexed like `lines`.
  const naturalHeights: number[] = [];
  let previousParagraph: { spaceAfter: import('./types.js').ShapeParagraph['spaceAfter']; lastNaturalHeight: number } | undefined;
  // The shared DrawingML break phase owns every soft opportunity. Excel keeps
  // the resource resolver, equation raster, line-box policy, and paint here.
  for (const p of txt.paragraphs) {
    const align = p.align || 'l';
    const marLpx = ((p.marL ?? 0) / EMU_PER_PX) * cs;
    const marRpx = ((p.marR ?? 0) / EMU_PER_PX) * cs;
    const firstLineIndent = Math.max(0, ((p.indent ?? 0) / EMU_PER_PX) * cs);
    const paraW = Math.max(0, innerW - marLpx - marRpx);
    type ShapeStyle = {
      kind: 'text' | 'math';
      font: string;
      color: string;
      pxSize: number;
      render?: MathRender;
      ascent?: number;
      descent?: number;
      /** m:oMathPara display equation (own line, no first-line indent). */
      display?: boolean;
      /** Authored face of a text run, for an empty line's fallback ascent. */
      face?: string;
      /** Excel's natural line box of a text run, when the body uses it. */
      ratios?: ShapeRunLineRatios;
    };
    const input: DrawingMlInputRun<ShapeStyle>[] = [];
    let lastTextPt = 0;
    for (const run of p.runs) {
      if (run.type === 'break') { input.push({ type: 'break' }); continue; }
      if (run.type === 'math') {
        const render = mathRenders.get(run.nodes);
        if (!render) continue;
        const pxSize = (run.fontSize ?? (lastTextPt || DEFAULT_FONT_SIZE)) * PT_TO_PX * cs;
        const ascent = render.ascentEm * pxSize;
        const descent = render.descentEm * pxSize;
        const style: ShapeStyle = {
          kind: 'math', font: '', color: run.color ?? '#000000', pxSize,
          render, ascent, descent, display: run.display === true,
        };
        input.push({ type: 'object', width: render.widthEm * pxSize, style, payload: style, display: run.display });
        continue;
      }
      lastTextPt = run.size > 0 ? run.size : DEFAULT_FONT_SIZE;
      const { font, px } = textFont(run);
      input.push({ type: 'text', text: run.text,
        style: { kind: 'text', font, color: run.color ?? '#000000', pxSize: px, face: run.fontFace,
          ratios: metricBody ? runLineRatios.get(run) : undefined } });
    }
    const broken = breakDrawingMlText(input, {
      maxWidth: wrap ? paraW : Infinity,
      firstLineIndent,
      measureText(text, style) {
        ctx.font = style.font;
        return ctx.measureText(text).width;
      },
      sameStyle: (left, right) => left.kind === right.kind
        && left.font === right.font && left.color === right.color
        && left.pxSize === right.pxSize,
      tabStops: p.tabStops?.map((tab) => ({
        pos: tab.pos / EMU_PER_PX * cs, algn: tab.algn,
      })),
      defaultTabSize: (p.defTabSz ?? 914400) / EMU_PER_PX * cs,
      tabStartPen: (index) => marLpx + (index === 0 ? firstLineIndent : 0),
    });
    // A display equation always closes the pending line first. When nothing
    // is pending (paragraph start, after <a:br>, a line feed or another display
    // equation) that closed line is empty but still reserves one line height.
    // An empty run is already a (blank) line of its own in the break phase.
    const blankBeforeDisplay: boolean[] = [];
    let pending = false;
    for (const item of input) {
      if (item.type === 'break') { pending = false; continue; }
      if (item.type === 'object') {
        if (item.display) { blankBeforeDisplay.push(!pending); pending = false; } else pending = true;
        continue;
      }
      if (item.text === '') { pending = true; continue; }
      item.text.split('\n').forEach((piece, s) => {
        if (s > 0) pending = false;
        if (piece.replace(/\r/gu, '') !== '') pending = true;
      });
    }
    let displayIndex = 0;
    // Empty lines take the nearest preceding text run's size and face.
    const runIndexOf = new Map<ShapeStyle, number>();
    input.forEach((item, idx) => { if (item.type === 'text') runIndexOf.set(item.style, idx); });
    let fallback: ShapeStyle | undefined;
    let fallbackIndex = -1;
    const noteText = (style: ShapeStyle): void => {
      const idx = runIndexOf.get(style) ?? -1;
      if (idx >= fallbackIndex) { fallback = style; fallbackIndex = idx; }
    };
    const emptyLineBox = (): { height: number; ascent: number } => {
      const pxSize = fallback?.pxSize ?? DEFAULT_FONT_SIZE * PT_TO_PX * cs;
      const face = fallback?.face;
      if (fallback?.ratios) {
        return {
          height: pxSize * (fallback.ratios.ascentRatio + fallback.ratios.descentRatio),
          ascent: pxSize * fallback.ratios.ascentRatio,
        };
      }
      return {
        height: pxSize * 1.2,
        ascent: measuredAscent(`${pxSize}px ${fontStackFor(face, undefined, '',
          officeRoute(ctx, face), googleSubstitutesByContext.get(ctx) === true,
          undefined, contextRegularAlias(ctx, face))}`, pxSize),
      };
    };
    const reduction = txt.autoFit === 'norm' ? txt.lnSpcReduction ?? 0 : 0;
    const spaceLine = excelShapeSpacing(p.spaceLine);
    const lineHeightOf = (natural: number): number => drawingMlLineHeight(
      natural, spaceLine, PT_TO_PX * cs, reduction,
    );
    // A metric line box keeps its natural ascent/descent split, which
    // a:lnSpc then re-divides (drawingMlSpacedLineBox).
    const metricLine = (ascent: number, descent: number): { height: number; ascent: number } => {
      const box = drawingMlSpacedLineBox({ ascent, descent }, spaceLine, PT_TO_PX * cs, reduction);
      return { height: box.ascent + box.descent, ascent: box.ascent };
    };
    const firstLineIndex = lines.length;
    for (const [index, brokenLine] of broken.entries()) {
      const segs: Seg[] = [];
      let naturalHeight = 0;
      let ascent = 0;
      let metricAscent = 0;
      let metricDescent = 0;
      let hasMath = false;
      let displayMath = false;
      for (const part of brokenLine.segments) {
        if (part.type === 'tab') {
          segs.push({ kind: 'tab', w: part.width });
          continue;
        }
        if (part.type === 'object') {
          const style = part.style;
          if (!style.render) continue;
          const mathAscent = style.ascent ?? 0;
          const mathDescent = style.descent ?? 0;
          segs.push({ kind: 'math', render: style.render, color: style.color,
            w: part.width, ascent: mathAscent, descent: mathDescent });
          naturalHeight = Math.max(naturalHeight, mathAscent + mathDescent);
          ascent = Math.max(ascent, mathAscent);
          hasMath = true;
          if (style.display) displayMath = true;
          continue;
        }
        const style = part.style;
        segs.push({ kind: 'text', text: part.text, font: style.font,
          color: style.color, w: part.width });
        naturalHeight = Math.max(naturalHeight, style.pxSize * 1.2);
        ascent = Math.max(ascent, measuredAscent(style.font, style.pxSize));
        if (style.ratios) {
          metricAscent = Math.max(metricAscent, style.pxSize * style.ratios.ascentRatio);
          metricDescent = Math.max(metricDescent, style.pxSize * style.ratios.descentRatio);
        }
        noteText(style);
      }
      if (displayMath && blankBeforeDisplay[displayIndex++]) {
        const blank = emptyLineBox();
        lines.push({ segs: [], align, height: lineHeightOf(blank.height), ascent: blank.ascent,
          hasMath: false, leftInset: marLpx, availW: paraW, gapBefore: 0 });
        naturalHeights.push(blank.height);
      }
      // Runs that paint nothing (trimmed or wrapped spaces) still size their line.
      for (const idx of brokenLine.hiddenRuns ?? []) {
        const item = input[idx];
        if (item.type !== 'text') continue;
        naturalHeight = Math.max(naturalHeight, item.style.pxSize * 1.2);
        ascent = Math.max(ascent, measuredAscent(item.style.font, item.style.pxSize));
        if (item.style.ratios) {
          metricAscent = Math.max(metricAscent, item.style.pxSize * item.style.ratios.ascentRatio);
          metricDescent = Math.max(metricDescent, item.style.pxSize * item.style.ratios.descentRatio);
        }
        noteText(item.style);
      }
      let height: number;
      if (metricBody) {
        // Office shares one baseline across the runs of a line and unions their
        // ascents and descents (#1604 controls: Arial+Meiryo, Arial+MS Gothic,
        // 11+40 pt runs in one line).
        if (metricAscent + metricDescent === 0) {
          const blank = emptyLineBox();
          metricAscent = blank.ascent;
          metricDescent = blank.height - blank.ascent;
        }
        naturalHeight = metricAscent + metricDescent;
        ({ height, ascent } = metricLine(metricAscent, metricDescent));
      } else {
        if (naturalHeight === 0) ({ height: naturalHeight, ascent } = emptyLineBox());
        height = lineHeightOf(naturalHeight);
      }
      // Classify this line by its own equation: an inline equation on a
      // paragraph's first line keeps the first-line indent even when a later
      // equation in the paragraph is display math.
      const isDisplayMath = displayMath;
      lines.push({
        segs, align, height, ascent, hasMath,
        leftInset: isDisplayMath ? marLpx : marLpx + (index === 0 ? firstLineIndent : 0),
        availW: isDisplayMath ? paraW : paraW - (index === 0 ? firstLineIndent : 0),
        gapBefore: 0,
      });
      naturalHeights.push(naturalHeight);
    }
    // §21.1.2.2.10 spcBef / §21.1.2.2.9 spcAft. Excel adds the previous
    // paragraph's spcAft and this paragraph's spcBef (#1604: 12 + 18 pt gave
    // 30 pt), and ignores spcBef on the body's first paragraph. A percentage
    // refers to the natural single-line height of the adjacent line. Each
    // value is first rounded to a whole unit (excelShapeSpacing).
    if (lines.length > firstLineIndex) {
      if (previousParagraph && firstLineIndex > 0) {
        lines[firstLineIndex].gapBefore = drawingMlParagraphSpacing(
          excelShapeSpacing(previousParagraph.spaceAfter), previousParagraph.lastNaturalHeight, PT_TO_PX * cs,
        ) + drawingMlParagraphSpacing(excelShapeSpacing(p.spaceBefore), naturalHeights[firstLineIndex], PT_TO_PX * cs);
      }
      previousParagraph = { spaceAfter: p.spaceAfter, lastNaturalHeight: naturalHeights[lines.length - 1] };
    }
  }

  // Canvas px size of a CSS font string built by textFont.
  const fontPx = (font: string): number => Number(/(\d+(?:\.\d+)?)px/u.exec(font)?.[1] ?? 0);
  // Offset from a 'middle' paint position to that run's alphabetic baseline.
  const middleToAlphabetic = (seg: { font: string; text: string }): number => {
    const prevFont = ctx.font;
    const prevBaseline = ctx.textBaseline;
    ctx.font = seg.font;
    ctx.textBaseline = 'alphabetic';
    const alphabetic = ctx.measureText(seg.text).actualBoundingBoxAscent;
    ctx.textBaseline = 'middle';
    const middle = ctx.measureText(seg.text).actualBoundingBoxAscent;
    ctx.font = prevFont;
    ctx.textBaseline = prevBaseline;
    const offset = alphabetic - middle;
    return Number.isFinite(offset) ? offset : 0;
  };

  // Total text block height
  const blockH = lines.reduce((s, l) => s + l.gapBefore + l.height, 0);

  // Vertical anchor — ECMA-376 §20.1.7.2 <a:bodyPr anchor>.
  // For 'ctr' we intentionally skip Math.max(0,...) so the block stays
  // visually centered even when blockH exceeds innerH (text clips at edge).
  const y0 = drawingMlBlockTop(txt.anchor, rect, blockH);

  let lineTop = y0;
  for (const line of lines) {
    lineTop += line.gapBefore;
    const totalW = line.segs.reduce((s, seg) => s + seg.w, 0);
    // Per-line region: the left edge is padLeft + the paragraph's left inset
    // (marL, plus first-line indent on the first line), and alignment happens
    // within the line's available width (paraW). ECMA-376 §21.1.2.2.7.
    const base = padLeft + line.leftInset;
    const x = drawingMlLineX(line.align, base, line.availW, totalW);
    const paintSegments = line.segs.map((seg) => ({ type: 'object' as const, style: seg, width: seg.w }));

    if (line.hasMath) {
      // A line containing an equation aligns text AND the math raster to a
      // shared alphabetic baseline (= lineTop + ascent), so the equation sits
      // on the same baseline as adjacent text. (Pure-text lines keep the
      // simpler 'middle' baseline below.)
      //
      // For a lone display equation in a top-anchored box (anchor="t", tIns=0),
      // line.ascent === seg.ascent, so `baseline - seg.ascent === lineTop` and
      // the raster is drawn TOP-FLUSH to the box — matching how Excel autofits
      // the box to the equation and top-anchors it (issue #877). Any size
      // residual against the authored spAutoFit box is NOT a downward shift and
      // is NOT fixable here: it is a construct-specific difference between the
      // two math typesetting systems (MathJax + STIX Two Math vs Office +
      // Cambria Math MATH-table constants — display-operator variant ladders,
      // limit/fraction shift minima). Measured Office/ours ink-height ratios
      // even flip sign per construct (~1.17x for a display sum with limits,
      // ~0.89x for display fractions + radicals, ~1.0x for plain scripts), so
      // no font-metrics-derived scale factor exists and emulating Cambria
      // Math's constants is out of policy. Adjudicated won't-fix in #877 — see
      // the measurement comment there before re-litigating.
      ctx.textBaseline = 'alphabetic';
      const baseline = lineTop + line.ascent;
      paintDrawingMlLine(paintSegments, x, baseline, (part, penX) => {
        const seg = part.style;
        if (seg.kind === 'text') {
          ctx.font = seg.font;
          ctx.fillStyle = seg.color;
          ctx.fillText(seg.text, penX, baseline);
        } else if (seg.kind === 'math') {
          const img = cachedTintedMathRaster(seg.render, seg.color);
          ctx.drawImage(img, penX, baseline - seg.ascent, seg.w, seg.ascent + seg.descent);
        }
      });
    } else if (metricBody) {
      // Excel's metric line box: one alphabetic baseline at the line ascent.
      ctx.textBaseline = 'alphabetic';
      const baseline = lineTop + line.ascent;
      paintDrawingMlLine(paintSegments, x, baseline, (part, penX) => {
        const seg = part.style;
        if (seg.kind === 'text') {
          ctx.font = seg.font;
          ctx.fillStyle = seg.color;
          ctx.fillText(seg.text, penX, baseline);
        }
      });
    } else {
      // Ordinary 1.2 em line box, centred on the line's largest run. Office
      // puts every run of a line on one baseline (#1604 mixed-run controls),
      // so runs in another font or size share that run's alphabetic baseline
      // instead of each being centred on its own em box.
      const drawY = lineTop + line.height / 2;
      const textSegs = line.segs.filter((seg): seg is Extract<Seg, { kind: 'text' }> => seg.kind === 'text');
      const lead = textSegs.reduce<Extract<Seg, { kind: 'text' }> | undefined>((best, seg) => (
        !best || fontPx(seg.font) > fontPx(best.font) ? seg : best), undefined);
      const shared = lead && textSegs.some((seg) => seg.font !== lead.font)
        ? drawY + middleToAlphabetic(lead) : undefined;
      ctx.textBaseline = shared === undefined ? 'middle' : 'alphabetic';
      const y = shared ?? drawY;
      paintDrawingMlLine(paintSegments, x, y, (part, penX) => {
        const seg = part.style;
        if (seg.kind === 'text') {
          ctx.font = seg.font;
          ctx.fillStyle = seg.color;
          ctx.fillText(seg.text, penX, y);
        }
      });
    }
    lineTop += line.height;
  }
}

function fillAndStroke(
  ctx: CanvasRenderingContext2D,
  shape: ShapeInfo,
  width: number,
  height: number,
  cs: number,
): void {
  fillShape(ctx, shape, width, height, cs);
  strokeShapePath(ctx, shape, width, height, cs);
}

/** Excel's printed DrawingML pattern stays on the page axes through shape
 * rotation, reflection and a rotated group. The fill itself is phased from
 * the unrotated shape anchor. Counter-transform only the pattern coordinates;
 * the shape geometry continues to follow the authored transform. The six
 * baseline/rotation/flip/group controls all exported an axis-aligned 8 pt PDF
 * tile, including the rotated and reflected cases. */
function axisAlignedPatternTransform(
  shape: ShapeInfo,
  width: number,
  height: number,
): DOMMatrix2DInit | undefined {
  if (shape.rot === 0 && !shape.flipH && !shape.flipV) return undefined;
  const angle = shape.rot * Math.PI / 180;
  const cosine = Math.cos(angle);
  const sine = Math.sin(angle);
  const fx = shape.flipH ? -1 : 1;
  const fy = shape.flipV ? -1 : 1;
  // Inverse of the local centre rotation/reflection used by drawShape.
  const a = fx * cosine;
  const b = -fy * sine;
  const c = fx * sine;
  const d = fy * cosine;
  const cx = width / 2;
  const cy = height / 2;
  return { a, b, c, d, e: cx - a * cx - c * cy, f: cy - b * cx - d * cy };
}

/** Resolve once per shape, including all fill-bearing custom subpaths. */
function resolveSpreadsheetShapeFill(
  ctx: CanvasRenderingContext2D,
  shape: ShapeInfo,
  width: number,
  height: number,
  cs: number,
): ReturnType<typeof resolveFill> {
  const fill = shape.fill ?? (shape.fillColor
    ? { fillType: 'solid' as const, color: shape.fillColor }
    : null);
  // Excel's direct PDF export uses the same 8 pt tile artwork and cell size
  // as PowerPoint, but its print pattern matrix starts at each shape's X and
  // an 8 pt grid measured from the PDF page bottom. The sheet viewer has no
  // print-page origin, so its phase is anchored to this local shape frame.
  // One point is 4/3 CSS pixels at native zoom; cellScale changes the sheet
  // coordinate system without an additional canvas scale for shape painting.
  return resolveFill(
    fill, ctx, 0, 0, width, height, shape.rot, PT_TO_PX * cs,
    axisAlignedPatternTransform(shape, width, height),
    (target, x, y, w, h) => appendSpreadsheetShapeOutline(target, shape, x, y, w, h),
  );
}

/** Fill the current path; reuse one raster across a custom shape's subpaths. */
function fillShape(
  ctx: CanvasRenderingContext2D, shape: ShapeInfo,
  width: number, height: number, cs: number,
  paint = resolveSpreadsheetShapeFill(ctx, shape, width, height, cs),
): boolean {
  if (!paint) return false;
  ctx.fillStyle = paint;
  ctx.fill();
  return true;
}

function shapeStroke(shape: ShapeInfo): Stroke | null {
  if (!shape.strokeColor || shape.strokeWidth <= 0) return null;
  return {
    color: shape.strokeColor,
    width: shape.strokeWidth,
    ...(shape.strokeFill ? { fill: shape.strokeFill } : {}),
    ...(shape.strokeDashStyle ? { dashStyle: shape.strokeDashStyle } : {}),
    ...(shape.strokeCustomDash?.length ? { customDash: shape.strokeCustomDash } : {}),
    ...(shape.strokeLineCap ? { lineCap: shape.strokeLineCap } : {}),
    ...(shape.strokeLineJoin ? { lineJoin: shape.strokeLineJoin } : {}),
    ...(shape.strokeMiterLimit !== undefined ? { miterLimit: shape.strokeMiterLimit } : {}),
    ...(shape.strokeAlignment ? { alignment: shape.strokeAlignment } : {}),
    ...(shape.strokeCmpd ? { cmpd: shape.strokeCmpd } : {}),
    ...(shape.strokeHeadEnd ? { headEnd: shape.strokeHeadEnd } : {}),
    ...(shape.strokeTailEnd ? { tailEnd: shape.strokeTailEnd } : {}),
  };
}

function strokeShapePath(
  ctx: CanvasRenderingContext2D,
  shape: ShapeInfo,
  width: number,
  height: number,
  cs: number,
): void {
  const stroke = shapeStroke(shape);
  if (!stroke) return;
  applyStroke(ctx, stroke, 1 / EMU_PER_PX);
  // Match DOCX/PPTX stroke hosts: use the authored box, not the fill-bearing
  // silhouette (which excludes decorative paths that may still be stroked).
  // Shade coordinates use that host box; raster coverage comes from the
  // recorded stroke path, including its actual transformed caps and joins.
  if (stroke.fill) {
    const paint = resolveFill(
      stroke.fill, ctx, 0, 0, width, height, shape.rot, PT_TO_PX * cs,
      axisAlignedPatternTransform(shape, width, height),
      undefined, currentStrokeBounds(ctx),
    );
    if (paint) ctx.strokeStyle = paint;
  }
  ctx.stroke();
}

// ═══════════════════════════════════════════════════════════════════════════
// Chart rendering — delegated to @silurus/ooxml-core's unified renderer.
//
// The parser now emits the canonical `ChartModel` directly (see the Rust
// `From<ChartData> for ChartModel`), so there is no TS adapter here anymore —
// `anchor.chart` is passed straight to `renderChart`.
// ═══════════════════════════════════════════════════════════════════════════

// ── renderCharts ────────────────────────────────────────────────────────────

function renderCharts(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  colAxis: GridAxisGeometry,
  rowAxis: GridAxisGeometry,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
  rtl: boolean,
  canvasW: number,
  loadedImages?: Map<string, CanvasImageSource | null>,
  anchors: readonly ChartAnchor[] = ws.charts,
  threeD?: ChartThreeDRenderer,
  regionMap?: ChartRegionMapRenderer,
  chartEx?: ChartExRenderer,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;

  const scrollOriginSheetX = sheetXForCol(colAxis, startCol);
  const scrollOriginSheetY = sheetYForRow(rowAxis, startRow);
  const clipX = sheetAnchoredRectX(scrollAreaX, scrollAreaW, canvasW, rtl);

  for (const anchor of anchors) {
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;
    const toCol1   = anchor.toCol   + 1;
    const toRow1   = anchor.toRow   + 1;

    const shX1 = sheetXForCol(colAxis, fromCol1) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const shY1 = sheetYForRow(rowAxis, fromRow1) + (anchor.fromRowOff * cs) / EMU_PER_PX;
    const shX2 = sheetXForCol(colAxis, toCol1) + (anchor.toColOff * cs) / EMU_PER_PX;
    const shY2 = sheetYForRow(rowAxis, toRow1) + (anchor.toRowOff * cs) / EMU_PER_PX;

    const cw = shX2 - shX1;
    const ch = shY2 - shY1;
    if (cw <= 0 || ch <= 0) continue;

    const logicalCx = scrollAreaX + (shX1 - scrollOriginSheetX) - scrollOffsetX;
    const cx = sheetAnchoredRectX(logicalCx, cw, canvasW, rtl);
    const cy = scrollAreaY + (shY1 - scrollOriginSheetY) - scrollOffsetY;

    if (cx + cw < clipX || cx > clipX + scrollAreaW) continue;
    if (cy + ch < scrollAreaY || cy > scrollAreaY + scrollAreaH) continue;

    ctx.save();
    ctx.beginPath();
    ctx.rect(clipX, scrollAreaY, scrollAreaW, scrollAreaH);
    ctx.clip();

    // Render the chart in its natural 96-DPI coordinate space, then let Canvas
    // scale the complete paint by the worksheet zoom. Passing a smaller rect
    // and only shrinking `ptToPx` leaves renderer-owned gaps/padding at fixed
    // screen pixels; that can repack or entirely suppress a manual legend at a
    // zoom step even though Excel scales the chart as one proportional object.
    ctx.save();
    if (cs !== 1) ctx.scale(cs, cs);
    // `anchor.chart` is already the canonical ChartModel emitted by the Rust
    // parser (`ooxml_common::chart::ChartModel`) — the former `adaptChartData`
    // default/mapping logic now lives in the parser's `From<ChartData>`.
    renderChart(
      ctx,
      anchor.chart,
      cs === 1
        ? { x: cx, y: cy, w: cw, h: ch }
        : { x: cx / cs, y: cy / cs, w: cw / cs, h: ch / cs },
      PT_TO_PX,
      0,
      threeD,
      regionMap,
      fill => loadedImages?.get(chartImageFillKey(fill)),
      chartEx,
      // Excel PDF controls at 75/100/200% view zoom and two scroll positions
      // retain the same page-grid tile matrix. Moving the chart 5 pt moves its
      // fill rectangle but leaves that matrix fixed. Keep the chart pattern on
      // the viewport grid (the existing main policy), while the CTM above
      // scales each cell with worksheet zoom on screen.
    );
    ctx.restore();
    ctx.restore();
  }
}
