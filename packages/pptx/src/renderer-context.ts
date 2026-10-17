// Render-context types and low-level helpers shared by the PPTX slide
// renderer modules: unit conversion, decode-target keys and options, and the
// fetch-image callback shape. Split from renderer.ts (#1857).

import type { CjkLang } from '@silurus/ooxml-core';
import type {
  PictureElement,
  MediaElement,
  TableElement,
  ChartElement,
  TextBody,
  Paragraph,
} from './types';
import { asBullet } from './types';
import {
  buildCustomPath as buildCustomPathCore,
  hexToRgba as hexToRgbaCore,
  EMU_PER_PT as PT_TO_EMU,
  duotoneCacheKey,
  type BlipPixelEffects,
  preferVectorBlip,
  sourceRasterTargetSize,
} from '@silurus/ooxml-core';
import type { SvgImageSource, DecodedImageTargetPlan } from '@silurus/ooxml-core';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

/** Theme font context threaded through the render call chain. */
export interface RenderContext {
  cjkFallback?: CjkLang;
  officeFontRoutes?: Readonly<Record<string, import('@silurus/ooxml-core').OfficeFontFallbackRoute>>;
  /** Normalized authored families with any installed face (presence probes). */
  officeInstalledFamilies?: readonly string[];
  /** Presentation load explicitly enabled Google Fonts substitutions. */
  googleSubstitutes?: boolean;
  themeMajorFont: string | null;
  themeMinorFont: string | null;
  /** Lower-cased authored family → this presentation's isolated FontFace alias. */
  embeddedFontAliases?: ReadonlyMap<string, string>;
  /** Isolated FontFace alias → lower-cased authored family for fallback policy. */
  embeddedFontAuthoredFamilies?: ReadonlyMap<string, string>;
  embeddedFontTuples?: ReadonlySet<string>;
  /** Theme hyperlink colour as a 6-char hex (no leading #), or null. */
  themeHlinkColor?: string | null;
  /**
   * Device-pixel ratio the slide is being drawn at (the `dpr` passed to
   * `ctx.scale(dpr, dpr)` in the top render fn). Threaded so axis-aligned
   * thin line strokes (table grid borders, text underline / strike-through)
   * can apply {@link crispOffset} and render crisp on a DPR=1 display.
   * Always set — `renderSlide` populates it from the real dpr; the in-module
   * default literals use 1. Required so the crisp-line call sites need no
   * `?? 1` fallback (aligns with docx's required `RenderState.dpr`).
   */
  dpr: number;
  /**
   * Contrast-aware default text colour for THIS slide's synthetic SmartArt
   * fallback shape (issue #805), pre-derived by `renderSlide` from the
   * slide-background luminance via {@link smartArtFallbackTextColor}. Null /
   * absent = keep the ordinary theme default. Consumed only by
   * {@link shapeDefaultTextColor} for shapes matching
   * {@link isSmartArtFallbackShape}; every other shape ignores it.
   */
  smartArtFallbackTextColor?: string | null;
  /**
   * Picture-bullet image sources prepared for this render pass. Display-sized
   * bitmap decodes use a resolution-specific core cache key, so synchronous
   * text paint consumes the exact prefetch result rather than peeking only the
   * native-resolution key. The render-pass lease owns bitmap lifetime; Window
   * SVG bullets remain count-bounded HTMLImageElements in the SVG cache.
   */
  pictureBulletImages?: ReadonlyMap<string, SvgImageSource | null>;
  /** Shape blip-fill sources settled before synchronous shape/text painting. */
  shapeFillImages?: ReadonlyMap<string, PreparedShapeFill | null>;
}

/** Information about a rendered text segment for building a transparent selection overlay. */
export interface PptxTextRunInfo {
  /** Paint-order index for distinguishing displayed text containers, including
   * recovery previews. This is not a source-model locator for hit testing. */
  renderElementIndex?: number;
  /** Original slide-model element index. Absent for recovery-preview text,
   * which has no element in the source tree used by hit testing. */
  elementIndex?: number;
  /** Provenance of that rendered element. */
  origin?: import('./types.js').SlideElementOrigin;
  /**
   * Source shape's DrawingML `cNvPr@id`. The identifier is stable and unique
   * within its slide. Absent for parser-synthesized shapes that have no cNvPr.
   */
  shapeId?: string;
  text: string;
  /** X position in CSS px, relative to the shape's top-left corner. */
  inShapeX: number;
  /** Y position (top of line box) in CSS px, relative to the shape's top-left corner. */
  inShapeY: number;
  /** Measured text width in CSS px. */
  w: number;
  /** Line height in CSS px. */
  h: number;
  /** Font size in CSS px. */
  fontSize: number;
  /** CSS `font` shorthand used for canvas drawing (e.g. `"bold 16px Arial"`). */
  font: string;
  /** Shape's left edge in canvas CSS px. */
  shapeX: number;
  /** Shape's top edge in canvas CSS px. */
  shapeY: number;
  /** Shape's width in canvas CSS px. */
  shapeW: number;
  /** Shape's height in canvas CSS px. */
  shapeH: number;
  /** Shape rotation in degrees (clockwise). */
  rotation: number;
  /** Mirror the selection frame horizontally before any text-body rotation. */
  shapeFlipH?: boolean;
  /** Mirror the selection frame vertically before any text-body rotation. */
  shapeFlipV?: boolean;
  /**
   * Zero-based logical grid address when this run belongs to a DrawingML table cell.
   * This is semantic identity only; consumers decide how cell boundaries affect
   * selection, search, or serialization.
   */
  tableCell?: Readonly<{ row: number; column: number }>;
  /** Cell-owned physical x-axis overflow, normalized so an omitted
   * `horzOverflow` is `clip`. Present only for table-cell runs of an effective
   * horizontal body (so never with textBodyRotation); its absence never implies
   * clipping, even with tableCell. shapeX/Y/W/H describe the same cell paint
   * frame. Logical text remains complete, so worker/search consumers clip
   * visible geometry without losing it.
   */
  cellHorzOverflow?: 'clip' | 'overflow';
  /**
   * Additional rotation from a vertical text body (`vert="vert"` → 90,
   * `vert="vert270"` → -90). The CSS overlay must add this to `rotation`.
   */
  textBodyRotation?: number;
  /**
   * Resolved hyperlink target for this run (IX1), classified into the shared
   * {@link HyperlinkTarget} shape. Present only for runs whose `<a:rPr>` carried
   * an `<a:hlinkClick>`; the overlay makes such spans clickable. The glyph
   * drawing (colour + underline) is unaffected — this is metadata for the
   * transparent overlay only.
   */
  hyperlink?: HyperlinkTarget;
}

export type TextRunCallback = (run: PptxTextRunInfo) => void;

/**
 * Convert EMU to canvas pixels.
 * scale = canvasWidthPx / slideWidthEMU  (so that slideWidth EMU == canvasWidth px)
 */
export function emuToPx(emu: number, scale: number): number {
  return emu * scale;
}

export function rasterTargetOptions(
  widthCssPx: number,
  heightCssPx: number,
  dpr: number,
  srcRect?: PictureElement['srcRect'],
): { targetWidthPx: number; targetHeightPx: number } | undefined {
  const target = sourceRasterTargetSize(
    Math.abs(widthCssPx * dpr),
    Math.abs(heightCssPx * dpr),
    srcRect,
  );
  return target
    ? { targetWidthPx: target.width, targetHeightPx: target.height }
    : undefined;
}

export type PlannedRasterOptions = {
  targetWidthPx: number;
  targetHeightPx: number;
  maxRetainedPixels: number;
};

/** The decode-time pixel transform of a blip: its ordered CT_Blip effects
 *  (with the duotone at its position) when it carries any, otherwise just the
 *  duotone, so a duotone-only or plain picture keeps its existing cache key. */
export function pixelTransform(blip: {
  readonly duotone?: import('@silurus/ooxml-core').Duotone | null;
  readonly blipEffects?: readonly import('@silurus/ooxml-core').BlipEffect[] | null;
}): import('@silurus/ooxml-core').Duotone | BlipPixelEffects | undefined {
  return blip.blipEffects?.length
    ? { effects: blip.blipEffects, duotone: blip.duotone ?? null }
    : blip.duotone ?? undefined;
}

export function imagePlanKey(
  path: string,
  duotone?: import('@silurus/ooxml-core').Duotone | BlipPixelEffects | null,
): string {
  return duotoneCacheKey(path, duotone);
}

export function plannedRasterOptions(
  plan: DecodedImageTargetPlan,
  key: string,
): PlannedRasterOptions | undefined {
  const target = plan.targets.get(key);
  return target ? {
    targetWidthPx: target.width,
    targetHeightPx: target.height,
    // A restricted variant key prevents a larger cached zoom level from being
    // substituted into a pass whose aggregate plan charged only this surface.
    maxRetainedPixels: target.maxRetainedPixels,
  } : undefined;
}

export function pictureResourcePath(picture: PictureElement): string {
  return preferVectorBlip(picture)
    ? picture.svgImagePath as string
    : picture.imagePath;
}

export function bulletTargetPx(
  body: TextBody,
  paragraph: Paragraph,
  scale: number,
  dpr: number,
): number | undefined {
  const bullet = asBullet(paragraph.bullet);
  if (bullet.type !== 'blip') return undefined;
  let firstRunSizePt: number | undefined;
  for (const run of paragraph.runs) {
    if (run.type === 'text' && run.fontSize != null) {
      firstRunSizePt = run.fontSize;
      break;
    }
  }
  const basePt = firstRunSizePt
    ?? paragraph.defFontSize
    ?? body.defaultFontSize
    ?? 18;
  const sizePt = bullet.sizePts != null
    ? bullet.sizePts
    : basePt * ((bullet.sizePct ?? 100) / 100);
  return Math.max(1, sizePt * PT_TO_EMU * scale * dpr);
}

export const hexToRgba = hexToRgbaCore;

export interface PreparedShapeFill {
  image: SvgImageSource;
  intrinsicSize?: Readonly<{ width: number; height: number }>;
}

/**
 * Build a canvas path from custGeom path commands.
 * Coordinates are in [0,1] relative to the shape bounding box;
 * the renderer maps them to canvas pixels.
 * Tracks pen position so arcTo can compute the ellipse centre correctly.
 */
export const buildCustomPath = buildCustomPathCore;

// The lazy image-byte source closure (one stable identity per Presentation, so
// it namespaces the shared decoded-bitmap cache per deck).
export type FetchImage = (path: string, mime: string) => Promise<Blob>;

// ─── Chart rendering ────────────────────────────────────────────────────────
// Chart rendering is delegated to @silurus/ooxml-core's unified renderer.
// See renderChart(ctx, el, scale) below.


// ─── Table rendering ─────────────────────────────────────────────────────────

export function applyFrameTransform(
  ctx: CanvasRenderingContext2D,
  el: Pick<TableElement | ChartElement | MediaElement, 'x' | 'y' | 'width' | 'height' | 'rotation' | 'flipH' | 'flipV'>,
  scale: number,
): void {
  if (el.rotation === 0 && !el.flipH && !el.flipV) return;
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);
  ctx.translate(x + w / 2, y + h / 2);
  ctx.rotate((el.rotation * Math.PI) / 180);
  if (el.flipH) ctx.scale(-1, 1);
  if (el.flipV) ctx.scale(1, -1);
  ctx.translate(-(x + w / 2), -(y + h / 2));
}
