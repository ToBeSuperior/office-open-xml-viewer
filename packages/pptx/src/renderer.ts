
import { pptxSlideCjkFallback } from './google-fonts.js';
import type { CjkLang } from '@silurus/ooxml-core';
import type {
  Slide,
  SlideElement,
  ChartElement,
  RenderOptions,
  DimOptions,
  ImageFill,
} from './types';
import { asBullet } from './types';
import {
  renderChart,
  withPatternCoordinateSpace,
  EMU_PER_PT as PT_TO_EMU,
  isHTMLCanvas,
  defaultDpr,
  clampCanvasSize,
  getCachedSvgImageByPath,
  getCachedBitmapByPath,
  getCachedDuotoneBitmapByPath,
  chartImageFillKey,
  chartImageFillUsageSize,
  collectChartImageFillUsages,
  collectChartImageFillUsagesForCharts,
  withBitmapCacheLease,
  inspectCachedRasterSource,
  preferVectorBlip,
  isOoxmlDecodedImageLimitError,
  OoxmlResourceLimitError,
  isTiffDecodeError,
  isOptionalImageCodecUnavailableError,
} from '@silurus/ooxml-core';
import type {
  MathRenderer,
  ChartThreeDRenderer,
  ChartRegionMapRenderer,
  ChartExRenderer,
  TiffRenderer,
  SvgBlobDecoder,
  SvgImageSource,
} from '@silurus/ooxml-core';
import { smartArtFallbackTextColor } from './smartart-fallback-contrast';
import {
  rasterTargetOptions,
  pixelTransform,
  imagePlanKey,
  emuToPx,
  plannedRasterOptions,
  applyFrameTransform,
  type RenderContext,
  type PreparedShapeFill,
  type TextRunCallback,
} from './renderer-context';
import { failedMath, prepareSlideMath, elementTextBodies } from './renderer-math';
import { pictureBulletSizePt } from './renderer-text-layout';
import {
  shapeFillKey,
  NO_PAINT,
  prepareBackground,
  type PreparedPaint,
} from './renderer-shape-fill';
import { renderShape } from './renderer-shape';
import {
  posterFetchImage,
  preparePicture,
  prepareMedia,
  type PosterFetchImage,
} from './renderer-picture';
import { renderTable } from './renderer-table';
import { slideMayDecodeImages, planSlideImages, slideImageFillUsages } from './renderer-image-plan';
import { drawParseErrorPlaceholder } from './renderer-error-page';

// renderer.ts keeps its historical export surface (internal session, viewer,
// presentation, and tests import from here); the implementations live in the
// renderer-* modules split out for #1857.
export { renderTable } from './renderer-table';
export { getPosterBitmap } from './renderer-picture';
export { reflectedShapeTextRotation, shapeTextRotation } from './renderer-shape';
export {
  paragraphHasRenderableContent,
  resolveBulletLabel,
  renderTextBody,
} from './renderer-text-body';
export {
  splitScene3dShapeTransform,
  transformedEffectBounds,
  projectedEffectGeometry,
  cacheDevicePaint,
  buildBevelInputs,
} from './renderer-shape-effects';
export {
  resolveShapeFill,
  tileAnchorOffset,
  tileSourceExtent,
  paintPreparedShapeImageFill,
  applyStroke,
} from './renderer-shape-fill';
export { paintHighlight } from './renderer-text-paint';
export {
  cssFontStack,
  buildFont,
  naturalWidthExceedsBbox,
  paragraphInputRuns,
  layoutParagraph,
} from './renderer-text-layout';
export { prepareSlideMath } from './renderer-math';
export { applyFrameTransform } from './renderer-context';
export type { RenderContext, PptxTextRunInfo, TextRunCallback } from './renderer-context';

// The decoded raster/metafile bitmap cache now lives in core
// (`getCachedBitmapByPath` / `peekCachedBitmapByPath` / `dropDecodedBitmapCache`),
// shared verbatim with docx and xlsx. Re-exported under the historical pptx
// names so the presentation teardown and the bullet-draw tests keep their import
// surface; the synchronous picture-bullet draw reads a warmed bitmap through
// `peekCachedBitmap`. See core/src/image/bitmap-image-by-path.ts.
export {
  getCachedBitmapByPath as getCachedBitmap,
  peekCachedBitmapByPath as peekCachedBitmap,
  dropDecodedBitmapCache as dropImageBitmapCache,
  // Second-layer duotone (§20.1.8.23) recolour cache; dropped alongside the base
  // bitmap cache on deck teardown.
  dropDuotoneBitmapCache,
} from '@silurus/ooxml-core';

// ===== Public API =====

export type { RenderOptions } from './types';

/**
 * Overlay a translucent rectangle over the whole slide so a hidden slide reads
 * faintly (PowerPoint's hidden-slide thumbnail look). A pure mechanism — it
 * never inspects whether the slide is hidden; the caller ({@link PptxViewer}'s
 * `'dim'` mode) decides when to apply it. `widthPx`/`heightPx` are CSS px; the
 * ctx is already `scale(dpr, dpr)`d, so a fillRect in CSS px covers the canvas.
 */
export function applyDimOverlay(
  ctx: CanvasRenderingContext2D,
  dim: DimOptions,
  widthPx: number,
  heightPx: number,
): void {
  ctx.save();
  ctx.globalAlpha = dim.opacity;
  ctx.fillStyle = dim.color;
  ctx.fillRect(0, 0, widthPx, heightPx);
  ctx.restore();
}

/**
 * Raw {@link renderSlide} options: shared paint controls plus the optional math
 * engine and hidden-slide dim overlay. Viewer constructors normally inject the
 * math engine once through their load options; direct render callers can pass
 * the same public engine here.
 */
export type SlideRenderOptions = RenderOptions & {
  math?: MathRenderer;
  threeD?: ChartThreeDRenderer;
  regionMap?: ChartRegionMapRenderer;
  chartEx?: ChartExRenderer;
  tiff?: TiffRenderer;
  dim?: DimOptions;
};

/** Presentation-owned font aliases are intentionally kept out of the public
 * render options. They are an implementation detail of embedded-font lifetime
 * isolation, not a caller-configurable rendering policy. */
type InternalSlideRenderOptions = SlideRenderOptions & {
  cjkFallback?: CjkLang;
  officeFontRoutes?: Readonly<Record<string, import('@silurus/ooxml-core').OfficeFontFallbackRoute>>;
  /** Normalized authored families with any installed face (presence probes). */
  officeInstalledFamilies?: readonly string[];
  googleSubstitutes?: boolean;
  embeddedFontAliases?: ReadonlyMap<string, string>;
  embeddedFontAuthoredFamilies?: ReadonlyMap<string, string>;
  embeddedFontTuples?: ReadonlySet<string>;
  svgDecoder?: SvgBlobDecoder;
};

/**
 * Per-canvas monotonic render token for the {@link renderSlide} cancellation
 * guard. A WeakMap keyed on the canvas replaces the previous property monkey-
 * patch (`canvas.__pptxRenderToken`), so no non-standard field is written onto
 * the caller's canvas and the `as unknown as` cast is gone. WeakMap keys are
 * held weakly, so a discarded canvas is collected normally.
 */
const renderTokens = new WeakMap<HTMLCanvasElement | OffscreenCanvas, number>();

/** Invalidate an in-flight main-thread render before restoring or reusing its
 * caller-owned target. The renderer observes the same token after every await. */
export function invalidatePptxRenderTarget(
  target: HTMLCanvasElement | OffscreenCanvas,
): void {
  renderTokens.set(target, (renderTokens.get(target) ?? 0) + 1);
}

/**
 * Render a single slide onto a <canvas> element.
 * Returns the canvas for convenience.
 */
export async function renderSlide(
  canvas: HTMLCanvasElement | OffscreenCanvas,
  slide: Slide,
  slideWidth: number,
  slideHeight: number,
  opts: SlideRenderOptions = {},
  onTextRun?: TextRunCallback,
): Promise<HTMLCanvasElement | OffscreenCanvas> {
  return renderSlideWithEmbeddedFonts(
    canvas,
    slide,
    slideWidth,
    slideHeight,
    opts,
    onTextRun,
  );
}

/** @internal Presentation/worker entry that carries presentation-scoped font
 * aliases without widening {@link renderSlide}'s public options. */
export async function renderSlideWithEmbeddedFonts(
  canvas: HTMLCanvasElement | OffscreenCanvas,
  slide: Slide,
  slideWidth: number,
  slideHeight: number,
  opts: InternalSlideRenderOptions = {},
  onTextRun?: TextRunCallback,
): Promise<HTMLCanvasElement | OffscreenCanvas> {
  // The core document admission queue serializes image-bearing paints for this
  // deck, so concurrent slide renders cannot each consume the full budget. Its
  // render-pass lease also pins decoded bitmaps: LRU eviction removes the cache
  // entry immediately but defers GPU close until this pass ends, so drawImage
  // never receives a closed bitmap between an await and the actual draw.
  const bitmapOwner = opts.fetchImage
    ?? (opts.fetchMedia ? posterFetchImage(opts.fetchMedia) : undefined);
  // Claim the target before queue admission. A newer zoom/navigation request
  // can supersede this one while it waits behind another image-bearing paint;
  // the stale job must then exit before resizing or clearing the canvas.
  const myToken = (renderTokens.get(canvas) ?? 0) + 1;
  renderTokens.set(canvas, myToken);
  const superseded = () => renderTokens.get(canvas) !== myToken;
  const paint = () => superseded()
    ? Promise.resolve(canvas)
    : renderSlideLeased(
      canvas,
      slide,
      slideWidth,
      slideHeight,
      opts,
      onTextRun,
      bitmapOwner,
      superseded,
    );
  return bitmapOwner && slideMayDecodeImages(slide)
    ? withBitmapCacheLease(bitmapOwner, opts.imageResources, paint)
    : paint();
}

/** {@link renderSlide}'s body, verbatim; runs under the caller's render-pass
 *  lease. */
async function renderSlideLeased(
  canvas: HTMLCanvasElement | OffscreenCanvas,
  slide: Slide,
  slideWidth: number,
  slideHeight: number,
  opts: InternalSlideRenderOptions = {},
  onTextRun?: TextRunCallback,
  bitmapOwner?: PosterFetchImage,
  superseded: () => boolean = () => false,
): Promise<HTMLCanvasElement | OffscreenCanvas> {
  // Cancellation guard. renderSlide is async (it awaits image / equation decode),
  // so rapid navigation can start a newer render of the SAME canvas before this
  // one finishes. The caller stamps a per-canvas token; once a newer render
  // supersedes us, we stop before clearing the canvas so only the latest
  // render's output survives. The paint itself never yields, so two renders'
  // draws cannot interleave.
  const targetWidth = opts.width ?? ((isHTMLCanvas(canvas) ? canvas.offsetWidth : 0) || 960);
  const scale = targetWidth / slideWidth;
  const canvasW = Math.round(targetWidth);
  const canvasH = Math.round(slideHeight * scale);

  const dpr = opts.dpr ?? defaultDpr();
  // Clamp the backing store to browser canvas limits (RB5). A huge target width
  // (or large dpr × slide size) can exceed the per-axis / total-area cap, at
  // which point the browser silently allocates a smaller-or-empty buffer and the
  // slide renders blank. `clampCanvasSize` scales BOTH axes by one factor (≤ 1)
  // so the aspect ratio is kept; we fold that into the effective dpr, keep the
  // CSS box at its intended size, and the browser stretches the (slightly
  // lower-res) backing store to fill it — a visible slide beats a blank one.
  const clamped = clampCanvasSize(canvasW * dpr, canvasH * dpr);
  const effectiveDpr = clamped.clamped ? dpr * clamped.scale : dpr;
  // Resolve equation previews before image target planning, so a selected
  // fallback uses the same bounded decode plan as ordinary slide pictures.
  if (opts.math) {
    await prepareSlideMath(slide, opts.math);
    if (superseded()) return canvas;
  }
  // Keep model locators separate from paint-order indices. Preview elements
  // exist only in this projection; source hit testing still uses the input slide.
  let sourceIndices: (number | undefined)[] | undefined;
  const sourceSlide = slide;
  if (opts.math) {
    // Resolve whole MCE branches atomically, including branches that emit more
    // than one shape. A prefix count keeps failure checks linear in the slide's
    // emitted elements rather than rescanning overlapping branch ranges.
    const failures = [0];
    const previews = new Map<number, { end: number; elements: SlideElement[] }>();
    for (const element of slide.elements) {
      let failed = false;
      for (const body of elementTextBodies(element)) {
        if (body.paragraphs.some(para => para.runs.some(run => run.type === 'math' && failedMath.has(run.nodes)))) {
          failed = true;
          break;
        }
      }
      failures.push(failures[failures.length - 1] + (failed ? 1 : 0));
    }
    slide.elements.forEach((element, index) => {
      if ((element.type !== 'shape' && element.type !== 'table') || !element.mathFallback) return;
      const fallback = element.mathFallback;
      const start = index - fallback.before;
      const end = index + fallback.after + 1;
      if (start >= 0 && end <= slide.elements.length && failures[end] > failures[start]
        && end > (previews.get(start)?.end ?? start)) {
        previews.set(start, { end, elements: fallback.elements });
      }
    });
    const elements: SlideElement[] = [];
    sourceIndices = [];
    for (let i = 0; i < slide.elements.length;) {
      const preview = previews.get(i);
      if (preview) {
        for (const element of preview.elements) {
          elements.push(element);
          sourceIndices.push(undefined);
        }
        i = preview.end;
      } else {
        sourceIndices.push(i);
        elements.push(slide.elements[i++]);
      }
    }
    slide = { ...slide, elements };
  }

  const plannedImages = planSlideImages(
    slide,
    canvasW,
    canvasH,
    scale,
    effectiveDpr,
    opts.imageResources,
    opts.fetchImage,
    opts.fetchMedia,
    bitmapOwner,
    opts.tiff,
  );
  const imagePlan = await plannedImages;
  // Resizing the backing store clears it, so it starts the slide's paint. Every
  // asynchronous input is resolved before this is called, and from here to the
  // end of the paint nothing yields (see PreparedPaint).
  const beginPaint = (): CanvasRenderingContext2D => {
    canvas.width = clamped.width;
    canvas.height = clamped.height;
    // CSS size only applies to the visible HTMLCanvasElement (not OffscreenCanvas)
    if (isHTMLCanvas(canvas)) {
      canvas.style.width = `${canvasW}px`;
      // Mirror the docx renderer: when callers use `renderSlide(canvas, ...)`
      // directly without the {@link PptxViewer} wrapper, set `display:block`
      // as a safety net so the inline-element baseline does not leave a
      // descender gap below the canvas. Respect any user-specified value.
      if (!canvas.style.display) canvas.style.display = 'block';
    }

    const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null;
    if (!ctx) throw new Error('Could not get 2D context');
    // Use the effective dpr (folded with any clamp factor) so drawing fills the
    // clamped backing store and crisp-offset math stays aligned with it.
    ctx.scale(effectiveDpr, effectiveDpr);
    return ctx;
  };

  // RB7 partial degradation: a slide whose part failed to parse (see the Rust
  // `broken_slide`) carries `parseError` and no elements. Paint a visible error
  // placeholder — correctly sized so navigation geometry is unchanged — instead
  // of a blank frame, and stop. Healthy slides (no parseError) are unaffected.
  if (slide.parseError) {
    drawParseErrorPlaceholder(beginPaint(), canvasW, canvasH, slide.slideNumber, slide.parseError);
    return canvas;
  }

  const defaultTextColor = slide.theme ? slide.theme.defaultTextColor : opts.defaultTextColor;
  const themeDefaultColor = defaultTextColor
    ? `#${defaultTextColor}`
    : '#000000';

  const pictureBulletImages = new Map<string, SvgImageSource | null>();
  const shapeFillImages = new Map<string, PreparedShapeFill | null>();
  const rc: RenderContext = {
    cjkFallback: opts.cjkFallback ? pptxSlideCjkFallback(slide, slide.theme ? slide.theme.majorFont : opts.majorFont ?? null, slide.theme ? slide.theme.minorFont : opts.minorFont ?? null, opts.cjkFallback) : undefined,
    themeMajorFont: slide.theme ? slide.theme.majorFont : opts.majorFont ?? null,
    themeMinorFont: slide.theme ? slide.theme.minorFont : opts.minorFont ?? null,
    themeHlinkColor: slide.theme ? slide.theme.hlinkColor : opts.hlinkColor ?? null,
    embeddedFontAliases: opts.embeddedFontAliases,
    embeddedFontAuthoredFamilies: opts.embeddedFontAuthoredFamilies,
    embeddedFontTuples: opts.embeddedFontTuples,
    officeFontRoutes: opts.officeFontRoutes,
    officeInstalledFamilies: opts.officeInstalledFamilies,
    googleSubstitutes: opts.googleSubstitutes,
    // The backing store may have been clamped below `canvasSize × dpr`; downstream
    // crisp-offset math must use the SAME effective dpr the ctx was scaled by.
    dpr: effectiveDpr,
    // Issue #805 — legible default for the synthetic SmartArt fallback shape's
    // null-colour runs, derived once per slide from the background luminance.
    smartArtFallbackTextColor: smartArtFallbackTextColor(slide.background, themeDefaultColor),
    pictureBulletImages,
    shapeFillImages,
  };

  // Steps painted before the slide's elements, in paint order. A step with a
  // failure stops further preparation, exactly as the former interleaved paint
  // stopped at the throwing await.
  const leadingSteps: PreparedPaint[] = [];
  const failed = () => leadingSteps.some((step) => step.failure);
  leadingSteps.push(await prepareBackground(
    slide.background,
    canvasW,
    canvasH,
    scale,
    opts.fetchImage,
    opts.tiff,
    opts.svgDecoder,
    imagePlan,
  ));
  if (superseded()) return canvas;

  const slideNumber = slide.slideNumber;

  // Resolve every picture and media poster concurrently, so the paint cost is
  // max(decode) instead of sum; the paint below consumes them in z-order. Each
  // preparation settles to a step (never rejects), so one that is abandoned
  // after an earlier fatal failure cannot surface an unhandled rejection.
  const elementPreparations: Promise<PreparedPaint | null>[] = failed()
    ? []
    : slide.elements.map((el): Promise<PreparedPaint | null> => {
      if (el.type === 'picture') {
        return preparePicture(
          el,
          scale,
          opts.fetchImage,
          opts.tiff,
          effectiveDpr,
          opts.svgDecoder,
          imagePlan,
        );
      }
      if (el.type === 'media') {
        return prepareMedia(
          el,
          scale,
          opts.fetchMedia,
          opts.skipMediaControls,
          bitmapOwner,
          opts.tiff,
          effectiveDpr,
          opts.svgDecoder,
          imagePlan,
        );
      }
      return Promise.resolve(null);
    }).map((preparation) => preparation.catch(
      (error: unknown): PreparedPaint => ({ ...NO_PAINT, failure: { error } }),
    ));

  const chartMarkerImages = new Map<string, CanvasImageSource | null>();
  // Picture bullets (`<a:buBlip>`, §21.1.2.4.2) and chart picture markers are
  // drawn inside synchronous layout/paint paths, so decode both before drawing.
  // Resolve every bullet image up front (deduped by path) and preserve each
  // exact prefetch result: a display-sized decode lives under a resolution-
  // specific cache key that the synchronous native-key peek cannot see.
  // Missing/failed decodes resolve to null and the marker is simply skipped.
  if (!failed() && opts.fetchImage) {
    const fetchImage = opts.fetchImage;
    const shapeFills = new Map<string, {
      fill: ImageFill;
      widthPt: number;
      heightPt: number;
      targetWidthPx?: number;
      targetHeightPx?: number;
      preserveNaturalSize: boolean;
      hasSourceCrop: boolean;
    }>();
    for (const { fill, width, height } of slideImageFillUsages(slide)) {
      if (!(width > 0) || !(height >= 0)) continue;
      const key = shapeFillKey(fill);
      const prior = shapeFills.get(key);
      const widthPt = width / PT_TO_EMU;
      const heightPt = height / PT_TO_EMU;
      const fr = fill.fillRect ?? {};
      const target = rasterTargetOptions(
        emuToPx(width, scale) * (1 - (fr.l ?? 0) - (fr.r ?? 0)),
        emuToPx(height, scale) * (1 - (fr.t ?? 0) - (fr.b ?? 0)),
        effectiveDpr,
        fill.srcRect,
      );
      shapeFills.set(key, prior
        ? {
            fill: prior.fill,
            widthPt: Math.max(prior.widthPt, widthPt),
            heightPt: Math.max(prior.heightPt, heightPt),
            targetWidthPx: Math.max(prior.targetWidthPx ?? 0, target?.targetWidthPx ?? 0) || undefined,
            targetHeightPx: Math.max(prior.targetHeightPx ?? 0, target?.targetHeightPx ?? 0) || undefined,
            preserveNaturalSize: prior.preserveNaturalSize || fill.tile != null,
            hasSourceCrop: prior.hasSourceCrop || fill.srcRect != null,
          }
        : {
            fill, widthPt, heightPt,
            targetWidthPx: target?.targetWidthPx,
            targetHeightPx: target?.targetHeightPx,
            preserveNaturalSize: fill.tile != null,
            hasSourceCrop: fill.srcRect != null,
          });
    }
    const bulletPaths = new Map<string, { mimeType: string; targetHeightPx?: number }>();
    // A chart whose frame or derived decode size is non-positive or non-finite
    // cannot paint an image safely. Exclude it before aggregate source gating
    // so it cannot suppress visible preloads or force a native-size decode.
    const chartDescriptors: Array<{
      element: ChartElement;
      frame: Parameters<typeof chartImageFillUsageSize>[1];
      usages: Array<{
        usage: ReturnType<typeof collectChartImageFillUsages>[number];
        size: NonNullable<ReturnType<typeof chartImageFillUsageSize>>;
      }>;
    }> = [];
    for (const element of slide.elements) {
      if (element.type !== 'chart'
        || !Number.isFinite(element.width) || element.width <= 0
        || !Number.isFinite(element.height) || element.height <= 0) continue;
      const frame = {
        widthPt: element.width / PT_TO_EMU,
        heightPt: element.height / PT_TO_EMU,
        targetWidthPx: emuToPx(element.width, scale) * effectiveDpr,
        targetHeightPx: emuToPx(element.height, scale) * effectiveDpr,
      };
      const usages = [] as typeof chartDescriptors[number]['usages'];
      let valid = true;
      for (const usage of collectChartImageFillUsages(element.chart)) {
        const size = chartImageFillUsageSize(usage, frame);
        if (!size) {
          valid = false;
          break;
        }
        usages.push({ usage, size });
      }
      if (valid) chartDescriptors.push({ element, frame, usages });
    }
    const chartFillMap = new Map<string, {
      fill: ReturnType<typeof collectChartImageFillUsages>[number]['fill'];
      widthPt: number;
      heightPt: number;
      targetWidthPx?: number;
      targetHeightPx?: number;
      maxRetainedPixels?: number;
      preserveNaturalSize: boolean;
      hasSourceCrop: boolean;
    }>();
    for (const usage of collectChartImageFillUsagesForCharts(
      chartDescriptors.map(({ element }) => element.chart),
      (usage, chartIndex) => chartImageFillUsageSize(
        usage,
        chartDescriptors[chartIndex]!.frame,
      ) != null,
    )) {
      const { fill } = usage;
      const key = chartImageFillKey(fill);
      if (!chartFillMap.has(key)) chartFillMap.set(key, {
        fill,
        widthPt: 0,
        heightPt: 0,
        preserveNaturalSize: usage.preserveNaturalSize,
        hasSourceCrop: usage.hasSourceCrop,
      });
    }
    for (const { usages } of chartDescriptors) {
      for (const { usage, size } of usages) {
        const { fill } = usage;
        const key = chartImageFillKey(fill);
        const prior = chartFillMap.get(key);
        if (!prior) continue;
        // ECMA-376 §20.1.8.58: tile geometry is based on the source's natural
        // pixel dimensions and authored DPI. A display-sized decode would alter
        // the tile size/alignment/repetition count. Because one decoded source
        // is shared by key, any tiled consumer makes the whole shared entry
        // native-sized, even when another chart stretches the same blip.
        const preserveNaturalSize = prior.preserveNaturalSize || usage.preserveNaturalSize;
        const hasSourceCrop = prior.hasSourceCrop || usage.hasSourceCrop;
        const planned = preserveNaturalSize || pixelTransform(fill)
          ? undefined
          : plannedRasterOptions(imagePlan, imagePlanKey(fill.imagePath, pixelTransform(fill)));
        const vector = !pixelTransform(fill)
          && (fill.mimeType === 'image/svg+xml' || preferVectorBlip({
            svgImagePath: fill.svgImagePath,
            srcRect: hasSourceCrop ? true : null,
          }));
        // A picture fill may cover a marker, chart/plot area, wall, or floor.
        // The owning chart frame is the smallest format-derived upper bound
        // common to all consumers. Core has already aggregated each same-chart
        // occurrence's source crop and stretch fillRect into these factors.
        chartFillMap.set(key, {
          ...prior,
          widthPt: Math.max(prior.widthPt, size.widthPt),
          heightPt: Math.max(prior.heightPt, size.heightPt),
          targetWidthPx: preserveNaturalSize
            ? undefined
            : planned?.targetWidthPx
              ?? (vector
                ? Math.max(prior.targetWidthPx ?? 0, size.targetWidthPx ?? 0) || undefined
                : undefined),
          targetHeightPx: preserveNaturalSize
            ? undefined
            : planned?.targetHeightPx
              ?? (vector
                ? Math.max(prior.targetHeightPx ?? 0, size.targetHeightPx ?? 0) || undefined
                : undefined),
          maxRetainedPixels: planned?.maxRetainedPixels,
          preserveNaturalSize,
          hasSourceCrop,
        });
      }
    }
    for (const el of slide.elements) {
      if (el.type !== 'shape' || !el.textBody) continue;
      for (const para of el.textBody.paragraphs) {
        const b = asBullet(para.bullet);
        if (b.type === 'blip') {
          // normAutofit never enlarges text, so the unshrunk authored marker
          // height is a sufficient preload target even when layout later picks
          // a smaller fontScale. Aggregate the largest use of a shared part.
          const targetHeightPx = pictureBulletSizePt(el.textBody, para, b)
            * PT_TO_EMU * scale * effectiveDpr;
          const existing = bulletPaths.get(b.imagePath);
          if (!existing) {
            bulletPaths.set(b.imagePath, {
              mimeType: b.mimeType,
              ...(Number.isFinite(targetHeightPx) && targetHeightPx > 0
                ? { targetHeightPx }
                : {}),
            });
          } else if (Number.isFinite(targetHeightPx) && targetHeightPx > 0) {
            existing.targetHeightPx = Math.max(existing.targetHeightPx ?? 0, targetHeightPx);
          }
        }
      }
    }
    if (bulletPaths.size > 0 || chartFillMap.size > 0 || shapeFills.size > 0) {
      const bulletPromises: Promise<void>[] = [...bulletPaths].map(async (
        [path, { mimeType, targetHeightPx }],
      ) => {
        try {
          const target = mimeType === 'image/svg+xml'
            ? targetHeightPx === undefined
              ? {}
              : { targetWidthPx: 1, targetHeightPx }
            : plannedRasterOptions(imagePlan, imagePlanKey(path)) ?? {};
          const image = mimeType === 'image/svg+xml'
            ? await getCachedSvgImageByPath(path, fetchImage, {
                ...target,
                workerDecoder: opts.svgDecoder,
              })
            : await getCachedBitmapByPath(path, mimeType, fetchImage, {
                tiff: opts.tiff,
                ...target,
              });
          pictureBulletImages.set(path, image);
        } catch (error) {
          if (isOptionalImageCodecUnavailableError(error, 'tiff')) {
            pictureBulletImages.set(path, null);
            return;
          }
          if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error) || isTiffDecodeError(error)) throw error;
          pictureBulletImages.set(path, null);
        }
      });
      const chartPromises: Promise<unknown>[] = [...chartFillMap].map(async ([key, entry]) => {
        const {
          fill, widthPt, heightPt, targetWidthPx, targetHeightPx,
          maxRetainedPixels, hasSourceCrop,
        } = entry;
        const target = targetWidthPx && targetHeightPx
          ? {
              targetWidthPx,
              targetHeightPx,
              ...(maxRetainedPixels === undefined ? {} : { maxRetainedPixels }),
            }
          : undefined;
        try {
          const decodeFallback = () => fill.mimeType === 'image/svg+xml'
            ? pixelTransform(fill) ? Promise.resolve(null) : getCachedSvgImageByPath(fill.imagePath, fetchImage, {
                ...(target ?? {}),
                workerDecoder: opts.svgDecoder,
              })
            : getCachedDuotoneBitmapByPath(
                fill.imagePath, fill.mimeType, pixelTransform(fill), fetchImage,
                {
                  widthPt,
                  heightPt,
                  ...(target ?? {}),
                  failClosedOnDuotoneFailure: true,
                  tiff: opts.tiff,
                },
              );
          let bitmap: CanvasImageSource | null;
          const blip = {
            svgImagePath: fill.svgImagePath,
            srcRect: hasSourceCrop ? true : null,
          };
          if (!pixelTransform(fill) && preferVectorBlip(blip)) {
            try {
              bitmap = await getCachedSvgImageByPath(blip.svgImagePath, fetchImage, {
                ...(target ?? {}),
                workerDecoder: opts.svgDecoder,
              });
            } catch (error) {
              if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error)) throw error;
              bitmap = await decodeFallback();
            }
          } else {
            bitmap = await decodeFallback();
          }
          chartMarkerImages.set(key, bitmap);
        } catch (error) {
          if (isOptionalImageCodecUnavailableError(error, 'tiff')) {
            chartMarkerImages.set(key, null);
            return;
          }
          if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error) || isTiffDecodeError(error)) throw error;
          chartMarkerImages.set(key, null);
        }
      });
      const shapePromises: Promise<void>[] = [...shapeFills].map(async (
        [key, {
          fill, widthPt, heightPt, targetWidthPx, targetHeightPx,
          preserveNaturalSize, hasSourceCrop,
        }],
      ) => {
        const vector = !pixelTransform(fill)
          && (fill.mimeType === 'image/svg+xml' || preferVectorBlip({
            svgImagePath: fill.svgImagePath,
            srcRect: hasSourceCrop ? true : null,
          }));
        const planned = !preserveNaturalSize && !pixelTransform(fill)
          ? plannedRasterOptions(imagePlan, imagePlanKey(fill.imagePath, pixelTransform(fill)))
          : undefined;
        const rawTarget = !preserveNaturalSize && targetWidthPx && targetHeightPx
          ? { targetWidthPx, targetHeightPx }
          : undefined;
        const target = vector ? rawTarget : planned;
        try {
          const intrinsicSize = preserveNaturalSize
            ? (await inspectCachedRasterSource(fill.imagePath, fill.mimeType, fetchImage))
                .dimensions ?? undefined
            : undefined;
          const fallback = () => fill.mimeType === 'image/svg+xml'
            ? pixelTransform(fill)
              ? Promise.resolve(null)
              : getCachedSvgImageByPath(fill.imagePath, fetchImage, {
                  ...(target ?? {}), workerDecoder: opts.svgDecoder,
                })
            : getCachedDuotoneBitmapByPath(
                fill.imagePath, fill.mimeType, pixelTransform(fill), fetchImage,
                {
                  widthPt,
                  heightPt,
                  ...(target ?? {}),
                  failClosedOnDuotoneFailure: true,
                  tiff: opts.tiff,
                  svgDecoder: opts.svgDecoder,
                },
              );
          let image: SvgImageSource | null;
          if (!pixelTransform(fill) && preferVectorBlip({
            svgImagePath: fill.svgImagePath,
            srcRect: hasSourceCrop ? true : null,
          })) {
            try {
              image = await getCachedSvgImageByPath(fill.svgImagePath as string, fetchImage, {
                ...(target ?? {}), workerDecoder: opts.svgDecoder,
              });
            } catch (error) {
              if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error)) throw error;
              image = await fallback();
            }
          } else {
            image = await fallback();
          }
          shapeFillImages.set(key, image ? { image, intrinsicSize } : null);
        } catch (error) {
          if (isOptionalImageCodecUnavailableError(error, 'tiff')) {
            shapeFillImages.set(key, null);
            return;
          }
          if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error) || isTiffDecodeError(error)) throw error;
          shapeFillImages.set(key, null);
        }
      });
      try {
        await Promise.all([...bulletPromises, ...chartPromises, ...shapePromises]);
      } catch (error) {
        leadingSteps.push({ ...NO_PAINT, failure: { error } });
      }
      if (superseded()) return canvas;
    }
  }

  // Consume the preparations in paint order, as the former sequential paint
  // awaited them. At the first fatal failure, stop: steps after it are
  // abandoned (never awaited or painted), so a later hanging fetch cannot delay
  // the rejection, and an earlier fatal failure always takes precedence.
  const elementSteps: (PreparedPaint | null)[] = [];
  if (!failed()) {
    for (const preparation of elementPreparations) {
      const step = await preparation;
      // A newer render of this canvas started while we awaited an image —
      // stop before clearing the canvas so we don't paint this stale slide.
      if (superseded()) return canvas;
      elementSteps.push(step);
      if (step?.failure) break;
    }
  }
  if (superseded()) return canvas;

  // ---- Paint: synchronous from here on, so the whole slide is one task. ----
  const ctx = beginPaint();
  for (const step of leadingSteps) {
    step.paint(ctx);
    if (step.failure) throw step.failure.error;
  }
  // renderSlide installed only this device scale on the root context. Keep it
  // as the slide frame while each element adds local rotations or reflections.
  const slidePatternFrame = { a: effectiveDpr, d: effectiveDpr };
  for (const [elementIndex, el] of slide.elements.entries()) {
    // A text-run callback may start a newer render of this canvas; stop so we
    // don't paint this (now stale) slide over the newer one.
    if (superseded()) return canvas;
    const sourceIndex = sourceIndices ? sourceIndices[elementIndex] : elementIndex;
    const elementTextRun: TextRunCallback | undefined = onTextRun
      ? (run) => {
          if (sourceIndex === undefined) {
            // Preview ids belong to the unselected MCE branch, so neither its
            // shape id nor a projected array index can address the source model.
            onTextRun({
              ...run,
              renderElementIndex: elementIndex,
              elementIndex: undefined,
              origin: undefined,
              shapeId: undefined,
            });
          } else {
            onTextRun({
              ...run,
              renderElementIndex: elementIndex,
              elementIndex: sourceIndex,
              origin: sourceSlide.elementSources?.[sourceIndex]?.origin ?? 'slide',
            });
          }
        }
      : undefined;
    if (el.type === 'shape') {
      withPatternCoordinateSpace(ctx, slidePatternFrame, () =>
        renderShape(ctx, el, scale, themeDefaultColor, slideNumber, rc, elementTextRun, opts.fetchImage));
    } else if (el.type === 'picture' || el.type === 'media') {
      const step = elementSteps[elementIndex];
      if (step) {
        withPatternCoordinateSpace(ctx, slidePatternFrame, () => step.paint(ctx));
        if (step.failure) throw step.failure.error;
      }
    } else if (el.type === 'table') {
      withPatternCoordinateSpace(ctx, slidePatternFrame, () =>
        renderTable(ctx, el, scale, slideNumber, rc, elementTextRun));
    } else if (el.type === 'chart') {
      // OOXML: 1pt = 12700 EMU. The slide renderer's `scale` is px-per-EMU,
      // so PT_TO_EMU * scale gives pixels-per-point at the current display size.
      const chartPtToPx = PT_TO_EMU * scale;
      // `el.chart` is already the canonical ChartModel emitted by the Rust
      // parser (`ooxml_common::chart::ChartModel`) — no per-field adapter.
      withPatternCoordinateSpace(ctx, slidePatternFrame, () => {
        ctx.save();
        applyFrameTransform(ctx, el, scale);
        renderChart(
          ctx,
          el.chart,
          {
            x: emuToPx(el.x, scale),
            y: emuToPx(el.y, scale),
            w: emuToPx(el.width, scale),
            h: emuToPx(el.height, scale),
          },
          chartPtToPx,
          el.rotation,
          opts.threeD,
          opts.regionMap,
          fill => chartMarkerImages.get(
            chartImageFillKey(fill),
          ),
          opts.chartEx,
        );
        ctx.restore();
      });
    }
  }

  // Hidden-slide dimming (a render mechanism; the caller decides WHEN — see
  // PptxViewer's 'dim' mode). Overlay AFTER all content so it reads faintly.
  if (superseded()) return canvas;
  if (opts.dim) applyDimOverlay(ctx, opts.dim, canvasW, canvasH);

  return canvas;
}
