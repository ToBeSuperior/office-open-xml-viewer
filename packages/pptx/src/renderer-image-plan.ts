// Per-slide decoded-image planning for the PPTX renderer: collects every
// picture, image fill, and bullet decode target before any decode starts.
// Split from renderer.ts (#1857).

import type { Slide, SlideElement, Fill, ImageFill } from './types';
import { asBullet } from './types';
import {
  EMU_PER_PT as PT_TO_EMU,
  chartImageFillUsageSize,
  collectChartImageFillUsages,
  normalizeImageResourceOptions,
  planDecodedImageTargets,
  inspectCachedRasterSource,
  isBrowserResizableRasterMimeType,
  isDecodeTargetResizableRasterFormat,
  preferVectorBlip,
  isOoxmlDecodedImageLimitError,
  OoxmlResourceLimitError,
} from '@silurus/ooxml-core';
import type { DecodedBitmapCacheOwner, DecodedImageTargetDemand } from '@silurus/ooxml-core';
import type {
  TiffRenderer,
  ImageResourceOptions,
  DecodedImageTargetPlan,
} from '@silurus/ooxml-core';
import {
  rasterTargetOptions,
  pixelTransform,
  imagePlanKey,
  emuToPx,
  bulletTargetPx,
} from './renderer-context';
import { shapeImageFillModeIsPaintable } from './renderer-shape-fill';
import { posterFetchImage } from './renderer-picture';

/** Build the complete visible-slide browser-raster demand before decoding.
 * Every display target comes from authored geometry, crop, and effective DPR.
 * Source inspection preserves native output while it fits, then lets the
 * shared planner make the aggregate native/display/quality decision. */
export async function planSlideImages(
  slide: Slide,
  canvasW: number,
  canvasH: number,
  scale: number,
  dpr: number,
  imageResources?: ImageResourceOptions,
  fetchImage?: (path: string, mime: string) => Promise<Blob>,
  fetchMedia?: (path: string) => Promise<Blob>,
  bitmapOwner?: DecodedBitmapCacheOwner,
  tiff?: TiffRenderer,
): Promise<DecodedImageTargetPlan> {
  const policy = normalizeImageResourceOptions(imageResources);
  const demands: DecodedImageTargetDemand[] = [];
  const pending: Array<Promise<DecodedImageTargetDemand | null>> = [];
  const push = (
    key: string,
    target: ReturnType<typeof rasterTargetOptions>,
    imagePath: string,
    mimeType: string,
    loader: ((path: string, mime: string) => Promise<Blob>) | undefined,
    retainedSurfaceCount = 1,
    owner: DecodedBitmapCacheOwner | undefined = loader,
  ) => {
    if (!target || !loader) return;
    // Common browser rasters have a bounded, axis-wise display target already.
    // Admit that target immediately so each source inspection can pipeline into
    // its decode instead of placing an all-slide inspection barrier before the
    // first bitmap. The decode boundary still sniffs the bytes and enforces all
    // non-configurable source, dimension, and per-surface ceilings.
    if (isBrowserResizableRasterMimeType(mimeType)
      && (policy.resolution === 'display' || policy.strategy === 'adaptive')) {
      demands.push({ key, ...target, retainedSurfaceCount });
      return;
    }
    pending.push(inspectCachedRasterSource(imagePath, mimeType, loader, owner)
      .then((inspection) => inspection.dimensions
        && isDecodeTargetResizableRasterFormat(inspection.format, tiff !== undefined)
        ? {
            key,
            ...target,
            sourceWidthPx: inspection.dimensions.width,
            sourceHeightPx: inspection.dimensions.height,
            retainedSurfaceCount,
          }
        : null)
      .catch((error: unknown) => {
        if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error)) throw error;
        return null;
      }));
  };

  const background = slide.background;
  if (background?.fillType === 'image' && background.imagePath
    && !background.tile && !pixelTransform(background)) {
    const fr = background.fillRect ?? {};
    const width = canvasW * (1 - (fr.l ?? 0) - (fr.r ?? 0));
    const height = canvasH * (1 - (fr.t ?? 0) - (fr.b ?? 0));
    push(
      imagePlanKey(background.imagePath, pixelTransform(background)),
      rasterTargetOptions(width, height, dpr, background.srcRect),
      background.imagePath,
      background.mimeType,
      fetchImage,
      1,
    );
  }

  for (const element of slide.elements) {
    for (const { fill, width, height } of elementImageFillUsages(element)) {
      if (!fill.tile && !pixelTransform(fill)) {
        const fr = fill.fillRect ?? {};
        const targetWidth = emuToPx(width, scale) * (1 - (fr.l ?? 0) - (fr.r ?? 0));
        const targetHeight = emuToPx(height, scale) * (1 - (fr.t ?? 0) - (fr.b ?? 0));
        const vector = fill.mimeType === 'image/svg+xml' || preferVectorBlip(fill);
        if (!vector) push(
          imagePlanKey(fill.imagePath, pixelTransform(fill)),
          rasterTargetOptions(targetWidth, targetHeight, dpr, fill.srcRect),
          fill.imagePath,
          fill.mimeType,
          fetchImage,
          1,
        );
      }
    }
    if (element.type === 'picture') {
      const vector = preferVectorBlip(element) || element.mimeType === 'image/svg+xml';
      if (!vector && !pixelTransform(element)) {
        push(
          imagePlanKey(element.imagePath, pixelTransform(element)),
          rasterTargetOptions(
            emuToPx(element.width, scale),
            emuToPx(element.height, scale),
            dpr,
            element.srcRect,
          ),
          element.imagePath,
          element.mimeType,
          fetchImage,
          1,
        );
      }
    } else if (element.type === 'media' && element.posterPath) {
      const loader = fetchMedia ? posterFetchImage(fetchMedia) : undefined;
      push(
        imagePlanKey(element.posterPath),
        rasterTargetOptions(
          emuToPx(element.width, scale),
          emuToPx(element.height, scale),
          dpr,
        ),
        element.posterPath,
        element.posterMimeType || 'application/octet-stream',
        loader,
        1,
        bitmapOwner ?? loader,
      );
    } else if (element.type === 'chart') {
      const frame = {
        widthPt: element.width / PT_TO_EMU,
        heightPt: element.height / PT_TO_EMU,
        targetWidthPx: emuToPx(element.width, scale) * dpr,
        targetHeightPx: emuToPx(element.height, scale) * dpr,
      };
      for (const usage of collectChartImageFillUsages(element.chart)) {
        const fill = usage.fill;
        const size = chartImageFillUsageSize(usage, frame);
        const vector = fill.mimeType === 'image/svg+xml' || preferVectorBlip({
          svgImagePath: fill.svgImagePath,
          srcRect: usage.hasSourceCrop ? true : null,
        });
        if (!vector && !pixelTransform(fill) && !usage.preserveNaturalSize
          && size?.targetWidthPx && size.targetHeightPx) {
          push(
            imagePlanKey(fill.imagePath, pixelTransform(fill)),
            {
              targetWidthPx: size.targetWidthPx,
              targetHeightPx: size.targetHeightPx,
            },
            fill.imagePath,
            fill.mimeType,
            fetchImage,
            1,
          );
        }
      }
    } else if (element.type === 'shape') {
      if (!element.textBody) continue;
      for (const paragraph of element.textBody.paragraphs) {
        const bullet = asBullet(paragraph.bullet);
        if (bullet.type !== 'blip') continue;
        const size = bulletTargetPx(element.textBody, paragraph, scale, dpr);
        if (!size || !fetchImage) continue;
        pending.push(inspectCachedRasterSource(
          bullet.imagePath,
          bullet.mimeType,
          fetchImage,
        ).then((inspection) => inspection.dimensions
          && inspection.dimensions.width > 0
          && inspection.dimensions.height > 0
          && isDecodeTargetResizableRasterFormat(inspection.format, tiff !== undefined)
          ? {
              key: imagePlanKey(bullet.imagePath),
              targetWidthPx: size * inspection.dimensions.width / inspection.dimensions.height,
              targetHeightPx: size,
              sourceWidthPx: inspection.dimensions.width,
              sourceHeightPx: inspection.dimensions.height,
            }
          : null).catch((error: unknown) => {
            if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error)) throw error;
            return null;
          }));
      }
    }
  }
  demands.push(...(await Promise.all(pending))
    .filter((demand): demand is DecodedImageTargetDemand => demand !== null));
  return planDecodedImageTargets(demands, policy);
}

/** ECMA-376 §§20.1.4.2.25/27 and 21.1.3.17 use the same
 * EG_FillProperties as shapes. Styles and bands have already cascaded into
 * cell.fill; only merge anchors own a painted fill. Use authored grid extents
 * for decode demand (the unchanged text measurement owns the final paint box).
 * A zero-minimum row uses the authored frame as its decode target, without
 * changing or predicting text-driven growth. If both are zero, preparation
 * uses a budget-checked native decode: text can still grow the painted row. */
function* elementImageFillUsages(element: SlideElement): Generator<{
  fill: ImageFill; width: number; height: number;
}> {
  const usage = (fill: Fill | null | undefined, width: number, height: number) =>
    fill?.fillType === 'image' && shapeImageFillModeIsPaintable(fill)
      ? { fill, width, height }
      : null;
  if (element.type === 'shape') {
    const own = usage(element.fill, element.width, element.height);
    if (own && element.height > 0) yield own;
  } else if (element.type === 'table') {
    // Prefix sums keep merged-cell resource planning linear in the grid,
    // independent of the authored span lengths; no per-cell slice allocation.
    const colExtents = [0];
    for (const col of element.cols) colExtents.push(colExtents[colExtents.length - 1] + col);
    const rowExtents = [0];
    for (const row of element.rows) rowExtents.push(rowExtents[rowExtents.length - 1] + row.height);
    const width = colExtents[colExtents.length - 1];
    const height = rowExtents[rowExtents.length - 1] || element.height;
    const background = usage(element.background, width, height);
    if (background) yield background;
    for (let ri = 0; ri < element.rows.length; ri++) {
      const row = element.rows[ri];
      for (let ci = 0; ci < row.cells.length; ci++) {
        const cell = row.cells[ci];
        if (cell.hMerge || cell.vMerge || cell.fill?.fillType !== 'image'
          || !shapeImageFillModeIsPaintable(cell.fill)) continue;
        const cellWidth = (colExtents[Math.min(ci + (cell.gridSpan || 1), element.cols.length)] ?? 0)
          - (colExtents[ci] ?? 0);
        const cellHeight = (rowExtents[Math.min(ri + (cell.rowSpan || 1), element.rows.length)]
          - rowExtents[ri]) || element.height;
        const own = usage(cell.fill, cellWidth, cellHeight);
        if (own) yield own;
      }
    }
  }
}

export function* slideImageFillUsages(slide: Slide) {
  for (const element of slide.elements) yield* elementImageFillUsages(element);
}

export function slideMayDecodeImages(slide: Slide): boolean {
  if (slide.background?.fillType === 'image') return true;
  return slide.elements.some((element) => {
    if (element.type === 'picture') return true;
    if (element.type === 'media') return !!element.posterPath;
    if (element.type === 'chart') return collectChartImageFillUsages(element.chart).length > 0;
    if (!elementImageFillUsages(element).next().done) return true;
    if ((element.type === 'shape' || element.type === 'table') && element.mathFallback?.elements.length) return true;
    return element.type === 'shape' && (
      element.fill?.fillType === 'image'
      || !!element.textBody?.paragraphs.some(paragraph => asBullet(paragraph.bullet).type === 'blip')
    );
  });
}
