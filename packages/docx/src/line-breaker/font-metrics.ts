import type { CanvasFontRoute, ResolvedFontMetric } from '@silurus/ooxml-core';
import { localFontMetricCoversText, normalizeFontMetricFamily } from '@silurus/ooxml-core';
import type { FontResolution } from '../layout/font-service.js';
import { referenceFontLineMetrics } from '../reference-font-line-metrics.js';

/** Vertical admission shared by ordinary text and numbering markers. Resource
 * identity, tuple ambiguity and glyph coverage are resolved before this gate.
 * Canvas sides remain the fallback; exact spacing suppresses design geometry. */
export function selectedFontLineMetric(
  selected: FontResolution | undefined,
  localFont: ResolvedFontMetric | undefined,
  naturalMetricAllowed: boolean,
  referenceMetricAllowed = true,
) {
  const resourceMetric = (naturalMetricAllowed || localFont?.designAscentRatio == null)
    && (localFont?.lineHeightRatio != null || localFont?.designAscentRatio != null
      || localFont?.eastAsianLineHeightRatio != null) ? localFont : undefined;
  const referenceMetric = naturalMetricAllowed && referenceMetricAllowed && !resourceMetric
    && mayUseAuthoredReferenceVerticalMetric(selected)
    ? referenceFontLineMetrics(selected.requestedFamily, selected.weight, selected.style)
    : undefined;
  return { resourceMetric, referenceMetric, lineMetric: resourceMetric ?? referenceMetric };
}

/** A subset face lends design metrics only to spans whose every scalar it owns.
 * A FontFace tuple can win CSS selection while Canvas silently paints missing
 * glyphs from another face. Ranges are normalized once at the service boundary.
 *
 * `paintsGlyphs` is false only for a probe that selects a face without painting
 * it: the paragraph mark (§17.3.1.29). Painted-glyph ownership exists to keep a
 * span from borrowing the geometry of a face Canvas did not draw it with; a
 * mark draws nothing, so its box is the selected resource's own line box even
 * when that face lacks the script-matched probe glyph (for example an
 * East Asian slot authored with a Latin face). Unicode-range subsets still
 * gate the match, because they decide which registered resource is selected. */
export function metricCoversText(
  metric: ResolvedFontMetric, text: string, route?: CanvasFontRoute, paintsGlyphs = true,
): boolean {
  if (paintsGlyphs && metric.requiresGlyphCoverage
    && !localFontMetricCoversText(metric, text, route)) return false;
  const ranges = metric.unicodeRanges;
  if (ranges === undefined || text.length === 0) return true;
  for (const scalar of text) {
    const codePoint = scalar.codePointAt(0)!;
    let low = 0;
    let high = ranges.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (ranges[middle][1] < codePoint) low = middle + 1;
      else high = middle;
    }
    if (low >= ranges.length || ranges[low][0] > codePoint) return false;
  }
  return true;
}


/** An authored-family key alone cannot prove the selected route. Match the
 * chosen face tuple and acquisition source before admitting its metric;
 * a substitute may preserve the request while painting another face. */
export function metricMatchesSelectedFont(
  metric: ResolvedFontMetric,
  selected: FontResolution,
  text: string,
  paintsGlyphs = true,
): boolean {
  if (resourceMetricKey(metric) !== resolutionMetricKey(selected)
    || !metricCoversText(metric, text, selected.route, paintsGlyphs)) return false;
  const identity = metric.sourceIdentity;
  if (selected.source === 'embedded') {
    // The production loader routes Canvas to a resource-derived private family.
    // Its metric must come from that same registration, including when another
    // document embeds different bytes under the same authored family/style.
    return selected.resourceIdentity === undefined
      ? identity?.startsWith('embedded:') === true
      : identity === selected.resourceIdentity;
  }
  if (selected.source === 'local' || selected.source === 'substitute') {
    // The same CSS tuple may be supplied by an unrelated caller or another
    // loaded face. Only the resource that created this resolver route may lend
    // vertical geometry; a matching family name is insufficient.
    return identity !== undefined && identity === selected.resourceIdentity;
  }
  return false;
}


/** Multiple resources may be registered under one CSS family/style. The
 * authored alias does not control Canvas's choice among overlapping faces;
 * only geometry shared by every covering resource is safe to admit. */
export function sameResourceGeometry(a: ResolvedFontMetric, b: ResolvedFontMetric): boolean {
  return a.farEastCodePage === b.farEastCodePage
    && a.lineHeightRatio === b.lineHeightRatio
    && a.designAscentRatio === b.designAscentRatio
    && a.designDescentRatio === b.designDescentRatio
    && a.lineGapRatio === b.lineGapRatio
    && a.eastAsianLineHeightRatio === b.eastAsianLineHeightRatio
    && a.fontBoxRatio === b.fontBoxRatio;
}


export type MetricTupleIndex = ReadonlyMap<string, readonly ResolvedFontMetric[]>;

export const metricTupleIndexes = new WeakMap<object, MetricTupleIndex>();


export function selectedMetricKey(
  family: string,
  weight: number,
  style: string,
): string {
  return `${normalizeFontMetricFamily(family)}\0${weight}\0${style}`;
}


const resourceMetricKeys = new WeakMap<ResolvedFontMetric, string>();

function resourceMetricKey(metric: ResolvedFontMetric): string {
  const cached = resourceMetricKeys.get(metric);
  if (cached !== undefined) return cached;
  const key = selectedMetricKey(metric.family, metric.weight ?? 400, metric.style ?? 'normal');
  if (Object.isFrozen(metric)) resourceMetricKeys.set(metric, key);
  return key;
}

const resolutionMetricKeys = new WeakMap<FontResolution, string>();

/** Retained resolutions are immutable and shared by every span of a tuple.
 * Local aliases can be long: normalize once rather than at every line-metric
 * and catalog-admission lookup. Mutable custom resolutions stay uncached. */
function resolutionMetricKey(selected: FontResolution): string {
  const cached = resolutionMetricKeys.get(selected);
  if (cached !== undefined) return cached;
  const key = selectedMetricKey(selected.resolvedFamily, selected.weight, selected.style);
  if (Object.isFrozen(selected)) resolutionMetricKeys.set(selected, key);
  return key;
}


/** A frozen service snapshot is immutable for the document lifetime, so one
 * tuple/source index can serve every paragraph. A custom mutable service gets
 * one temporary index per synchronous buildSegments call, then can change its
 * entries before the next call. */
export function indexedFontMetrics(
  metrics: Readonly<Record<string, ResolvedFontMetric>>,
): MetricTupleIndex {
  const frozenRoot = Object.isFrozen(metrics);
  if (frozenRoot) {
    const retained = metricTupleIndexes.get(metrics);
    if (retained) return retained;
  }
  const values = Object.values(metrics);
  const cacheable = frozenRoot && values.every((metric) => Object.isFrozen(metric));
  const mutable = new Map<string, ResolvedFontMetric[]>();
  for (const metric of values) {
    const key = resourceMetricKey(metric);
    const candidates = mutable.get(key);
    if (candidates) candidates.push(metric);
    else mutable.set(key, [metric]);
  }
  for (const candidates of mutable.values()) Object.freeze(candidates);
  if (cacheable) metricTupleIndexes.set(metrics, mutable);
  return mutable;
}


export function selectResourceMetric(
  index: MetricTupleIndex,
  selected: FontResolution,
  probeText: string,
  paintsGlyphs = true,
): ResolvedFontMetric | undefined {
  if (selected.source === 'native') {
    // The measured line describes the route's first available face, not the
    // absent authored one; it lends geometry only to clusters that face draws
    // (or, for a paragraph mark that draws nothing, to the face it selects).
    const metric = selected.routeLineMetric;
    return metric && metricCoversText(metric, probeText, selected.route, paintsGlyphs)
      ? metric : undefined;
  }
  if (selected.source !== 'embedded' && selected.source !== 'local'
    && selected.source !== 'substitute') return undefined;
  const candidates = index.get(resolutionMetricKey(selected)) ?? [];
  let first: ResolvedFontMetric | undefined;
  for (const metric of candidates) {
    if (!metricMatchesSelectedFont(metric, selected, probeText, paintsGlyphs)) continue;
    first ??= metric;
  }
  if (!first) return undefined;
  // Source identity proves which resource was requested, but not which of the
  // overlapping CSS faces Canvas actually paints. Compare every covering
  // resource in the tuple, including resources acquired from other sources.
  for (const metric of candidates) {
    if (metricCoversText(metric, probeText, selected.route, paintsGlyphs)
      && !sameResourceGeometry(first, metric)) return undefined;
  }
  return first;
}


/** Width inference has a stricter, independent projection. Conflicting xAvg
 * metadata cannot revoke a valid vertical metric for the same selected face. */
export function selectResourceAverageWidthRatio(
  index: MetricTupleIndex,
  selected: FontResolution,
  probeText: string,
): number | undefined {
  if (selected.source !== 'embedded' && selected.source !== 'local'
    && selected.source !== 'substitute') return undefined;
  const candidates = index.get(resolutionMetricKey(selected)) ?? [];
  const withSpace = `${probeText} `;
  const first = candidates.find((metric) => metricMatchesSelectedFont(metric, selected, withSpace));
  const ratio = first?.averageCharWidthRatio;
  if (ratio == null || ratio <= 0) return undefined;
  return candidates.every((metric) => !metricCoversText(metric, withSpace, selected.route)
    || metric.averageCharWidthRatio === ratio) ? ratio : undefined;
}


export function mayUseExactLocalReferenceWidthMetric(
  selected: FontResolution | undefined,
): selected is FontResolution {
  // The width floor requires an exact selected face. A native CSS family list
  // cannot establish which fallback Canvas painted or measured the glyph.
  return selected?.source === 'local'
    && selected.resourceIdentity?.startsWith('office-local:') === true;
}


export function mayUseAuthoredReferenceVerticalMetric(
  selected: FontResolution | undefined,
): selected is FontResolution {
  // Layout follows the face that actually paints on the host. A native CSS
  // family list whose authored face is not installed paints with a browser
  // fallback, so the authored face's catalog line box would size lines for
  // glyphs drawn by a different face; that span keeps its own Canvas box, or
  // the painting face's measured normal line (FontResolution.routeLineMetric,
  // admitted by selectResourceMetric for the clusters that face draws).
  // See LIBRARY_ABSENT_FACE_FALLBACK_LINE_BOX (layout/line-compatibility.ts).
  // A loaded local() tuple uses the catalog only when neither parsed nor
  // prepared host measurements supply the required geometry. A registered
  // substitute is a different selected face: if its resource metric is
  // unavailable, use its Canvas line box in both axes instead of importing the
  // authored font's vertical geometry.
  // A native resolution of an installed family is present: the installed
  // family paints (CSS font matching), so its catalog geometry may apply.
  return (selected?.source === 'native' && selected.installedFamily === true)
    || mayUseExactLocalReferenceWidthMetric(selected);
}

/** A prepared host face that failed coverage must fall through to the span's
 * native measurement, never re-enter through its authored catalog profile. */
export function allowsReferenceFontMetric(
  index: MetricTupleIndex, selected: FontResolution | undefined,
): boolean {
  return !selected || !(index.get(resolutionMetricKey(selected)) ?? []).some(metric => metric.requiresGlyphCoverage);
}
