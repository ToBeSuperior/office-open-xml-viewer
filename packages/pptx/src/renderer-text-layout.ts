// PPTX text layout: layout segment/line model, CSS font-stack and canvas font
// building, run measurement, input-run construction, and paragraph line
// breaking. Split from renderer.ts (#1857).

import { ownValue } from '@silurus/ooxml-core/internal/own-key';
import { quoteCssFontFamily } from '@silurus/ooxml-core/internal/css-font-family';
import {
  arabicFallbackFamilies,
  containsArabicFallbackText,
  officeFaceWebAlias,
  quoteCssFontFamilies as quoteAll,
} from '@silurus/ooxml-core/internal/font-stack';
import { FONT_TRACKING_SENTINEL } from '@silurus/ooxml-core/internal/font-measurement-sentinels';
import { preparedPowerPointText } from './font-display-text.js';
import type { CjkLang } from '@silurus/ooxml-core';
import { normalizeLocalFontMetricFamily } from '@silurus/ooxml-core';
import { containsHanScript } from '@silurus/ooxml-core/internal/script-preload-accumulator';
import type { Fill, TextBody, Paragraph, BlipBullet } from './types';
import { asBullet } from './types';
import {
  resolveFill as resolveFillCore,
  withInheritedPatternScope,
  EMU_PER_PT as PT_TO_EMU,
  createAuxCanvas,
  classifyCjkFont,
  classifyFontGeneric,
  googleCjkFontAlias,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
  isCjkBreakChar,
  isComplexScriptCodePoint,
  graphemeClusterOffsets,
  symbolFontToUnicode,
  isSymbolFontFamily,
  GOOGLE_FONT_SUBSTITUTES,
  substituteEntryCoversText,
} from '@silurus/ooxml-core';
import type { MathNode } from '@silurus/ooxml-core';
import type { HyperlinkTarget } from '@silurus/ooxml-core';
import { paintDistanceAwareReflectionBlur } from './reflection-blur';
import { classifyPptxHyperlink } from './hyperlink';
import {
  powerPointFaceMetrics,
  officeFaceWeight,
  type PowerPointFaceMetrics,
} from './powerpoint-line-metrics.js';
import {
  COMPLEX_SCRIPT_DEFAULT_FACES,
  complexScriptDefaultFace,
  eastAsianDefaultFaces,
  isEastAsianSlotSymbol,
} from './east-asian-default.js';
import {
  breakDrawingMlText,
  measureDrawingMlAdvance,
  type DrawingMlInputRun,
} from '@silurus/ooxml-core/internal/drawingml-text';
import { emuToPx, hexToRgba, type RenderContext } from './renderer-context';
import { mathRenders, failedMath } from './renderer-math';

export type LayoutSegment = {
  text: string;
  font: string;
  /** Inline DrawingML TAB, classified UAX#9 S during visual ordering (#916). */
  isTab?: true;
  /** PowerPoint's line metrics for this segment's face (see
   * powerPointFaceMetrics); undefined when the face is not resolvable. */
  lineMetric?: PowerPointFaceMetrics;
  /** Metrics of the run's latin face. PowerPoint sizes a line by the run's
   * latin face even when an East Asian segment draws none of its glyphs
   * (#1610 powerpoint-line-supplement-3); null when unresolved. */
  lineMetricLatin?: PowerPointFaceMetrics | null;
  /** The authored face that draws this segment's glyphs (latin, ea, cs or
   * symbol slot after theme resolution). Stacked vertical text looks its
   * cell and vertical glyphs up by face (`stacked-faces.ts`). */
  faceFamily?: string;
  /** The run's latin face: it sizes a stacked column even where an East Asian
   * segment draws none of its glyphs, as it sizes a horizontal line (#1610). */
  faceFamilyLatin?: string;
  /** The face of an a:br or endParaRPr mark (text ''), kept so the mark can be
   * re-sized to the run it follows (see `layoutParagraph`). */
  markFace?: { family: string; bold: boolean; italic: boolean };
  /** Baseline displacement of this segment under pPr@fontAlgn t / ctr / b,
   * relative to the line's `metricAscent` (powerPointFontAlgnOffset). */
  fontAlgnOffsetPx?: number;
  /** Reading-frame gap resolved against a:tabLst immediately before paint. */
  tabWidthPx?: number;
  sizePx: number;
  /**
   * Actual glyph size used for paint and width measurement. PowerPoint renders
   * any non-zero DrawingML baseline run at about 65% of its authored size while
   * retaining `sizePx` for line height and baseline displacement.
   */
  drawSizePx?: number;
  color: string;
  patternFill?: Extract<Fill, { fillType: 'pattern' }>;
  noFill?: boolean;
  underline: boolean;
  /** OOXML rPr @u value when not the default "sng": "dbl"/"dotted"/"wavy"/etc. */
  underlineStyle?: string;
  /** rgba() colour for the underline when uFill overrides the text colour. */
  underlineColor?: string;
  /** Explicit rPr > uFill; otherwise the underline follows the glyph paint. */
  underlineFill?: Fill;
  underlineLine?: import('@silurus/ooxml-core').TextOutline;
  underlineLineNoFill?: boolean;
  strikethrough: boolean;
  /** Two parallel strike lines (rPr strike="dblStrike"). */
  strikeDouble?: boolean;
  /** Extra inter-character spacing in px (already scaled). */
  letterSpacingPx?: number;
  /** Internal identity of the authored run; keeps rPr@spc within that run. */
  sourceRunId?: number;
  /** Cascaded `rPr@lang` of the authored run; selects kinsoku across seams. */
  lang?: string;
  /** Spacing boundary before this segment when one authored run changed font. */
  leadingLetterSpacingPx?: number;
  baseline?: number;
  /** Run-level glyph drop shadow (rPr > effectLst > outerShdw). */
  shadow?: import('@silurus/ooxml-core').Shadow;
  /** Run-level mirrored glyph reflection (rPr > effectLst > reflection). */
  reflection?: import('@silurus/ooxml-core').Reflection;
  /** Run-level glyph outline (rPr > a:ln). Width in EMU; renderer scales. */
  outline?: import('@silurus/ooxml-core').TextOutline;
  /**
   * Run-level text highlight (rPr > a:highlight, ECMA-376 §21.1.2.3.4) as an
   * rgba() string. Renderer fills a background box behind the glyphs.
   */
  highlight?: string;
  /**
   * Resolved hyperlink target for this segment's run (IX1). Carried so the
   * onTextRun callback can attach it to the overlay span. Does not affect glyph
   * drawing (colour + underline already handled via `color`/`underline`).
   */
  hyperlink?: HyperlinkTarget;
  /**
   * Present when this segment is an OMML equation drawn as an image instead of
   * text (`text` is then ""). Box metrics are in px at the current scale.
   */
  math?: {
    nodes: MathNode[];
    display: boolean;
    width: number;
    ascent: number;
    descent: number;
    fallbackText?: string;
    unavailable?: boolean;
  };
};

export type TextPaint = string | CanvasGradient | CanvasPattern;

/** Resolve a run once per destination canvas. Effects and decorations consume
 * this same paint; a reflection's auxiliary canvas resolves the same source
 * again inside its inherited slide coordinate scope. */
export function resolveSegmentTextPaint(
  ctx: CanvasRenderingContext2D,
  seg: LayoutSegment,
  x: number,
  baseline: number,
  scale: number,
): TextPaint {
  return seg.patternFill
    ? resolveFillCore(seg.patternFill, ctx, x, baseline, 0, 0, 0, scale * PT_TO_EMU) ?? seg.color
    : seg.color;
}

export interface LayoutLine {
  segments: LayoutSegment[];
  /** ECMA-376 §21.1.2.2.1 — this line is terminated by a MANUAL line break
   *  (`<a:br>`). In a `just` paragraph it is the end of a logical line and is
   *  left-aligned, not stretched — like the paragraph's last line (§20.1.10.59).
   *  `dist`/`thaiDist` still fill every line, including these. */
  endsWithBreak?: boolean;
  /** The a:br / endParaRPr marks that size this line with text (see
   * `followingMark`). They take part in the line metrics only: they draw
   * nothing and are not layout pieces. */
  metricMarks?: LayoutSegment[];
}

/**
 * Resolve OOXML theme font references (e.g. "+mn-ea", "+mj-lt") to CSS-safe font names.
 * Canvas will silently ignore an invalid CSS font string, keeping whatever font was set before —
 * leading to wrong text size. Map theme references to generic families as a safe fallback.
 */
export function normalizeFontFamily(family: string | null, rc: RenderContext): string {
  const isolated = (name: string): string =>
    rc.embeddedFontAliases?.get(name.trim().toLowerCase()) ?? name;
  if (!family) return isolated(rc.themeMinorFont ?? 'sans-serif');
  if (family.startsWith('+')) {
    // +mn-lt = minor Latin, +mj-lt = major Latin, +mn-ea = minor East Asian, +mj-ea = major East Asian
    if (family === '+mj-lt' || family === '+mj-ea' || family === '+mj-cs') {
      return isolated(rc.themeMajorFont ?? 'sans-serif');
    }
    // +mn-lt, +mn-ea, +mn-cs, or any other + prefix → minor font
    return isolated(rc.themeMinorFont ?? 'sans-serif');
  }
  // OOXML typeface sometimes appends ",<generic>" hint (e.g. "Wingdings,Sans-Serif").
  // Strip it so the CSS font name resolves to the actual named font.
  const primary = family.split(',')[0].trim();
  if (!primary) return rc.themeMinorFont ?? 'sans-serif';
  return isolated(primary);
}

export function hasNamedFontFamily(family: string | null | undefined): boolean {
  return !!family?.trim();
}

/** CSS generic font families — must NOT be quoted in a canvas font string. */
const CSS_GENERIC_FAMILIES = new Set([
  'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui',
]);

/** Infer a CSS generic fallback from a named font so missing fonts degrade
 *  consistently. Delegates the serif/sans/mono decision to the shared core
 *  classifier (single source of truth); maps it to the CSS generic keyword. */
function genericFallback(family: string): string {
  const g = classifyFontGeneric(family);
  return g === 'mono' ? 'monospace' : g === 'serif' ? 'serif' : 'sans-serif';
}

/**
 * True when `family` names an Arabic-script face. Its Arabic runs place the
 * script fallbacks before CJK and other script tails.
 *
 * These fallback faces (esp. Noto Naskh Arabic) also carry serif-style *Latin*
 * glyphs, so appending them to every font stack made Latin text in an
 * uninstalled Latin/CJK face fall
 * into the serif Naskh face instead of degrading to the sans-serif generic.
 * Gating on actual Arabic text lets Latin/CJK runs degrade to their generic.
 */
function isArabicScriptFace(family: string): boolean {
  // Faces we explicitly substitute to a Noto Arabic web font are Arabic-script.
  if (officeFaceWebAlias(family)?.includes('Arabic')) return true;
  const l = family.toLowerCase();
  // Common Arabic-script family-name markers (Latin transliterations + Arabic).
  return /arabic|naskh|kufi|nastaliq|amiri|scheherazade|lateef|aldhabi|urdu|farsi|العرب|[؀-ۿ]/.test(l);
}

/**
 * Build the CSS font-family LIST for a (already-normalized, non-generic) face:
 * named face + optional web alias + script Noto fallbacks +
 * inferred generic. The script fallbacks are:
 *
 * - Arabic runs: matching-class Arabic Notos enter only that run's stack.
 * - CJK faces (Noto KR/SC/TC/JP, ordered by the document's CJK language so
 *   shared Han glyphs take the right shapes — see core/fonts/scripts.ts).
 * - Non-CJK scripts (Cyrillic via Noto Sans/Serif, Thai, Devanagari, Hebrew)
 *   appended unconditionally — no Han collision, browser picks per glyph.
 *
 * Exported for unit testing the fallback ordering.
 */
export function cssFontStack(
  normalized: string,
  authoredFamily = normalized,
  fallback?: CjkLang,
  text = '',
  googleSubstitutes = false,
): string {
  const generic = genericFallback(authoredFamily);
  const arabicText = containsArabicFallbackText(text);
  // Optional published web aliases for authored Office faces enter the stack
  // only with `useGoogleFonts`; the authored family stays first.
  const alias = googleSubstitutes ? officeFaceWebAlias(authoredFamily) : undefined;
  // Script-scoped visual substitutes (core substitute-script.ts) join only a
  // run that contains their script. A CSS stack cannot split a run, but the
  // authored family stays first, so an installed authored face still wins.
  const sub = alias && substituteEntryCoversText(
    ownValue(GOOGLE_FONT_SUBSTITUTES, authoredFamily.toLowerCase()), text, 'any',
  ) ? alias : undefined;
  const subPart = sub ? `${quoteCssFontFamily(sub)}, ` : '';
  const googleAlias = googleCjkFontAlias(authoredFamily);
  const aliasPart = googleAlias ? `${quoteCssFontFamily(googleAlias)}, ` : '';
  // Arabic faces lead with script fallbacks only for Arabic runs.
  const arabicFamilies = arabicFallbackFamilies(generic === 'serif');
  const arabicPart = arabicText ? `${quoteAll(arabicFamilies)}, ` : '';
  if (isArabicScriptFace(authoredFamily) && arabicText) {
    const cjk = fallback ? cjkFallbackChain(fallback, 'sans') : [];
    return `${quoteCssFontFamily(normalized)}, ${subPart}${arabicPart}${cjk.length ? `${quoteAll(cjk)}, ` : ''}${generic}`;
  }
  const variant: 'sans' | 'serif' = generic === 'serif' ? 'serif' : 'sans';
  const authoredCjk = classifyCjkFont(authoredFamily);
  const cjk = authoredCjk ?? fallback;
  const cjkFamilies = cjk
    ? cjkFallbackChain(cjk, variant).filter((name) => name !== googleAlias)
    : [];
  const cjkPart = cjkFamilies.length > 0 ? `${quoteAll(cjkFamilies)}, ` : '';
  const nonCjk = variant === 'serif' ? NON_CJK_SERIF_FALLBACKS : NON_CJK_SANS_FALLBACKS;
  const nonCjkPart = `${quoteAll(nonCjk)}, `;
  return authoredCjk
    ? `${quoteCssFontFamily(normalized)}, ${subPart}${aliasPart}${cjkPart}${nonCjkPart}${arabicPart}${generic}`
    : `${quoteCssFontFamily(normalized)}, ${subPart}${aliasPart}${nonCjkPart}${cjkPart}${arabicPart}${generic}`;
}

/**
 * Stable string key for a {@link HyperlinkTarget} so adjacent same-format runs
 * that carry the SAME link still merge into one segment (classification mints a
 * fresh object per run, so identity comparison would over-split). `undefined`
 * (no link) and any two links with the same kind + destination compare equal.
 */
function hyperlinkKey(t: HyperlinkTarget | undefined): string {
  if (!t) return '';
  return t.kind === 'external' ? `e:${t.url}` : `i:${t.ref}`;
}

/**
 * Map weight words in a face name to the OpenType OS/2.usWeightClass name
 * classes: https://learn.microsoft.com/en-us/typography/opentype/spec/os2#usweightclass
 * This name-based substitution policy does not read catalog weights: a named
 * weight can be a separate RIBBI family whose regular face is recorded as 400
 * (e.g. Franklin Gothic Medium), which loses the design weight on substitution.
 * Names without a weight word keep the run's regular/bold selection; explicit
 * run bold takes precedence in buildFont. Never tune for a substitute's stroke.
 */
function namedFaceWeight(family: string): number | null {
  const name = family.toLowerCase();
  if (/\bthin\b/.test(name)) return 100;
  if (/\b(extra[- ]?light|ultra[- ]?light)\b/.test(name)) return 200;
  if (/\blight\b/.test(name)) return 300;
  if (/\b(black|heavy)\b/.test(name)) return 900;
  if (/\b(extra[- ]?bold|ultra[- ]?bold)\b/.test(name)) return 800;
  if (/\b(semi[- ]?bold|demi[- ]?bold)\b/.test(name)) return 600;
  if (/\bbold\b/.test(name)) return 700;
  if (/\bmedium\b/.test(name)) return 500;
  if (/\b(normal|regular)\b/.test(name)) return 400;
  return null;
}

/**
 * Paint one text segment's DrawingML reflection in a bbox-sized device-pixel
 * surface. Shape/picture reflections deliberately use the full-canvas core
 * helper for long-standing VRT stability, but doing that once per text segment
 * would allocate a slide-sized canvas for every reflected run. Text reflection
 * is new, so keep its working set proportional to the glyph bounds.
 *
 * The callback paints in absolute device coordinates after the crop offset has
 * been folded into its transform. The final mirror is therefore composited at
 * identity, exactly like the shape effect pipeline.
 */
export function applyTextRunReflection(
  liveCtx: CanvasRenderingContext2D,
  paintText: (ctx: CanvasRenderingContext2D) => void,
  bbox: { x: number; y: number; w: number; h: number },
  reflection: import('@silurus/ooxml-core').Reflection,
  effectScale: number,
  liveTransform: DOMMatrix,
  deviceW: number,
  deviceH: number,
): void {
  const blur = Math.max(0, reflection.blur * effectScale);
  const margin = Math.ceil(blur * 3) + 2;
  const cropX = Math.max(0, Math.floor(bbox.x - margin));
  const cropY = Math.max(0, Math.floor(bbox.y - margin));
  const cropRight = Math.min(deviceW, Math.ceil(bbox.x + bbox.w + margin));
  const cropBottom = Math.min(deviceH, Math.ceil(bbox.y + bbox.h + margin));
  const cropW = Math.max(1, cropRight - cropX);
  const cropH = Math.max(1, cropBottom - cropY);
  const source = createAuxCanvas(cropW, cropH);
  const sourceCtx = source?.getContext('2d') as CanvasRenderingContext2D | null;
  if (!source || !sourceCtx) return;

  sourceCtx.save();
  sourceCtx.setTransform(
    liveTransform.a,
    liveTransform.b,
    liveTransform.c,
    liveTransform.d,
    liveTransform.e - cropX,
    liveTransform.f - cropY,
  );
  // The reflection source is cropped in device space. Carry the slide pattern
  // frame into that crop so a patterned glyph keeps its original tile phase.
  withInheritedPatternScope(liveCtx, sourceCtx, () => paintText(sourceCtx), { x: cropX, y: cropY });
  sourceCtx.restore();

  const localTop = bbox.y - cropY;
  const localBottom = localTop + bbox.h;
  let aux = source;
  let auxCtx = sourceCtx;
  if (blur > 0) {
    const progressive = createAuxCanvas(cropW, cropH);
    const progressiveCtx = progressive?.getContext('2d') as CanvasRenderingContext2D | null;
    if (progressive && progressiveCtx) {
      aux = progressive;
      auxCtx = progressiveCtx;
    }
  }
  if (aux !== source) {
    paintDistanceAwareReflectionBlur(
      auxCtx,
      source,
      { x: bbox.x - cropX, y: localTop, w: bbox.w, h: bbox.h },
      blur,
      cropW,
    );
  }
  const stPos = Math.max(0, Math.min(1, reflection.stPos));
  const endPos = Math.max(0, Math.min(1, reflection.endPos));
  // Chromium's OffscreenCanvas clears the destination for a gradient-filled
  // `destination-in` pass (the equivalent HTMLCanvas path works), making
  // worker-rendered reflections disappear. This surface contains only freshly
  // rasterized text, so it is origin-clean: apply the alpha ramp directly to
  // its pixel alpha. The work is bounded to the cropped glyph box, not the
  // entire slide.
  const span = Math.max(1, localBottom - localTop);
  try {
    const pixels = auxCtx.getImageData(0, 0, cropW, cropH);
    for (let row = 0; row < cropH; row++) {
      const position = Math.max(0, Math.min(1, (localBottom - (row + 0.5)) / span));
      let alpha: number;
      if (position <= stPos) alpha = reflection.stA;
      else if (position >= endPos || endPos <= stPos) alpha = reflection.endA;
      else {
        const t = (position - stPos) / (endPos - stPos);
        alpha = reflection.stA + (reflection.endA - reflection.stA) * t;
      }
      const factor = Math.max(0, Math.min(1, alpha));
      for (let offset = row * cropW * 4 + 3; offset < (row + 1) * cropW * 4; offset += 4) {
        pixels.data[offset] = Math.round(pixels.data[offset] * factor);
      }
    }
    auxCtx.putImageData(pixels, 0, 0);
  } catch {
    // Text surfaces should remain origin-clean. If a host canvas implementation
    // cannot expose pixels, keep a uniformly faded reflection instead of
    // dropping the authored effect entirely.
    auxCtx.save();
    auxCtx.globalCompositeOperation = 'destination-in';
    auxCtx.fillStyle = `rgba(0,0,0,${Math.max(0, Math.min(1, reflection.stA))})`;
    auxCtx.fillRect(0, 0, cropW, cropH);
    auxCtx.restore();
  }

  const dist = reflection.dist * effectScale;
  const dirRad = (reflection.dir * Math.PI) / 180;
  const bottom = bbox.y + bbox.h;
  liveCtx.save();
  liveCtx.setTransform(1, 0, 0, 1, 0, 0);
  liveCtx.translate(
    bbox.x + Math.cos(dirRad) * dist,
    bottom + Math.sin(dirRad) * dist,
  );
  liveCtx.scale(reflection.sx, reflection.sy);
  liveCtx.translate(-bbox.x, -bottom);
  liveCtx.drawImage(aux as CanvasImageSource, cropX, cropY);
  liveCtx.restore();
}

export function buildFont(
  bold: boolean,
  italic: boolean,
  sizePx: number,
  family: string,
  rc: RenderContext,
  text = '',
  hasNamedFamily = true,
  extraFamilies: readonly string[] = [],
): string {
  const style  = italic ? 'italic ' : '';
  const normalized = normalizeFontFamily(family, rc);
  const { authoredFamily, officeRoute: selectedRoute } = selectOfficeFace(normalized, bold, italic, rc);
  const inferredWeight = namedFaceWeight(authoredFamily);
  const weight = bold ? 'bold ' : inferredWeight ? `${inferredWeight} ` : '';
  // ECMA-376 §19.3.1.52: master txStyles is an optional source of text style.
  // The theme-minor CSS fallback when every font slot is absent is renderer
  // policy, not evidence that an exact resource was selected. Preserve that
  // established fallback until a run or inherited style names a face.
  const officeRoute = hasNamedFamily ? selectedRoute : undefined;
  const fallback = containsHanScript(text)
    ? rc.cjkFallback ?? classifyCjkFont(rc.themeMajorFont) ?? classifyCjkFont(rc.themeMinorFont) ?? undefined
    : undefined;
  if (CSS_GENERIC_FAMILIES.has(normalized)) {
    const families = fallback ? cjkFallbackChain(fallback, normalized === 'serif' ? 'serif' : 'sans') : [];
    const latin = normalized === 'monospace' ? ['Courier New', 'Liberation Mono']
      : normalized === 'serif' ? [...NON_CJK_SERIF_FALLBACKS, 'Times New Roman', 'Liberation Serif']
      : [...NON_CJK_SANS_FALLBACKS, 'Arial', 'Helvetica', 'Liberation Sans'];
    return `${style}${weight}${sizePx}px ${families.length ? `${quoteAll([...latin, ...families])}, ` : ''}${normalized}`;
  }
  const stack = cssFontStack(
    officeRoute?.family ?? normalized, authoredFamily, fallback, text,
    rc.googleSubstitutes === true,
  );
  // Application-default faces follow the primary face in their fallback
  // order (issue #1627), ahead of the generic web fallbacks.
  const extras = extraFamilies.filter((name) => name !== normalized);
  // Match the complete escaped primary string. A callback keeps $&/$' in
  // authored family names literal instead of treating them as replace tokens.
  return `${style}${weight}${sizePx}px ${extras.length
    ? stack.replace(/^("(?:[^"\\]|\\.)*"), /u, (_match, first: string) => `${first}, ${quoteAll(extras)}, `) : stack}`;
}

/**
 * The PowerPoint line metrics of a resolved family (see
 * powerPointFaceMetrics). A document-embedded face has its own bytes, which
 * the reference catalog does not describe, so it has none, and neither does
 * a CSS generic family.
 */
export function lineMetricFor(
  family: string,
  bold: boolean,
  italic: boolean,
  rc: RenderContext,
): PowerPointFaceMetrics | undefined {
  if (CSS_GENERIC_FAMILIES.has(family)) return undefined;
  // Layout follows the face that paints, selected exactly as buildFont does.
  // Catalog geometry describes a face only when that face is present (an
  // exact local() route was loaded, or its family is installed); an absent
  // authored face paints with a browser fallback, so the run keeps the
  // ordinary line model instead of the absent face's ascent share.
  const normalized = normalizeFontFamily(family, rc);
  const { authoredFamily, embedded, officeRoute } = selectOfficeFace(normalized, bold, italic, rc);
  if (officeRoute) return powerPointFaceMetrics(embedded ? authoredFamily : family, bold, italic);
  // Deck-embedded bytes (or CSS synthesis from them) paint the embedded alias;
  // the catalog does not describe those bytes, and the alias shadows any
  // installed system family of the same authored name.
  if (embedded || !familyInstalled(authoredFamily, rc)) return undefined;
  return powerPointFaceMetrics(family, bold, italic);
}

/**
 * The face selection shared by painting (buildFont) and line metrics: the
 * authored family behind a deck-embedded alias, and the exact local() route
 * that paints this tuple. A tuple the deck embeds paints its own bytes; a
 * missing embedded style falls to the exact installed route when one loaded.
 */
function selectOfficeFace(normalized: string, bold: boolean, italic: boolean, rc: RenderContext): {
  authoredFamily: string;
  embedded: boolean;
  officeRoute: import('@silurus/ooxml-core').OfficeFontFallbackRoute | undefined;
} {
  const embeddedAuthored = rc.embeddedFontAuthoredFamilies?.get(normalized);
  const authoredFamily = embeddedAuthored ?? normalized;
  const inferredWeight = namedFaceWeight(authoredFamily);
  const embeddedTuple = `${authoredFamily.toLowerCase()}:${bold || inferredWeight === 700 ? 700 : 400}:${italic ? 'italic' : 'normal'}`;
  const officeRoute = rc.embeddedFontTuples?.has(embeddedTuple)
    ? undefined
    : rc.officeFontRoutes?.[officeRouteKey(authoredFamily, officeFaceWeight(authoredFamily, bold, italic), italic)];
  return { authoredFamily, embedded: embeddedAuthored !== undefined, officeRoute };
}

/** Whether an installed face of the family paints the authored name (see
 * AcquiredOfficeRoutes.installed): CSS font matching then paints that
 * installed family, synthesizing weight/style. */
function familyInstalled(family: string, rc: RenderContext): boolean {
  return rc.officeInstalledFamilies?.includes(normalizeLocalFontMetricFamily(family)) === true;
}

/** The core office-fallback route key for an authored family tuple. */
function officeRouteKey(family: string, weight: number, italic: boolean): string {
  const normalized = normalizeLocalFontMetricFamily(family);
  return weight === 400 && !italic ? normalized : `${normalized}:${weight}:${italic ? 'italic' : 'normal'}`;
}

/**
 * Lay out a paragraph into display lines.
 * Handles:
 *  - Explicit line breaks (TextRun type='break')
 *  - Space-based word wrap (Latin text)
 *  - Character-level wrap fallback for CJK / words wider than container
 *  - Tab stops (right-aligned and left-aligned)
 *
 * @param marLPx  Paragraph left margin in canvas px (used for tab stop position calculation)
 */
/** True when the paragraph carries a leading marker (char / autoNum / picture
 *  bullet) that occupies the hanging gutter. A picture (`blip`) bullet lives on
 *  the PPTX-widened {@link Bullet}, so narrow via {@link asBullet} rather than
 *  comparing the raw union member. */
export function paragraphHasBullet(para: Paragraph): boolean {
  return (
    para.bullet.type === 'char' ||
    para.bullet.type === 'autoNum' ||
    asBullet(para.bullet).type === 'blip'
  );
}

/** Authored picture-bullet height before Canvas scale and normAutofit shrink.
 * Shared by preload and text layout so their cache target cannot drift from the
 * marker geometry. ECMA-376 §21.1.2.4.10 absolute size wins over §21.1.2.4.9
 * percentage; with neither, §21.1.2.4.13 inherits the first run/default size. */
export function pictureBulletSizePt(
  body: TextBody,
  para: Paragraph,
  bullet: BlipBullet,
): number {
  let firstRunSizePt: number | null = null;
  for (const run of para.runs) {
    if (run.type === 'text' && run.fontSize != null) {
      firstRunSizePt = run.fontSize;
      break;
    }
  }
  const baseSizePt = firstRunSizePt ?? para.defFontSize ?? body.defaultFontSize ?? 18;
  return bullet.sizePts != null
    ? bullet.sizePts
    : baseSizePt * ((bullet.sizePct ?? 100) / 100);
}

/** First-line indent (ECMA-376 §21.1.2.2.7 `a:pPr@indent`) resolved to the px
 *  amount the FIRST line's TEXT is shifted right / narrowed by. A positive indent
 *  on a non-bullet paragraph narrows the first line and shifts its start right;
 *  a negative indent extends that line left of marL. Only a BULLETED paragraph
 *  resolves to 0: `indent` is the marker's hanging gutter, positioned separately
 *  (`bulletX = textX + raw indentPx`), not a change to the text's wrap width.
 *  Shared by the spAutoFit measurement ({@link naturalWidthExceedsBbox}), the
 *  wrap budget ({@link layoutParagraph}) and the draw-side `textXOffset`, so the
 *  three paths can never disagree. */
export function firstLineIndentPxFor(hasBullet: boolean, indentPx: number): number {
  return hasBullet ? 0 : indentPx;
}

// The core predicate covers the RTL cs axis. DrawingML also routes Indic and
// Southeast Asian shaping scripts through a:cs when that font slot is present.
const INDIC_CS_GLYPH_RE = /[\p{Script=Devanagari}\p{Script=Thai}\p{Script=Bengali}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Gujarati}\p{Script=Gurmukhi}\p{Script=Oriya}\p{Script=Sinhala}\p{Script=Khmer}\p{Script=Lao}\p{Script=Myanmar}\p{Script=Tibetan}]/u;
/** PowerPoint paints non-zero DrawingML baseline runs at about 65% of their
 * authored size. The authored size still owns line height and baseline offset;
 * this helper is only for the glyph font used by measure and paint. */
function baselineDrawSizePx(authoredSizePx: number, baseline: number | undefined): number {
  return baseline != null && baseline !== 0 ? authoredSizePx * 0.65 : authoredSizePx;
}

/**
 * Decide whether `<a:spAutoFit/>` should let text wrap based on the paragraphs'
 * natural single-line width. Returns true when at least one paragraph would
 * exceed the available text width if rendered without wrapping — that's the
 * cue PowerPoint uses to switch from "grow shape horizontally" to "wrap and
 * grow shape vertically" mode (ECMA-376 §20.1.10.5).
 *
 * The measurement is intentionally conservative: it sums all run widths in a
 * paragraph at the run's own font size, ignoring tab stops and explicit
 * `\t` runs. That matches what PowerPoint compares — "one continuous line
 * of glyphs" — and avoids over-eager wrap when a paragraph is barely wider
 * than the bbox due to font-measurement drift.
 *
 * Known limitation (issue #1006): because tab jumps are ignored here, a tab-led
 * paragraph inside a `spAutoFit` shape whose raw glyph sum fits the box will
 * take the grow-not-wrap path even if its tab-jumped extent would overflow. The
 * adjudicated fixtures use `noAutofit`; spAutoFit + tab is unadjudicated and left
 * as documented behaviour rather than guessed.
 */
export function naturalWidthExceedsBbox(
  ctx: CanvasRenderingContext2D,
  body: TextBody,
  bw: number,
  lPad: number,
  rPad: number,
  scale: number,
  rc: RenderContext,
  slideNumber?: number,
): boolean {
  const bodyDefaultFontPx = (body.defaultFontSize ?? 18) * PT_TO_EMU * scale;
  for (const para of body.paragraphs) {
    const marLPx = emuToPx(para.marL, scale);
    const marRPx = emuToPx(para.marR, scale);
    const indentPx = emuToPx(para.indent, scale);
    // The signed first-line indent is consumed only by a NON-bullet paragraph
    // (firstLineIndentPxFor) — the SAME amount the wrap and draw
    // passes use, so the measurement can't disagree with what actually renders.
    const firstLineIndent = firstLineIndentPxFor(paragraphHasBullet(para), indentPx);
    const textMaxW = bw - lPad - rPad - marLPx - marRPx - firstLineIndent;
    // Measure the input runs that layoutParagraph breaks and paint draws, so
    // each character is measured in its own font slot (ea/cs/sym faces and
    // their application defaults), not in the run's Latin face.
    const { input } = paragraphInputRuns(
      para,
      para.defFontSize != null ? para.defFontSize * PT_TO_EMU * scale : bodyDefaultFontPx,
      '', scale, body.defaultBold ?? false, body.defaultItalic ?? false, 1, slideNumber, rc,
    );
    let lineW = 0;
    let previous: LayoutSegment | undefined;
    for (const item of input) {
      if (item.type !== 'text') {
        previous = undefined;
        continue;
      }
      ctx.font = item.style.font;
      const letterSpacingPx = item.style.letterSpacingPx ?? 0;
      lineW += measureTextAdvance(ctx, item.text, letterSpacingPx);
      // A run split at a font-slot seam keeps the spacing between its parts.
      if (previous?.sourceRunId === item.style.sourceRunId) lineW += letterSpacingPx;
      previous = item.style;
      if (lineW > textMaxW) return true;
    }
  }
  return false;
}

/** Number of Unicode code points in `s` (NOT UTF-16 code units). Used only by
 *  text paths that deliberately paint per code point (WordArt / vertical text)
 *  and by the fallback for Canvas implementations without native
 *  `letterSpacing`. Native horizontal text uses {@link measureTextAdvance} so
 *  the browser's shaping-cluster boundaries remain authoritative. */
export function codePointCount(s: string): number {
  let n = 0;
  for (const _ of s) n++;
  return n;
}

const nativeLetterSpacingSupport = new WeakMap<CanvasRenderingContext2D, boolean>();

/** Detect real Canvas letter-spacing support, rather than accepting an expando
 * or an inert test-double property. The result is stable for a context and is
 * shared by measurement and paint so their fallback choice cannot diverge. */
export function hasNativeLetterSpacing(ctx: CanvasRenderingContext2D): boolean {
  const cached = nativeLetterSpacingSupport.get(ctx);
  if (cached != null) return cached;

  const spacingCtx = ctx as CanvasRenderingContext2D & { letterSpacing?: string };
  const previous = spacingCtx.letterSpacing;
  if (typeof previous !== 'string') {
    nativeLetterSpacingSupport.set(ctx, false);
    return false;
  }

  let supported = false;
  try {
    spacingCtx.letterSpacing = '0px';
    const natural = ctx.measureText(FONT_TRACKING_SENTINEL).width;
    spacingCtx.letterSpacing = '1px';
    const tracked = ctx.measureText(FONT_TRACKING_SENTINEL).width;
    supported = Number.isFinite(natural) && Number.isFinite(tracked) && tracked !== natural;
  } catch {
    supported = false;
  } finally {
    try { spacingCtx.letterSpacing = previous; } catch { /* inert implementation */ }
  }
  nativeLetterSpacingSupport.set(ctx, supported);
  return supported;
}

/** Measure the Canvas advance with the tracking derived from DrawingML
 * `rPr@spc`. Native `measureText()` uses the same shaping clusters as paint,
 * which matters for combining sequences, emoji ZWJ runs, flags, and Indic
 * conjuncts. Canvas includes one trailing spacing unit in its reported width;
 * DrawingML spacing is only between characters in the run, so remove that
 * terminal unit. Older/inert implementations use the same n-1 code-point
 * approximation as the manual paint fallback. */
export function measureTextAdvance(
  ctx: CanvasRenderingContext2D,
  text: string,
  letterSpacingPx: number,
): number {
  const spacingCtx = ctx as CanvasRenderingContext2D & { letterSpacing?: string };
  const previous = spacingCtx.letterSpacing;
  if (letterSpacingPx !== 0 && hasNativeLetterSpacing(ctx)) {
    try {
      spacingCtx.letterSpacing = `${letterSpacingPx}px`;
      const tracked = ctx.measureText(text).width;
      if (Number.isFinite(tracked)) {
        return text.length > 0 ? tracked - letterSpacingPx : tracked;
      }
    } finally {
      try { spacingCtx.letterSpacing = previous; } catch { /* inert implementation */ }
    }
  }
  const internalBoundaries = Math.max(0, codePointCount(text) - 1);
  return ctx.measureText(text).width + letterSpacingPx * internalBoundaries;
}

/**
 * PowerPoint's input phase for the shared DrawingML text phases: resolves the
 * presentation theme, run formatting, fields, symbols and equation rasters
 * into core input runs. Horizontal layout (`layoutParagraph`) and stacked
 * vertical layout (`stacked-text.ts`) break the same runs.
 */
export function paragraphInputRuns(
  para: Paragraph,
  defaultFontSizePx: number,
  defaultColor: string,
  scale: number,
  defaultBold: boolean,
  defaultItalic: boolean,
  fontScale: number,
  slideNumber: number | undefined,
  rc: RenderContext,
  /** Horizontal layout only: see {@link powerPointKinsokuAcrossRuns}. */
  quoteSeams = false,
  measureCtx?: CanvasRenderingContext2D,
): {
  input: DrawingMlInputRun<LayoutSegment>[];
  sameStyle: (a: LayoutSegment, b: LayoutSegment) => boolean;
  /** Line-metric mark of each a:br, keyed by its input index. */
  breakMarks: Map<number, LayoutSegment>;
} {
  const input: DrawingMlInputRun<LayoutSegment>[] = [];
  // Grapheme clusters are segmented once over the paragraph's text, not per
  // run, so an extender (combining mark, variation selector, ZWJ, trailing
  // jamo) that opens a run joins its base in the previous run's last segment.
  // A cluster is one glyph and can carry only one format, so the carried
  // extenders take the base run's formatting (font slot, colour, link,
  // spacing). A line break or equation (an LF in the joined text) ends a
  // cluster. The boundary list is walked with one forward pointer, so the
  // phase stays linear in the paragraph length however the runs are cut.
  const runTexts = para.runs.map((run) => {
    if (run.type !== 'text') return null;
    return preparedPowerPointText(run, slideNumber);
  });
  const runStarts: number[] = [];
  let joinedText = '';
  for (const text of runTexts) {
    runStarts.push(joinedText.length);
    joinedText += text ?? '\n';
  }
  const clusterBounds = graphemeClusterOffsets(joinedText);
  clusterBounds.push(joinedText.length);
  let boundIndex = 0;
  /** The first cluster boundary at or after `pos`; the pointer only moves forward. */
  const boundaryFrom = (pos: number): number => {
    while (boundIndex < clusterBounds.length && clusterBounds[boundIndex] < pos) boundIndex++;
    return boundIndex < clusterBounds.length ? clusterBounds[boundIndex] : joinedText.length;
  };
  // The previous text run's last emitted segment, the base of a carried cluster.
  let seam: { text: string } | null = null;
  // The line-metric mark of every a:br, keyed by its input index (see
  // `followingMark` in layoutParagraph).
  const breakMarks = new Map<number, LayoutSegment>();
  // Runtime recovery is library policy; Part 3 §§9.3–9.4 defines a complete
  // MCE alternative, never a per-equation preview. Prefix counts make nested
  // branch failure checks linear. The outermost failed range wins (the parser
  // emits inner alternatives first, including equal ranges), suppressing
  // even successfully converted Choice equations, without rewriting the model
  // or shifting the source run IDs of surrounding text. Both horizontal and
  // stacked layout consume this same projection.
  let previews: Map<number, { end: number; text: string }> | undefined;
  if (para.mathFallbacks?.length) {
    previews = new Map();
    const failures = [0];
    for (const run of para.runs) {
      failures.push(failures[failures.length - 1]
        + (run.type === 'math' && failedMath.has(run.nodes) ? 1 : 0));
    }
    for (const fallback of para.mathFallbacks) {
      const { start, end } = fallback;
      if (Number.isInteger(start) && Number.isInteger(end) && start >= 0
        && end <= para.runs.length && end > start && para.runs[start].type === 'math'
        && failures[end] > failures[start] && end >= (previews.get(start)?.end ?? start)) {
        previews.set(start, fallback);
      }
    }
  }
  let previewEnd = 0;
  for (const [sourceRunId, run] of para.runs.entries()) {
    if (sourceRunId < previewEnd) continue;
    const preview = previews?.get(sourceRunId);
    if (preview) previewEnd = preview.end;
    if (run.type !== 'text') seam = null;
    if (run.type === 'break') {
      const sizePx = run.fontSize != null
        ? run.fontSize * PT_TO_EMU * scale * fontScale : defaultFontSizePx;
      const family = normalizeFontFamily(run.fontFamily ?? para.defFontFamily ?? null, rc);
      const bold = run.bold ?? para.defBold ?? defaultBold;
      const italic = run.italic ?? para.defItalic ?? defaultItalic;
      const style: LayoutSegment | undefined = run.fontSize != null || run.fontFamily != null
        || run.bold != null || run.italic != null
        ? {
            text: '', sizePx, color: defaultColor,
            font: buildFont(bold, italic, sizePx, family, rc, ''),
            underline: false, strikethrough: false,
            lineMetric: lineMetricFor(family, bold, italic, rc),
          }
        : undefined;
      breakMarks.set(input.length, {
        text: '', sizePx, color: defaultColor,
        font: buildFont(bold, italic, sizePx, family, rc, ''),
        underline: false, strikethrough: false,
        lineMetric: lineMetricFor(family, bold, italic, rc),
        markFace: { family, bold, italic },
      });
      input.push({ type: 'break', style });
      continue;
    }
    if (run.type === 'math') {
      const render = preview ? undefined : mathRenders.get(run.nodes);
      const emPx = run.fontSize != null
        ? run.fontSize * PT_TO_EMU * scale * fontScale : defaultFontSizePx;
      const unavailable = !!preview || (!render && failedMath.has(run.nodes));
      const fallbackText = preview?.text ?? (unavailable ? run.fallbackText : undefined);
      if (measureCtx) measureCtx.font = `${emPx}px sans-serif`;
      const width = render ? render.widthEm * emPx : unavailable
        ? measureCtx?.measureText(fallbackText ?? 'Equation unavailable').width ?? 0 : 0;
      const style: LayoutSegment = {
        text: '', font: `${emPx}px sans-serif`, sizePx: emPx,
        color: run.color ? hexToRgba(run.color) : defaultColor,
        underline: false, strikethrough: false,
        math: {
          nodes: run.nodes, display: run.display, width,
          ascent: render ? render.ascentEm * emPx : unavailable ? emPx : 0,
          descent: render ? render.descentEm * emPx : 0,
          fallbackText, unavailable,
        },
      };
      input.push({ type: 'object', width, style, display: run.display });
      continue;
    }

    const sizePx = run.fontSize != null
      ? run.fontSize * PT_TO_EMU * scale * fontScale : defaultFontSizePx;
    const drawSizePx = baselineDrawSizePx(sizePx, run.baseline ?? undefined);
    const family = normalizeFontFamily(run.fontFamily ?? para.defFontFamily ?? null, rc);
    const familySym = run.fontFamilySym ? normalizeFontFamily(run.fontFamilySym, rc) : null;
    const bold = run.bold ?? para.defBold ?? defaultBold;
    const italic = run.italic ?? para.defItalic ?? defaultItalic;
    let rawText = runTexts[sourceRunId] ?? '';
    // The parser resolves the ea/cs faces, theme script fonts included. An
    // empty slot takes PowerPoint's application default, never the Latin
    // face (issue #1627). The default tier is chosen from the run's whole
    // East Asian text, before any leading cluster joins the previous seam.
    const eaDefaults = run.fontFamilyEa ? [] : eastAsianDefaultFaces(
      run.fontFamily ?? para.defFontFamily ?? rc.themeMinorFont ?? null,
      [...rawText].filter((ch) => isCjkBreakChar(ch.codePointAt(0) ?? 0)).join(''),
    );
    const familyEa = run.fontFamilyEa ? normalizeFontFamily(run.fontFamilyEa, rc) : eaDefaults[0];
    const familyCs = run.fontFamilyCs ? normalizeFontFamily(run.fontFamilyCs, rc) : null;
    // Offset of rawText's first code unit in the joined paragraph text.
    let runOffset = runStarts[sourceRunId];
    const runEnd = runOffset + rawText.length;
    if (seam && rawText) {
      const firstEnd = Math.min(boundaryFrom(runOffset), runEnd);
      if (firstEnd > runOffset) {
        seam.text += rawText.slice(0, firstEnd - runOffset);
        rawText = rawText.slice(firstEnd - runOffset);
        runOffset = firstEnd;
      }
    }
    const eaFont = buildFont(bold, italic, drawSizePx, familyEa, rc, rawText, true,
      eaDefaults.slice(1));
    // Font stacks whose visual substitute is script-scoped must be resolved
    // from the grapheme cluster they paint, not the whole source run. A mixed
    // Arabic/Latin DrawingML run is still one authored run, but Noto Arabic is
    // only a substitute for its Arabic glyphs; letting the Arabic half enable
    // that face for the whole run also paints and measures the Latin half in
    // Noto Arabic. The cluster loop below is already the font-slot boundary, so
    // caching by the resulting stack preserves clusters without inventing a
    // second text segmentation policy.
    const stackFonts = new Map<string, string>();
    const stackFontFor = (face: string, text: string, named: boolean, extras: readonly string[] = []) => {
      // `buildFont` depends on text only through these script predicates. Keep
      // cache cardinality bounded by face/script class, not grapheme count.
      const key = `${face}\0${named ? 1 : 0}\0${containsHanScript(text) ? 1 : 0}`
        + `\0${containsArabicFallbackText(text) ? 1 : 0}\0${extras.join('\0')}`;
      const cached = stackFonts.get(key);
      if (cached !== undefined) return cached;
      const built = buildFont(bold, italic, drawSizePx, face, rc, text, named, extras);
      stackFonts.set(key, built);
      return built;
    };
    // An empty cs slot draws each complex-script character in its script's
    // application default; the defaults also follow an authored cs face.
    const csFontFor = (face: string, text: string) =>
      stackFontFor(face, text, true, COMPLEX_SCRIPT_DEFAULT_FACES);
    const letterSpacingPx = (run.letterSpacing ?? 0) * PT_TO_EMU * scale;
    const color = run.color ? hexToRgba(run.color)
      : run.hyperlink && rc.themeHlinkColor ? hexToRgba(rc.themeHlinkColor) : defaultColor;
    const baseStyle: LayoutSegment = {
      // `emitGroup` installs the cluster-scoped resolved font.
      text: '', font: '', sizePx, drawSizePx, color,
      // PowerPoint's default hyperlink theme colour masks pattFill. Reapplying
      // the text fill writes hlinkClr="tx" and restores the authored pattern.
      patternFill: run.hyperlink && !run.hyperlinkUsesTextFill ? undefined : run.patternFill,
      noFill: run.noFill,
      underline: run.underline || run.hyperlink !== undefined,
      underlineStyle: run.underlineStyle,
      underlineColor: run.underlineColor ? hexToRgba(run.underlineColor) : undefined,
      underlineFill: run.underlineFill,
      underlineLine: run.underlineLine,
      underlineLineNoFill: run.underlineLineNoFill,
      strikethrough: run.strikethrough,
      strikeDouble: run.strikeDouble === true,
      letterSpacingPx: letterSpacingPx || undefined,
      sourceRunId,
      lang: run.lang,
      baseline: run.baseline ?? undefined,
      shadow: run.shadow,
      reflection: run.reflection,
      outline: run.outline,
      highlight: run.highlight ? hexToRgba(run.highlight) : undefined,
      hyperlink: classifyPptxHyperlink(run.hyperlink),
    };
    // rPr/ea and rPr/sym are presentation-only font slots. Split solely where
    // the selected font changes, leaving run seams and break policy to core.
    let group = '';
    let groupFont = '';
    let groupShare: PowerPointFaceMetrics | undefined;
    let groupFamily = family;
    const latinShare = lineMetricFor(family, bold, italic, rc) ?? null;
    const emitGroup = () => {
      if (group) {
        input.push({ type: 'text', text: group,
          style: { ...baseStyle, font: groupFont, lineMetric: groupShare,
            lineMetricLatin: latinShare, faceFamily: groupFamily, faceFamilyLatin: family } });
      }
      group = '';
    };
    // Slots are chosen per grapheme cluster from its base character, so a
    // combining mark, variation selector, ZWJ or other extender stays in its
    // base's font segment and a cluster never straddles two segments (one
    // stacked cell, one shaped horizontal glyph). An empty ea/cs slot draws
    // in PowerPoint's application default (issue #1627).
    let clusterStart = 0;
    let emitted = false;
    let lastCluster = '';
    const runQuoteSeams = quoteSeams && QUOTE_SEAM_LANG_RE.test(run.lang ?? '');
    while (clusterStart < rawText.length) {
      const clusterEnd = Math.min(boundaryFrom(runOffset + clusterStart + 1), runEnd) - runOffset;
      const cluster = rawText.slice(clusterStart, clusterEnd);
      clusterStart = clusterEnd;
      emitted = true;
      const ch = String.fromCodePoint(cluster.codePointAt(0) ?? 0);
      let glyph = cluster;
      const cp = ch.codePointAt(0) ?? 0;
      // An East Asian-slot symbol resolves its stack from its own cluster, so
      // the run's other scripts never pick its fallback face.
      const eaSymbol = !isCjkBreakChar(cp) && !!run.fontFamilyEa && isEastAsianSlotSymbol(cp, run.lang);
      const eaGlyph = eaSymbol || isCjkBreakChar(cp);
      const csGlyph = isComplexScriptCodePoint(cp) || INDIC_CS_GLYPH_RE.test(ch);
      const csFace = csGlyph ? familyCs ?? complexScriptDefaultFace(ch) : family;
      let font = eaSymbol ? stackFontFor(familyEa, cluster, true)
        : eaGlyph ? eaFont : csGlyph ? csFontFor(csFace, cluster) : stackFontFor(
          family, cluster, hasNamedFontFamily(run.fontFamily ?? para.defFontFamily),
        );
      let face = eaGlyph ? familyEa : csFace;
      let share = lineMetricFor(face, bold, italic, rc);
      if (/[\uf020-\uf0ff]/u.test(ch) && (familySym != null || isSymbolFontFamily(family))) {
        const symbolFamily = familySym ?? family;
        const mapped = symbolFontToUnicode(ch, symbolFamily);
        glyph = mapped + cluster.slice(ch.length);
        font = buildFont(bold, italic, drawSizePx,
          mapped === ch ? symbolFamily : 'sans-serif', rc, glyph);
        share = undefined;
        face = mapped === ch ? symbolFamily : 'sans-serif';
      }
      // A Latin-slot curly quote next to an ideograph is a slot seam even when
      // both slots resolve to this font, in the measured en-US/ja-JP runs only
      // (see QUOTE_SEAM_LANG_RE). The pieces keep one style, so measure and
      // paint coalesce them; only the break phase sees the seam.
      const quoteSeam = runQuoteSeams && (
        (isKinsokuQuote(cluster) && isKinsokuIdeograph(lastCluster))
        || (isKinsokuQuote(lastCluster) && isKinsokuIdeograph(cluster)));
      if (group && (font !== groupFont || share !== groupShare || quoteSeam)) emitGroup();
      group += glyph;
      groupFont = font;
      groupShare = share;
      groupFamily = face;
      lastCluster = cluster;
    }
    emitGroup();
    if (emitted) {
      const last = input[input.length - 1];
      seam = last?.type === 'text' ? last : null;
    }
  }

  const sameStyle = (a: LayoutSegment, b: LayoutSegment): boolean =>
    a.font === b.font && a.color === b.color && a.patternFill === b.patternFill
    && a.noFill === b.noFill && a.sizePx === b.sizePx
    && a.drawSizePx === b.drawSizePx && a.underline === b.underline
    && a.underlineStyle === b.underlineStyle
    && a.underlineColor === b.underlineColor
    && a.underlineFill === b.underlineFill
    && a.underlineLine === b.underlineLine
    && a.underlineLineNoFill === b.underlineLineNoFill
    && a.strikethrough === b.strikethrough && a.strikeDouble === b.strikeDouble
    && a.letterSpacingPx === b.letterSpacingPx && a.baseline === b.baseline
    && a.shadow === b.shadow && a.reflection === b.reflection
    && a.outline === b.outline && a.highlight === b.highlight
    && hyperlinkKey(a.hyperlink) === hyperlinkKey(b.hyperlink)
    && (!a.letterSpacingPx || a.sourceRunId === b.sourceRunId)
    // Every face that sizes the PowerPoint line box is part of the key: two
    // runs drawn in the same face but carrying different latin slots must
    // stay apart, or the merged segment would drop one slot's share and a
    // purely visual difference (colour) would decide the line height.
    && a.lineMetric === b.lineMetric
    && a.lineMetricLatin === b.lineMetricLatin;
  return { input, sameStyle, breakMarks };
}

/**
 * Whether kinsoku continues across an input-run seam (authored runs, or the
 * font-slot segments of one run). DrawingML specifies no kinsoku rule beyond
 * a:pPr@eaLnBrk (§21.1.2.2.7), and rPr@lang (§21.1.2.3.9) only names the run
 * language. Observed PowerPoint behavior (issue #1653, controls K00–K23):
 * Arial in every slot, eaLnBrk="1", 20 pt, boxes 24/30/40 pt wide, with a
 * curly quote next to ideographs in one run and split over two runs. Each
 * two-run case breaks like its one-run twin. Under ja-JP, an opening quote
 * never ends a line and a closing quote never starts one, even at the seam.
 * Under en-US, neither rule applies to the quotes. Under en-US our curly
 * quotes paint in the Latin slot; under ja/ko/zh with a resolved East Asian
 * face they paint in the East Asian slot (isEastAsianSlotSymbol). The
 * horizontal input adapter keeps a seam in both measured languages even when
 * the quote and the ideograph share one font, so the language alone decides
 * it. Limits: only
 * the observed curly-quote/ideograph seams where both runs are Japanese join;
 * other punctuation keeps its old boundary. Retraction never crosses another
 * seam unless that seam meets this same rule. Mixed-language seams and runs
 * without lang keep the seam boundary, as C08 observed; a ko/zh quote that
 * shares the ideographs' East Asian face has no seam and keeps in-run kinsoku. Vertical (stacked) text and XLSX shapes have no such controls and
 * are not wired.
 */
const JAPANESE_LANG_RE = /^ja(?:[-_]|$)/i;
/** The one-face quote seam is added only for the two run languages the
 * controls measured, compared case-insensitively. Any other or absent lang
 * keeps one segment, so its quotes keep the in-run kinsoku they had before. */
const QUOTE_SEAM_LANG_RE = /^(?:en-US|ja-JP)$/i;
function isKinsokuQuote(text: string): boolean {
  return text === '“' || text === '”';
}
function isKinsokuIdeograph(text: string): boolean {
  return /^\p{Script=Han}/u.test(text);
}
function powerPointKinsokuAcrossRuns(
  left: LayoutSegment, right: LayoutSegment, leftText: string, rightText: string,
): boolean {
  return JAPANESE_LANG_RE.test(left.lang ?? '') && JAPANESE_LANG_RE.test(right.lang ?? '')
    && ((isKinsokuQuote(leftText) && isKinsokuIdeograph(rightText))
      || (isKinsokuIdeograph(leftText) && isKinsokuQuote(rightText)));
}

/**
 * PowerPoint adapter for the shared DrawingML text phases. It resolves the
 * presentation theme, run formatting, fields, symbols, and equation rasters;
 * the core owns all soft break decisions. The resulting LayoutLine retains
 * PowerPoint's paint metadata for renderTextBody.
 */
export function layoutParagraph(
  ctx: CanvasRenderingContext2D,
  para: Paragraph,
  maxWidthPx: number,
  defaultFontSizePx: number,
  defaultColor: string,
  scale: number,
  marLPx: number,
  defaultBold: boolean = false,
  defaultItalic: boolean = false,
  fontScale: number = 1.0,
  slideNumber?: number,
  rc: RenderContext = { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
  firstLineIndentPx: number = 0,
): LayoutLine[] {
  const { input, sameStyle, breakMarks } = paragraphInputRuns(
    para, defaultFontSizePx, defaultColor, scale, defaultBold, defaultItalic, fontScale, slideNumber, rc, true, ctx,
  );
  const marRPx = emuToPx(para.marR, scale);
  const broken = breakDrawingMlText(input, {
    maxWidth: maxWidthPx,
    firstLineIndent: firstLineIndentPx,
    measureText(text, style) {
      ctx.font = style.font;
      return measureDrawingMlAdvance(ctx, text, style.letterSpacingPx ?? 0);
    },
    sameStyle,
    boundaryAdvance: (left, right) => left.sourceRunId === right.sourceRunId
      ? right.letterSpacingPx ?? 0 : 0,
    tabStops: (para.tabStops ?? []).map((stop) => ({
      pos: emuToPx(stop.pos, scale), algn: stop.algn,
    })),
    defaultTabSize: emuToPx(para.defTabSz ?? 914400, scale),
    tabStartPen: (lineIndex) => para.rtl
      ? marRPx : marLPx + (lineIndex === 0 ? firstLineIndentPx : 0),
    nonMonotoneMeasure: input.some((item) => item.type === 'text' && (item.style.letterSpacingPx ?? 0) < 0),
    eastAsianLineBreak: para.eaLnBrk !== false,
    kinsokuAcrossRuns: powerPointKinsokuAcrossRuns,
  });
  const end = para.endRunProperties;
  const endSizePx = end?.fontSize != null
    ? end.fontSize * PT_TO_EMU * scale * fontScale : defaultFontSizePx;
  const endFamily = normalizeFontFamily(end?.fontFamily ?? para.defFontFamily ?? null, rc);
  const endBold = end?.bold ?? para.defBold ?? defaultBold;
  const endItalic = end?.italic ?? para.defItalic ?? defaultItalic;
  const endStyle: LayoutSegment | undefined = end && (
    end.fontSize != null || end.fontFamily != null || end.bold != null || end.italic != null)
    ? {
        text: '', sizePx: endSizePx, color: defaultColor,
        font: buildFont(endBold, endItalic, endSizePx, endFamily, rc, ''),
        underline: false, strikethrough: false,
        lineMetric: lineMetricFor(endFamily, endBold, endItalic, rc),
      }
    : undefined;
  const endMark: LayoutSegment | undefined = end
    ? {
        text: '', sizePx: endSizePx, color: defaultColor,
        font: buildFont(endBold, endItalic, endSizePx, endFamily, rc, ''),
        underline: false, strikethrough: false,
        lineMetric: lineMetricFor(endFamily, endBold, endItalic, rc),
        markFace: { family: endFamily, bold: endBold, italic: endItalic },
      }
    : undefined;
  // A line-break or end-of-paragraph mark after text sizes its line with its
  // face at the size of the run it follows, never at its own size (PowerPoint
  // reference-export controls, both line models and fontAlgn base / t / b):
  // - #1636: a:br and endParaRPr authoring a face, at 16 / 40 / 80 pt after
  //   24 pt and 60 pt runs, in the run's face or another;
  // - #1663: the same marks without a face of their own (a:br without rPr,
  //   lang-only or size-only rPr / endParaRPr) contribute the face they
  //   inherit from the list style. An omitted endParaRPr contributes nothing.
  // A mark alone on its line keeps its own size (the empty-line rules above).
  const followingMark = (mark: LayoutSegment, lineSegments: readonly LayoutSegment[]): LayoutSegment => {
    const previous = lineSegments[lineSegments.length - 1];
    if (!previous || !mark.markFace || previous.sizePx === mark.sizePx) return { ...mark, text: '' };
    const { family, bold, italic } = mark.markFace;
    return {
      ...mark, text: '', sizePx: previous.sizePx,
      font: buildFont(bold, italic, previous.sizePx, family, rc, ''),
    };
  };
  return broken.map((line, lineIndex): LayoutLine => {
    const isLastLine = lineIndex === broken.length - 1;
    const content: LayoutSegment[] = [
      // endParaRPr formats only the empty insertion line after the final
      // character/break (§21.1.2.2.2); it never replaces existing run paint.
      ...(isLastLine && line.segments.length === 0 && endStyle
        ? [endStyle] : []),
      // Office L07/L08: an empty line opened by a line feed inside a run keeps
      // that run's size; a zero-width segment carries it to the line metrics.
      ...(line.segments.length === 0 && line.lineFeedRun !== undefined
      && !(isLastLine && endStyle)
      && input[line.lineFeedRun]?.type === 'text'
      ? [{ ...(input[line.lineFeedRun] as { style: LayoutSegment }).style, text: '' }]
      : line.segments.map((part, index): LayoutSegment => {
      if (part.type === 'text') {
        const previous = line.segments[index - 1];
        const leadingLetterSpacingPx = previous?.type === 'text'
          && previous.style.sourceRunId === part.style.sourceRunId
          ? part.style.letterSpacingPx : undefined;
        return { ...part.style, text: part.text, leadingLetterSpacingPx };
      }
      if (part.type === 'tab') return { ...part.style, text: '', isTab: true, tabWidthPx: part.width };
      return { ...part.style, text: '' };
    })),
    ];
    const endsInBreak = line.endBreakRun !== undefined && input[line.endBreakRun]?.type === 'break';
    const breakStyle = endsInBreak ? (input[line.endBreakRun!] as { style?: LayoutSegment }).style : undefined;
    const breakMark = endsInBreak ? breakMarks.get(line.endBreakRun!) : undefined;
    const hasText = line.segments.length > 0;
    const metricMarks = hasText
      ? [
          ...(breakMark ? [followingMark(breakMark, content)] : []),
          ...(isLastLine && endMark ? [followingMark(endMark, content)] : []),
        ]
      : [];
    return {
      segments: [...content, ...(!hasText && breakStyle ? [{ ...breakStyle, text: '' }] : [])],
      ...(metricMarks.length > 0 ? { metricMarks } : {}),
      ...(line.endsWithBreak ? { endsWithBreak: true } : {}),
    };
  });
}
