// XLSX font routing and grid metrics: per-worksheet Office font routes, CSS
// font stacks, Max Digit Width, header sizes, and the RTL x transform.
// Split from renderer.ts (#1857).

import { inheritCellFontLanguage, usesJapaneseCellFonts } from './internal/cell-font-language.js';
import type { CjkLang } from '@silurus/ooxml-core';
import { quoteCssFontFamily } from '@silurus/ooxml-core/internal/css-font-family';
import {
  arabicFallbackFamilies,
  containsArabicFallbackText,
} from '@silurus/ooxml-core/internal/font-stack';
import { containsHanScript } from '@silurus/ooxml-core/internal/script-preload-accumulator';
import { activeFontSet, isHTMLCanvas, loadedGoogleRegularAliases } from '@silurus/ooxml-core';
import type { Worksheet } from './types.js';
import {
  PT_TO_PX,
  classifyCjkFont,
  classifyFontGeneric,
  googleCjkFontAlias,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
} from '@silurus/ooxml-core';
import { officeRequestKey } from './shape-office-line.js';
import { XLSX_GOOGLE_FONTS } from './google-fonts.js';
import { GridGeometry } from './internal/grid-geometry.js';
import { MDW_FALLBACK } from './internal/grid-metrics.js';

// Calibri is the workbook default font in Excel. When absent, keep the
// authored sans class rather than silently selecting another named face.
// `useGoogleFonts` explicitly permits the legacy Calibri → Carlito webfont
// substitute; that optional route is threaded per worksheet/canvas below.
// Caladea is likewise an opt-in web alias for Cambria. No other named face
// inherits these aliases merely because its authored font is unavailable.
// Arabic fallbacks enter a run's stack only when its text contains Arabic;
// Noto Naskh Arabic also carries Latin glyphs, so putting it in every sans
// stack would change the fallback class of ordinary Latin text.
// The trailing non-CJK Noto faces (Hebrew / Thai / Devanagari, plus "Noto Sans"
// for Cyrillic) extend the same per-glyph fallback idea to the other
// non-Latin, non-CJK scripts: any such codepoint resolves to a real web font
// (loaded opt-in via useGoogleFonts) instead of an OS face or tofu. CJK is NOT
// appended here — shared Han glyphs differ in shape per language, so the
// correct Noto CJK is chosen per cell from the cell's font name; see
// fontStackFor() / cssTailFor().
const NON_CJK_SANS_TAIL = NON_CJK_SANS_FALLBACKS.map((n) => `"${n}"`).join(', ');
const NON_CJK_SERIF_TAIL = NON_CJK_SERIF_FALLBACKS.map((n) => `"${n}"`).join(', ');
export const DEFAULT_FONT_FAMILY =
  `"Calibri", Arial, ${NON_CJK_SANS_TAIL}, sans-serif`;
const GOOGLE_DEFAULT_FONT_FAMILY =
  `"Calibri", "Carlito", Arial, ${NON_CJK_SANS_TAIL}, sans-serif`;
// The authored face is prepended by fontStackFor. Only the corresponding base
// text face gets its advance-width substitute; other named faces use the
// generic chain for their font class. This is font-resource routing policy,
// not a claim that the substitutes reproduce every Excel line break.
const NAMED_SANS_TAIL =
  `Arial, Helvetica, "Liberation Sans", ${NON_CJK_SANS_TAIL}, sans-serif`;
const NAMED_SERIF_TAIL =
  `"Times New Roman", "Liberation Serif", ${NON_CJK_SERIF_TAIL}, serif`;
const CALIBRI_TAIL = `"Carlito", ${NAMED_SANS_TAIL}`;
const CAMBRIA_TAIL = `"Caladea", ${NAMED_SERIF_TAIL}`;
type OfficeRoutes = Readonly<Record<string, import('@silurus/ooxml-core').OfficeFontFallbackRoute>>;
export const officeRoutesByContext = new WeakMap<object, OfficeRoutes>();
const officeRoutesByWorksheet = new WeakMap<Worksheet, OfficeRoutes>();
type NormalFontBinding = {
  tupleKey: string | null;
  route: import('@silurus/ooxml-core').OfficeFontFallbackRoute | undefined;
  hasDeclaredFace: boolean;
  canvasMdw?: number;
};
const normalFontBindingByWorksheet = new WeakMap<Worksheet, NormalFontBinding>();
export const googleSubstitutesByContext = new WeakMap<object, boolean>();
const googleSubstitutesByWorksheet = new WeakMap<Worksheet, boolean>();
const regularAliasesByContext = new WeakMap<object, ReadonlyMap<string, string>>();
const regularAliasesByWorksheet = new WeakMap<Worksheet, ReadonlyMap<string, string>>();
export type ResolvedThemeCellFont = {
  family: string | null; weight: number; fallbackAlias?: string; syntheticBold?: true;
};
export type ThemeCellFontState = Pick<Worksheet, 'themeJapaneseMajorFont' | 'themeJapaneseMinorFont'> & {
  japanese: boolean;
  cache: Map<string, ResolvedThemeCellFont>;
};
// A live caller-owned canvas must not retain a worksheet's cell/resource graph
// after workbook destruction or cache eviction. Bind only two theme names, the
// session language snapshot, and at most eight resolved style combinations.
export const themeCellFontByContext = new WeakMap<object, ThemeCellFontState>();
const themeCellFontByWorksheet = new WeakMap<Worksheet, ThemeCellFontState>();

function bindThemeCellFontState(ctx: CanvasRenderingContext2D, worksheet: Worksheet): void {
  const japanese = usesJapaneseCellFonts(worksheet);
  let state = themeCellFontByWorksheet.get(worksheet);
  if (!state || state.japanese !== japanese
      || state.themeJapaneseMajorFont !== worksheet.themeJapaneseMajorFont
      || state.themeJapaneseMinorFont !== worksheet.themeJapaneseMinorFont) {
    state = { themeJapaneseMajorFont: worksheet.themeJapaneseMajorFont,
      themeJapaneseMinorFont: worksheet.themeJapaneseMinorFont, japanese, cache: new Map() };
    inheritCellFontLanguage(worksheet, state);
    themeCellFontByWorksheet.set(worksheet, state);
  }
  themeCellFontByContext.set(ctx, state);
}

function bindNormalFontState(
  worksheet: Worksheet, routes: OfficeRoutes | undefined,
  fontSet: FontFaceSet | null, googleSubstitutes: boolean,
  measureCtx?: CanvasRenderingContext2D,
  regularAliases: ReadonlyMap<string, string> = new Map(),
): void {
  const tupleKey = worksheet.defaultFontFamily
    ? officeRequestKey({ family: worksheet.defaultFontFamily,
      weight: worksheet.defaultFontBold ? 700 : 400,
      style: worksheet.defaultFontItalic ? 'italic' : 'normal' })
    : null;
  const route = tupleKey === null ? undefined : routes?.[tupleKey];
  const hasDeclaredFace = worksheet.defaultFontFamily
    ? hasDeclaredFamilyFace(worksheet.defaultFontFamily, fontSet) : false;
  const canvasMdw = measureCtx && worksheet.defaultFontFamily && worksheet.defaultFontSize
    ? computeMdw(worksheet.defaultFontFamily, worksheet.defaultFontSize, route,
      googleSubstitutes, worksheet.defaultFontBold ? 700 : 400,
      worksheet.defaultFontItalic ? 'italic' : 'normal', measureCtx,
      regularAliases.get(worksheet.defaultFontFamily.trim().toLocaleLowerCase('en-US')))
    : undefined;
  const next = { tupleKey, route, hasDeclaredFace, canvasMdw };
  const old = normalFontBindingByWorksheet.get(worksheet);
  // Retained route maps can gain entries in place, and a popup can use another
  // FontFaceSet for the same worksheet. Either change can alter MDW while the
  // worksheet's cached grid geometry still has the previous column widths.
  if (!old || old.tupleKey !== next.tupleKey
      || old.route !== next.route || old.hasDeclaredFace !== next.hasDeclaredFace
      || old.canvasMdw !== next.canvasMdw) GridGeometry.invalidate(worksheet);
  normalFontBindingByWorksheet.set(worksheet, next);
}

/** The renderer's exact-resource route is scoped to a canvas/worksheet, not a
 * process-wide family alias: popups and workers have distinct FontFaceSets. */
export function bindXlsxOfficeFontRoutes(
  ctx: CanvasRenderingContext2D,
  worksheet: Worksheet,
  routes?: OfficeRoutes,
  googleSubstitutes = false,
): void {
  bindThemeCellFontState(ctx, worksheet);
  if (officeRoutesByWorksheet.get(worksheet) !== routes ||
      googleSubstitutesByWorksheet.get(worksheet) !== googleSubstitutes) GridGeometry.invalidate(worksheet);
  googleSubstitutesByContext.set(ctx, googleSubstitutes);
  const fontSet = (isHTMLCanvas(ctx.canvas) ? ctx.canvas.ownerDocument?.fonts : undefined)
    ?? activeFontSet();
  const regularAliases = googleSubstitutes
    ? loadedGoogleRegularAliases(fontSet, XLSX_GOOGLE_FONTS) : new Map<string, string>();
  regularAliasesByContext.set(ctx, regularAliases);
  regularAliasesByWorksheet.set(worksheet, regularAliases);
  bindNormalFontState(worksheet, routes, fontSet, googleSubstitutes, ctx, regularAliases);
  googleSubstitutesByWorksheet.set(worksheet, googleSubstitutes);
  if (routes) {
    officeRoutesByContext.set(ctx, routes);
    officeRoutesByWorksheet.set(worksheet, routes);
  } else {
    officeRoutesByContext.delete(ctx);
    officeRoutesByWorksheet.delete(worksheet);
  }
}

export function contextRegularAlias(ctx: CanvasRenderingContext2D, name: string | null | undefined): string | undefined {
  return name ? regularAliasesByContext.get(ctx)?.get(name.trim().toLocaleLowerCase('en-US')) : undefined;
}

export function bindXlsxWorksheetOfficeFontRoutes(
  worksheet: Worksheet, routes?: OfficeRoutes, googleSubstitutes = false,
): void {
  if (officeRoutesByWorksheet.get(worksheet) !== routes ||
      googleSubstitutesByWorksheet.get(worksheet) !== googleSubstitutes) GridGeometry.invalidate(worksheet);
  if (routes) officeRoutesByWorksheet.set(worksheet, routes);
  else officeRoutesByWorksheet.delete(worksheet);
  googleSubstitutesByWorksheet.set(worksheet, googleSubstitutes);
  const fontSet = activeFontSet();
  const regularAliases = googleSubstitutes
    ? loadedGoogleRegularAliases(fontSet, XLSX_GOOGLE_FONTS) : new Map<string, string>();
  regularAliasesByWorksheet.set(worksheet, regularAliases);
  bindNormalFontState(worksheet, routes, fontSet, googleSubstitutes, undefined, regularAliases);
}

export function officeRoute(
  ctx: CanvasRenderingContext2D,
  name: string | null | undefined,
  bold = false,
  italic = false,
): import('@silurus/ooxml-core').OfficeFontFallbackRoute | undefined {
  // A positively loaded exact local route must serve both the Normal-font
  // digit measurement and matching cell paint. Restricting this lookup to
  // Calibri let a non-Calibri Normal route narrow columns while its cells
  // still painted with a wider CSS fallback.
  const key = officeRequestKey({ family: name?.trim() || 'Calibri',
    weight: bold ? 700 : 400, style: italic ? 'italic' : 'normal' });
  return officeRoutesByContext.get(ctx)?.[key];
}

/** Keep a viewer-owned column metric after local font binding. The main realm
 * owns hit testing; a worker may have a different FontFaceSet and must not
 * silently replace the scalar while rendering the same worksheet. */
export function pinXlsxGridGeometry(worksheet: Worksheet, mdw?: number): void {
  if (mdw === undefined) return;
  if (!Number.isFinite(mdw) || mdw <= 0) {
    throw new Error('XLSX maximum digit width must be a finite positive number');
  }
  GridGeometry.forWorksheet(worksheet, mdw);
}
// Monospace counterpart: a monospaced cell font the host lacks degrades to a
// monospace generic rather than the proportional sans default.
const DEFAULT_MONO_FONT_FAMILY = `"Courier New", "Liberation Mono", monospace`;

/**
 * CSS font-family TAIL (everything after the cell's named face) for an xlsx
 * cell. An unnamed cell keeps the workbook default chain. Named Calibri and
 * Cambria get their corresponding web alias only with `useGoogleFonts`;
 * other named faces use a fallback of the same generic class. For a CJK face the matching Noto CJK
 * leads so shared Han glyphs take the document language's shapes; see
 * core/fonts/scripts.ts. Exported for unit testing.
 */
export function cssTailFor(name: string | null | undefined, fallback?: CjkLang, googleSubstitutes = false, text = ''): string {
  const authored = name?.trim() || null;
  const cjk = authored ? classifyCjkFont(authored) : null;
  const generic = classifyFontGeneric(authored); // 'serif' | 'sans' | 'mono'
  const appendScriptFallbacks = (base: string, includeRegional = true): string => {
    const families = [
      ...(includeRegional && fallback
        ? cjkFallbackChain(fallback, generic === 'serif' ? 'serif' : 'sans') : []),
      ...(containsArabicFallbackText(text) ? arabicFallbackFamilies(generic === 'serif') : []),
    ];
    if (families.length === 0) return base;
    const split = base.lastIndexOf(',');
    return `${base.slice(0, split)}, ${families.map((family) => `"${family}"`).join(', ')}${base.slice(split)}`;
  };
  if (!cjk) {
    // An explicit opt-in can use the legacy webfont substitute. Otherwise a
    // missing authored Calibri face falls through to the same sans class.
    const base = !authored ? (googleSubstitutes ? GOOGLE_DEFAULT_FONT_FAMILY : DEFAULT_FONT_FAMILY)
      : authored.toLowerCase() === 'calibri' ? (googleSubstitutes ? CALIBRI_TAIL : NAMED_SANS_TAIL)
      : authored.toLowerCase() === 'cambria' ? (googleSubstitutes ? CAMBRIA_TAIL : NAMED_SERIF_TAIL)
      : generic === 'serif' ? NAMED_SERIF_TAIL
      : generic === 'mono' ? DEFAULT_MONO_FONT_FAMILY : NAMED_SANS_TAIL;
    return appendScriptFallbacks(base);
  }
  const serif = generic === 'serif';
  const googleAlias = googleCjkFontAlias(authored);
  const cjkFamilies = [
    ...(googleAlias ? [googleAlias] : []),
    ...cjkFallbackChain(cjk, serif ? 'serif' : 'sans')
      .filter((family) => family !== googleAlias),
  ];
  const cjkPart = cjkFamilies
    .map((n) => `"${n}"`)
    .join(', ');
  const cjkPrefix = cjkPart ? `${cjkPart}, ` : '';
  // CJK Noto leads for the authored region; if that face lacks a Latin glyph,
  // use the matching generic Latin class without selecting an unrelated Office
  // substitute. The named CJK face still leads the complete stack.
  return appendScriptFallbacks(`${cjkPrefix}${serif ? NAMED_SERIF_TAIL : NAMED_SANS_TAIL}`, false);
}

/** Full CSS font-family list for a cell font name (named face first). The
 * consumer-selected region is relevant only to Han glyphs; authored CJK font
 * names remain authoritative for every string. */
export function fontStackFor(
  name: string | null | undefined,
  cjkFallback?: CjkLang,
  text = '',
  route?: import('@silurus/ooxml-core').OfficeFontFallbackRoute,
  googleSubstitutes = false,
  fallbackAlias?: string,
  regularAlias?: string,
): string {
  const normalized = name?.trim();
  const fallback = containsHanScript(text) ? cjkFallback : undefined;
  const selected = route && (normalized?.toLocaleLowerCase('en-US') || 'calibri')
    === route.requestedFamily.trim().toLocaleLowerCase('en-US') ? route : undefined;
  const aliasPart = fallbackAlias && fallbackAlias.toLocaleLowerCase('en-US') !== normalized?.toLocaleLowerCase('en-US')
    ? `${quoteCssFontFamily(fallbackAlias)}, ` : '';
  const regularPart = regularAlias && regularAlias.toLocaleLowerCase('en-US') !== normalized?.toLocaleLowerCase('en-US')
    ? `${quoteCssFontFamily(regularAlias)}, ` : '';
  return selected
    ? `${quoteCssFontFamily(selected.family)}, ${aliasPart}${regularPart}${cssTailFor(normalized, fallback, googleSubstitutes, text)}`
    : normalized ? `${quoteCssFontFamily(normalized)}, ${aliasPart}${regularPart}${cssTailFor(normalized, fallback, googleSubstitutes, text)}` : cssTailFor(null, fallback, googleSubstitutes, text);
}

export const DEFAULT_FONT_SIZE = 11;
// Fallback Max Digit Width when the workbook's default font isn't known.
// Calibri 11 pt at 96 DPI ≈ 8 px (Canvas2D measurement), matching the EMU
// offsets Excel 365 writes into <xdr:twoCellAnchor>. ECMA-376 §18.3.1.13
// defines MDW as the maximum rendered width among the digits 0-9 of the
// workbook's default font; measured in Excel, that is `<fonts>[0]`, not the
// Normal cell style's font (which only sets the automatic row-height
// baseline). The value depends on that font and size (e.g. Meiryo UI 10 pt
// yields MDW ≈ 6 px).
export const HEADER_W = 50;
export const HEADER_H = 22;

/**
 * Sheet-level right-to-left horizontal mirror (ECMA-376 §18.3.1.87
 * `<sheetView rightToLeft>`). The grid is always laid out left-to-right
 * internally; an RTL sheet is produced by mirroring every horizontal extent
 * about the canvas width. A left-anchored rect `[x, x + w]` maps to
 * `[canvasW - x - w, canvasW - x]`.
 *
 * Because the LTR layout puts the row-header strip at `[0, HEADER_W]` and the
 * cell area at `[HEADER_W, canvasW]`, mirroring moves the header strip to the
 * right edge and the cell area to `[0, canvasW - HEADER_W]`, matching Excel.
 *
 * This is the SINGLE source of truth for the RTL x transform. The Canvas
 * renderer, the selection overlay, and pointer hit-testing must all use it so
 * that a cell drawn at screen-x is the same cell a click at screen-x resolves
 * to, at every scroll offset. `canvasW` is the CSS-pixel width of the drawing
 * surface (`canvasArea.clientWidth`), identical to `ctx.canvas.width / dpr`.
 *
 * The transform is an involution: applying it to a screen point recovers the
 * logical-LTR point (`rtlMirrorX(rtlMirrorX(x, w, W), w, W) === x`), so the
 * same function serves both cell→px (draw) and px→cell (hit-test, with w = 0
 * for a point).
 */
export function rtlMirrorX(x: number, w: number, canvasW: number): number {
  return canvasW - x - w;
}

/** Map a logical-LTR anchored-object rectangle into sheet display space.
 * Sheet RTL mirrors placement only; the object contents and width stay intact. */
export function sheetAnchoredRectX(
  x: number,
  width: number,
  canvasWidth: number,
  rtl: boolean,
): number {
  return rtl ? rtlMirrorX(x, width, canvasWidth) : x;
}

/** Measure the Max Digit Width (ECMA-376 §18.3.1.13) for an arbitrary font
 *  using Canvas2D. The maximum of `measureText('0'..'9').width` is taken.
 *  Round to integer CSS pixels, using Windows Excel as the layout reference
 *  independently of the viewer OS.
 *
 *  A bound worksheet caches the measured scalar for its current owner; each
 *  rebind refreshes it after font preflight, without retaining that Canvas.
 *  Direct calls measure afresh in the current realm. */
export function computeMdw(
  family: string, sizePt: number,
  route?: import('@silurus/ooxml-core').OfficeFontFallbackRoute,
  googleSubstitutes = false,
  weight: 400 | 700 = 400,
  style: 'normal' | 'italic' = 'normal',
  ownerCtx?: CanvasRenderingContext2D,
  regularAlias?: string,
): number {
  const sizePx = sizePt * PT_TO_PX;
  // A viewer-owned context resolves CSS faces in its own Document (which may
  // be a popup). Direct calls without an owner use an off-DOM canvas.
  const canvas = ownerCtx ? null : (typeof OffscreenCanvas !== 'undefined')
    ? new OffscreenCanvas(1, 1)
    : (typeof document !== 'undefined' ? document.createElement('canvas') : null);
  const ctx = ownerCtx ?? canvas?.getContext('2d');
  if (!ctx) return MDW_FALLBACK;
  // Quote the family so multi-word names like "Meiryo UI" parse as one face.
  const stylePrefix = weight !== 400 || style !== 'normal' ? `${style} ${weight} ` : '';
  if (ownerCtx) ctx.save();
  try {
    ctx.font = `${stylePrefix}${sizePx}px ${fontStackFor(family, undefined, '', route, googleSubstitutes, undefined, regularAlias)}`;
    let mdw = 0;
    for (const d of '0123456789') {
      const w = ctx.measureText(d).width;
      if (w > mdw) mdw = w;
    }
    return quantizeMdw(mdw);
  } finally {
    // The caller may continue measuring or painting with this same context.
    if (ownerCtx) ctx.restore();
  }
}

function quantizeMdw(widthPx: number): number {
  return Math.round(widthPx) || MDW_FALLBACK;
}

function hasDeclaredFamilyFace(family: string, fontSet: FontFaceSet | null): boolean {
  if (!fontSet || typeof fontSet[Symbol.iterator] !== 'function') return false;
  const requested = family.trim().toLocaleLowerCase('en-US');
  for (const face of fontSet) {
    if (face.family.trim().replace(/^(['"])(.*)\1$/u, '$2').toLocaleLowerCase('en-US') === requested) return true;
  }
  return false;
}

/** Resolve the Max Digit Width from the workbook's default font face
 *  (`<fonts>[0]`, as Excel sizes columns). ECMA-376 §18.3.1.13 defines the
 *  source digit metric; when that face is unavailable, its catalog hmtx digit
 *  width cannot be used alongside a wider Canvas fallback without clipping
 *  cell text. The retained
 *  exact local/application face still wins when available. If the parser did
 *  not identify a default font, use the conventional 8 px fallback. */
export function getMdwForWorksheet(ws: Pick<Worksheet,
  'defaultFontFamily' | 'defaultFontSize' | 'defaultFontBold' | 'defaultFontItalic'>): number {
  if (!ws.defaultFontFamily || !ws.defaultFontSize) return MDW_FALLBACK;
  const weight = ws.defaultFontBold ? 700 : 400;
  const style = ws.defaultFontItalic ? 'italic' : 'normal';
  const tupleKey = officeRequestKey({ family: ws.defaultFontFamily, weight, style });
  const route = officeRoutesByWorksheet.get(ws as Worksheet)?.[tupleKey];
  const boundFont = normalFontBindingByWorksheet.get(ws as Worksheet);
  return boundFont?.canvasMdw ?? computeMdw(
    ws.defaultFontFamily, ws.defaultFontSize,
    route,
    googleSubstitutesByWorksheet.get(ws as Worksheet) === true,
    weight,
    style,
    undefined,
    regularAliasesByWorksheet.get(ws as Worksheet)?.get(ws.defaultFontFamily.trim().toLocaleLowerCase('en-US')),
  );
}

/** Worksheet-lifetime geometry snapshot. Font loading completes before a
 * worksheet reaches paint; explicit model mutations invalidate this snapshot. */
export function getGridGeometryForWorksheet(ws: Worksheet): GridGeometry {
  return GridGeometry.forWorksheetMeasured(ws, () => getMdwForWorksheet(ws));
}
