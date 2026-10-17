// XLSX cell text: cell style resolution, canvas fonts, plain and rich text
// layout, wrapping, decorations, and phonetic bands. Split from renderer.ts
// (#1857).

import {
  wrapSpreadsheetCellParagraph,
  layoutSpreadsheetCellRichLines,
  type SpreadsheetCellRichSeg,
  type SpreadsheetCellRichLine,
} from './spreadsheet-cell.js';
import { usesJapaneseCellFonts } from './internal/cell-font-language.js';
import type { CjkLang } from '@silurus/ooxml-core';
import { findReferenceFontMetrics } from '@silurus/ooxml-core';
import type {
  Worksheet,
  Styles,
  Cell,
  CellValue,
  CellFont,
  CellFill,
  Border,
  CellXf,
  Run,
  PhoneticRun,
  PhoneticProperties,
  PhoneticAlignment,
} from './types.js';
import { placePhoneticRuns } from './phonetic.js';
import { crispOffset, PT_TO_PX, hexToRgba } from '@silurus/ooxml-core';
import { computeLineVisualOrder, resolveCellBidi } from './bidi-line.js';
import {
  themeCellFontByContext,
  fontStackFor,
  officeRoute,
  googleSubstitutesByContext,
  contextRegularAlias,
  DEFAULT_FONT_SIZE,
  type ResolvedThemeCellFont,
  type ThemeCellFontState,
} from './renderer-font-metrics.js';

/** Vertical pixel metric for cell text: point size → device px at the current
 *  cell scale `cs`. `factor` is the line-height / char-height multiplier (1.2
 *  for wrapped lines, 1.1 for stacked chars, 1.0 for decoration/baseline
 *  offsets). Centralizes the `* cs` factor so a new vertical-metric draw site
 *  can't silently omit it. (Glyph SIZE uses buildFont's floored variant.)
 *
 *  An authored family name cannot establish the selected face's font tables.
 *  Keep cell sizing independent of the former family-keyed correction table
 *  until Excel resource identity and line allocation are established. */
export function vMetricPx(sizePt: number, cs: number, factor = 1): number {
  return Math.round(sizePt * PT_TO_PX * factor * cs);
}

/** Excel for Mac 16.113.2 and Windows Excel 16.0.20430 (Japanese UI,
 * with Japanese editing enabled on Windows) select the theme's Jpan face for
 * the whole scheme-marked cell, including Latin glyphs. Controlled Excel PDFs
 * covered minor/major, Calibri/Arial authored names, Latin/Japanese/mixed text,
 * regular/bold/italic, Gothic/Mincho themes, absent Jpan, and absent scheme.
 * Missing scheme or Jpan retained the authored name; major Light/Demibold
 * retained their named weight even with bold. Italic retained the theme family.
 *
 * SpreadsheetML name/scheme (§18.8.29/§18.8.35) and DrawingML's supplemental
 * font (CT_SupplementalFont) encode the choices, not Excel's UI/editing language.
 * This is observed Office behavior, not a normative rule for every locale.
 * Snapshot navigator.language starting with 'ja' when opening the workbook as
 * the declared proxy on every OS. Layout and paint keep that choice until the
 * workbook is re-opened, including sheets loaded later and worker projections.
 * Non-Japanese Windows UI remains unverified. Keeping the authored name for
 * non-ja viewers is library policy: default Office themes commonly include
 * Jpan. The proxy cannot recover a producer's different UI/editing language
 * from the workbook. */
export function resolvedThemeCellFont(
  font: Pick<CellFont, 'name' | 'scheme' | 'bold' | 'italic'>,
  worksheet?: Pick<Worksheet, 'themeJapaneseMajorFont' | 'themeJapaneseMinorFont'>,
): ResolvedThemeCellFont {
  const fallback = { family: font.name, weight: font.bold ? 700 : 400 };
  if (!usesJapaneseCellFonts(worksheet)) return fallback;
  const themeFace = themeFaceForCellFont(font, worksheet);
  if (!themeFace) return fallback;

  // The pinned Office face catalog retains full/PostScript aliases. Keep the
  // authored theme name first: an installed face must win. A catalog alias
  // may follow for hosts that expose the same family under its PostScript
  // base. Use this fixed catalog on every host; its source is not an OS gate.
  // This is name-table metadata, not proof of the selected font bytes.
  const profiles = findReferenceFontMetrics(themeFace, { source: 'office-mac' });
  // Observed Japanese-UI Excel rule (not specified by CT_FontScheme): use
  // the requested weight only if the theme name has that exact catalog face;
  // otherwise retain its base weight. Regular families use 400; a name that
  // identifies a single weight keeps that weight. Ambiguous names without a
  // Regular face use CSS normal (400) as library policy; unknown names retain
  // the authored bold request because the catalog supplies no face evidence.
  // Weight selection is independent of italic availability: these controls
  // have no italic faces, but italic cells still retain the selected weight.
  const weights = new Set(profiles.map((profile) => profile.weight));
  const baseWeight = weights.size === 1 ? profiles[0].weight : 400;
  const weight = !profiles.length || weights.has(fallback.weight)
    ? fallback.weight : baseWeight;
  const weightMatches = profiles.filter((profile) => profile.weight === weight);
  const styleMatches = weightMatches.filter((profile) => profile.style === (font.italic ? 'italic' : 'normal'));
  const matches = styleMatches.length ? styleMatches : weightMatches;
  // Office-produced controls covered Calibri/Arial, Latin/Japanese/mixed text:
  // minor Gothic bold selects its 700 face; major Gothic Light, minor Mincho
  // Regular, and major Mincho Demibold keep 300/400/600. PDF text mode 2 adds
  // fill+stroke for bold Light and minor Mincho, but not Demibold or real
  // Bold. Limit synthesis to catalog-proven Light/Regular weights (300/400);
  // unknown faces keep the browser's bold request. Other retained weights are
  // unverified. Do not request 700 to imitate the stroke: CSS may change face.
  const postScriptBases = new Set(matches.flatMap((profile) => profile.aliases)
    .filter((alias) => /-(?:Regular|Bold|Italic|BoldItalic|Light|Medium)$/i.test(alias))
    .map((alias) => alias.replace(/-(?:Regular|Bold|Italic|BoldItalic|Light|Medium)$/i, '')));
  const catalogFamilies = new Set(matches.map((profile) => profile.family));
  const alias = postScriptBases.size === 1 ? [...postScriptBases][0]
    : catalogFamilies.size === 1 ? [...catalogFamilies][0] : undefined;
  return {
    family: themeFace,
    weight,
    ...(font.bold && (weight === 300 || weight === 400) && profiles.length
      ? { syntheticBold: true as const } : {}),
    ...(alias && alias.toLocaleLowerCase('en-US') !== themeFace.toLocaleLowerCase('en-US')
      ? { fallbackAlias: alias } : {}),
  };
}

function themeFaceForCellFont(
  font: Pick<CellFont, 'scheme'>,
  worksheet?: Pick<Worksheet, 'themeJapaneseMajorFont' | 'themeJapaneseMinorFont'>,
): string | undefined {
  return font.scheme === 'major' ? worksheet?.themeJapaneseMajorFont
    : font.scheme === 'minor' ? worksheet?.themeJapaneseMinorFont : undefined;
}

function cachedThemeCellFont(font: CellFont, state?: ThemeCellFontState): ResolvedThemeCellFont {
  const themeFace = state?.japanese && themeFaceForCellFont(font, state);
  if (!state || !themeFace) return { family: font.name, weight: font.bold ? 700 : 400 };
  // Two theme faces × four requested bold/italic combinations. Neither the
  // cached values nor the state contain a reference back to the worksheet.
  const key = `${font.scheme}:${font.bold ? 1 : 0}:${font.italic ? 1 : 0}`;
  let selected = state.cache.get(key);
  if (!selected) {
    selected = resolvedThemeCellFont(font, state);
    if (state.cache.size < 8) state.cache.set(key, selected);
  }
  return selected;
}

/** Observed Japanese-UI Excel synthetic bold, separate from face selection.
 * Windows Excel 16.0.20430 PDF controls (Calibri/Arial authored names,
 * Latin/Japanese/mixed text) use 2 Tr and 0.31543 w at 11.04 pt for the retained
 * Light/Regular faces: stroke width = em / 35, independent of font name.
 * Demibold and real Bold are fill-only counterexamples. This is compatibility
 * paint, not a stroke prescribed by SpreadsheetML §18.8.2 b.
 *
 * These PDFs also have glyph-dependent TJ advance differences: Light Latin
 * deltas 0.002–0.011 em, Mincho mostly zero except b/space (~0.01 em), and CJK
 * zero. They do not establish a common tracking/advance rule. Retain Canvas
 * shaping rather than applying uniform spacing that would alter CJK and the
 * unchanged Latin advances. Exact Excel synthetic advance metrics remain
 * unsupported until a resource-based rule is established.
 */
export function paintCellText(
  ctx: CanvasRenderingContext2D, text: string, x: number, y: number,
  font: CellFont, cs = 1,
): void {
  ctx.fillText(text, x, y);
  if (!font.bold) return;
  const state = themeCellFontByContext.get(ctx);
  if (!state?.japanese || !cachedThemeCellFont(font, state).syntheticBold) return;
  ctx.save();
  ctx.strokeStyle = ctx.fillStyle;
  ctx.lineWidth = Math.max(1, Math.round(font.size * PT_TO_PX * cs)) / 35;
  // Borders/decorations may have left a dash or join on the context; PDF text
  // mode 2 uses a solid stroke with the default PDF join/miter settings.
  ctx.setLineDash([]);
  ctx.lineJoin = 'miter';
  ctx.miterLimit = 10;
  ctx.strokeText(text, x, y);
  ctx.restore();
}

export function buildFont(ctx: CanvasRenderingContext2D, font: CellFont, cs = 1, cjkFallback?: CjkLang, text = ''): string {
  const style = font.italic ? 'italic ' : '';
  const sizePx = Math.max(1, Math.round(font.size * PT_TO_PX * cs));
  const selected = cachedThemeCellFont(font, themeCellFontByContext.get(ctx));
  const weight = selected.weight === 400 ? '' : `${selected.weight} `;
  return `${style}${weight}${sizePx}px ${fontStackFor(
    selected.family, cjkFallback, text,
    officeRoute(ctx, selected.family, font.bold, font.italic),
    googleSubstitutesByContext.get(ctx) === true,
    selected.fallbackAlias,
    contextRegularAlias(ctx, selected.family),
  )}`;
}

/**
 * Draw the furigana (phonetic-hint) band across the top of a cell
 * (ECMA-376 §18.4.6 `<rPh>` / §18.4.3 `<phoneticPr>`).
 *
 * The caller must have already set the clip to the cell rect and drawn the base
 * text; this stamps the small reading glyphs ABOVE the base text, sitting in the
 * top strip Excel reserves when a phonetic row is auto-fitted (that expanded
 * height is already stored in the row's `ht`, so we never grow the row here —
 * we draw within the existing rectangle, which clips an over-tall reading the
 * same way Excel clips when a manual height is too small).
 *
 * `phoneticPr.fontId` (§18.18.32) selects the reading font from the style sheet;
 * out of bounds falls back to font 0 (§18.4.3). `alignment` (§18.18.56) and
 * `type` (§18.18.57) default to `left` / `fullwidthKatakana` when absent.
 * `baseLeftX` is the sheet-x of the base text's first glyph — the phonetic
 * band is positioned relative to the base glyphs it annotates.
 */
export function drawPhoneticBand(
  ctx: CanvasRenderingContext2D,
  runs: readonly PhoneticRun[],
  pr: PhoneticProperties | undefined,
  baseText: string,
  baseFontStr: string,
  styles: Styles,
  baseLeftX: number,
  cellTopY: number,
  cs: number,
  color: string,
  cjkFallback?: CjkLang,
): void {
  if (runs.length === 0) return;
  // §18.4.3: fontId selects the reading font; out of bounds → font 0.
  const fontId = pr?.fontId ?? 0;
  const phFont: CellFont | undefined = styles.fonts[fontId] ?? styles.fonts[0];
  if (!phFont) return;
  const alignment: PhoneticAlignment = pr?.alignment ?? 'left';

  ctx.save();
  ctx.font = buildFont(ctx, phFont, cs, cjkFallback, runs.map((run) => run.text).join(''));
  ctx.textBaseline = 'top';
  ctx.textAlign = 'left';
  ctx.fillStyle = color;

  // Sit the reading just below the cell's top padding.
  const y = cellTopY + Math.round(2 * cs);

  if (alignment === 'noControl') {
    // §18.18.56 noControl: NOT per word — lay the readings out left-to-right
    // from the base text's left edge, each at its own natural (reading-font)
    // width, in run order. Measured with the reading font already active.
    let cursor = baseLeftX;
    for (const run of runs) {
      paintCellText(ctx, run.text, cursor, y, phFont, cs);
      cursor += ctx.measureText(run.text).width;
    }
    ctx.restore();
    return;
  }

  // Per-word (left / center / distributed): position each hint over its base
  // span using BASE-font advances (so the band lines up with the glyphs it
  // annotates). `baseFontStr` is the font string the base text was drawn with.
  const placed = placePhoneticRuns(runs, baseText, baseLeftX, alignment, (s) =>
    measureInFont(ctx, s, baseFontStr),
  );

  for (const p of placed) {
    const naturalW = ctx.measureText(p.text).width;
    const cps = [...p.text];
    if (p.spread === 'distribute' && cps.length > 1 && naturalW < p.width) {
      // §18.18.56 distributed: spread the reading glyphs so the first hugs the
      // left edge and the last the right edge of the base span.
      const extra = (p.width - naturalW) / (cps.length - 1);
      try { (ctx as CanvasRenderingContext2D & { letterSpacing: string }).letterSpacing = `${extra}px`; } catch { /* ignore */ }
      paintCellText(ctx, p.text, p.x, y, phFont, cs);
      try { (ctx as CanvasRenderingContext2D & { letterSpacing: string }).letterSpacing = '0px'; } catch { /* ignore */ }
    } else if (p.spread === 'center') {
      // §18.18.56 center: centre the natural reading width over the base span.
      paintCellText(ctx, p.text, p.x + (p.width - naturalW) / 2, y, phFont, cs);
    } else {
      // §18.18.56 left: left-justified at the base span's left edge.
      paintCellText(ctx, p.text, p.x, y, phFont, cs);
    }
  }
  ctx.restore();
}

/** Measure `s` under `fontStr`, restoring the caller's current font afterward. */
export function measureInFont(ctx: CanvasRenderingContext2D, s: string, fontStr: string): number {
  const prev = ctx.font;
  ctx.font = fontStr;
  const w = ctx.measureText(s).width;
  ctx.font = prev;
  return w;
}

/**
 * Stroke a single or double horizontal text-decoration line
 * (underline / strikethrough). ECMA-376 §18.4.13 ST_UnderlineValues
 * "double" / "doubleAccounting" both render as two parallel lines with a
 * small gap. The two strokes straddle the requested baseline ±1px so the
 * pair reads as a ~3px-thick rule and the visual centre matches Excel.
 */
export function drawTextDecoLine(
  ctx: CanvasRenderingContext2D,
  x1: number, x2: number, y: number,
  color: string, double: boolean,
  dpr = 1,
): void {
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = 0.5;
  ctx.beginPath();
  // crispOffset snaps each horizontal stroke onto the device pixel grid so the
  // 0.5-px rule renders crisp at any dpr (matches docx/pptx underline/strike).
  if (double) {
    const yUp = y - 1, yDn = y + 1;
    const yu = yUp + crispOffset(yUp, 0.5, dpr);
    const yd = yDn + crispOffset(yDn, 0.5, dpr);
    ctx.moveTo(x1, yu); ctx.lineTo(x2, yu);
    ctx.moveTo(x1, yd); ctx.lineTo(x2, yd);
  } else {
    const ys = y + crispOffset(y, 0.5, dpr);
    ctx.moveTo(x1, ys); ctx.lineTo(x2, ys);
  }
  ctx.stroke();
  ctx.restore();
}

/**
 * Resolve a Run's font against a base Font. Per ECMA-376, a run's <rPr>
 * completely specifies bold/italic/underline/strike for that run, while
 * size/name fall back to the base when omitted. A run with no <rPr>
 * (run.font undefined) inherits the base entirely. An <rPr> without a
 * <color> is automatic (black) rather than the cell's color: Excel draws
 * such a run black even in a red cell.
 */
export function applyRunFont(base: CellFont, run: Run): CellFont {
  const rf = run.font;
  if (!rf) return base;
  return {
    bold: rf.bold,
    italic: rf.italic,
    underline: rf.underline,
    underlineStyle: rf.underlineStyle,
    strike: rf.strike,
    size: rf.size ?? base.size,
    color: rf.color ?? null,
    name: rf.name ?? base.name,
    vertAlign: rf.vertAlign,
  };
}

export function resolveXf(styles: Styles, styleIndex: number): { font: CellFont; fill: CellFill; border: Border; xf: CellXf } {
  const xf: CellXf = styles.cellXfs[styleIndex] ?? styles.cellXfs[0] ?? {
    fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false,
  };
  const font: CellFont = styles.fonts[xf.fontId] ?? { bold: false, italic: false, underline: false, strike: false, size: DEFAULT_FONT_SIZE, color: null, name: null };
  const fill: CellFill = styles.fills[xf.fillId] ?? { patternType: 'none', fgColor: null, bgColor: null };
  const border: Border = styles.borders[xf.borderId] ?? { left: null, right: null, top: null, bottom: null };
  return { font, fill, border, xf };
}

function columnStyleIndex(worksheet: Worksheet, col: number): number {
  const ranges = worksheet.colStyleRanges ?? [];
  for (let index = ranges.length - 1; index >= 0; index--) {
    const range = ranges[index];
    if (col >= range.min && col <= range.max) return range.styleIndex;
  }
  return 0;
}

export function effectiveCellStyleIndex(
  worksheet: Worksheet,
  cell: Cell | undefined,
  col: number,
): number {
  return cell?.styleIndex ?? columnStyleIndex(worksheet, col);
}

export function wrapTextLines(
  ctx: CanvasRenderingContext2D, text: string, maxWidth: number,
  measureWidth?: (value: string) => number,
): string[] {
  const lines: string[] = [];
  // Hard line breaks (\n from Alt+Enter) always split regardless of wrapText.
  for (const paragraph of text.split('\n')) {
    lines.push(...wrapParagraphLines(ctx, paragraph, maxWidth, measureWidth));
  }
  return lines;
}

/** SpreadsheetML `wrapText` uses the shared cell host hook. It retains cell
 * overflow and terminal-space policy; DrawingML text boxes use the distinct
 * `a:txBody` break phase below in `drawShapeText`. */
export function wrapParagraphLines(
  ctx: CanvasRenderingContext2D, paragraph: string, maxWidth: number,
  measureWidth: (value: string) => number = (value) => ctx.measureText(value).width,
): string[] {
  return wrapSpreadsheetCellParagraph(paragraph, maxWidth, measureWidth);
}

type RichSeg = SpreadsheetCellRichSeg<CellFont>;
type RichLine = SpreadsheetCellRichLine<CellFont>;

/** Font-resolving adapter for SpreadsheetML rich shared-string lines. */
export function layoutRichTextLines(
  ctx: CanvasRenderingContext2D,
  runs: Run[],
  baseFont: CellFont,
  cs: number,
  maxWidth: number,
  cjkFallback?: CjkLang,
): RichLine[] {
  return layoutSpreadsheetCellRichLines(
    ctx, runs, baseFont, maxWidth, applyRunFont,
    (font, text) => {
      ctx.font = buildFont(ctx, vertAlignDrawFont(font), cs, cjkFallback, text);
    },
    DEFAULT_FONT_SIZE,
  );
}

/** ECMA-376 §18.18.40 `general` horizontal alignment: text is left-aligned,
 *  numbers (including dates and times) right-aligned, and booleans centered.
 *  Parsers represent `general` as an absent `alignH`. */
export function generalHorizontalAlignment(type: CellValue['type']): 'left' | 'right' | 'center' {
  if (type === 'number') return 'right';
  if (type === 'bool') return 'center';
  return 'left';
}

/** Cell geometry + alignment shared by the rich-text draw helpers (wrap and
 *  non-wrap). `alignH`/`alignV` accept the raw `xf` strings; any value other
 *  than `right`/`center` anchors left, and other than `top`/`center` anchors
 *  bottom (matching the legacy single-line path's handling of `justify` /
 *  `centerContinuous` / etc.). */
export interface RichCellGeom {
  alignH: string;
  alignV: string;
  /** Cell top-left in canvas px (merge span included). */
  cx: number;
  cy: number;
  cellW: number;
  cellH: number;
  /** Left text inset (paddingX + indent) and the symmetric paddings. */
  leftPad: number;
  paddingX: number;
  paddingY: number;
}

/** Underline / strike decoration y for a run on a line, given the line's text
 *  baseline and its `textY`. With a `'top'` baseline `textY` is the line top, so
 *  the underline sits a full text height below it and the strike near mid-height;
 *  `'middle'` / `'bottom'` shift relative to the centre / bottom baseline. The
 *  caller adds any super/subscript `yShift`. One function so the single-line and
 *  multi-line paths place decorations identically.
 *
 *  The 0.55 / 0.35 / 0.5 font-size fractions (repeated by the single-run cell
 *  path in renderQuadrant) are library policy dating from the initial xlsx
 *  underline/strike support: SpreadsheetML `<u>` / `<strike>` (§18.4.13 /
 *  §18.4.10) give no geometry, no Excel measurement is recorded, and the font
 *  rule (OpenType post.underlinePosition / OS/2.yStrikeoutPosition) is not in
 *  the reference-font-metrics table. */
function decoYForBaseline(baseline: CanvasTextBaseline, textY: number, rSizePx: number): { underline: number; strike: number } {
  if (baseline === 'middle') return { underline: textY + Math.round(rSizePx * 0.55), strike: textY };
  if (baseline === 'bottom') return { underline: textY + 1, strike: textY - Math.round(rSizePx * 0.35) };
  return { underline: textY + rSizePx + 1, strike: textY + Math.round(rSizePx * 0.5) };
}

/** Resolve the vertical anchor for content that lays out as exactly one line.
 *  ECMA-376 Part 1 §18.8.1 makes wrapText a line-wrapping switch, while
 *  §18.18.88 defines vertical alignment over the cell height. A wrap-enabled
 *  value that still fits on one line therefore uses the same top/middle/bottom
 *  anchor as an unwrapped value; only a genuinely multi-line result owns a
 *  line-box block. */
export function singleLineVerticalAnchor(geom: RichCellGeom): {
  baseline: CanvasTextBaseline;
  textY: number;
} {
  const { alignV, cy, cellH, paddingY } = geom;
  if (alignV === 'top') return { baseline: 'top', textY: cy + paddingY };
  if (alignV === 'center') return { baseline: 'middle', textY: cy + cellH / 2 };
  return { baseline: 'bottom', textY: cy + cellH - paddingY };
}

/** The draw font for a run/segment: super/subscript renders at ~65% size
 *  (ECMA-376 §18.4.14 vertAlign / ST_VerticalAlignRun §22.9.2.17); everything
 *  else keeps its declared size. The *base* size still governs line height and
 *  the baseline shift — only the glyph (and the width it occupies) shrinks. */
function vertAlignDrawFont(font: CellFont): CellFont {
  return (font.vertAlign === 'superscript' || font.vertAlign === 'subscript')
    ? { ...font, size: font.size * 0.65 } : font;
}

/**
 * Draw a sequence of already-resolved, already-measured rich segments on one
 * line at `textY` under `baseline`, in bidi visual order (UAX#9 rule L2). The
 * caller supplies whether the cell needs the bidi pass and its base direction —
 * a wrapped cell's display lines share one paragraph direction, while hard-break
 * lines are independent paragraphs. Each segment draws with its own font/color,
 * ~65%-size super/subscript baseline shift (§18.4.14 vertAlign / §22.9.2.17), and
 * underline / strike decoration. `startX` is the line's left edge (the caller
 * resolves alignH from the line's measured width). The single shared per-segment
 * drawer for every rich path — non-wrap (via {@link drawRichLine}) and wrap.
 */
export function drawResolvedRichLine(
  ctx: CanvasRenderingContext2D,
  segs: RichSeg[],
  startX: number,
  textY: number,
  baseline: CanvasTextBaseline,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; needBidi?: boolean; baseRtl?: boolean },
  cjkFallback?: CjkLang,
): void {
  ctx.textAlign = 'left';
  ctx.textBaseline = baseline;
  const vis = opts.needBidi ? computeLineVisualOrder(segs, opts.baseRtl ?? false) : null;
  const dctx = ctx as CanvasRenderingContext2D & { direction: 'ltr' | 'rtl' };
  let x = startX;
  for (let vi = 0; vi < segs.length; vi++) {
    const i = vis ? vis.order[vi] : vi;
    if (vis) { try { dctx.direction = vis.rtl[i] ? 'rtl' : 'ltr'; } catch { /* ignore */ } }
    const seg = segs[i];
    const drawFont = vertAlignDrawFont(seg.font);
    ctx.font = buildFont(ctx, drawFont, cs, cjkFallback, seg.text);
    const segColor = opts.fontColor ?? seg.font.color;
    ctx.fillStyle = segColor ? hexToRgba(segColor) : '#000000';
    // Baseline shift for super/subscript, relative to the run's *base* size: up
    // for super, slightly down for sub so each sits at the right vertical band.
    const baseSizePx = vMetricPx(seg.font.size, cs);
    let yShift = 0;
    if (seg.font.vertAlign === 'superscript') yShift = -Math.round(baseSizePx * 0.35);
    else if (seg.font.vertAlign === 'subscript') yShift = Math.round(baseSizePx * 0.10);
    paintCellText(ctx, seg.text, x, textY + yShift, drawFont, cs);
    const rSizePx = vMetricPx(drawFont.size, cs);
    if (seg.font.underline || seg.font.strike) {
      const deco = decoYForBaseline(baseline, textY, rSizePx);
      if (seg.font.underline) {
        const stroke = segColor ? hexToRgba(segColor) : '#000000';
        const dbl = seg.font.underlineStyle === 'double' || seg.font.underlineStyle === 'doubleAccounting';
        drawTextDecoLine(ctx, x, x + seg.width, deco.underline + yShift, stroke, dbl, dpr);
      }
      if (seg.font.strike) {
        const syBase = deco.strike + yShift;
        const sy = syBase + crispOffset(syBase, 0.5, dpr);
        ctx.save();
        ctx.strokeStyle = segColor ? hexToRgba(segColor) : '#000000';
        ctx.lineWidth = 0.5;
        ctx.beginPath(); ctx.moveTo(x, sy); ctx.lineTo(x + seg.width, sy); ctx.stroke();
        ctx.restore();
      }
    }
    x += seg.width;
  }
  // Defensive reset (the per-cell ctx.restore() also restores direction).
  if (vis) { try { dctx.direction = 'ltr'; } catch { /* ignore */ } }
}

/**
 * Draw one already-grouped line of rich-text RUNS at `textY` under `baseline`,
 * positioned horizontally by `geom.alignH` over its measured width. Resolves each
 * run's font (super/subscript at ~65% size, §18.4.14) into a measured segment —
 * the segment keeps the run's *full* size (line height + baseline shift) but its
 * width is measured at the draw font, so a super/subscript run reserves its
 * reduced glyph width — then delegates the painting to {@link drawResolvedRichLine}.
 * Used by the non-wrap rich paths; the wrap path drives `drawResolvedRichLine`
 * directly from its laid-out segments.
 */
function drawRichSegments(
  ctx: CanvasRenderingContext2D,
  segs: RichSeg[],
  geom: RichCellGeom,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; readingOrder?: number },
  textY: number,
  baseline: CanvasTextBaseline,
  cjkFallback?: CjkLang,
): void {
  const { alignH, cx, cellW, leftPad, paddingX } = geom;
  const totalWidth = segs.reduce((a, s) => a + s.width, 0);
  let startX: number;
  if (alignH === 'right') startX = cx + cellW - paddingX - totalWidth;
  else if (alignH === 'center') startX = cx + cellW / 2 - totalWidth / 2;
  else startX = cx + leftPad;
  const { needBidi, baseRtl } = resolveCellBidi(opts.readingOrder, segs.map((s) => s.text).join(''));
  drawResolvedRichLine(ctx, segs, startX, textY, baseline, cs, dpr, { fontColor: opts.fontColor, needBidi, baseRtl }, cjkFallback);
}

function drawRichLine(
  ctx: CanvasRenderingContext2D,
  lineRuns: Run[],
  baseFont: CellFont,
  geom: RichCellGeom,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; readingOrder?: number },
  textY: number,
  baseline: CanvasTextBaseline,
  cjkFallback?: CjkLang,
): void {
  const segs: RichSeg[] = lineRuns.map((r) => {
    const font = applyRunFont(baseFont, r);
    ctx.font = buildFont(ctx, vertAlignDrawFont(font), cs, cjkFallback, r.text);
    return { text: r.text, font, width: ctx.measureText(r.text).width };
  });
  drawRichSegments(ctx, segs, geom, cs, dpr, opts, textY, baseline, cjkFallback);
}

/**
 * Draw a single line of rich text (no hard break) with per-run fonts, anchored
 * by the cell's alignV-dependent baseline — `'top'` at the top padding,
 * `'middle'` at the cell centre, `'bottom'` at the bottom padding — exactly how
 * Excel paints a one-line rich cell. Shared by the in-viewport draw path and the
 * off-screen-anchor merge pre-pass so a merged rich cell renders identically
 * whether or not its anchor is scrolled out of view.
 */
function drawSingleLineRichText(
  ctx: CanvasRenderingContext2D,
  runs: Run[],
  baseFont: CellFont,
  geom: RichCellGeom,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; readingOrder?: number } = {},
  cjkFallback?: CjkLang,
): void {
  const { baseline, textY } = singleLineVerticalAnchor(geom);
  drawRichLine(ctx, runs, baseFont, geom, cs, dpr, opts, textY, baseline, cjkFallback);
}

/**
 * Lay out and draw rich text with hard line breaks (LF / Alt+Enter) in a
 * non-wrapped cell. Runs are split at every LF into lines; a blank line from
 * consecutive / leading / trailing breaks reserves one single-line height (the
 * cell analog of PR #585 / docx #582), sized from the nearest preceding text
 * run. Each line is drawn at a `'top'` baseline (matching the wrap rich path)
 * and the whole block is anchored vertically by `alignV` over its summed height.
 */
function drawMultiLineRichText(
  ctx: CanvasRenderingContext2D,
  runs: Run[],
  baseFont: CellFont,
  geom: RichCellGeom,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; readingOrder?: number } = {},
  cjkFallback?: CjkLang,
): void {
  const { alignV, cy, cellH, paddingY } = geom;

  const lines = richHardBreakLineMetrics(runs, baseFont, cs);
  const totalH = lines.reduce((sum, line) => sum + line.heightPx, 0);

  let yy: number;
  if (alignV === 'top') yy = cy + paddingY;
  else if (alignV === 'center') yy = cy + (cellH - totalH) / 2;
  else yy = cy + cellH - totalH - paddingY;

  for (const line of lines) {
    // A blank line draws nothing but still reserves its height.
    if (line.runs.length > 0) drawRichLine(ctx, line.runs, baseFont, geom, cs, dpr, opts, yy, 'top', cjkFallback);
    yy += line.heightPx;
  }
}

/** Invalid date/time display follows Excel's hash fill (issue #1710).
 * Formatting APIs expose only '#'; painting repeats it in the visible cell
 * width. Bound the allocation by the target canvas, including huge/offscreen
 * merges, and keep the marker horizontal even for wrapped/rotated cells.
 */
export function drawHashFill(
  ctx: CanvasRenderingContext2D,
  cx: number, cy: number, cellW: number, cellH: number,
  alignV: string, font: CellFont, cs: number,
): void {
  const padding = 3;
  const left = Math.max(0, cx) + padding;
  const available = Math.min(cx + cellW, ctx.canvas.width) - padding - left;
  if (available <= 0) return;
  const hashWidth = ctx.measureText('#').width;
  const count = hashWidth > 0 ? Math.max(1, Math.floor(available / hashWidth)) : 1;
  ctx.textAlign = 'left';
  const { baseline, textY } = singleLineVerticalAnchor({
    alignH: 'left', alignV, cx, cy, cellW, cellH,
    leftPad: padding, paddingX: padding, paddingY: 2,
  });
  ctx.textBaseline = baseline;
  paintCellText(ctx, '#'.repeat(count), left, textY, font, cs);
}

/**
 * Draw rich text (mixed-font runs, ECMA-376 §18.4.4 r) in a NON-wrapped cell.
 * A break-free value is one alignV-anchored line ({@link drawSingleLineRichText});
 * a value with a hard break is laid out as multiple lines
 * ({@link drawMultiLineRichText}).
 *
 * §18.8.1 (CT_CellAlignment @wrapText) governs only soft-wrapping; it says
 * nothing about hard breaks. Rendering a literal LF (Alt+Enter, preserved in the
 * run text via §18.4.12 t @xml:space) as a line break even when wrapText is off
 * is undocumented Excel runtime behavior — matched here for parity with the
 * plain-text non-wrap path and the wrap rich-text path (#585).
 */
export function drawNonWrapRichText(
  ctx: CanvasRenderingContext2D,
  runs: Run[],
  baseFont: CellFont,
  geom: RichCellGeom,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; readingOrder?: number } = {},
  cjkFallback?: CjkLang,
): void {
  if (runs.some((r) => r.text.includes('\n'))) {
    drawMultiLineRichText(ctx, runs, baseFont, geom, cs, dpr, opts, cjkFallback);
  } else {
    drawSingleLineRichText(ctx, runs, baseFont, geom, cs, dpr, opts, cjkFallback);
  }
}

/** Lay out and draw a plain wrap-enabled cell value. A value that does not
 *  actually wrap remains a single line and uses the ordinary cell-height
 *  anchor; only two or more produced lines use the design line-box block. This
 *  helper is shared by visible cells and off-screen merged anchors so scrolling
 *  cannot select a different vertical-placement algorithm. */
export function drawWrappedPlainText(
  ctx: CanvasRenderingContext2D,
  text: string,
  textX: number,
  font: CellFont,
  geom: RichCellGeom,
  cs: number,
): void {
  const { alignV, cy, cellH, leftPad, paddingX, paddingY } = geom;
  const available = geom.cellW - leftPad - paddingX;
  // Wrap and paint with the same resolved Canvas face. A fixed-height cell
  // may clip lines when a substitute is wider, as it would with any wider
  // installed face; narrowing glyphs to an unrelated font's advances makes
  // the substitute itself look distorted and can conceal overflow.
  const lines = wrapTextLines(ctx, text, available);
  if (lines.length === 1) {
    const { baseline, textY } = singleLineVerticalAnchor(geom);
    ctx.textBaseline = baseline;
    paintCellText(ctx, lines[0], textX, textY, font, cs);
    return;
  }

  const lineH = vMetricPx(font.size, cs, 1.2);
  const totalTextH = lines.length * lineH;
  let startY: number;
  if (alignV === 'top') startY = cy + paddingY;
  else if (alignV === 'center') startY = cy + (cellH - totalTextH) / 2;
  else startY = cy + cellH - totalTextH - paddingY;
  ctx.textBaseline = 'top';
  for (let li = 0; li < lines.length; li++) {
    paintCellText(ctx, lines[li], textX, startY + li * lineH, font, cs);
  }
}

/**
 * Lay out and draw WRAP-mode rich text (mixed-font runs, §18.8.1 wrapText on).
 * `layoutRichTextLines` soft-wraps at word / CJK boundaries (and hard breaks),
 * and each resulting line is painted by the shared {@link drawResolvedRichLine}
 * (super/subscript §18.4.14, underline/strike, bidi).
 *
 * Base direction (§18.8.1 readingOrder) is resolved PER LF-delimited paragraph,
 * not once for the whole cell: a soft wrap continues the same bidi paragraph (its
 * display lines share a direction), but a hard break (Alt+Enter LF) starts a new
 * paragraph that, under Context reading order (0/absent), resolves its OWN base
 * direction from its OWN first strong character (UAX#9 P1–P3). Explicit LTR/RTL
 * (1/2) is cell-wide, so only Context varies between paragraphs. Each line carries
 * its paragraph index ({@link RichLine.para}) to pick the matching direction.
 *
 * Drives both the in-viewport draw path and the off-screen-anchor merge pre-pass
 * so a merged wrapped cell renders identically regardless of anchor visibility.
 */
export function drawWrappedRichText(
  ctx: CanvasRenderingContext2D,
  runs: Run[],
  baseFont: CellFont,
  geom: RichCellGeom,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; readingOrder?: number } = {},
  cjkFallback?: CjkLang,
): void {
  const { alignH, alignV, cx, cy, cellW, cellH, leftPad, paddingX, paddingY } = geom;
  const rLines = layoutRichTextLines(ctx, runs, baseFont, cs, cellW - leftPad - paddingX, cjkFallback);
  // layoutRichTextLines preserves leading, trailing, and consecutive hard-break
  // regions, so exactly one produced line also proves there was no hard break.
  if (rLines.length === 1) {
    const { baseline, textY } = singleLineVerticalAnchor(geom);
    drawRichSegments(ctx, rLines[0].segments, geom, cs, dpr, opts, textY, baseline, cjkFallback);
    return;
  }
  const totalH = rLines.reduce((s, l) => s + vMetricPx(l.maxFontSize, cs, 1.2), 0);
  let yy: number;
  if (alignV === 'top') yy = cy + paddingY;
  else if (alignV === 'center') yy = cy + (cellH - totalH) / 2;
  else yy = cy + cellH - totalH - paddingY;
  // Resolve the bidi base direction for each LF-delimited paragraph from its own
  // text (Context = first-strong; explicit 1/2 = cell-wide). `line.para` indexes
  // this, so soft-wrapped lines share their paragraph's direction while a hard
  // break gets its own. `resolveCellBidi` gates each so pure-LTR paragraphs keep
  // the exact pre-bidi path — the same gate the non-wrap path uses per line.
  const paraTexts = runs.map((r) => r.text).join('').split('\n');
  const paraBidi = paraTexts.map((t) => resolveCellBidi(opts.readingOrder, t));
  for (const line of rLines) {
    const totalW = line.segments.reduce((s, seg) => s + seg.width, 0);
    let xx: number;
    if (alignH === 'right') xx = cx + cellW - paddingX - totalW;
    else if (alignH === 'center') xx = cx + cellW / 2 - totalW / 2;
    else xx = cx + leftPad;
    const { needBidi, baseRtl } = paraBidi[line.para];
    drawResolvedRichLine(ctx, line.segments, xx, yy, 'top', cs, dpr, { fontColor: opts.fontColor, needBidi, baseRtl }, cjkFallback);
    yy += vMetricPx(line.maxFontSize, cs, 1.2);
  }
}

export function richHardBreakLineMetrics(
  runs: Run[],
  baseFont: CellFont,
  cs: number,
): Array<{ runs: Run[]; heightPx: number }> {
  const lineRuns: Run[][] = [[]];
  for (const run of runs) {
    const parts = run.text.split('\n');
    for (let index = 0; index < parts.length; index++) {
      if (index > 0) lineRuns.push([]);
      if (parts[index] !== '') lineRuns[lineRuns.length - 1].push({ ...run, text: parts[index] });
    }
  }

  let lastTextPt = baseFont.size;
  return lineRuns.map((line) => {
    if (line.length === 0) {
      return {
        runs: line,
        heightPx: vMetricPx(lastTextPt || DEFAULT_FONT_SIZE, cs, 1.2),
      };
    }
    let maxPt = 0;
    for (const run of line) {
      const font = applyRunFont(baseFont, run);
      if (font.size > maxPt) {
        maxPt = font.size;
      }
      lastTextPt = font.size;
    }
    return {
      runs: line,
      heightPx: vMetricPx(maxPt, cs, 1.2),
    };
  });
}
