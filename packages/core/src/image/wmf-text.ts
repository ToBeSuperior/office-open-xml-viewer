// ── WMF text output (META_TEXTOUT / META_EXTTEXTOUT) ────────────────────────
//
// [MS-WMF] references (v18.0):
//   2.3.3.5  META_EXTTEXTOUT   Y, X, StringLength, fwOpts, optional Rectangle
//                              (present with ETO_OPAQUE or ETO_CLIPPED), String
//                              padded to 16 bits, optional Dx array.
//   2.3.3.20 META_TEXTOUT      StringLength, String (padded), YStart, XStart.
//   2.2.1.2  Font object       Height, Width, Escapement, Orientation, Weight,
//                              Italic, Underline, StrikeOut, CharSet, …,
//                              PitchAndFamily, FaceName (NUL-terminated, at
//                              most 32 bytes including the NUL).
//   2.1.1.5  CharacterSet, 2.1.2.2 ExtTextOutOptions, 2.1.2.3 TextAlignmentMode,
//   2.1.1.8  FamilyFont / 2.1.1.24 PitchFont,
//   2.3.5.11 META_SETBKCOLOR, 2.3.5.15 META_SETBKMODE, 2.3.5.24 META_SETTEXTALIGN,
//   2.3.5.25 META_SETTEXTCHAREXTRA, 2.3.5.26 META_SETTEXTCOLOR,
//   2.3.5.27 META_SETTEXTJUSTIFICATION.
// WMF playback has no SETGRAPHICSMODE, so text is always GM_COMPATIBLE: the
// LOGFONT documentation states that lfEscapement then specifies both the
// escapement and the orientation (lfOrientation is not used), and glyphs stay
// upright on the device regardless of a mirrored logical mapping.
//
// Character spacing follows the GDI ExtTextOut model: an authored Dx array
// replaces the font's natural advances; SETTEXTCHAREXTRA is added to every
// character and SETTEXTJUSTIFICATION's BreakExtra is shared across the break
// (space) characters on display devices.
//
// Fail-closed (reported through the player's incomplete-metafile channel) only
// where neither MS-WMF nor the stored bytes define the result: no selected
// font (the playback DC default font is device-dependent, MS-WMF 3.1.5), a
// zero font height (font-mapper default), an undefined TA_UPDATECP current
// position, glyph-index strings (ETO_GLYPH_INDEX names glyphs of a font
// program we do not have), '@' vertical faces (font-mapper behaviour),
// undefined alignment codes, DEFAULT/OEM charsets with non-ASCII bytes (the
// code page is the playback system's locale), JOHAB (no decoder available),
// bytes undefined in the selected code page, and malformed records.

import { findReferenceFontMetrics } from '../fonts/reference-font-metrics.js';
import {
  hasMetafileSymbolRepertoire,
  metafileSymbolCodeToUnicode,
  SYMBOL_CHARSET_SUBSTITUTE_FACE,
} from '../fonts/metafile-symbol-encoding.js';
import { officeFaceWebAlias } from '../fonts/font-stack.js';
import { classifyFontGeneric, cjkFallbackChain, type CjkLang } from '../fonts/scripts.js';
import { quoteCssFontFamily } from '../internal/css-font-family.js';
import { referenceFontDecoration } from '../text/font-decoration.js';
import { DECORATION_LINE_WIDTH_EM, STRIKE_RISE_EM } from '../text/underline.js';

type AnyCtx = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/** MS-WMF 2.2.1.2 Font object, as authored by META_CREATEFONTINDIRECT. */
export interface WmfFont {
  kind: 'font';
  /** Signed Height: < 0 character (em) height, > 0 cell height, 0 default. */
  height: number;
  width: number;
  escapement: number;
  orientation: number;
  weight: number;
  italic: boolean;
  underline: boolean;
  strikeout: boolean;
  charset: number;
  pitchAndFamily: number;
  face: string;
  /** False when the record is too short or FaceName has no NUL terminator. */
  valid: boolean;
}

// ExtTextOutOptions (MS-WMF 2.1.2.2).
export const ETO_OPAQUE = 0x0002;
export const ETO_CLIPPED = 0x0004;
const ETO_GLYPH_INDEX = 0x0010;
const ETO_RTLREADING = 0x0080;
const ETO_PDY = 0x2000;
// TextAlignmentMode (MS-WMF 2.1.2.3).
export const TA_UPDATECP = 0x0001;
const TA_RIGHT = 0x0002;
const TA_CENTER = 0x0006;
const TA_BOTTOM = 0x0008;
const TA_BASELINE = 0x0018;
const TA_RTLREADING = 0x0100;
// BackgroundMode (MS-WMF 2.1.1.4).
const OPAQUE = 0x0002;

// CharacterSet (MS-WMF 2.1.1.5) → Windows code page decoders (WHATWG labels).
const ANSI_CHARSET = 0;
const DEFAULT_CHARSET = 1;
const SYMBOL_CHARSET = 2;
const JOHAB_CHARSET = 130;
const OEM_CHARSET = 255;
const SINGLE_BYTE_CODE_PAGES: Readonly<Record<number, string>> = {
  77: 'macintosh', // MAC_CHARSET, code page 10000
  161: 'windows-1253', // GREEK_CHARSET
  162: 'windows-1254', // TURKISH_CHARSET
  163: 'windows-1258', // VIETNAMESE_CHARSET
  177: 'windows-1255', // HEBREW_CHARSET
  178: 'windows-1256', // ARABIC_CHARSET
  186: 'windows-1257', // BALTIC_CHARSET
  204: 'windows-1251', // RUSSIAN_CHARSET
  222: 'windows-874', // THAI_CHARSET
  238: 'windows-1250', // EASTEUROPE_CHARSET
};
interface DoubleByteCodePage { label: string; lead: (byte: number) => boolean; cjk: CjkLang }
const DOUBLE_BYTE_CODE_PAGES: Readonly<Record<number, DoubleByteCodePage>> = {
  // SHIFTJIS_CHARSET, code page 932: lead bytes 0x81–0x9F and 0xE0–0xFC.
  128: { label: 'shift_jis', lead: (b) => (b >= 0x81 && b <= 0x9f) || (b >= 0xe0 && b <= 0xfc), cjk: 'jp' },
  // HANGUL_CHARSET 949, GB2312_CHARSET 936, CHINESEBIG5_CHARSET 950: 0x81–0xFE.
  129: { label: 'euc-kr', lead: (b) => b >= 0x81 && b <= 0xfe, cjk: 'kr' },
  134: { label: 'gbk', lead: (b) => b >= 0x81 && b <= 0xfe, cjk: 'sc' },
  136: { label: 'big5', lead: (b) => b >= 0x81 && b <= 0xfe, cjk: 'tc' },
};

// WHATWG Encoding §9.1, index-windows-1252.txt (identifier e56d49d9,
// 2024-09-18). Decode directly so WMF ANSI text is deterministic even on
// hosts whose TextDecoder delegates this label to an ISO-8859-1 codec.
const WINDOWS_1252_C1 = [
  0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021,
  0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008d, 0x017d, 0x008f,
  0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
  0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
] as const;

const decoders = new Map<string, TextDecoder | null>();
function decoderFor(label: string): TextDecoder | null {
  if (!decoders.has(label)) {
    let decoder: TextDecoder | null = null;
    try { decoder = new TextDecoder(label, { fatal: true }); } catch { decoder = null; }
    decoders.set(label, decoder);
  }
  return decoders.get(label) ?? null;
}

/** One drawn character: its Unicode text, the String bytes it came from (Dx
 *  entries are per byte), and the face class it must be drawn with. */
interface TextUnit {
  text: string;
  start: number;
  count: number;
  /** 'authored' = the selected face; 'symbol' = a symbol code converted to
   *  Unicode, drawn in a Unicode fallback (the symbol face would reinterpret it);
   *  'substitute' = a font code drawn in the font mapper's symbol substitute. */
  face: 'authored' | 'symbol' | 'substitute';
}

type Decoded = { units: TextUnit[] } | { problem: string };

/** Decode a byte string under a CharacterSet, or report why it is undefined.
 *  `hostFaceCovers` tells whether the authored face on this host supplies
 *  glyphs for a string (only consulted for symbol faces). */
function decodeBytes(
  bytes: Uint8Array,
  charset: number,
  face: string,
  hostFaceCovers: (text: string) => boolean,
): Decoded {
  const units: TextUnit[] = [];
  if (charset === SYMBOL_CHARSET) {
    // Symbol fonts are addressed through the F000 private range of their
    // symbol cmap. A face the host lacks is replaced by the font mapper's
    // symbol-charset substitute. A published repertoire converts a code to
    // Unicode; any other code stays font-coded in the face that draws it.
    const fontCoded = Array.from(bytes, (byte) => String.fromCharCode(0xf000 + byte)).join('');
    const substituted = !hasMetafileSymbolRepertoire(face) && bytes.length > 0 && !hostFaceCovers(fontCoded);
    const repertoire = substituted ? SYMBOL_CHARSET_SUBSTITUTE_FACE : face;
    bytes.forEach((byte, index) => {
      const mapped = metafileSymbolCodeToUnicode(byte, repertoire);
      units.push(mapped === undefined
        ? { text: fontCoded[index], start: index, count: 1, face: substituted ? 'substitute' : 'authored' }
        : { text: mapped, start: index, count: 1, face: 'symbol' });
    });
    return { units };
  }
  if (charset === ANSI_CHARSET) {
    bytes.forEach((byte, index) => units.push({
      text: String.fromCharCode(byte >= 0x80 && byte <= 0x9f ? WINDOWS_1252_C1[byte - 0x80] : byte),
      start: index, count: 1, face: 'authored',
    }));
    return { units };
  }
  if (charset === DEFAULT_CHARSET || charset === OEM_CHARSET) {
    // Every Windows ANSI and OEM code page shares ASCII; other bytes depend on
    // the playback system's locale.
    if (bytes.some((byte) => byte >= 0x80)) return { problem: 'locale-dependent character set' };
    bytes.forEach((byte, index) => units.push({ text: String.fromCharCode(byte), start: index, count: 1, face: 'authored' }));
    return { units };
  }
  const single = SINGLE_BYTE_CODE_PAGES[charset];
  const double = DOUBLE_BYTE_CODE_PAGES[charset];
  const decoder = single ? decoderFor(single) : double ? decoderFor(double.label) : null;
  if (!decoder) {
    return { problem: charset === JOHAB_CHARSET || single || double ? 'character set decoder unavailable' : 'undefined character set' };
  }
  for (let index = 0; index < bytes.length;) {
    const count = double && double.lead(bytes[index]) ? 2 : 1;
    if (index + count > bytes.length) return { problem: 'truncated double-byte character' };
    let text: string;
    try { text = decoder.decode(bytes.subarray(index, index + count)); } catch { return { problem: 'byte undefined in code page' }; }
    units.push({ text, start: index, count, face: 'authored' });
    index += count;
  }
  return { units };
}

/** MS-WMF 2.2.1.2: decode a CREATEFONTINDIRECT parameter block. */
export function parseWmfFont(params: Uint8Array): WmfFont {
  const dv = new DataView(params.buffer, params.byteOffset, params.byteLength);
  const fixed = params.byteLength >= 18;
  const i16 = (offset: number) => (fixed ? dv.getInt16(offset, true) : 0);
  const faceBytes = params.subarray(18, Math.min(params.byteLength, 50));
  const nul = faceBytes.indexOf(0);
  const charset = fixed ? params[13] : 0;
  const rawFace = faceBytes.subarray(0, nul < 0 ? faceBytes.length : nul);
  // FaceName is stored in the font's own code page for DBCS faces; decode it
  // with the CharacterSet's decoder when one applies, else as Latin-1.
  const labelled = DOUBLE_BYTE_CODE_PAGES[charset]?.label ?? SINGLE_BYTE_CODE_PAGES[charset];
  let face = String.fromCharCode(...rawFace);
  if (labelled) {
    try { face = decoderFor(labelled)?.decode(rawFace) ?? face; } catch { /* keep Latin-1 */ }
  }
  return {
    kind: 'font',
    height: i16(0),
    width: i16(2),
    escapement: i16(4),
    orientation: i16(6),
    weight: i16(8),
    italic: fixed && params[10] !== 0,
    underline: fixed && params[11] !== 0,
    strikeout: fixed && params[12] !== 0,
    charset,
    pitchAndFamily: fixed ? params[17] : 0,
    face,
    valid: fixed && nul > 0,
  };
}

/** CSS generic family from PitchAndFamily (MS-WMF 2.2.2.14: family in bits
 *  4–7, FamilyFont 2.1.1.8; pitch in bits 0–1, PitchFont 2.1.1.24). An
 *  FF_DONTCARE face falls back to the shared name classifier. */
function genericFamily(font: WmfFont): string {
  switch (font.pitchAndFamily & 0xf0) {
    case 0x10: return 'serif'; // FF_ROMAN
    case 0x20: return 'sans-serif'; // FF_SWISS
    case 0x30: return 'monospace'; // FF_MODERN
    case 0x40: return 'cursive'; // FF_SCRIPT
    case 0x50: return 'fantasy'; // FF_DECORATIVE
    default: break;
  }
  if ((font.pitchAndFamily & 0x3) === 1) return 'monospace'; // FIXED_PITCH
  const generic = classifyFontGeneric(font.face);
  return generic === 'mono' ? 'monospace' : generic === 'serif' ? 'serif' : 'sans-serif';
}

function authoredStack(font: WmfFont, generic: string): string {
  const names = [font.face];
  const alias = officeFaceWebAlias(font.face);
  if (alias) names.push(alias);
  const cjk = DOUBLE_BYTE_CODE_PAGES[font.charset]?.cjk;
  if (cjk) names.push(...cjkFallbackChain(cjk, generic === 'serif' ? 'serif' : 'sans'));
  return `${names.map(quoteCssFontFamily).join(', ')}, ${generic}`;
}

/** Converted symbol text uses a Unicode face of the authored family class.
 *  An unclassified symbol face keeps the serif class of the published Symbol
 *  repertoire (its Greek and mathematical glyphs are a serif design). */
function symbolStack(font: WmfFont): string {
  return (font.pitchAndFamily & 0xf0) === 0 ? 'serif' : genericFamily(font);
}

interface FaceMetrics {
  /** GDI tmAscent / tmDescent as fractions of the em. */
  ascentEm: number;
  descentEm: number;
  /** tmAveCharWidth as a fraction of the em (evaluated only for a non-zero Width). */
  avgCharWidthEm: () => number;
}

// OpenType OS/2 xAvgCharWidth (versions 0–2): weighted lowercase + space.
const AVG_WIDTH_WEIGHTS: ReadonlyArray<readonly [string, number]> = [
  ['a', 64], ['b', 14], ['c', 27], ['d', 35], ['e', 100], ['f', 20], ['g', 14],
  ['h', 42], ['i', 63], ['j', 3], ['k', 6], ['l', 35], ['m', 20], ['n', 56],
  ['o', 56], ['p', 17], ['q', 4], ['r', 49], ['s', 56], ['t', 71], ['u', 31],
  ['v', 10], ['w', 18], ['x', 3], ['y', 18], ['z', 2], [' ', 166],
];

const PROBE_EM = 1000;

/** GDI realizes TrueType text metrics from OS/2 usWinAscent/usWinDescent and
 *  xAvgCharWidth. Use the catalogued values for the authored face; otherwise
 *  measure the face Canvas actually selects on this host. */
function faceMetrics(ctx: AnyCtx, font: WmfFont, css: (px: number) => string, weight: number): FaceMetrics | null {
  const style = font.italic ? 'italic' : 'normal';
  const profile = findReferenceFontMetrics(font.face, { weight, style })[0]
    ?? findReferenceFontMetrics(font.face, { style })[0]
    ?? findReferenceFontMetrics(font.face)[0];
  if (profile && profile.unitsPerEm > 0) {
    const upm = profile.unitsPerEm;
    const [ascent, descent] = profile.win ?? [profile.hhea[0], -profile.hhea[1]];
    const avg = profile.xAvgCharWidth;
    return {
      ascentEm: ascent / upm,
      descentEm: descent / upm,
      avgCharWidthEm: () => (avg && avg > 0 ? avg / upm : measuredAverageWidth(ctx, css)),
    };
  }
  ctx.font = css(PROBE_EM);
  const probe = ctx.measureText('Hg');
  const ascent = probe.fontBoundingBoxAscent;
  const descent = probe.fontBoundingBoxDescent;
  if (!Number.isFinite(ascent) || !Number.isFinite(descent) || ascent + descent <= 0) return null;
  return {
    ascentEm: ascent / PROBE_EM,
    descentEm: descent / PROBE_EM,
    avgCharWidthEm: () => measuredAverageWidth(ctx, css),
  };
}

function measuredAverageWidth(ctx: AnyCtx, css: (px: number) => string): number {
  ctx.font = css(PROBE_EM);
  let total = 0;
  for (const [char, weight] of AVG_WIDTH_WEIGHTS) total += ctx.measureText(char).width * weight;
  return total / 1000 / PROBE_EM;
}

const MISSING_CONTROL_FAMILY = '__ooxml_wmf_missing_face_control_5b2e__';

/** Whether Canvas draws `text` from `face` itself on this host: compare its
 *  glyph geometry against a deliberately absent control family (CSS silently
 *  substitutes a missing family, so equal geometry means fallback). */
function hostFaceCovers(ctx: AnyCtx, face: string, text: string): boolean {
  const geometry = (family: string) => {
    ctx.font = `${PROBE_EM}px ${family}${quoteCssFontFamily(MISSING_CONTROL_FAMILY)}, serif`;
    const m = ctx.measureText(text);
    return [m.width, m.actualBoundingBoxAscent, m.actualBoundingBoxDescent, m.actualBoundingBoxLeft, m.actualBoundingBoxRight];
  };
  const selected = geometry(`${quoteCssFontFamily(face)}, `);
  const missing = geometry('');
  return !selected.every((value, index) => Object.is(value, missing[index]));
}

/** The DC state text output reads; PlayState in wmf.ts satisfies it. */
export interface WmfTextState {
  ctx: AnyCtx;
  orgX: number; orgY: number; extX: number; extY: number;
  vpX: number; vpY: number; vpW: number; vpH: number;
  haveExt: boolean;
  curFont: WmfFont | null;
  textColor: string;
  bkColor: string;
  bkMode: number;
  textAlign: number;
  charExtra: number;
  breakExtra: number;
  breakCount: number;
  currentX: number;
  currentY: number;
  currentPositionKnown: boolean;
}

export interface WmfTextOutput {
  x: number;
  y: number;
  bytes: Uint8Array;
  options: number;
  /** Rect (MS-WMF 2.2.2.18) Left, Top, Right, Bottom in logical units. */
  rect: readonly [number, number, number, number] | null;
  /** Dx (and with ETO_PDY interleaved Dy) values, one entry per String byte. */
  dx: Int16Array | null;
}

/** Draw one text output. Returns whether anything was painted, or the reason
 *  the output is undefined for this player. */
export function drawWmfText(s: WmfTextState, out: WmfTextOutput): { drew: boolean } | { problem: string } {
  const sx = s.vpW / s.extX;
  const sy = s.vpH / s.extY;
  if (!s.haveExt) return { problem: 'logical mapping undefined' };
  if (!(Number.isFinite(sx) && Number.isFinite(sy) && sx !== 0 && sy !== 0)) return { problem: 'degenerate mapping' };
  const font = s.curFont;
  if (!font) return { problem: 'no font selected' };
  if (!font.valid) return { problem: 'malformed font' };
  if (font.height === 0) return { problem: 'default font height' };
  if (font.face.startsWith('@')) return { problem: 'vertical font' };
  if (out.options & ETO_GLYPH_INDEX) return { problem: 'glyph index string' };
  const horizontal = s.textAlign & TA_CENTER;
  const vertical = s.textAlign & TA_BASELINE;
  if (horizontal === 0x4 || vertical === 0x10) return { problem: 'undefined text alignment' };
  if ((s.textAlign & TA_UPDATECP) && !s.currentPositionKnown) return { problem: 'current position undefined' };

  const { ctx } = s;
  const weight = font.weight === 0 ? 400 : Math.min(1000, Math.max(1, font.weight));
  const style = font.italic ? 'italic ' : '';
  const generic = genericFamily(font);
  const authored = authoredStack(font, generic);
  const symbol = symbolStack(font);
  const substitute = `${quoteCssFontFamily(SYMBOL_CHARSET_SUBSTITUTE_FACE)}, ${symbol}`;
  const css = (stack: string) => (px: number) => `${style}${weight} ${px}px ${stack}`;

  const decoded = decodeBytes(out.bytes, font.charset, font.face, (text) => hostFaceCovers(ctx, font.face, text));
  if ('problem' in decoded) return decoded;
  const { units } = decoded;
  const metrics = faceMetrics(ctx, font, css(authored), weight);
  if (!metrics) return { problem: 'font metrics unavailable' };
  // Font height maps through the logical y scale (MS-WMF 2.2.1.2 Height).
  const emLogical = font.height < 0 ? -font.height : font.height / (metrics.ascentEm + metrics.descentEm);
  const emPx = emLogical * Math.abs(sy);
  if (!(emPx > 0 && Number.isFinite(emPx))) return { problem: 'degenerate font size' };
  const ascentPx = metrics.ascentEm * emPx;
  const descentPx = metrics.descentEm * emPx;
  const hScale = font.width === 0 ? 1 : (Math.abs(font.width) * Math.abs(sx)) / (metrics.avgCharWidthEm() * emPx);
  if (!(hScale > 0 && Number.isFinite(hScale))) return { problem: 'degenerate font width' };
  const fontFor = (unit: TextUnit) =>
    css(unit.face === 'symbol' ? symbol : unit.face === 'substitute' ? substitute : authored)(emPx);

  // Baseline direction: escapement in tenths of a degree, counterclockwise on
  // the device. Logical lengths along / across it scale with the mapping.
  const theta = (font.escapement / 10) * (Math.PI / 180);
  const cos = Math.cos(theta);
  const sin = Math.sin(theta);
  const along = Math.hypot(cos * sx, sin * sy);
  const across = Math.hypot(sin * sx, cos * sy) * Math.sign(sy);

  // Advances along the baseline (device px) and perpendicular offsets.
  const pdy = (out.options & ETO_PDY) !== 0;
  const extraPx = s.charExtra * along;
  const breakPx = s.breakCount > 0 ? (s.breakExtra / s.breakCount) * along : 0;
  const spaced = out.dx !== null || extraPx !== 0 || breakPx !== 0;
  const advances: number[] = [];
  const rises: number[] = [];
  const measure = (unit: TextUnit) => {
    ctx.font = fontFor(unit);
    return ctx.measureText(unit.text).width * hScale;
  };
  for (const unit of units) {
    let advance = 0;
    let rise = 0;
    if (out.dx) {
      for (let b = unit.start; b < unit.start + unit.count; b++) {
        advance += out.dx[pdy ? b * 2 : b] * along;
        if (pdy) rise += out.dx[b * 2 + 1] * across;
      }
    } else {
      advance = measure(unit);
    }
    if (spaced) {
      advance += extraPx;
      if (unit.count === 1 && out.bytes[unit.start] === 0x20) advance += breakPx;
    }
    advances.push(advance);
    rises.push(rise);
  }

  // Draw groups: one Canvas call per run of same-face units when the natural
  // advances apply, else one call per character at its own pen position.
  interface Group { text: string; font: string; pen: number; rise: number; advance: number }
  const groups: Group[] = [];
  let pen = 0;
  let rise = 0;
  units.forEach((unit, index) => {
    const unitFont = fontFor(unit);
    const last = groups.at(-1);
    if (!spaced && last && last.font === unitFont) {
      last.text += unit.text;
      last.advance += advances[index];
    } else {
      groups.push({ text: unit.text, font: unitFont, pen, rise, advance: advances[index] });
    }
    pen += advances[index];
    rise += rises[index];
  });
  if (!spaced) {
    // Shape each same-face run as a whole (complex scripts, no kerning).
    let groupPen = 0;
    for (const group of groups) {
      ctx.font = group.font;
      group.pen = groupPen;
      group.advance = ctx.measureText(group.text).width * hScale;
      groupPen += group.advance;
    }
    pen = groupPen;
  }
  const total = pen;
  const totalRise = rise;

  const useCp = (s.textAlign & TA_UPDATECP) !== 0;
  const refX = useCp ? s.currentX : out.x;
  const refY = useCp ? s.currentY : out.y;
  const originX = s.vpX + (refX - s.orgX) * sx;
  const originY = s.vpY + (refY - s.orgY) * sy;
  const x0 = horizontal === TA_RIGHT ? -total : horizontal === TA_CENTER ? -total / 2 : 0;
  const baseline = vertical === TA_BASELINE ? 0 : vertical === TA_BOTTOM ? -descentPx : ascentPx;
  const rtl = (s.textAlign & TA_RTLREADING) !== 0 || (out.options & ETO_RTLREADING) !== 0;
  const decoration = (font.underline || font.strikeout)
    ? referenceFontDecoration(font.face, weight, font.italic ? 'italic' : 'normal')
    : null;

  ctx.save();
  try {
    if (out.rect && (out.options & (ETO_OPAQUE | ETO_CLIPPED))) {
      const [l, t, r, b] = out.rect;
      const ax = s.vpX + (l - s.orgX) * sx; const bx = s.vpX + (r - s.orgX) * sx;
      const ay = s.vpY + (t - s.orgY) * sy; const by = s.vpY + (b - s.orgY) * sy;
      const rx = Math.min(ax, bx); const ry = Math.min(ay, by);
      const rw = Math.abs(bx - ax); const rh = Math.abs(by - ay);
      if (out.options & ETO_OPAQUE) {
        ctx.fillStyle = s.bkColor;
        ctx.fillRect(rx, ry, rw, rh);
      }
      if (out.options & ETO_CLIPPED) {
        ctx.beginPath();
        ctx.rect(rx, ry, rw, rh);
        ctx.clip();
      }
    }
    // Local frame: x along the baseline, y down, origin at the reference point.
    let ox = originX;
    let oy = originY;
    if (theta !== 0) {
      ctx.translate(originX, originY);
      ctx.rotate(-theta);
      ox = 0;
      oy = 0;
    }
    if (s.bkMode === OPAQUE && total > 0) {
      ctx.fillStyle = s.bkColor;
      ctx.fillRect(ox + x0, oy + baseline - ascentPx, total, ascentPx + descentPx);
    }
    ctx.fillStyle = s.textColor;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'alphabetic';
    ctx.direction = rtl ? 'rtl' : 'ltr';
    // GDI TextOut/ExtTextOut apply no pair kerning.
    if ('fontKerning' in ctx) ctx.fontKerning = 'none';
    for (const group of groups) {
      if (!group.text) continue;
      const gx = ox + x0 + (rtl ? total - group.pen - group.advance : group.pen);
      const gy = oy + baseline + group.rise;
      ctx.font = group.font;
      if (hScale === 1) {
        ctx.fillText(group.text, gx, gy);
      } else {
        ctx.save();
        ctx.translate(gx, gy);
        ctx.scale(hScale, 1);
        ctx.fillText(group.text, 0, 0);
        ctx.restore();
      }
    }
    if (total > 0) {
      const thickness = decoration ? decoration.underlineThicknessEm * emPx : Math.max(1, emPx * DECORATION_LINE_WIDTH_EM);
      if (font.underline) {
        const top = decoration ? decoration.underlineTopEm * emPx : Math.max(2, thickness);
        ctx.fillRect(ox + x0, oy + baseline + top, total, thickness);
      }
      if (font.strikeout) {
        const strike = decoration ? decoration.strikeThicknessEm * emPx : thickness;
        const top = decoration ? decoration.strikeTopEm * emPx : -STRIKE_RISE_EM * emPx - strike / 2;
        ctx.fillRect(ox + x0, oy + baseline + top, total, strike);
      }
    }
  } finally {
    ctx.restore();
  }

  if (useCp && horizontal !== TA_CENTER) {
    // TA_LEFT moves the current position past the string; TA_RIGHT moves it to
    // the string's start; TA_CENTER leaves it unchanged.
    const signed = horizontal === TA_RIGHT ? -total : total;
    const dxDev = signed * cos + (horizontal === TA_RIGHT ? 0 : totalRise * sin);
    const dyDev = -signed * sin + (horizontal === TA_RIGHT ? 0 : totalRise * cos);
    s.currentX = refX + dxDev / sx;
    s.currentY = refY + dyDev / sy;
  }
  return { drew: units.length > 0 };
}
