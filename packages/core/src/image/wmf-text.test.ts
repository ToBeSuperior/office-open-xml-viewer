import { describe, expect, it, vi } from 'vitest';
import { playWmf } from './wmf.js';
import { decodeRasterOrMetafile, getIncompleteMetafileReport } from './raster-or-metafile.js';

// Synthetic MS-WMF byte fixtures for META_TEXTOUT / META_EXTTEXTOUT and the
// DC text state (MS-WMF 2.3.3.5, 2.3.3.20, 2.2.1.2, 2.3.5.x).

const u16 = (value: number) => [value & 0xff, (value >>> 8) & 0xff];
const u32 = (value: number) => [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff];

function record(fn: number, params: number[]): Uint8Array {
  return new Uint8Array([...u32(3 + params.length / 2), ...u16(fn), ...params]);
}

function join(...parts: Uint8Array[]): Uint8Array {
  const output = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) { output.set(part, offset); offset += part.length; }
  return output;
}

const header = () => new Uint8Array([...u16(1), ...u16(9), ...u16(0x300), ...u32(0), ...u16(8), ...u32(0), ...u16(0)]);
const eof = () => record(0, []);
const windowExt = (x = 100, y = 100) => join(record(0x020b, [...u16(0), ...u16(0)]), record(0x020c, [...u16(y), ...u16(x)]));
const bkMode = (mode: number) => record(0x0102, u16(mode));
const textAlign = (align: number) => record(0x012e, u16(align));
const select = (index: number) => record(0x012d, u16(index));
const moveTo = (x: number, y: number) => record(0x0214, [...u16(y), ...u16(x)]);

interface FontSpec {
  height?: number; width?: number; escapement?: number; orientation?: number; weight?: number;
  italic?: boolean; underline?: boolean; strikeout?: boolean; charset?: number; pitchAndFamily?: number;
  face?: string; faceBytes?: number[]; fullFaceField?: boolean;
}

/** META_CREATEFONTINDIRECT. MS-WMF 2.2.1.2 FaceName is NUL-terminated and at
 *  most 32 bytes; writers may store only the used bytes (fullFaceField false). */
function font(spec: FontSpec = {}): Uint8Array {
  const face = spec.faceBytes ?? [...Buffer.from(spec.face ?? 'Test Face', 'latin1')];
  const name = [...face, 0];
  if (spec.fullFaceField !== false) while (name.length < 32) name.push(0);
  if (name.length & 1) name.push(0);
  return record(0x02fb, [
    ...u16(spec.height ?? -20), ...u16(spec.width ?? 0), ...u16(spec.escapement ?? 0),
    ...u16(spec.orientation ?? 0), ...u16(spec.weight ?? 400),
    spec.italic ? 1 : 0, spec.underline ? 1 : 0, spec.strikeout ? 1 : 0, spec.charset ?? 0,
    0, 0, 0, spec.pitchAndFamily ?? 0,
    ...name,
  ]);
}

function extText(
  x: number, y: number, bytes: number[],
  { options = 0, rect, dx = [] }: { options?: number; rect?: [number, number, number, number]; dx?: number[] } = {},
): Uint8Array {
  return record(0x0a32, [
    ...u16(y), ...u16(x), ...u16(bytes.length), ...u16(options),
    ...(rect ? rect.flatMap(u16) : []),
    ...bytes, ...(bytes.length & 1 ? [0] : []),
    ...dx.flatMap(u16),
  ]);
}

function textOut(x: number, y: number, bytes: number[]): Uint8Array {
  return record(0x0521, [...u16(bytes.length), ...bytes, ...(bytes.length & 1 ? [0] : []), ...u16(y), ...u16(x)]);
}

const ascii = (text: string) => [...Buffer.from(text, 'latin1')];

/** A 1:1 logical-to-device picture, transparent background, one font in slot 0. */
function picture(fontRecord: Uint8Array, ...tail: Uint8Array[]): Uint8Array {
  return join(header(), windowExt(), bkMode(1), fontRecord, select(0), textAlign(0x18), ...tail, eof());
}

interface Op { op: string; args: unknown[]; font: string; fill: string; direction: string }

/** Deterministic Canvas: every glyph advances 0.5 em; the font box is 0.9/0.3
 *  em. Families listed as installed draw distinct glyph geometry. */
function context(installed: readonly string[] = []) {
  const ops: Op[] = [];
  const state = { font: '10px serif', fillStyle: '#000000', direction: 'inherit' };
  const px = () => Number(/([\d.]+)px/.exec(state.font)?.[1] ?? 10);
  const log = (op: string) => (...args: unknown[]) => {
    ops.push({ op, args, font: state.font, fill: String(state.fillStyle), direction: state.direction });
  };
  const target: Record<string, unknown> = {
    get font() { return state.font; },
    set font(value: string) { state.font = value; },
    get fillStyle() { return state.fillStyle; },
    set fillStyle(value: string) { state.fillStyle = value; },
    get direction() { return state.direction; },
    set direction(value: string) { state.direction = value; },
    textAlign: 'start', textBaseline: 'alphabetic', fontKerning: 'auto',
    measureText(text: string) {
      const size = px();
      const covered = installed.some((family) => state.font.includes(`"${family}"`));
      return {
        width: [...text].length * size * 0.5,
        fontBoundingBoxAscent: size * 0.9,
        fontBoundingBoxDescent: size * 0.3,
        actualBoundingBoxAscent: size * (covered ? 0.7 : 0.6),
        actualBoundingBoxDescent: 0, actualBoundingBoxLeft: 0, actualBoundingBoxRight: 0,
      };
    },
  };
  for (const name of ['fillText', 'fillRect', 'rect', 'clip', 'translate', 'rotate', 'scale', 'save', 'restore', 'beginPath']) {
    target[name] = log(name);
  }
  const ctx = target as unknown as CanvasRenderingContext2D;
  const texts = () => ops.filter((op) => op.op === 'fillText');
  return { ctx, ops, texts };
}

function play(file: Uint8Array, installed: readonly string[] = []) {
  const mock = context(installed);
  const report = vi.fn();
  const drew = playWmf(file, mock.ctx, 100, 100, report);
  return { ...mock, report, drew, unsupported: (report.mock.calls[0]?.[0] ?? []) as string[] };
}

describe('WMF text: Font object and DC state records', () => {
  it('draws after META_SETMAPPERFLAGS and a FaceName stored without its full 32-byte field', () => {
    // MathType-style records: SETMAPPERFLAGS only affects font matching, and
    // a 7-byte "Symbol\0" FaceName is a complete Font object.
    const file = join(
      header(), windowExt(), bkMode(1), textAlign(0x18),
      record(0x0231, u32(1)),
      font({ charset: 2, face: 'Symbol', fullFaceField: false }), select(0),
      extText(10, 50, [0x6c]),
      eof(),
    );
    const run = play(file);
    expect(run.unsupported).toEqual([]);
    expect(run.texts().map((op) => op.args)).toEqual([['λ', 10, 50]]);
  });

  it('maps a negative Height to the em and a positive Height to the cell (ascent + descent)', () => {
    const negative = play(picture(font({ height: -24 }), extText(0, 50, ascii('A'))));
    expect(negative.texts()[0].font).toBe('400 24px "Test Face", sans-serif');
    // Mock face box: 0.9 + 0.3 em, so a 24-unit cell is a 20px em.
    const positive = play(picture(font({ height: 24 }), extText(0, 50, ascii('A'))));
    expect(positive.texts()[0].font).toBe('400 20px "Test Face", sans-serif');
    const zero = play(picture(font({ height: 0 }), extText(0, 50, ascii('A'))));
    expect(zero.texts()).toEqual([]);
    expect(zero.unsupported).toEqual(['META_EXTTEXTOUT (default font height)']);
  });

  it('takes the generic family from PitchAndFamily and keeps the authored face first', () => {
    const cases: Array<[number, string]> = [[0x10, 'serif'], [0x20, 'sans-serif'], [0x30, 'monospace'], [0x40, 'cursive'], [0x50, 'fantasy'], [0x01, 'monospace']];
    for (const [pitchAndFamily, generic] of cases) {
      const run = play(picture(font({ pitchAndFamily, weight: 700, italic: true }), extText(0, 50, ascii('A'))));
      expect(run.texts()[0].font).toBe(`italic 700 20px "Test Face", ${generic}`);
    }
    // Office faces with a different-family web substitute keep it as the next entry.
    const calibri = play(picture(font({ face: 'Calibri', pitchAndFamily: 0x20 }), extText(0, 50, ascii('A'))));
    expect(calibri.texts()[0].font).toBe('400 20px "Calibri", "Carlito", sans-serif');
  });

  it('uses catalogued usWinAscent/usWinDescent for the authored face', () => {
    // Times New Roman: unitsPerEm 2048, usWinAscent 1825, usWinDescent 443.
    const top = play(picture(font({ face: 'Times New Roman', height: -2048 }), textAlign(0), extText(0, 0, ascii('A'))));
    expect(top.texts()[0].args.slice(1)).toEqual([0, 1825]);
    const bottom = play(picture(font({ face: 'Times New Roman', height: -2048 }), textAlign(8), extText(0, 0, ascii('A'))));
    expect(bottom.texts()[0].args.slice(1)).toEqual([0, -443]);
  });

  it.each([
    [0x18, 'left/baseline', 40, 50],
    [0x1a, 'right/baseline', 20, 50],
    [0x1e, 'center/baseline', 30, 50],
    [0x00, 'left/top', 40, 68],
    [0x08, 'left/bottom', 40, 44],
    [0x0e, 'center/bottom', 30, 44],
    [0x02, 'right/top', 20, 68],
  ])('aligns TextAlignmentMode %#x (%s) on the reference point', (align, _name, x, y) => {
    // 'AB' at 20px is 20px wide; the mock box is 18px ascent, 6px descent.
    const run = play(picture(font(), textAlign(align), extText(40, 50, ascii('AB'))));
    expect(run.texts().map((op) => op.args)).toEqual([['AB', x, y]]);
  });

  it.each([0x04, 0x10, 0x1c])('reports the undefined alignment code %#x', (align) => {
    const run = play(picture(font(), textAlign(align), extText(40, 50, ascii('A'))));
    expect(run.texts()).toEqual([]);
    expect(run.unsupported).toEqual(['META_EXTTEXTOUT (undefined text alignment)']);
  });

  it('draws from and advances the current position under TA_UPDATECP', () => {
    const run = play(picture(font(),
      moveTo(10, 50), textAlign(0x19),
      textOut(0, 0, ascii('AB')), // 10..30, CP → 30
      extText(0, 0, ascii('C')), // 30..40, CP → 40
      textAlign(0x1b), textOut(0, 0, ascii('D')), // right-aligned at 40: 30..40, CP → 30
      textAlign(0x1f), textOut(0, 0, ascii('EF')), // centered at 30: 20..40, CP unchanged
      textAlign(0x19), textOut(0, 0, ascii('G')), // 30..40
    ));
    expect(run.texts().map((op) => op.args)).toEqual([
      ['AB', 10, 50], ['C', 30, 50], ['D', 30, 50], ['EF', 20, 50], ['G', 30, 50],
    ]);
  });

  it('reports TA_UPDATECP without an established current position', () => {
    const run = play(picture(font(), textAlign(0x19), textOut(5, 5, ascii('A'))));
    expect(run.texts()).toEqual([]);
    expect(run.unsupported).toEqual(['META_TEXTOUT (current position undefined)']);
  });

  it('adds SETTEXTCHAREXTRA per character and SETTEXTJUSTIFICATION BreakExtra per break character', () => {
    const run = play(picture(font(),
      record(0x0108, u16(2)), // CharExtra 2
      record(0x020a, [...u16(2), ...u16(8)]), // BreakCount 2, BreakExtra 8 → 4 per space
      extText(0, 50, ascii('A B')),
    ));
    // Advances: A 10+2, space 10+2+4, B.
    expect(run.texts().map((op) => op.args)).toEqual([['A', 0, 50], [' ', 12, 50], ['B', 28, 50]]);
  });

  it('applies the text color, and paints the cell in OPAQUE background mode', () => {
    const run = play(join(header(), windowExt(), font(), select(0), textAlign(0x18),
      record(0x0209, u32(0x0000ff)), record(0x0201, u32(0x00ff00)), bkMode(2),
      extText(10, 50, ascii('AB')), eof()));
    expect(run.ops.find((op) => op.op === 'fillRect')).toMatchObject({ args: [10, 32, 20, 24], fill: '#00ff00' });
    expect(run.texts()[0]).toMatchObject({ args: ['AB', 10, 50], fill: '#ff0000' });
  });

  it('defaults to the Win32 OPAQUE mode with a white background', () => {
    const run = play(join(header(), windowExt(), font(), select(0), textAlign(0x18), extText(10, 50, ascii('A')), eof()));
    expect(run.ops.find((op) => op.op === 'fillRect')).toMatchObject({ args: [10, 32, 10, 24], fill: '#ffffff' });
  });

  it('keeps glyphs upright under a mirrored logical mapping (GM_COMPATIBLE)', () => {
    const file = join(header(), record(0x020b, [...u16(0), ...u16(0)]), record(0x020c, [...u16(-100), ...u16(100)]),
      bkMode(1), font(), select(0), textAlign(0x18), extText(10, -50, ascii('A')), eof());
    const run = play(file);
    expect(run.texts().map((op) => [op.args, op.font])).toEqual([[['A', 10, 50], '400 20px "Test Face", sans-serif']]);
    expect(run.ops.some((op) => op.op === 'scale' || op.op === 'rotate')).toBe(false);
  });

  it('rotates by Escapement (tenths of a degree, counterclockwise) and ignores Orientation', () => {
    const run = play(picture(font({ escapement: 900, orientation: 0 }), extText(30, 60, ascii('A'))));
    const transforms = run.ops.filter((op) => op.op === 'translate' || op.op === 'rotate').map((op) => [op.op, ...op.args]);
    expect(transforms).toEqual([['translate', 30, 60], ['rotate', -Math.PI / 2]]);
    expect(run.texts()[0].args).toEqual(['A', 0, 0]);
  });

  it('scales glyphs horizontally to a non-zero Width against the catalogued xAvgCharWidth', () => {
    // Times New Roman normal: xAvgCharWidth 821 / 2048 em.
    const run = play(picture(font({ face: 'Times New Roman', height: -2048, width: 1642 }), extText(0, 50, ascii('A'))));
    const scale = run.ops.find((op) => op.op === 'scale');
    expect(scale?.args[0]).toBeCloseTo(2);
    expect(scale?.args[1]).toBe(1);
  });

  it('draws Underline and StrikeOut from the catalogued post/OS/2 metrics', () => {
    // Times New Roman: underline -223/100, strikeout 530/102 design units.
    const run = play(picture(font({ face: 'Times New Roman', height: -2048, underline: true, strikeout: true }),
      extText(0, 3000, ascii('AB'))));
    const rects = run.ops.filter((op) => op.op === 'fillRect').map((op) => op.args);
    expect(rects).toEqual([[0, 3223, 2048, 100], [0, 2470, 2048, 102]]);
  });
});

describe('WMF text: META_EXTTEXTOUT options and Dx', () => {
  it('positions each character by its Dx entry', () => {
    const run = play(picture(font(), extText(5, 50, ascii('ABC'), { dx: [15, 7, 30] })));
    expect(run.texts().map((op) => op.args)).toEqual([['A', 5, 50], ['B', 20, 50], ['C', 27, 50]]);
  });

  it('sums the per-byte Dx entries of a double-byte character', () => {
    // Shift_JIS 0x82A0 'あ' then 'A': Dx has one entry per byte.
    const run = play(picture(font({ charset: 128 }), extText(0, 50, [0x82, 0xa0, 0x41], { dx: [12, 8, 9] })));
    expect(run.texts().map((op) => op.args)).toEqual([['あ', 0, 50], ['A', 20, 50]]);
    expect(run.texts()[0].font).toContain('"Noto Sans JP"');
  });

  it('applies interleaved vertical advances with ETO_PDY', () => {
    const run = play(picture(font(), extText(0, 50, ascii('AB'), { options: 0x2000, dx: [10, -5, 10, 0] })));
    expect(run.texts().map((op) => op.args)).toEqual([['A', 0, 50], ['B', 10, 45]]);
  });

  it('fills the ETO_OPAQUE rectangle with the background color and clips to ETO_CLIPPED', () => {
    const run = play(join(header(), windowExt(), bkMode(1), record(0x0201, u32(0xff0000)), font(), select(0), textAlign(0x18),
      extText(10, 50, ascii('A'), { options: 0x0006, rect: [5, 30, 60, 55] }), eof()));
    const ops = run.ops.filter((op) => ['fillRect', 'rect', 'clip', 'fillText'].includes(op.op));
    expect(ops.map((op) => [op.op, ...op.args])).toEqual([
      ['fillRect', 5, 30, 55, 25], ['rect', 5, 30, 55, 25], ['clip'], ['fillText', 'A', 10, 50],
    ]);
    expect(ops[0].fill).toBe('#0000ff');
  });

  it('lays out right-to-left reading order from the far edge', () => {
    for (const file of [
      picture(font(), extText(0, 50, ascii('AB'), { options: 0x0080, dx: [10, 30] })),
      picture(font(), textAlign(0x118), extText(0, 50, ascii('AB'), { dx: [10, 30] })),
    ]) {
      const run = play(file);
      expect(run.texts().map((op) => op.args)).toEqual([['A', 30, 50], ['B', 0, 50]]);
      expect(run.texts()[0].direction).toBe('rtl');
    }
  });

  it('reports glyph-index strings and malformed Dx arrays without drawing', () => {
    const glyphs = play(picture(font(), extText(0, 50, [1, 2], { options: 0x0010 })));
    expect(glyphs.unsupported).toEqual(['META_EXTTEXTOUT (glyph index string)']);
    const shortDx = play(picture(font(), extText(0, 50, ascii('AB'), { dx: [10] })));
    expect(shortDx.unsupported).toEqual(['META_EXTTEXTOUT (malformed Dx array)']);
    expect([...glyphs.texts(), ...shortDx.texts()]).toEqual([]);
  });

  it('rejects every in-bounds short EXTTEXTOUT parameter record', () => {
    const params = [...u16(24), ...u16(12), ...u16(3), ...u16(0), 65, 66, 67, 0];
    for (let length = 0; length < params.length; length += 2) {
      const run = play(picture(font(), record(0x0a32, params.slice(0, length))));
      expect(run.texts()).toEqual([]);
    }
  });
});

describe('WMF text: CharacterSet decoding', () => {
  it('decodes the complete Windows-1252 C1 index and keeps NUL as a character', () => {
    const bytes = [0x41, 0, 0x7f, ...Array.from({ length: 32 }, (_, index) => index + 0x80), 0xa0, 0xff];
    const expected = [
      0x41, 0, 0x7f,
      0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021,
      0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008d, 0x017d, 0x008f,
      0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
      0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
      0x00a0, 0x00ff,
    ];
    const run = play(picture(font(), extText(0, 50, bytes)));
    expect(Array.from(run.texts()[0].args[0] as string, (char) => char.codePointAt(0))).toEqual(expected);
  });

  it.each([
    [161, [0xe1, 0xe2], 'αβ'], // GREEK_CHARSET, code page 1253
    [204, [0xc4, 0xe0], 'Да'], // RUSSIAN_CHARSET, code page 1251
    [238, [0xb9, 0xe8], 'ąč'], // EASTEUROPE_CHARSET, code page 1250
    [177, [0xe0], 'א'], // HEBREW_CHARSET, code page 1255
    [77, [0x8a], 'ä'], // MAC_CHARSET, code page 10000
    [134, [0xc4, 0xe3], '你'], // GB2312_CHARSET, code page 936
    [1, ascii('Ab'), 'Ab'], // DEFAULT_CHARSET: ASCII is common to every code page
  ] as const)('decodes CharacterSet %i through its code page', (charset, bytes, text) => {
    const run = play(picture(font({ charset }), extText(0, 50, [...bytes])));
    expect(run.texts()[0].args[0]).toBe(text);
  });

  it.each([
    [1, [0xe9], 'locale-dependent character set'],
    [255, [0x80], 'locale-dependent character set'],
    [130, [0x88, 0x61], 'character set decoder unavailable'],
    [161, [0xd2], 'byte undefined in code page'],
    [128, [0x82], 'truncated double-byte character'],
    [3, [0x41], 'undefined character set'],
  ] as const)('reports CharacterSet %i bytes the code page leaves undefined', (charset, bytes, problem) => {
    const run = play(picture(font({ charset }), extText(0, 50, [...bytes])));
    expect(run.texts()).toEqual([]);
    expect(run.unsupported).toEqual([`META_EXTTEXTOUT (${problem})`]);
  });

  it('decodes a FaceName in the font code page', () => {
    // "ＭＳ 明朝" in Shift_JIS.
    const faceBytes = [0x82, 0x6c, 0x82, 0x72, 0x20, 0x96, 0xbe, 0x92, 0xa9];
    const run = play(picture(font({ charset: 128, faceBytes }), extText(0, 50, ascii('A'))));
    expect(run.texts()[0].font.startsWith('400 20px "ＭＳ 明朝"')).toBe(true);
  });
});

describe('WMF text: SYMBOL_CHARSET', () => {
  it('converts the published Symbol repertoire to Unicode in a fallback face', () => {
    const run = play(picture(font({ charset: 2, face: 'Symbol', italic: true, pitchAndFamily: 0x10 }),
      extText(0, 50, [0x6c, 0x6d, 0x73, 0x6a, 0x66, 0xd7, 0xf2, 0xe5, 0xa5, 0xb4, 0x2b, 0x2d, 0x3d])));
    expect(run.texts().map((op) => op.args[0])).toEqual(['λμσφϕ⋅∫∑∞×+−=']);
    expect(run.texts()[0].font).toBe('italic 400 20px serif');
  });

  it('decodes the published MT Extra ellipses', () => {
    const run = play(picture(font({ charset: 2, face: 'MT Extra' }), extText(0, 50, [0x4b, 0x4c, 0x4d, 0x4e, 0x4f])));
    expect(run.texts()[0].args[0]).toBe('…⋯⋮⋰⋱');
    expect(run.unsupported).toEqual([]);
  });

  it('keeps codes without a standard Unicode target font-coded in the authored face', () => {
    // 0xE6 (parenlefttp) maps only to an Adobe Corporate-Use code point.
    const run = play(picture(font({ charset: 2, face: 'Symbol' }), extText(0, 50, [0x61, 0xe6])));
    expect(run.texts().map((op) => [op.args[0], op.font])).toEqual([
      ['α', '400 20px serif'],
      ['', '400 20px "Symbol", sans-serif'],
    ]);
  });

  it('draws an installed symbol face through its F000 symbol cmap range', () => {
    const run = play(picture(font({ charset: 2, face: 'Private Symbols' }), extText(0, 50, [0x41, 0x42])), ['Private Symbols']);
    expect(run.texts().map((op) => [op.args[0], op.font])).toEqual([['', '400 20px "Private Symbols", sans-serif']]);
  });

  it('substitutes the Windows symbol-charset face for a symbol face the host lacks', () => {
    const run = play(picture(font({ charset: 2, face: 'Private Symbols', pitchAndFamily: 0x12 }), extText(0, 50, [0x31, 0x30, 0x61, 0xe6])));
    expect(run.texts().map((op) => [op.args[0], op.font])).toEqual([
      ['10α', '400 20px serif'],
      ['', '400 20px "Symbol", serif'],
    ]);
  });
});

describe('WMF text: undefined state and record bookkeeping', () => {
  it('reports text without a selected font, and vertical @ faces', () => {
    const none = play(join(header(), windowExt(), extText(0, 50, ascii('A')), eof()));
    expect(none.unsupported).toEqual(['META_EXTTEXTOUT (no font selected)']);
    const vertical = play(picture(font({ face: '@MS Mincho' }), extText(0, 50, ascii('A'))));
    expect(vertical.unsupported).toEqual(['META_EXTTEXTOUT (vertical font)']);
  });

  it('reports a FaceName without a NUL terminator inside its 32 bytes', () => {
    const unterminated = record(0x02fb, [...u16(-20), ...new Array(16).fill(0), ...new Array(32).fill(0x41)]);
    const run = play(picture(unterminated, extText(0, 50, ascii('A'))));
    expect(run.unsupported).toEqual(['META_EXTTEXTOUT (malformed font)']);
  });

  it('does not read a short CREATEFONTINDIRECT record from the next record', () => {
    for (const length of [0, 10, 16, 18]) {
      const run = play(picture(record(0x02fb, new Array(length).fill(0)), extText(0, 50, ascii('A'))));
      expect(run.texts()).toEqual([]);
    }
  });

  it('keeps drawing after state-neutral records but not after an unsupported state change', () => {
    const neutral = play(picture(font(), record(0x0418, [...u16(10), ...u16(10), ...u16(0), ...u16(0)]), extText(0, 50, ascii('A'))));
    expect(neutral.texts()).toHaveLength(1);
    const layout = play(picture(font(), record(0x0149, u32(1)), extText(0, 50, ascii('A'))));
    expect(layout.texts()).toEqual([]);
    expect(layout.unsupported).toEqual(['META_EXTTEXTOUT (text state undefined after an unsupported record)']);
  });

  it('gives palette and pattern-brush objects their object-table slots', () => {
    const palette = record(0x00f7, [...u16(0x300), ...u16(1), 0, 0, 0, 0]);
    const run = play(join(header(), windowExt(), bkMode(1), textAlign(0x18), palette,
      font({ height: -10 }), font({ height: -30 }), select(2), extText(0, 50, ascii('A')), eof()));
    expect(run.texts()[0].font).toBe('400 30px "Test Face", sans-serif');
  });

  it.each([[[1, 2], 1], [[0x57, 0x4d, 0x46, 0x43], 0]] as const)(
    'handles private comment %j (draws %i)',
    (data, draws) => {
      const comment = record(0x0626, [...u16(15), ...u16(data.length), ...data, ...(data.length & 1 ? [0] : [])]);
      const run = play(picture(font(), comment, extText(0, 50, ascii('A'))));
      expect(run.texts()).toHaveLength(draws);
    },
  );

  it('propagates a reported gap through bitmap decode and strict refusal', async () => {
    const file = picture(font(), extText(24, 24, ascii('A')), font({ height: 0 }), select(1), extText(0, 0, ascii('B')));
    const mock = context();
    const close = vi.fn();
    vi.stubGlobal('OffscreenCanvas', class {
      constructor(readonly width: number, readonly height: number) {}
      getContext() { return mock.ctx; }
    });
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ({ width: 200, height: 100, close })));
    try {
      const blob = () => new Blob([file as Uint8Array<ArrayBuffer>]);
      const bitmap = await decodeRasterOrMetafile(blob(), { widthPt: 100, heightPt: 50 });
      expect(mock.texts().map((op) => op.args[0])).toEqual(['A']);
      expect(getIncompleteMetafileReport(bitmap)).toEqual({
        format: 'wmf', unsupported: ['META_EXTTEXTOUT (default font height)'],
      });
      await expect(decodeRasterOrMetafile(blob(), { incompleteMetafile: 'reject' })).rejects.toMatchObject({
        code: 'ooxml-incomplete-metafile', format: 'wmf',
      });
      expect(close).toHaveBeenCalledOnce();
    } finally { vi.unstubAllGlobals(); }
  });
});
