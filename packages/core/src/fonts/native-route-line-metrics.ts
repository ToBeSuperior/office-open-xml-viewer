import { retainFace, releaseFaces } from './font-registry.js';
import { withFontCeiling } from './preload.js';
import type { CanvasFontBoxProbeContext } from './canvas-font-box.js';
import { quoteCssFontFamily } from '../internal/css-font-family.js';

/** One CSS family list exactly as a format's resolver hands it to Canvas, with
 * the requested weight/style tuple. */
export interface NativeRouteLineRequest {
  readonly familyList: string;
  readonly weight: number;
  readonly style: 'normal' | 'italic';
}

/** The normal line of a native route's first available face, measured in the
 * DOM realm. Plain data: it crosses the worker boundary and the layout realm
 * re-verifies it ({@link verifyNativeRouteNormalLine}) before admitting it. */
export interface NativeRouteNormalLine extends NativeRouteLineRequest {
  /** Index in the family list of the first available family. */
  readonly primaryIndex: number;
  /** That family's name (unquoted). */
  readonly primaryFamily: string;
  /** The family list up to and including the first available family. */
  readonly primaryFamilyList: string;
  /** Probe em whose measurement scaled linearly. */
  readonly emPx: number;
  readonly ascentRatio: number;
  readonly descentRatio: number;
  /** CSS normal line minus the Canvas font box; signed. */
  readonly lineGapRatio: number;
}

/** Format policy: the probe em for an identified first available family, or
 * undefined to decline measuring that family. */
export type NativeRoutePrimaryEm = (
  primaryFamily: string, weight: number, style: 'normal' | 'italic',
) => number | undefined;

// ── Route sentinels ─────────────────────────────────────────────────────────
//
// A CSS family list names faces by family, and a family that is not installed
// is skipped silently. Neither which entry Canvas uses as the first available
// face nor which face draws a given cluster is observable by name. Two
// generated registrations make both observable. Each maps every Unicode scalar
// (cmap format 13: many-to-one ranges, the format of a last-resort font) to one
// glyph, and they differ only in their advance and font box:
//
//   sentinel A: advance 0.50 em, box 0.70 / 0.20 em
//   sentinel B: advance 0.75 em, box 0.60 / 0.30 em
//
// Appending a sentinel to a family-list prefix ends font fallback there: every
// cluster the prefix cannot draw is drawn by the sentinel, never by a later
// family or a system fallback face.
//
// - First available family: `prefix, A` and `prefix, B` report the same Canvas
//   font box exactly when some family of the prefix is available (the box is
//   the first available face's); otherwise they report the two sentinel
//   boxes. The first available family is the shortest prefix whose boxes agree.
// - Ownership: with the prefix ending at that family, a cluster is drawn by the
//   prefix's faces in both `prefix, A` and `prefix, B`, or (partly) by a
//   sentinel in each. Equal advances therefore prove that no sentinel glyph,
//   i.e. no face after the prefix, drew any part of the cluster; any sentinel
//   glyph changes the advance by 0.25 em. Families before the first available
//   one are system families that matched nothing: callers decline a route
//   when one of them is a registered FontFace (a unicode-range resource can
//   draw some clusters without being the first available face).
//
// This proves which faces of the list draw a cluster without a coverage table,
// a family-name guess or size-adjust (WebKit keeps ink extents under
// size-adjust, so the size-adjust oracle of local-line-metrics declines there).
// Both registrations validate themselves before use; an engine that cannot
// load or measure them declines.

const SENTINELS = [
  { family: '__ooxml_route_sentinel_a_6d2e91__', advance: 500, ascent: 700, descent: 200 },
  { family: '__ooxml_route_sentinel_b_6d2e91__', advance: 750, ascent: 600, descent: 300 },
] as const;
const SENTINEL_UPM = 1000;
const SENTINEL_FAMILIES = SENTINELS.map((sentinel) => sentinel.family);
const PROBE_EM = 100;

/** Minimal TrueType face: .notdef plus one box glyph that every scalar maps to. */
function sentinelFontBytes(advance: number, ascent: number, descent: number, name: string): Uint8Array {
  const u16 = (value: number) => [(value >> 8) & 255, value & 255];
  const i16 = (value: number) => u16(value & 0xffff);
  const u32 = (value: number) => [(value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255];
  const zeros = (count: number) => new Array<number>(count).fill(0);
  const inkTop = ascent >> 1;
  const glyph = [...i16(1), ...i16(0), ...i16(0), ...i16(advance), ...i16(inkTop), ...u16(3), ...u16(0), 1, 1, 1, 1];
  const points = [[0, 0], [advance, 0], [advance, inkTop], [0, inkTop]];
  const xs: number[] = [];
  const ys: number[] = [];
  let previousX = 0;
  let previousY = 0;
  for (const [x, y] of points) {
    xs.push(...i16(x! - previousX));
    ys.push(...i16(y! - previousY));
    previousX = x!;
    previousY = y!;
  }
  glyph.push(...xs, ...ys);
  while (glyph.length % 4) glyph.push(0);
  const nameRecords: number[] = [];
  const nameStorage: number[] = [];
  for (const [id, value] of [[1, name], [2, 'Regular'], [4, name], [6, name]] as const) {
    const bytes = [...value].flatMap((character) => u16(character.charCodeAt(0)));
    nameRecords.push(...u16(3), ...u16(1), ...u16(0x409), ...u16(id), ...u16(bytes.length), ...u16(nameStorage.length));
    nameStorage.push(...bytes);
  }
  // Surrogate code points are not scalars; every other value maps to glyph 1.
  const groups = [[0x0000, 0xd7ff], [0xe000, 0x10ffff]] as const;
  const cmap13 = [...u16(13), ...u16(0), ...u32(16 + 12 * groups.length), ...u32(0), ...u32(groups.length)];
  for (const [start, end] of groups) cmap13.push(...u32(start), ...u32(end), ...u32(1));
  const tables: Record<string, number[]> = {
    'OS/2': [
      ...u16(4), ...i16(advance), ...u16(400), ...u16(5), ...u16(0), ...zeros(20),
      ...i16(50), ...i16(250), ...i16(0), ...zeros(10), ...zeros(16), 0x4f, 0x4f, 0x58, 0x4d,
      ...u16(0x00c0), ...u16(0x20), ...u16(0xffff),
      ...i16(ascent), ...i16(-descent), ...i16(0), ...u16(ascent), ...u16(descent),
      ...u32(1), ...u32(0), ...i16(inkTop), ...i16(inkTop), ...u16(0), ...u16(0x20), ...u16(1),
    ],
    cmap: [...u16(0), ...u16(1), ...u16(3), ...u16(10), ...u32(12), ...cmap13],
    glyf: glyph,
    head: [
      ...u32(0x00010000), ...u32(0x00010000), ...u32(0), ...u32(0x5f0f3cf5), ...u16(0x000b),
      ...u16(SENTINEL_UPM), ...zeros(16), ...i16(0), ...i16(0), ...i16(advance), ...i16(inkTop),
      ...u16(0), ...u16(8), ...i16(2), ...i16(0), ...i16(0),
    ],
    hhea: [
      ...u32(0x00010000), ...i16(ascent), ...i16(-descent), ...i16(0), ...u16(advance),
      ...i16(0), ...i16(0), ...i16(advance), ...i16(1), ...i16(0), ...i16(0), ...zeros(8), ...i16(0), ...u16(2),
    ],
    hmtx: [...u16(advance), ...i16(0), ...u16(advance), ...i16(0)],
    loca: [...u16(0), ...u16(0), ...u16(glyph.length / 2)],
    maxp: [...u32(0x00010000), ...u16(2), ...u16(4), ...u16(1), ...zeros(4), ...u16(2), ...zeros(16)],
    name: [...u16(0), ...u16(4), ...u16(6 + nameRecords.length), ...nameRecords, ...nameStorage],
    post: [...u32(0x00030000), ...zeros(4), ...i16(-100), ...i16(50), ...zeros(20)],
  };
  const tags = Object.keys(tables).sort();
  let selector = 0;
  while ((1 << (selector + 1)) <= tags.length) selector += 1;
  const searchRange = (1 << selector) * 16;
  const checksum = (bytes: ArrayLike<number>) => {
    let sum = 0;
    for (let index = 0; index < bytes.length; index += 4) {
      sum = (sum + ((((bytes[index] ?? 0) << 24) >>> 0) + ((bytes[index + 1] ?? 0) << 16)
        + ((bytes[index + 2] ?? 0) << 8) + (bytes[index + 3] ?? 0))) >>> 0;
    }
    return sum;
  };
  const directory: number[] = [];
  const body: number[] = [];
  let offset = 12 + 16 * tags.length;
  let headOffset = 0;
  for (const tag of tags) {
    const data = tables[tag]!;
    if (tag === 'head') headOffset = offset;
    directory.push(...[...tag].map((character) => character.charCodeAt(0)),
      ...u32(checksum(data)), ...u32(offset), ...u32(data.length));
    body.push(...data);
    while (body.length % 4) body.push(0);
    offset = 12 + 16 * tags.length + body.length;
  }
  const font = new Uint8Array([
    ...u32(0x00010000), ...u16(tags.length), ...u16(searchRange), ...u16(selector),
    ...u16(tags.length * 16 - searchRange), ...directory, ...body,
  ]);
  font.set(u32((0xb1b0afba - checksum(font)) >>> 0), headOffset + 8);
  return font;
}

function newProbeContext(owner?: Document | null): CanvasFontBoxProbeContext | null {
  if (owner) return owner.createElement('canvas').getContext('2d');
  return typeof OffscreenCanvas !== 'undefined'
    ? new OffscreenCanvas(1, 1).getContext('2d')
    : typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;
}

/** The sentinel geometry is a design contract: verify that this realm loaded
 * and applies both registrations before any result depends on them. */
function sentinelsWork(context: CanvasFontBoxProbeContext): boolean {
  try {
    return SENTINELS.every((sentinel) => {
      context.font = `${PROBE_EM}px ${quoteCssFontFamily(sentinel.family)}`;
      const measured = context.measureText('A');
      const scale = PROBE_EM / SENTINEL_UPM;
      return Math.abs(measured.width - sentinel.advance * scale) < 0.01
        && Math.abs(measured.fontBoundingBoxAscent - sentinel.ascent * scale) < 0.01
        && Math.abs(measured.fontBoundingBoxDescent - sentinel.descent * scale) < 0.01;
    });
  } catch {
    return false;
  }
}

/** Register both sentinels in `fontSet` (refcounted) and validate them in the
 * realm whose canvases use that set. Returns the retained faces for the
 * caller's release list, or null (nothing retained) when unavailable. */
export async function loadNativeRouteSentinels(fontSet: FontFaceSet | null): Promise<FontFace[] | null> {
  if (!fontSet || typeof FontFace === 'undefined') return null;
  const faces: FontFace[] = [];
  try {
    for (const sentinel of SENTINELS) {
      faces.push(retainFace(`native-route-sentinel:${sentinel.family}`, fontSet, () => {
        const face = new FontFace(sentinel.family, sentinelFontBytes(
          sentinel.advance, sentinel.ascent, sentinel.descent, sentinel.family.replace(/_/g, ''),
        ) as BufferSource);
        fontSet.add(face);
        return face;
      }).face);
    }
    await withFontCeiling(Promise.all(faces.map((face) => face.load())));
    const owner = typeof document !== 'undefined' && document.fonts === fontSet ? document : null;
    const context = newProbeContext(owner);
    if (faces.every((face) => face.status === 'loaded') && context && sentinelsWork(context)) return faces;
  } catch { /* Unavailable sentinels decline every native route. */ }
  releaseFaces(faces);
  return null;
}

// ── Family lists ────────────────────────────────────────────────────────────

const GENERIC_FAMILIES = new Set([
  'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui',
  'ui-serif', 'ui-sans-serif', 'ui-monospace', 'ui-rounded', 'math', 'emoji', 'fangsong',
]);

interface FamilyEntry {
  /** Source text of the entry, as written in the list. */
  readonly text: string;
  /** Family name (quotes and escapes removed). */
  readonly name: string;
  readonly generic: boolean;
}

/** Split a CSS font-family list into entries, honouring quoted strings and
 * escapes (CSS Syntax 3 §4.3.5). Unquoted multi-word names collapse their
 * whitespace (CSS Fonts 4 §2.1). Returns null for malformed lists. */
function parseFamilyList(list: string): FamilyEntry[] | null {
  const entries: FamilyEntry[] = [];
  let index = 0;
  while (index < list.length) {
    while (index < list.length && /\s/.test(list[index]!)) index += 1;
    const start = index;
    let name = '';
    let quoted = false;
    if (list[index] === '"' || list[index] === "'") {
      quoted = true;
      const quote = list[index]!;
      index += 1;
      let closed = false;
      while (index < list.length) {
        const character = list[index]!;
        if (character === '\\') {
          const hex = /^[0-9a-fA-F]{1,6}\s?/.exec(list.slice(index + 1));
          if (hex) {
            name += String.fromCodePoint(Math.min(parseInt(hex[0].trim(), 16) || 0xfffd, 0x10ffff));
            index += 1 + hex[0].length;
          } else {
            name += list[index + 1] ?? '';
            index += 2;
          }
          continue;
        }
        if (character === quote) { closed = true; index += 1; break; }
        name += character;
        index += 1;
      }
      if (!closed) return null;
      while (index < list.length && /\s/.test(list[index]!)) index += 1;
    } else {
      while (index < list.length && list[index] !== ',') index += 1;
      name = list.slice(start, index).trim().replace(/\s+/g, ' ');
    }
    const text = list.slice(start, index).trim();
    if (!name || (index < list.length && list[index] !== ',')) return null;
    entries.push({ text, name, generic: !quoted && GENERIC_FAMILIES.has(name.toLowerCase()) });
    index += 1;
  }
  return entries.length > 0 ? entries : null;
}

function sameFamilyList(a: readonly FamilyEntry[], b: readonly FamilyEntry[]): boolean {
  return a.length === b.length && a.every((entry, index) =>
    entry.name === b[index]!.name && entry.generic === b[index]!.generic);
}

function registeredFamilies(fontSet: FontFaceSet | null | undefined): Set<string> {
  const names = new Set<string>();
  if (!fontSet || typeof fontSet[Symbol.iterator] !== 'function') return names;
  for (const face of fontSet) {
    const family = parseFamilyList(face.family)?.[0]?.name ?? face.family;
    names.add(family.trim().toLowerCase());
  }
  return names;
}

/** Shortest prefix whose first available face is not a sentinel, judged by
 * the Canvas font box (see the sentinel notes above). */
function primaryIndex(
  context: CanvasFontBoxProbeContext,
  entries: readonly FamilyEntry[],
  weight: number,
  style: 'normal' | 'italic',
  memo: Map<string, boolean>,
): number | undefined {
  let prefix = '';
  for (let index = 0; index < entries.length; index += 1) {
    prefix = index === 0 ? entries[0]!.text : `${prefix}, ${entries[index]!.text}`;
    const key = `${style} ${weight} ${prefix}`;
    let available = memo.get(key);
    if (available === undefined) {
      const boxes = SENTINEL_FAMILIES.map((family) => {
        context.font = `${style} ${weight} ${PROBE_EM}px ${prefix}, ${quoteCssFontFamily(family)}`;
        const box = context.measureText('');
        return [box.fontBoundingBoxAscent, box.fontBoundingBoxDescent] as const;
      });
      available = boxes[0]![0] === boxes[1]![0] && boxes[0]![1] === boxes[1]![1];
      memo.set(key, available);
    }
    if (available) return index;
  }
  return undefined;
}

/** First available family of a route in this realm, or undefined when the
 * route cannot be judged: a generic keyword (an unnamed face), or a registered
 * FontFace at or before it (an application or web resource whose unicode-range
 * coverage, not installation, decides which clusters it draws). */
function identifiedPrimary(
  context: CanvasFontBoxProbeContext,
  entries: readonly FamilyEntry[],
  weight: number,
  style: 'normal' | 'italic',
  registered: ReadonlySet<string>,
  memo: Map<string, boolean>,
): number | undefined {
  const index = primaryIndex(context, entries, weight, style, memo);
  if (index === undefined || entries[index]!.generic) return undefined;
  for (let entry = 0; entry <= index; entry += 1) {
    if (registered.has(entries[entry]!.name.toLowerCase())) return undefined;
  }
  return index;
}

interface NormalBox { readonly height: number; readonly ascent: number; readonly descent: number }

/** offsetHeight is an integer CSS pixel and Canvas boxes carry sub-pixel
 * quantization (same bounds as the exact-local measurement). */
function scalesLinearly(full: NormalBox, half: NormalBox): boolean {
  return Math.abs(full.height - 2 * half.height) <= 3
    && Math.abs(full.ascent - 2 * half.ascent) <= 1
    && Math.abs(full.descent - 2 * half.descent) <= 1;
}

const LINEAR_PROBE_ATTEMPTS = 4;

/**
 * Measure the CSS normal line of each route's first available face, in the
 * document realm, before layout. Canvas exposes the font box but no leading;
 * a `line-height: normal` strut (U+200B, a closed shadow tree, computed family
 * list checked entry by entry) supplies the missing leading, exactly as the
 * exact-local measurement does for one face. Every strut of a batch is
 * appended once and read once, so a document costs one style/layout pass per
 * linear-scale attempt, never one per route or per segment.
 *
 * Requires the sentinels to be loaded in `owner.fonts`. A request whose first
 * available family cannot be identified, that `primaryEm` declines, whose
 * computed list differs, or whose box does not scale linearly yields null.
 */
export function measureNativeRouteNormalLines(
  requests: readonly NativeRouteLineRequest[],
  owner: Document,
  primaryEm: NativeRoutePrimaryEm,
): (NativeRouteNormalLine | null)[] {
  const results: (NativeRouteNormalLine | null)[] = requests.map(() => null);
  const context = owner.body ? newProbeContext(owner) : null;
  const view = owner.defaultView;
  if (!context || !view || !sentinelsWork(context)) return results;
  const registered = registeredFamilies(owner.fonts);
  const memo = new Map<string, boolean>();
  interface Pending {
    readonly request: number;
    readonly entries: FamilyEntry[];
    readonly primary: number;
    em: number;
  }
  let pending: Pending[] = [];
  requests.forEach((request, index) => {
    const entries = parseFamilyList(request.familyList);
    if (!entries) return;
    const primary = identifiedPrimary(context, entries, request.weight, request.style, registered, memo);
    if (primary === undefined) return;
    const em = primaryEm(entries[primary]!.name, request.weight, request.style);
    if (em === undefined || !Number.isFinite(em) || em <= 0) return;
    pending.push({ request: index, entries, primary, em });
  });
  for (let attempt = 0; attempt < LINEAR_PROBE_ATTEMPTS && pending.length > 0; attempt += 1) {
    const host = owner.createElement('div');
    host.style.cssText = 'all:initial;position:fixed;visibility:hidden;contain:strict;';
    if (typeof host.attachShadow !== 'function') return results;
    const root = host.attachShadow({ mode: 'closed' });
    const struts = pending.map((item) => [item.em, item.em / 2].map((em) => {
      const request = requests[item.request]!;
      const strut = owner.createElement('span');
      strut.style.cssText = `all:initial;display:block;white-space:pre;font:${request.style} ${request.weight} ${em}px ${request.familyList};line-height:normal;`;
      strut.textContent = '​';
      root.append(strut);
      return strut;
    }));
    const next: Pending[] = [];
    owner.body.append(host);
    try {
      pending.forEach((item, index) => {
        const request = requests[item.request]!;
        const boxes = struts[index]!.map((strut, half): NormalBox | null => {
          const computed = parseFamilyList(view.getComputedStyle(strut).fontFamily);
          if (!computed || !sameFamilyList(computed, item.entries)) return null;
          const em = half ? item.em / 2 : item.em;
          context.font = `${request.style} ${request.weight} ${em}px ${request.familyList}`;
          const box = context.measureText('');
          const height = strut.offsetHeight;
          return Number.isFinite(height) && height > 0
            && Number.isFinite(box.fontBoundingBoxAscent) && box.fontBoundingBoxAscent >= 0
            && Number.isFinite(box.fontBoundingBoxDescent) && box.fontBoundingBoxDescent >= 0
            ? { height, ascent: box.fontBoundingBoxAscent, descent: box.fontBoundingBoxDescent }
            : null;
        });
        const [full, half] = boxes;
        if (!full || !half) return;
        if (!scalesLinearly(full, half)) {
          // Engines may clamp large sizes: halve and retry, else decline.
          item.em /= 2;
          next.push(item);
          return;
        }
        results[item.request] = Object.freeze({
          familyList: request.familyList,
          weight: request.weight,
          style: request.style,
          primaryIndex: item.primary,
          primaryFamily: item.entries[item.primary]!.name,
          primaryFamilyList: item.entries.slice(0, item.primary + 1).map((entry) => entry.text).join(', '),
          emPx: item.em,
          ascentRatio: full.ascent / item.em,
          descentRatio: full.descent / item.em,
          lineGapRatio: (full.height - full.ascent - full.descent) / item.em,
        });
      });
    } catch {
      // An unavailable measurement declines the remaining routes.
      next.length = 0;
    } finally {
      host.remove();
    }
    pending = next;
  }
  return results;
}

/** Admission in the realm that lays out (a worker has no DOM, so its lines are
 * measured by the owning document). Each route must have the same first
 * available family here, identified under the same rules, with the same
 * Canvas font box at the measured em; otherwise this realm draws with a
 * different face and the measurement does not describe it. */
export function verifyNativeRouteNormalLines(
  lines: readonly NativeRouteNormalLine[],
  fontSet: FontFaceSet | null,
): boolean[] {
  const context = newProbeContext();
  if (!context || !sentinelsWork(context)) return lines.map(() => false);
  const registered = registeredFamilies(fontSet);
  const memo = new Map<string, boolean>();
  return lines.map((line) => {
    const entries = parseFamilyList(line.familyList);
    if (!entries || entries[line.primaryIndex]?.name !== line.primaryFamily) return false;
    if (identifiedPrimary(context, entries, line.weight, line.style, registered, memo) !== line.primaryIndex) {
      return false;
    }
    context.font = `${line.style} ${line.weight} ${line.emPx}px ${line.familyList}`;
    const box = context.measureText('');
    return Math.abs(box.fontBoundingBoxAscent - line.ascentRatio * line.emPx) < 1e-3
      && Math.abs(box.fontBoundingBoxDescent - line.descentRatio * line.emPx) < 1e-3;
  });
}

// ── Ownership ───────────────────────────────────────────────────────────────

/** The two font strings whose advances are compared for one route: the route
 * truncated after its first available family, then one sentinel. */
export function nativeRouteOwnershipFonts(
  primaryFamilyList: string, weight: number, style: 'normal' | 'italic',
): readonly [string, string] {
  return [
    `${style} ${weight} ${PROBE_EM}px ${primaryFamilyList}, ${quoteCssFontFamily(SENTINEL_FAMILIES[0])}`,
    `${style} ${weight} ${PROBE_EM}px ${primaryFamilyList}, ${quoteCssFontFamily(SENTINEL_FAMILIES[1])}`,
  ];
}

const ownershipContexts: (CanvasFontBoxProbeContext | null)[] = [];
const ownershipFonts: (string | undefined)[] = [];

/** Whether the route's first available family (or an absent family before
 * it) draws every part of `cluster`; undefined when the realm cannot judge.
 * Each sentinel has its own private context, so a font string is parsed only
 * when the route changes. */
export function nativeRouteClusterOwned(
  fonts: readonly [string, string], cluster: string,
): boolean | undefined {
  const widths: number[] = [];
  for (let index = 0; index < 2; index += 1) {
    ownershipContexts[index] ??= newProbeContext();
    const context = ownershipContexts[index];
    if (!context) return undefined;
    if (ownershipFonts[index] !== fonts[index]) {
      context.font = fonts[index]!;
      ownershipFonts[index] = fonts[index];
    }
    widths.push(context.measureText(cluster).width);
  }
  return Number.isFinite(widths[0]) && widths[0] === widths[1];
}

/** The sentinels are registered and applied in this realm (once per prepared
 * route; an unavailable sentinel would let fallback faces pass as owned). */
export function nativeRouteSentinelsWork(): boolean {
  const context = newProbeContext();
  return context !== null && sentinelsWork(context);
}
