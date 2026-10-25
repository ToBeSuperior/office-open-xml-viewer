/**
 * CSS font sources: `css-api` (one stylesheet per served family) and
 * `font-face-css` (one stylesheet describing many families), plus the
 * hardened `@font-face` parser and bounded stylesheet fetch they share.
 *
 * Stylesheets are untrusted input (a third-party CDN, a mirror, an internal
 * service). The parser is a single linear scan over CSS Syntax Level 3
 * structure (comments, strings, url tokens and `()`/`[]`/`{}` blocks; no
 * recursion, so nesting depth costs nothing) and interprets only the
 * `font-family`, `src`, `font-weight`, `font-style`, `font-stretch` and
 * `unicode-range` descriptors of top-level `@font-face` rules. Everything
 * else (other at-rules, qualified rules, `@font-face` nested in conditional
 * rules, `!important` or block-valued descriptors) is skipped. Library policy
 * on top of CSS Fonts Level 4:
 *
 * - only `url()` sources are accepted: `local()` would let a stylesheet probe
 *   or paint host fonts, which is the host-fallback path's job. The first
 *   `url()` whose `format()` is absent or one of woff2 / woff / truetype /
 *   opentype is used; entries with `tech()` are skipped (support cannot be
 *   evaluated here);
 * - relative URLs resolve against the stylesheet's final URL after redirects
 *   (#1497); results must be http(s) / blob: / data:, and `data:` URLs are
 *   charged their decoded size against the face ceiling (PR 1 policy);
 * - an invalid `unicode-range` is dropped as CSS drops an invalid declaration,
 *   so the face covers every scalar and is loaded whenever its family is;
 * - an `@font-face` block left unterminated at end of input is dropped rather
 *   than closed (a truncated stylesheet must not register a partial rule);
 * - the stylesheet is bounded by `HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES` while
 *   it is read and its rules by `HARD_MAX_FONT_SOURCE_DECLARATIONS`; both
 *   reject with `OoxmlError('font-source-limit')`. The body is decoded as
 *   UTF-8 (`@charset` and `Content-Type` charset are not interpreted).
 */
import { fontSourceLimitError } from '../errors/font-source-limit.js';
import { OoxmlError } from '../errors/ooxml-error.js';
import {
  HARD_MAX_FONT_SOURCE_DECLARATIONS,
  HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES,
} from '../worker/resource-policy.generated.js';
import { claimFromIndex, indexFontDeclarations } from './files-source.js';
import type { FamilyResolution, FontSourceFailure, FontStylesheet, SourceCache } from './source-cache.js';
import {
  cssApiRequests,
  validateFontFaceDeclaration,
  type CssApiSource,
  type FontFaceCssSource,
  type FontFaceDeclaration,
  type FontSourceCredentials,
} from './source.js';
import { fontFamilyKey } from './substitutions.js';

// ---- unicode-range ------------------------------------------------------------------

/**
 * Parse a CSS `unicode-range` value (CSS Fonts 4 §4.5: comma-separated
 * `U+hex`, `U+hex-hex` or `U+hex?` wildcard). Returns null unless every
 * component is recognised and within U+0–U+10FFFF with start ≤ end: a range
 * is an exclusion hint, never a support certificate, so an unparseable value
 * means "covers everything".
 */
export function parseUnicodeRange(value: string | undefined): readonly (readonly [number, number])[] | null {
  if (!value?.trim()) return null;
  const ranges: [number, number][] = [];
  for (const component of value.split(',')) {
    const part = component.trim();
    let lo: number, hi: number;
    const span = /^U\+([0-9A-F]{1,6})(?:-([0-9A-F]{1,6}))?$/i.exec(part);
    const wildcard = /^U\+([0-9A-F]{0,5})(\?{1,6})$/i.exec(part);
    if (span) {
      lo = parseInt(span[1], 16);
      hi = parseInt(span[2] ?? span[1], 16);
    } else if (wildcard && wildcard[1].length + wildcard[2].length <= 6) {
      lo = parseInt(wildcard[1] + '0'.repeat(wildcard[2].length), 16);
      hi = parseInt(wildcard[1] + 'F'.repeat(wildcard[2].length), 16);
    } else {
      return null;
    }
    if (lo > hi || hi > 0x10ffff) return null;
    ranges.push([lo, hi]);
  }
  return ranges;
}

// ---- CSS scanning -----------------------------------------------------------------

const isWhitespace = (unit: number) => unit === 0x20 || unit === 0x09 || unit === 0x0a || unit === 0x0c || unit === 0x0d;
const isNameUnit = (unit: number) =>
  (unit >= 0x30 && unit <= 0x39) || (unit >= 0x41 && unit <= 0x5a) || (unit >= 0x61 && unit <= 0x7a)
  || unit === 0x2d || unit === 0x5f || unit >= 0x80;

/** Index after a comment starting at `at` (`/*`), or the end of input. */
function skipComment(css: string, at: number): number {
  const end = css.indexOf('*/', at + 2);
  return end < 0 ? css.length : end + 2;
}

/** Index after a string starting at the quote at `at`. A string ends at its
 * unescaped quote, at an unescaped newline (bad string) or at end of input. */
function skipString(css: string, at: number): number {
  const quote = css.charCodeAt(at);
  let index = at + 1;
  while (index < css.length) {
    const unit = css.charCodeAt(index);
    if (unit === quote) return index + 1;
    if (unit === 0x0a || unit === 0x0c || unit === 0x0d) return index;
    index += unit === 0x5c /* \ */ ? 2 : 1;
  }
  return css.length;
}

/** When `at` starts an unquoted `url(` token, the index after it; else -1.
 * A quoted `url("…")` is an ordinary function and returns -1. */
function skipUnquotedUrl(css: string, at: number): number {
  if (css.slice(at, at + 4).toLowerCase() !== 'url(') return -1;
  if (at > 0 && (isNameUnit(css.charCodeAt(at - 1)) || css.charCodeAt(at - 1) === 0x5c)) return -1;
  let index = at + 4;
  while (index < css.length && isWhitespace(css.charCodeAt(index))) index++;
  const first = css.charCodeAt(index);
  if (first === 0x22 || first === 0x27) return -1;
  while (index < css.length) {
    const unit = css.charCodeAt(index);
    if (unit === 0x29 /* ) */) return index + 1;
    index += unit === 0x5c ? 2 : 1;
  }
  return css.length;
}

const CLOSER: Readonly<Record<string, string>> = { '{': '}', '(': ')', '[': ']' };

/**
 * Scan component values from `at` until one of `stops` appears outside any
 * block, string, comment or url token. Returns the stop's index, or the end of
 * input. Blocks are tracked with an explicit stack of expected closers (a
 * mismatched closer is an ordinary token, as in CSS Syntax §5.4.8).
 */
function scanUntil(css: string, at: number, stops: string): number {
  const expected: string[] = [];
  let index = at;
  while (index < css.length) {
    const char = css[index];
    const unit = css.charCodeAt(index);
    if (char === '/' && css[index + 1] === '*') { index = skipComment(css, index); continue; }
    if (unit === 0x22 || unit === 0x27) { index = skipString(css, index); continue; }
    if (unit === 0x5c) { index += 2; continue; }
    if (char === 'u' || char === 'U') {
      const after = skipUnquotedUrl(css, index);
      if (after >= 0) { index = after; continue; }
    }
    if (expected.length === 0 && stops.includes(char)) return index;
    if (char === '{' || char === '(' || char === '[') expected.push(CLOSER[char]);
    else if (expected.length && char === expected[expected.length - 1]) expected.pop();
    index++;
  }
  return css.length;
}

/** Index of the `}` closing the block opened at `at`, or -1 at end of input. */
function blockEnd(css: string, at: number): number {
  const end = scanUntil(css, at + 1, '}');
  return end < css.length ? end : -1;
}

/** `text` without comments outside strings and url tokens. */
function stripComments(text: string): string {
  let out = '';
  let index = 0;
  while (index < text.length) {
    const unit = text.charCodeAt(index);
    if (text[index] === '/' && text[index + 1] === '*') { index = skipComment(text, index); out += ' '; continue; }
    let next = index + 1;
    if (unit === 0x22 || unit === 0x27) next = skipString(text, index);
    else if (unit === 0x5c) next = index + 2;
    else if (text[index] === 'u' || text[index] === 'U') {
      const after = skipUnquotedUrl(text, index);
      if (after >= 0) next = after;
    }
    out += text.slice(index, next);
    index = next;
  }
  return out;
}

/** Split `text` at top-level occurrences of `separator`. */
function splitTopLevel(text: string, separator: string): string[] {
  const parts: string[] = [];
  let start = 0;
  while (start <= text.length) {
    const end = scanUntil(text, start, separator);
    parts.push(text.slice(start, end));
    start = end + 1;
  }
  return parts;
}

const isNewline = (char: string | undefined) => char === '\n' || char === '\r' || char === '\f';
const isHexDigit = (unit: number) =>
  (unit >= 0x30 && unit <= 0x39) || (unit >= 0x41 && unit <= 0x46) || (unit >= 0x61 && unit <= 0x66);
const REPLACEMENT = String.fromCharCode(0xfffd);

/**
 * Consume an escaped code point (CSS Syntax §4.3.7); `at` is the backslash,
 * already known to start a valid escape (not followed by a newline). One to
 * six hex digits are followed by one optional white space, which belongs to
 * the escape (CR LF counts as one, as input preprocessing makes it one LF);
 * zero, a surrogate or a value beyond U+10FFFF is U+FFFD. A backslash at end
 * of input is U+FFFD; otherwise the next code point stands for itself (NUL is
 * U+FFFD after preprocessing).
 */
function consumeEscape(text: string, at: number): { value: string; end: number } {
  let index = at + 1;
  if (index >= text.length) return { value: REPLACEMENT, end: index };
  let digits = 0;
  while (digits < 6 && index + digits < text.length && isHexDigit(text.charCodeAt(index + digits))) digits++;
  if (digits === 0) {
    const code = text.codePointAt(index) as number;
    return { value: code === 0 ? REPLACEMENT : String.fromCodePoint(code), end: index + (code > 0xffff ? 2 : 1) };
  }
  const code = parseInt(text.slice(index, index + digits), 16);
  index += digits;
  if (text.startsWith('\r\n', index)) index += 2;
  else if (index < text.length && isWhitespace(text.charCodeAt(index))) index++;
  const value = code === 0 || (code >= 0xd800 && code <= 0xdfff) || code > 0x10ffff
    ? REPLACEMENT : String.fromCodePoint(code);
  return { value, end: index };
}

/** Decode the escapes of a CSS string's contents: an escaped newline is a line
 * continuation (removed) and a backslash at end of input contributes nothing
 * (CSS Syntax §4.3.5). */
function unescapeCssString(text: string): string {
  if (!text.includes('\\')) return text;
  let out = '';
  let index = 0;
  while (index < text.length) {
    const char = text[index];
    if (char !== '\\') { out += char; index++; continue; }
    const next = text[index + 1];
    if (next === undefined) break;
    if (isNewline(next)) { index += text.startsWith('\r\n', index + 1) ? 3 : 2; continue; }
    const escape = consumeEscape(text, index);
    out += escape.value;
    index = escape.end;
  }
  return out;
}

/** The contents of `text` when it is exactly one terminated CSS string
 * (quotes included), or null. */
function cssString(text: string): string | null {
  const quote = text[0];
  if (quote !== '"' && quote !== "'") return null;
  let index = 1;
  while (index < text.length) {
    const char = text[index];
    if (char === quote) return index === text.length - 1 ? unescapeCssString(text.slice(1, -1)) : null;
    if (char === '\n' || char === '\f' || char === '\r') return null;
    index += char === '\\' ? 2 : 1;
  }
  return null;
}

/** Reserved words a `<family-name>` cannot be spelled as unquoted (CSS-wide
 * keywords, `default`, and the generic families). */
const RESERVED_FAMILY_IDENTS = new Set([
  'inherit', 'initial', 'unset', 'revert', 'revert-layer', 'default',
  'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'math', 'emoji', 'fangsong',
]);

const isNameStart = (char: string | undefined) =>
  char !== undefined && (/[A-Za-z_]/.test(char) || char.charCodeAt(0) >= 0x80);
const startsValidEscape = (text: string, at: number) => text[at] === '\\' && !isNewline(text[at + 1]);

/** Whether an identifier starts at `at` (CSS Syntax §4.3.9). */
function startsIdentifier(text: string, at: number): boolean {
  if (text[at] === '-') {
    return text[at + 1] === '-' || isNameStart(text[at + 1]) || startsValidEscape(text, at + 1);
  }
  return isNameStart(text[at]) || startsValidEscape(text, at);
}

/** The decoded identifiers of a white-space separated identifier sequence,
 * tokenized as CSS does (the white space ending a hex escape belongs to the
 * escape), or null when anything else appears. */
function identifierSequence(text: string): string[] | null {
  const idents: string[] = [];
  let index = 0;
  while (index < text.length) {
    if (isWhitespace(text.charCodeAt(index))) { index++; continue; }
    if (!startsIdentifier(text, index)) return null;
    let ident = '';
    while (index < text.length) {
      if (startsValidEscape(text, index)) {
        const escape = consumeEscape(text, index);
        ident += escape.value;
        index = escape.end;
      } else if (isNameUnit(text.charCodeAt(index))) {
        ident += text[index++];
      } else {
        break;
      }
    }
    if (index < text.length && !isWhitespace(text.charCodeAt(index))) return null;
    idents.push(ident);
  }
  return idents;
}

/** `font-family` descriptor: one string, or identifiers joined by single spaces. */
function parseFamily(value: string): string | null {
  const text = value.trim();
  if (text.startsWith('"') || text.startsWith("'")) {
    const family = cssString(text)?.trim();
    return family ? family : null;
  }
  const idents = identifierSequence(text);
  if (!idents?.length) return null;
  if (idents.length === 1 && RESERVED_FAMILY_IDENTS.has(idents[0].toLowerCase())) return null;
  const family = idents.join(' ').trim();
  return family || null;
}

const SUPPORTED_FORMATS = new Set(['woff2', 'woff', 'truetype', 'opentype']);

/** One CSS function at the start of `text`: its lower-cased name, argument
 * text and the remainder, or null. */
function leadingFunction(text: string): { name: string; args: string; rest: string } | null {
  const match = /^([A-Za-z-]+)\(/.exec(text);
  if (!match) return null;
  const close = scanUntil(text, match[0].length, ')');
  if (close >= text.length) return null;
  return { name: match[1].toLowerCase(), args: text.slice(match[0].length, close), rest: text.slice(close + 1).trim() };
}

/** The URL of one `src` entry when it is an acceptable `url()`, else null. */
function srcEntryUrl(entry: string): string | null {
  let text = entry.trim();
  let url: string | null = null;
  const unquotedEnd = skipUnquotedUrl(text, 0);
  if (unquotedEnd >= 0) {
    const raw = text.slice(4, unquotedEnd - 1).trim();
    // A url token may not contain quotes, `(`, white space or non-printables.
    if (unquotedEnd === text.length && text.at(-1) !== ')') return null;
    // Escaped code points are part of the URL.
    if (/["'(\s\u0000-\u0008\u000b\u000e-\u001f\u007f]/.test(raw.replace(/\\[^\n\r\f]/g, ''))) return null;
    url = unescapeCssString(raw);
    text = text.slice(unquotedEnd).trim();
  } else {
    const call = leadingFunction(text);
    if (!call || call.name !== 'url') return null;
    url = cssString(call.args.trim());
    text = call.rest;
  }
  if (!url) return null;
  while (text) {
    const call = leadingFunction(text);
    if (!call) return null;
    if (call.name === 'tech') return null;
    if (call.name !== 'format') return null;
    const formats = call.args.split(',').map((argument) => {
      const trimmed = argument.trim();
      return (cssString(trimmed) ?? trimmed).toLowerCase();
    });
    if (!formats.some((format) => SUPPORTED_FORMATS.has(format))) return null;
    text = call.rest;
  }
  return url;
}

const DESCRIPTOR_VALUE = /^[A-Za-z0-9.+\-% ]+$/;

/** `font-weight` / `font-style` / `font-stretch`: keywords, numbers, angles
 * and percentages only; the browser validates the grammar itself. */
function plainDescriptor(value: string): string | null {
  const text = value.trim().replace(/[ \t\n\f\r]+/g, ' ');
  return text && DESCRIPTOR_VALUE.test(text) ? text : null;
}

/** Declarations of one `@font-face` body; null when the rule is unusable. */
function parseFontFaceBody(body: string, stylesheetUrl: string, location: string): FontFaceDeclaration | null {
  let family: string | null = null;
  let src: string | null = null;
  const descriptors: { weight?: string; style?: string; stretch?: string; unicodeRange?: string } = {};
  for (const raw of splitTopLevel(body, ';')) {
    const declaration = stripComments(raw);
    const colon = scanUntil(declaration, 0, ':{');
    if (colon >= declaration.length || declaration[colon] !== ':') continue;
    const name = declaration.slice(0, colon).trim().toLowerCase();
    const value = declaration.slice(colon + 1);
    // A block-valued or !important descriptor is invalid in @font-face.
    if (scanUntil(value, 0, '{') < value.length || /!\s*important\s*$/i.test(value)) continue;
    switch (name) {
      case 'font-family': family = parseFamily(value) ?? family; break;
      case 'src': {
        const first = splitTopLevel(value, ',').map(srcEntryUrl).find((url) => url !== null);
        if (first) src = first;
        break;
      }
      case 'font-weight': descriptors.weight = plainDescriptor(value) ?? descriptors.weight; break;
      case 'font-style': descriptors.style = plainDescriptor(value) ?? descriptors.style; break;
      case 'font-stretch': descriptors.stretch = plainDescriptor(value) ?? descriptors.stretch; break;
      case 'unicode-range': {
        const text = value.trim();
        if (parseUnicodeRange(text)) descriptors.unicodeRange = text;
        break;
      }
      default: break;
    }
  }
  if (!family || !src) return null;
  let resolved = src;
  if (!/^\s*data:/i.test(src)) {
    try {
      resolved = new URL(src, stylesheetUrl).href;
    } catch {
      return null;
    }
  }
  try {
    return validateFontFaceDeclaration({ family, url: resolved, ...descriptors }, location);
  } catch (error) {
    // A limit breach (oversized URL or data: face) fails closed; any other
    // invalid rule is dropped like an invalid CSS declaration.
    if (error instanceof OoxmlError) throw error;
    return null;
  }
}

function utf8ByteLength(text: string): number {
  let bytes = 0;
  for (let index = 0; index < text.length; index++) {
    const unit = text.charCodeAt(index);
    if (unit < 0x80) bytes += 1;
    else if (unit < 0x800) bytes += 2;
    else if (unit >= 0xd800 && unit <= 0xdbff && index + 1 < text.length
      && text.charCodeAt(index + 1) >= 0xdc00 && text.charCodeAt(index + 1) <= 0xdfff) { bytes += 4; index++; }
    else bytes += 3;
  }
  return bytes;
}

/**
 * Parse the `@font-face` rules of an untrusted stylesheet into validated
 * declarations, in stylesheet order. `stylesheetUrl` is the absolute final
 * URL the stylesheet was served from (relative font URLs resolve against it).
 */
export function parseFontFaceCss(css: string, stylesheetUrl: string): FontFaceDeclaration[] {
  const size = utf8ByteLength(css);
  if (size > HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES) {
    throw fontSourceLimitError('stylesheet', HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES, size);
  }
  const declarations: FontFaceDeclaration[] = [];
  let index = 0;
  while (index < css.length) {
    const unit = css.charCodeAt(index);
    if (isWhitespace(unit)) { index++; continue; }
    if (css[index] === '/' && css[index + 1] === '*') { index = skipComment(css, index); continue; }
    if (css.startsWith('<!--', index)) { index += 4; continue; }
    if (css.startsWith('-->', index)) { index += 3; continue; }
    const isAtRule = unit === 0x40;
    let nameEnd = index + 1;
    while (isAtRule && nameEnd < css.length && isNameUnit(css.charCodeAt(nameEnd))) nameEnd++;
    // An at-rule prelude ends at `;` or `{`; a qualified rule's only at `{`.
    const stop = scanUntil(css, isAtRule ? nameEnd : index, isAtRule ? ';{' : '{');
    if (stop >= css.length) break;
    if (css[stop] === ';') { index = stop + 1; continue; }
    const end = blockEnd(css, stop);
    const isFontFace = isAtRule && css.slice(index + 1, nameEnd).toLowerCase() === 'font-face'
      && stripComments(css.slice(nameEnd, stop)).trim() === '';
    if (isFontFace && end >= 0) {
      const declaration = parseFontFaceBody(css.slice(stop + 1, end), stylesheetUrl,
        `stylesheet @font-face #${declarations.length + 1}`);
      if (declaration) {
        if (declarations.length >= HARD_MAX_FONT_SOURCE_DECLARATIONS) {
          throw fontSourceLimitError('declarations', HARD_MAX_FONT_SOURCE_DECLARATIONS, declarations.length + 1);
        }
        declarations.push(declaration);
      }
    }
    if (end < 0) break;
    index = end + 1;
  }
  return declarations;
}

// ---- bounded stylesheet fetch ---------------------------------------------------

interface ByteReader {
  read(): Promise<{ done: boolean; value?: Uint8Array }>;
  cancel(reason?: unknown): Promise<void>;
}

interface StylesheetResponse {
  readonly ok: boolean;
  readonly status?: number;
  readonly url?: string;
  readonly headers?: { get(name: string): string | null };
  readonly body?: { getReader(): ByteReader; cancel?(reason?: unknown): Promise<void> } | null;
  text(): Promise<string>;
}

/** Cancel an unread response body (the caller also aborts the request), so a
 * rejected response stops downloading. */
function discardBody(response: StylesheetResponse): void {
  try {
    response.body?.cancel?.().catch(() => undefined);
  } catch {
    /* a locked or consumed body: the request abort still stops it */
  }
}

async function readBoundedText(response: StylesheetResponse): Promise<string> {
  const limit = HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES;
  const declared = Number(response.headers?.get('content-length') ?? NaN);
  if (Number.isSafeInteger(declared) && declared > limit) {
    discardBody(response);
    throw fontSourceLimitError('stylesheet', limit, declared);
  }
  const reader = typeof response.body?.getReader === 'function' ? response.body.getReader() : null;
  if (!reader) {
    const text = await response.text();
    const size = utf8ByteLength(text);
    if (size > limit) throw fontSourceLimitError('stylesheet', limit, size);
    return text;
  }
  const decoder = new TextDecoder('utf-8');
  let size = 0;
  let text = '';
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    if (!value) continue;
    size += value.byteLength;
    if (size > limit) {
      reader.cancel().catch(() => undefined);
      throw fontSourceLimitError('stylesheet', limit, size);
    }
    text += decoder.decode(value, { stream: true });
  }
  return text + decoder.decode();
}

/**
 * Fetch and parse one stylesheet in the painting realm. Credentials are
 * `'same-origin'` unless the source opted into `'omit'`; never `'include'`,
 * never custom headers. A network error, a non-OK status, a final URL that is
 * not http(s) or a stylesheet without any usable rule is a failure (host
 * fallback); a limit breach rejects.
 */
export async function fetchFontStylesheet(
  url: string,
  credentials: FontSourceCredentials | undefined,
  signal: AbortSignal,
): Promise<FontStylesheet | FontSourceFailure> {
  if (typeof fetch !== 'function') return { ok: false, reason: 'fetch is unavailable' };
  // A request of our own, so every early rejection below aborts the download
  // even while the caller's signal stays live.
  const request = new AbortController();
  const onAbort = () => request.abort();
  signal.addEventListener('abort', onAbort, { once: true });
  if (signal.aborted) request.abort();
  let complete = false;
  let response: StylesheetResponse;
  let text: string;
  try {
    response = await fetch(url, { credentials: credentials ?? 'same-origin', signal: request.signal }) as StylesheetResponse;
    if (!response.ok) {
      discardBody(response);
      return { ok: false, reason: `HTTP ${response.status ?? 'error'}` };
    }
    text = await readBoundedText(response);
    complete = true;
  } catch (error) {
    if (error instanceof OoxmlError) throw error;
    if (signal.aborted) throw error;
    return { ok: false, reason: error instanceof Error ? error.message : String(error) };
  } finally {
    signal.removeEventListener('abort', onAbort);
    if (!complete) request.abort();
  }
  const finalUrl = response.url || url;
  if (!/^https?:/i.test(finalUrl)) return { ok: false, reason: 'stylesheet redirected off http(s)' };
  const declarations = parseFontFaceCss(text, finalUrl);
  if (declarations.length === 0) return { ok: false, reason: 'no usable @font-face rule' };
  return { ok: true, declarations, finalUrl };
}

function joinStylesheet(
  cache: SourceCache,
  url: string,
  credentials: FontSourceCredentials | undefined,
  signal: AbortSignal,
): Promise<FontStylesheet | FontSourceFailure> {
  return cache.stylesheets.join(url, (taskSignal) => fetchFontStylesheet(url, credentials, taskSignal), signal);
}

// ---- adapters -----------------------------------------------------------------------

/**
 * `css-api`: one request per served family (the demanded family, then each
 * declared substitute's `use`, filtered by the allow-list), issued in
 * parallel. The best-precedence request that yields rules of its served family
 * claims the demanded family; when every request fails the family is final
 * (failed); when the allow-list admits no candidate the family is absent.
 */
async function resolveCssApiFamily(
  source: CssApiSource,
  cache: SourceCache,
  family: string,
  signal: AbortSignal,
): Promise<FamilyResolution> {
  const requests = cssApiRequests(source, family);
  if (requests.length === 0) return { kind: 'absent' };
  const sheets = await Promise.all(requests.map((request) =>
    joinStylesheet(cache, request.url, source.credentials, signal)));
  const reasons: string[] = [];
  for (let index = 0; index < requests.length; index++) {
    const request = requests[index];
    const sheet = sheets[index];
    if (!sheet.ok) { reasons.push(sheet.reason); continue; }
    // A per-family stylesheet answers only for the family it was requested for.
    const key = fontFamilyKey(request.family);
    const declarations = sheet.declarations.filter((declaration) => fontFamilyKey(declaration.family) === key);
    if (declarations.length === 0) { reasons.push(`no @font-face rule for ${request.family}`); continue; }
    return {
      kind: 'claimed',
      served: declarations[0].family,
      ...(request.substitution ? { substitution: request.substitution } : {}),
      declarations,
    };
  }
  return { kind: 'failed', reason: reasons.join('; ') };
}

/** `font-face-css`: fetch the stylesheet once per source identity, then answer
 * every family from its index. A failed stylesheet fails every family asked of
 * it (no other source is consulted). */
async function resolveFontFaceCssFamily(
  source: FontFaceCssSource,
  cache: SourceCache,
  family: string,
  signal: AbortSignal,
): Promise<FamilyResolution> {
  const sheet = await joinStylesheet(cache, source.url, source.credentials, signal);
  if (!sheet.ok) return { kind: 'failed', reason: sheet.reason };
  cache.index ??= indexFontDeclarations(sheet.declarations);
  return claimFromIndex(cache.index, family, source.substitutes);
}

/** Resolve demanded families against a CSS source through the shared
 * resolution cache (one in-flight resolution per source identity × family). */
export async function resolveCssSource(
  source: CssApiSource | FontFaceCssSource,
  cache: SourceCache,
  families: readonly string[],
  signal: AbortSignal,
): Promise<ReadonlyMap<string, FamilyResolution>> {
  const entries = await Promise.all(families.map(async (family) => {
    const key = fontFamilyKey(family);
    const resolution = await cache.families.join(key, (taskSignal) => source.kind === 'css-api'
      ? resolveCssApiFamily(source, cache, family, taskSignal)
      : resolveFontFaceCssFamily(source, cache, family, taskSignal), signal);
    return [key, resolution] as const;
  }));
  return new Map(entries);
}
