import { afterEach, describe, expect, it, vi } from 'vitest';
import { OoxmlError } from '../errors/ooxml-error.js';
import {
  HARD_MAX_FONT_SOURCE_DECLARATIONS,
  HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES,
} from '../worker/resource-policy.generated.js';
import { fetchFontStylesheet, parseFontFaceCss, parseUnicodeRange } from './css-source.js';

const SHEET = 'https://cdn.internal.example/styles/office.css';

const face = (body: string) => `@font-face { ${body} }`;

function limitQuota(run: () => unknown): string | undefined {
  try {
    run();
  } catch (error) {
    if (error instanceof OoxmlError && error.code === 'font-source-limit') return error.violation?.quota;
    throw error;
  }
  return undefined;
}

describe('parseFontFaceCss', () => {
  it('reads machine-generated per-family CSS with every descriptor', () => {
    const css = `/* latin */
@font-face {
  font-family: 'Carlito';
  font-style: italic;
  font-weight: 700;
  font-stretch: 75% 125%;
  font-display: swap;
  src: url(https://fonts.gstatic.example/s/carlito/x.woff2) format('woff2');
  unicode-range: U+0000-00FF, U+0131, U+02??;
}`;
    expect(parseFontFaceCss(css, SHEET)).toEqual([{
      family: 'Carlito',
      url: 'https://fonts.gstatic.example/s/carlito/x.woff2',
      style: 'italic',
      weight: '700',
      stretch: '75% 125%',
      unicodeRange: 'U+0000-00FF, U+0131, U+02??',
    }]);
  });

  it.each([
    ['url(../files/a.woff2)', 'https://cdn.internal.example/files/a.woff2'],
    ["url('/files/a.woff2')", 'https://cdn.internal.example/files/a.woff2'],
    ['url( "a.woff2" )', 'https://cdn.internal.example/styles/a.woff2'],
    ['url(//assets.internal.example/a.woff2)', 'https://assets.internal.example/a.woff2'],
    ['url(a\\ b.woff2)', 'https://cdn.internal.example/styles/a%20b.woff2'],
  ])('resolves %s against the stylesheet URL', (src, expected) => {
    expect(parseFontFaceCss(face(`font-family: X; src: ${src};`), SHEET)[0]?.url).toBe(expected);
  });

  it('accepts only url() sources: local() skipped, first supported format chosen', () => {
    const [declaration] = parseFontFaceCss(face(`font-family: X; src:
      local("X Regular"), url(x.eot) format("embedded-opentype"), url(x.var.woff2) format(woff2) tech(variations),
      url(x.woff) format("woff"), url(x.ttf);`), SHEET);
    expect(declaration.url).toBe('https://cdn.internal.example/styles/x.woff');
    // local()-only rules register nothing.
    expect(parseFontFaceCss(face('font-family: X; src: local(X);'), SHEET)).toEqual([]);
  });

  it('keeps data: URLs verbatim (including unquoted ";base64,") and accepts blob:', () => {
    const data = 'data:font/woff2;base64,d09GMgABAAAAAA==';
    expect(parseFontFaceCss(face(`font-family: X; src: url(${data}) format("woff2");`), SHEET)[0].url).toBe(data);
    const blob = 'blob:https://app.example/1f0e';
    expect(parseFontFaceCss(face(`font-family: X; src: url("${blob}");`), SHEET)[0].url).toBe(blob);
  });

  it.each([
    'javascript:alert(1)',
    'file:///etc/fonts/x.ttf',
    'ftp://files.example/x.ttf',
    'chrome://x/y.woff',
  ])('drops a rule whose source resolves to %s', (url) => {
    expect(parseFontFaceCss(face(`font-family: X; src: url("${url}");`), SHEET)).toEqual([]);
  });

  it('interprets only top-level @font-face rules', () => {
    const css = `
@charset "utf-8";
@import url(other.css);
body { font-family: "A"; src: url(no.woff2); }
@media print { @font-face { font-family: Nested; src: url(nested.woff2); } }
@supports (display: grid) { .a { color: red } }
<!-- @font-face { font-family: Cdo; src: url(cdo.woff2); } -->
@FONT-FACE { font-family: Upper; src: url(upper.woff2); }`;
    expect(parseFontFaceCss(css, SHEET).map((entry) => entry.family)).toEqual(['Cdo', 'Upper']);
  });

  it('is not misled by braces, semicolons and comments inside strings and url tokens', () => {
    const css = `
/* @font-face { font-family: Commented; src: url(c.woff2); } */
@font-face {
  font-family: "Brace } ; Name";
  src: /* } */ url("a;b}.woff2");
}
@font-face { font-family: Second; src: url(x/*not-a-comment*/.woff2); }`;
    expect(parseFontFaceCss(css, SHEET)).toEqual([
      { family: 'Brace } ; Name', url: 'https://cdn.internal.example/styles/a;b%7D.woff2' },
      { family: 'Second', url: 'https://cdn.internal.example/styles/x/*not-a-comment*/.woff2' },
    ]);
  });

  it('decodes family escapes and rejects reserved or malformed names', () => {
    const family = (value: string) => parseFontFaceCss(face(`font-family: ${value}; src: url(a.woff2);`), SHEET)[0]?.family;
    expect(family('"Noto \\53 ans"')).toBe('Noto Sans');
    expect(family('Noto   Sans  JP')).toBe('Noto Sans JP');
    expect(family('serif')).toBeUndefined();
    expect(family('inherit')).toBeUndefined();
    expect(family('1Font')).toBeUndefined();
    expect(family('"unterminated')).toBeUndefined();
  });

  it('consumes escapes per CSS Syntax §4.3.7', () => {
    const family = (value: string) => parseFontFaceCss(face(`font-family: ${value}; src: url(a.woff2);`), SHEET)[0]?.family;
    // The white space ending a hex escape belongs to the escape.
    expect(family('\\49 nter')).toBe('Inter');
    expect(family('\\49\r\nnter')).toBe('Inter');
    expect(family('\\49  nter')).toBe('I nter');
    expect(family('\\000049nter')).toBe('Inter');
    expect(family('A\\0 B')).toBe('A\uFFFDB');
    expect(family('A\\D800 B')).toBe('A\uFFFDB');
    expect(family('A\\110000 B')).toBe('A\uFFFDB');
    expect(family('A\\1F600 B')).toBe('A\u{1F600}B');
    expect(family('A\\\u{1F600}B')).toBe('A\u{1F600}B');
    expect(family('A\\ B')).toBe('A B');
    expect(family('"A\\49 nter"')).toBe('AInter');
    // A string line continuation is removed.
    expect(family('"Line\\\nCont"')).toBe('LineCont');
    // A backslash at end of input is U+FFFD in an identifier.
    expect(parseFontFaceCss('@font-face { src: url(a.woff2); font-family: X\\}', SHEET)).toEqual([]);
    expect(parseFontFaceCss('@font-face { src: url(a.woff2); font-family: X\\', SHEET)).toEqual([]);
  });

  it('drops invalid, !important and block-valued descriptors as CSS does', () => {
    const [declaration] = parseFontFaceCss(face(`
      font-family: X; src: url(a.woff2);
      font-weight: 400; font-weight: 700 !important;
      font-style: {italic};
      unicode-range: U+0000-00FF; unicode-range: U+0041-0030;`), SHEET);
    // An invalid later declaration leaves the earlier valid one in place.
    expect(declaration).toEqual({
      family: 'X', url: 'https://cdn.internal.example/styles/a.woff2', weight: '400', unicodeRange: 'U+0000-00FF',
    });
    const [noRange] = parseFontFaceCss(face('font-family: X; src: url(a.woff2); unicode-range: U+110000;'), SHEET);
    expect(noRange.unicodeRange).toBeUndefined();
  });

  it('drops an @font-face block left unterminated at end of input', () => {
    expect(parseFontFaceCss('@font-face { font-family: X; src: url(a.woff2);', SHEET)).toEqual([]);
    expect(parseFontFaceCss('@font-face { font-family: X; src: url(a.woff2); } @font-face { font-family: Y', SHEET)
      .map((entry) => entry.family)).toEqual(['X']);
  });

  it('scans deeply nested and unbalanced input linearly without recursion', () => {
    const depth = 200_000;
    const css = `.a ${'{'.repeat(depth)}${'}'.repeat(depth)} ${'('.repeat(depth)} ;
      ${face('font-family: After; src: url(a.woff2);')}`;
    const started = performance.now();
    // The open parentheses of the hostile prelude swallow the rest of the sheet.
    expect(parseFontFaceCss(css, SHEET)).toEqual([]);
    expect(parseFontFaceCss(`${'['.repeat(depth)}`, SHEET)).toEqual([]);
    expect(performance.now() - started).toBeLessThan(2000);
  });

  it('rejects a stylesheet beyond the byte ceiling and rules beyond the declaration ceiling', () => {
    expect(limitQuota(() => parseFontFaceCss(' '.repeat(HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES + 1), SHEET)))
      .toBe('stylesheet');
    // Counted as UTF-8 bytes, not UTF-16 units.
    expect(limitQuota(() => parseFontFaceCss('あ'.repeat(HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES / 3 + 1), SHEET)))
      .toBe('stylesheet');
    const rules = face('font-family: X; src: url(a.woff2);').repeat(HARD_MAX_FONT_SOURCE_DECLARATIONS + 1);
    expect(limitQuota(() => parseFontFaceCss(rules, SHEET))).toBe('declarations');
    expect(parseFontFaceCss(face('font-family: X; src: url(a.woff2);').repeat(HARD_MAX_FONT_SOURCE_DECLARATIONS), SHEET))
      .toHaveLength(HARD_MAX_FONT_SOURCE_DECLARATIONS);
  });

  it('rejects an overlong face URL instead of dropping it', () => {
    expect(limitQuota(() => parseFontFaceCss(face(`font-family: X; src: url(${'a'.repeat(9000)}.woff2);`), SHEET)))
      .toBe('url');
  });
});

describe('parseUnicodeRange', () => {
  it('accepts spans and wildcards and rejects anything else', () => {
    expect(parseUnicodeRange('U+0-7F, u+30??')).toEqual([[0, 0x7f], [0x3000, 0x30ff]]);
    expect(parseUnicodeRange('U+7F-0')).toBeNull();
    expect(parseUnicodeRange('U+0-110000')).toBeNull();
    expect(parseUnicodeRange('U+0000 - 00FF')).toBeNull();
    expect(parseUnicodeRange('U+???????')).toBeNull();
    expect(parseUnicodeRange('')).toBeNull();
  });
});

describe('fetchFontStylesheet', () => {
  const G = globalThis as { fetch?: unknown };
  const originalFetch = G.fetch;
  afterEach(() => { G.fetch = originalFetch; });

  function streamed(chunks: Uint8Array[], init: { url?: string; contentLength?: number } = {}) {
    let cancelled = false;
    let bodyCancelled = false;
    let index = 0;
    return {
      response: {
        ok: true, status: 200, url: init.url ?? '',
        headers: { get: (name: string) => name === 'content-length' && init.contentLength !== undefined ? String(init.contentLength) : null },
        body: {
          getReader: () => ({
            read: async () => index < chunks.length ? { done: false, value: chunks[index++] } : { done: true },
            cancel: async () => { cancelled = true; },
          }),
          cancel: async () => { bodyCancelled = true; },
        },
        text: async () => { throw new Error('body must be streamed'); },
      },
      cancelled: () => cancelled,
      bodyCancelled: () => bodyCancelled,
      reads: () => index,
    };
  }

  it('decodes a streamed body and resolves relative URLs against the final URL', async () => {
    const css = face("font-family: 'Interé'; src: url(../fonts/i.woff2);");
    const bytes = new TextEncoder().encode(css);
    const { response } = streamed([bytes.slice(0, 20), bytes.slice(20)], { url: 'https://cdn2.example/v2/a.css' });
    const fetchMock = vi.fn(async () => response);
    G.fetch = fetchMock;
    const result = await fetchFontStylesheet('https://fonts.example/a.css', undefined, new AbortController().signal);
    expect(result).toEqual({
      ok: true,
      finalUrl: 'https://cdn2.example/v2/a.css',
      declarations: [{ family: 'Interé', url: 'https://cdn2.example/fonts/i.woff2' }],
    });
    expect(fetchMock).toHaveBeenCalledWith('https://fonts.example/a.css',
      expect.objectContaining({ credentials: 'same-origin' }));
  });

  it('stops reading once the byte ceiling is exceeded', async () => {
    const chunk = new Uint8Array(256 * 1024).fill(0x20);
    const stream = streamed([chunk, chunk, chunk, chunk, chunk, chunk, chunk, chunk]);
    let requestSignal: AbortSignal | undefined;
    G.fetch = async (_url: string, init: RequestInit) => { requestSignal = init.signal ?? undefined; return stream.response; };
    const caller = new AbortController();
    await expect(fetchFontStylesheet('https://fonts.example/a.css', 'omit', caller.signal))
      .rejects.toMatchObject({ code: 'font-source-limit', violation: { quota: 'stylesheet' } });
    expect(stream.reads()).toBe(5);
    expect(stream.cancelled()).toBe(true);
    // The request is aborted although the caller's signal stays live.
    expect(requestSignal?.aborted).toBe(true);
    expect(caller.signal.aborted).toBe(false);
  });

  it('rejects a declared Content-Length beyond the ceiling before reading', async () => {
    const stream = streamed([], { contentLength: HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES + 1 });
    let requestSignal: AbortSignal | undefined;
    G.fetch = async (_url: string, init: RequestInit) => { requestSignal = init.signal ?? undefined; return stream.response; };
    await expect(fetchFontStylesheet('https://fonts.example/a.css', undefined, new AbortController().signal))
      .rejects.toMatchObject({ code: 'font-source-limit' });
    expect(stream.reads()).toBe(0);
    expect(stream.bodyCancelled()).toBe(true);
    expect(requestSignal?.aborted).toBe(true);
  });

  it('reports transport failures and empty stylesheets as failures, not errors', async () => {
    const signal = new AbortController().signal;
    G.fetch = async () => ({ ok: false, status: 404, text: async () => '' });
    expect(await fetchFontStylesheet('https://fonts.example/a.css', undefined, signal)).toEqual({ ok: false, reason: 'HTTP 404' });
    G.fetch = async () => { throw new TypeError('Failed to fetch'); };
    expect(await fetchFontStylesheet('https://fonts.example/a.css', undefined, signal)).toEqual({ ok: false, reason: 'Failed to fetch' });
    G.fetch = async () => ({ ok: true, status: 200, text: async () => 'body { color: red }' });
    expect(await fetchFontStylesheet('https://fonts.example/a.css', undefined, signal))
      .toEqual({ ok: false, reason: 'no usable @font-face rule' });
  });
});
