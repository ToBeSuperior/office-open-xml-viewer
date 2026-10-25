import { afterEach, describe, expect, it, vi } from 'vitest';
import { OoxmlError } from '../errors/ooxml-error.js';
import {
  HARD_MAX_FONT_SOURCE_DECLARATIONS,
  HARD_MAX_FONT_SOURCE_FACE_BYTES,
  HARD_MAX_FONT_SOURCE_URL_LENGTH,
} from '../worker/resource-policy.generated.js';
import {
  FONT_SOURCE_PROTOCOL,
  assertFontSourceDescriptors,
  cssApi,
  cssApiRequests,
  fontFaceCss,
  fontFiles,
  fontResolver,
  fontResolverFunction,
  validateFontFaceDeclaration,
  type FontResolveFn,
} from './source.js';
import { OFFICE_FONT_SUBSTITUTIONS } from './substitutions.js';

function limitOf(run: () => unknown): unknown {
  try {
    run();
  } catch (error) {
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe('font-source-limit');
    return (error as OoxmlError).violation;
  }
  throw new Error('expected a font-source limit');
}

afterEach(() => vi.unstubAllGlobals());

describe('font source descriptors', () => {
  it('produce frozen, structured-clone-safe descriptors a worker re-validates identically', () => {
    const bytes = new Uint8Array([0, 1, 0, 0]).buffer;
    const resolver = fontResolver(() => []);
    const sources = [
      cssApi({ urlTemplate: 'https://fonts.example/css?family={family}:{query}', query: '400,700',
        substitutes: OFFICE_FONT_SUBSTITUTIONS, credentials: 'omit' }),
      fontFaceCss({ url: 'https://cdn.example/office.css' }),
      fontFiles({ faces: [
        { family: 'Carlito', url: 'https://cdn.example/Carlito.woff2', weight: '400', unicodeRange: 'U+0-FF' },
        { family: 'Meiryo', data: bytes },
      ] }),
      resolver,
    ];
    for (const source of sources) {
      expect(source.protocol).toBe(FONT_SOURCE_PROTOCOL);
      expect(Object.isFrozen(source)).toBe(true);
    }
    // The posted copy is a different object, so the worker validates it; the
    // result must describe the same source (no application function inside).
    const crossed = assertFontSourceDescriptors(structuredClone(sources));
    expect(crossed).toEqual(sources);
    expect(crossed[3]).toEqual({ protocol: FONT_SOURCE_PROTOCOL, kind: 'font-resolver', id: resolver.id });
    // Factory descriptors pass through without copying their face bytes.
    expect(assertFontSourceDescriptors(sources)[2]).toBe(sources[2]);
  });

  it('resolves relative application URLs against the calling realm, never in the receiving realm', () => {
    vi.stubGlobal('document', { baseURI: 'https://app.example/viewer/index.html' });
    const css = fontFaceCss({ url: '../fonts/office.css' });
    const files = fontFiles({ faces: [{ family: 'Carlito', url: '/fonts/Carlito.woff2' }] });
    expect(css.url).toBe('https://app.example/fonts/office.css');
    expect(files.faces[0].url).toBe('https://app.example/fonts/Carlito.woff2');
    expect(() => assertFontSourceDescriptors([{ ...structuredClone(css), url: '/fonts/office.css' }]))
      .toThrow(TypeError);
  });

  it('rejects malformed descriptors with TypeError before any load', () => {
    const invalid: Array<() => unknown> = [
      () => cssApi({ urlTemplate: 'https://fonts.example/css' }),
      () => cssApi({ urlTemplate: 'https://fonts.example/css?family={family}&{subset}' }),
      () => cssApi({ urlTemplate: 'https://fonts.example/css?family={family}:{query}' }),
      () => cssApi({ urlTemplate: '/css?family={family}' }),
      () => cssApi({ urlTemplate: 'ftp://fonts.example/css?family={family}' }),
      () => cssApi({ urlTemplate: 'https://fonts.example/css?family={family}', families: { Lato: true, lato: true } }),
      () => cssApi({ urlTemplate: 'https://fonts.example/css?family={family}', credentials: 'include' as 'omit' }),
      () => fontFaceCss({ url: 'blob:https://app.example/1' }),
      () => fontFaceCss({ url: 'https://cdn.example/a.css', headers: {} } as never),
      () => fontFiles({ faces: [] }),
      () => fontFiles({ faces: [{ family: 'A' }] }),
      () => fontFiles({ faces: [{ family: 'A', url: 'https://a.example/a.woff2', data: new ArrayBuffer(4) }] }),
      () => fontFiles({ faces: [{ family: 'A', data: new Uint8Array(4) as unknown as ArrayBuffer }] }),
      () => fontFiles({ faces: [{ family: 'A', data: new ArrayBuffer(0) }] }),
      () => fontFiles({ faces: [{ family: ' A', url: 'https://a.example/a.woff2' }] }),
      () => fontFiles({ faces: [{ family: 'A', url: 'javascript:alert(1)' }] }),
      () => fontFiles({ faces: [{ family: 'A', url: 'https://a.example/a.woff2' }],
        substitutes: [{ for: 'Calibri', use: 'Carlito', script: 'latin' as 'arabic' }] }),
      () => fontResolver('resolve' as unknown as FontResolveFn),
      () => assertFontSourceDescriptors({}),
      () => assertFontSourceDescriptors([{ protocol: 'other/v1', kind: 'font-files', faces: [] }]),
      () => assertFontSourceDescriptors([{ protocol: FONT_SOURCE_PROTOCOL, kind: 'font-resolver', id: 'x' }]),
    ];
    for (const run of invalid) expect(run).toThrow(TypeError);
  });

  it('fails closed with font-source-limit beyond the generated ceilings', () => {
    expect(limitOf(() => fontFiles({ faces: [{ family: 'A',
      data: new ArrayBuffer(HARD_MAX_FONT_SOURCE_FACE_BYTES + 1) }] })))
      .toEqual({ resource: 'font-source', quota: 'face', metric: 'bytes',
        limit: HARD_MAX_FONT_SOURCE_FACE_BYTES, observed: HARD_MAX_FONT_SOURCE_FACE_BYTES + 1 });
    const long = `https://a.example/${'a'.repeat(HARD_MAX_FONT_SOURCE_URL_LENGTH)}`;
    expect(limitOf(() => fontFaceCss({ url: long }))).toMatchObject({ quota: 'url', metric: 'length' });
    expect(limitOf(() => fontFiles({ faces: Array.from({ length: HARD_MAX_FONT_SOURCE_DECLARATIONS + 1 },
      () => ({ family: 'A', url: 'https://a.example/a.woff2' })) })))
      .toMatchObject({ quota: 'declarations', metric: 'entry-count' });
    // A data: URL is charged against the face byte ceiling (base64 expansion),
    // not the URL length ceiling, so small inline faces are accepted.
    const inline = `data:font/woff2;base64,${'A'.repeat(HARD_MAX_FONT_SOURCE_URL_LENGTH * 2)}`;
    expect(fontFiles({ faces: [{ family: 'A', url: inline }] }).faces[0].url).toBe(inline);
    const oversized = `data:font/woff2;base64,${'A'.repeat(Math.ceil(HARD_MAX_FONT_SOURCE_FACE_BYTES * 4 / 3))}`;
    expect(limitOf(() => validateFontFaceDeclaration({ family: 'A', url: oversized })))
      .toMatchObject({ quota: 'face', metric: 'bytes', limit: HARD_MAX_FONT_SOURCE_FACE_BYTES });
  });

  it('keeps the resolver function in this realm and shares one id per function', () => {
    const resolve: FontResolveFn = () => [];
    const first = fontResolver(resolve);
    const second = fontResolver(resolve, { substitutes: OFFICE_FONT_SUBSTITUTIONS });
    const other = fontResolver(() => []);
    expect(second.id).toBe(first.id);
    expect(other.id).not.toBe(first.id);
    expect(fontResolverFunction(second)).toBe(resolve);
    expect(fontResolverFunction(structuredClone(first))).toBeUndefined();
  });
});

describe('css-api request composition', () => {
  it('requests the authored family first, then declared substitutes, encoding the family only', () => {
    const source = cssApi({
      urlTemplate: 'https://fonts.example/css?family={family}:{query}',
      query: '400,700',
      substitutes: OFFICE_FONT_SUBSTITUTIONS,
    });
    expect(cssApiRequests(source, 'Calibri')).toEqual([
      { family: 'Calibri', url: 'https://fonts.example/css?family=Calibri:400,700' },
      { family: 'Carlito', url: 'https://fonts.example/css?family=Carlito:400,700',
        substitution: { for: 'Calibri', use: 'Carlito' } },
    ]);
    expect(cssApiRequests(source, 'Noto Sans & Co')[0].url)
      .toBe('https://fonts.example/css?family=Noto+Sans+%26+Co:400,700');
  });

  it('restricts requests to the allow-list, using its spelling and per-family query', () => {
    const source = cssApi({
      urlTemplate: 'https://fonts.example/css?family={family}:{query}',
      query: 'wght@400',
      families: { 'Open Sans': true, 'Libre Franklin': 'wght@400;600' },
      substitutes: [{ for: 'Franklin Gothic Book', use: 'Libre Franklin' }],
    });
    expect(cssApiRequests(source, 'OPEN SANS')).toEqual([
      { family: 'Open Sans', url: 'https://fonts.example/css?family=Open+Sans:wght@400' },
    ]);
    expect(cssApiRequests(source, 'franklin gothic book')).toEqual([
      { family: 'Libre Franklin', url: 'https://fonts.example/css?family=Libre+Franklin:wght@400;600',
        substitution: { for: 'Franklin Gothic Book', use: 'Libre Franklin' } },
    ]);
    expect(cssApiRequests(source, 'Comic Sans MS')).toEqual([]);
  });
});
