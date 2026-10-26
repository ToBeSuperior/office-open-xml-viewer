# Migrating from 0.89 to 0.90: font sources

Version 0.90 replaces the boolean `useGoogleFonts` load option with `fonts`, an
ordered list of font sources. Migration is required only if you pass
`useGoogleFonts`. The default is unchanged: with no sources, the library
requests no fonts from any network host and the canvas uses the fonts installed
on the viewer's machine.

`useGoogleFonts` still works in 0.90, with a deprecation warning. It is removed
in 0.91.0.

## Replace `useGoogleFonts`

| 0.89 | 0.90 |
| --- | --- |
| `useGoogleFonts: true` | `fonts: [googleFonts()]` |
| `useGoogleFonts: false` | omit the option |

```ts
import { DocxViewer, googleFonts } from '@silurus/ooxml/docx';

const viewer = new DocxViewer(canvas, { fonts: [googleFonts()] });
```

The factories are exported from `@silurus/ooxml` and from the
`@silurus/ooxml/docx`, `@silurus/ooxml/xlsx` and `@silurus/ooxml/pptx` entries.
They are the same objects in every entry. `fonts` works the same way for DOCX,
XLSX and PPTX, on every engine `load()` and every Viewer, in `mode: 'main'` and
`mode: 'worker'`.

`fonts: [googleFonts()]` sends the same stylesheet requests as
`useGoogleFonts: true`. In 0.90, DOCX and XLSX also load only the
`unicode-range` subsets that intersect the characters the document paints, as
PPTX already did.

While the alias exists:

- `useGoogleFonts: true` is converted to `fonts: [googleFonts()]` before
  loading starts.
- `useGoogleFonts: false` does nothing, but still warns.
- `useGoogleFonts: true` together with `fonts` throws a `TypeError`. Add
  `googleFonts()` to `fonts` instead.
- The warning is printed once per JavaScript realm (page, popup or worker):

  ```text
  [ooxml] useGoogleFonts is deprecated and will be removed in 0.91.0. Replace it with fonts: [googleFonts()] — import { googleFonts } from '@silurus/ooxml' (also exported from the docx/xlsx/pptx entries). A custom CSS origin is googleFonts({ cssOrigin }). See docs/migration-0.90.md.
  ```

The VS Code extension keeps its `ooxmlViewer.useGoogleFonts` setting. It now
passes `fonts: [googleFonts()]`, so its users see no warning.

## Font source kinds

### A Google-compatible mirror or proxy: `googleFonts({ cssOrigin })`

```ts
fonts: [googleFonts({ cssOrigin: 'https://fonts.googleapis.cn' })]
fonts: [googleFonts({ cssOrigin: 'https://fonts.internal.example:8443' })]
```

`cssOrigin` is a scheme, host and optional port (`http:` or `https:`). One
trailing slash is accepted. A path, query or user name throws a `TypeError`
when `googleFonts()` is called, before any load. The requests keep the Google
CSS2 path and query (`/css2?family=…&display=swap`), so the origin must serve
that API. No request goes to `fonts.googleapis.com`.

Font file URLs come from the returned CSS. Relative URLs are resolved against
the stylesheet's final URL, after redirects. Two documents that use different
origins share no stylesheet, cache entry or registered face.

### Another CSS API: `cssApi()`

```ts
import { cssApi, officeFontSubstitutions } from '@silurus/ooxml';

fonts: [cssApi({
  urlTemplate: 'https://fonts.example/css?family={family}:{query}',
  query: '400,700',
  substitutes: officeFontSubstitutions,
})]
```

- `urlTemplate` is an absolute `http(s)` URL. It must contain `{family}` and
  may contain `{query}`.
- `{family}` is percent-encoded, with `+` for spaces.
- `{query}` is inserted verbatim from `query` or from the `families`
  allow-list, never from document content. `query` is required when the
  template has `{query}`, unless every `families` entry supplies its own query.
- `families` is an optional allow-list (`{ 'Open Sans': true, Lato: 'wght@400' }`).
  When present, only those families are requested.

The source sends one stylesheet request per served family.

### A self-hosted stylesheet: `fontFaceCss()`

```ts
fonts: [fontFaceCss({ url: '/fonts/office.css', substitutes: officeFontSubstitutions })]
```

The stylesheet may describe any number of families. It is fetched once and
shared by every document that uses the same configuration. Only `@font-face`
rules with `url()` sources are read. `local()` sources and everything outside
`@font-face` are ignored.

### Bundled files: `fontFiles()`

```ts
fonts: [fontFiles({
  faces: [
    { family: 'Carlito', url: '/fonts/Carlito-Regular.woff2', weight: '400' },
    { family: 'Carlito', url: '/fonts/Carlito-Bold.woff2', weight: '700' },
    { family: 'Noto Sans JP', url: '/fonts/NotoSansJP-kana.woff2', unicodeRange: 'U+3000-30FF' },
    { family: 'Brand', url: 'data:font/woff2;base64,…' },
    { family: 'Brand Display', data: brandDisplayBytes },
  ],
  substitutes: officeFontSubstitutions,
})]
```

Each face has exactly one of these:

- `url`: an `http(s)`, `blob:` or `data:` URL.
- `data`: a non-empty `ArrayBuffer`. Views, `SharedArrayBuffer`s and
  resizable buffers are rejected.

`weight`, `style`, `stretch` and `unicodeRange` take standard CSS
`@font-face` descriptor values. A `data:` URL is accepted and its decoded size
counts against the same limits as a `data` buffer.

Relative `url` values in `fontFiles()` and `fontFaceCss()` are resolved against
the page's base URL when the factory is called. Build these descriptors on the
page, not in a worker.

### Authenticated or private services: `fontResolver()`

```ts
import { fontResolver, googleFonts, officeFontSubstitutions } from '@silurus/ooxml';

const privateFonts = fontResolver(async (demand, { signal }) => {
  const response = await fetch('/api/fonts/resolve', {
    method: 'POST',
    body: JSON.stringify({ families: demand.families }),
    signal,
  });
  return response.json(); // FontFaceDeclaration[]: { family, url | data, … }
}, { substitutes: officeFontSubstitutions });

fonts: [privateFonts, googleFonts({ cssOrigin: 'https://fonts.internal.example' })]
```

The function always runs on the main thread. In worker mode, the worker asks
the main thread to run it, so tokens and network clients stay in your
application.

The `demand` argument contains:

- `families`: the authored families the host does not have.
- `tuples`: the weight and style of each authored family.
- `fallbackFamilies`: the library's script fallback families for the text,
  such as Noto families for CJK.
- `scalars`: the Unicode scalars painted, or `'all'`.
- `substituteScripts`: optional. The scripts each family's text contains.

The function returns `FontFaceDeclaration` objects with the same fields as
`fontFiles()` faces. URLs must be absolute. A resolver may ignore any demand
field and may answer only some families. `signal` is aborted when the document
is destroyed or the 15 s font deadline expires.

If the answer is not an array or contains a malformed declaration, the families
keep their host fallback. A `console.warn` reports the problem. If the
function throws, the families also keep their host fallback, and the warning
includes the error message.

## Chaining

Sources are consulted per family, in option order:

1. Each source tries the family itself, then the `use` family of each of its
   matching `substitutes`.
2. The first source that has a face for the family claims it. Later sources
   are never consulted for that family.
3. A source that does not serve the family passes it to the next source.
   Examples: a family outside a `cssApi` allow-list, a family missing from a
   stylesheet or file list, or a family a resolver did not answer.
4. A source that tries a family and fails also makes the family final, with no
   retry anywhere else. Failures include an unreachable or empty stylesheet, an
   HTTP error, a resolver error and the 15 s deadline. A `fontFaceCss`
   stylesheet that fails to load fails every family it is asked for.
5. A face that is claimed but fails to load also keeps its host fallback. It
   does not re-enter the chain.

Each failure keeps the family on the fonts installed on the viewer's machine
and prints one `console.warn` listing the families. There is no strict mode
and no callback. A failing mirror therefore never falls through to the public
Google origin: that origin is contacted only when `googleFonts()` without
`cssOrigin` is in the list. Failures are not cached, so a document opened later
retries the same source.

What paints, highest first:

1. A font embedded in the document.
2. An installed face with the authored name. Installed families are never
   requested from a source.
3. A source face for the authored family.
4. A source face for a declared substitute.
5. The host's fallback chain.

## Substitutions

`officeFontSubstitutions` (`OFFICE_FONT_SUBSTITUTIONS` inside the library) is
the library's list of openly licensed stand-ins for Office faces:

- Calibri → Carlito
- Cambria → Caladea
- Franklin Gothic Book and Franklin Gothic Medium → Libre Franklin
- several Arabic faces → Noto Naskh Arabic or Noto Sans Arabic, for Arabic
  text only

Any source can use it, or a list of your own, through `substitutes`:

```ts
substitutes: [...officeFontSubstitutions, { for: 'Corporate Sans', use: 'Inter' }]
```

An entry is `{ for, use, script? }`:

- `for`: the authored family, matched case-insensitively.
- `use`: the family the source serves.
- `script`: optional. `'arabic'` limits the substitute to Arabic text.

The renderers use a substitution only when its face actually loaded.
`googleFonts()` already includes `officeFontSubstitutions`, plus Google's names
for the Noto CJK families.

## Requirements and limits

### Network, CORS and cookies

- Fonts from another origin need CORS headers (`Access-Control-Allow-Origin`).
  This applies to stylesheets and font files, as with Google Fonts today.
- Stylesheets are fetched with `credentials: 'same-origin'` by default.
  `credentials: 'omit'` is available on `cssApi()` and `fontFaceCss()`.
  `'include'` and custom headers are not supported.
- Font files are fetched by the browser, which follows the CSS font-loading
  rules.
- Cookies are never sent to another origin. For an authenticated service on
  another origin, use `fontResolver` and return `data` bytes or a signed URL.
- A hosted source receives the end user's IP address and User-Agent.

### Content Security Policy

- Stylesheet origins (`cssApi`, `fontFaceCss`, `googleFonts`) need
  `connect-src`. The library downloads the stylesheet with `fetch()`; it does
  not inject a `<link>`.
- Font file origins need `font-src`. Add `data:` or `blob:` to `font-src` when
  you use those URLs.
- `data` buffers are not fetched, so they need no `font-src` entry.
- For Google Fonts, the policy is `connect-src https://fonts.googleapis.com`
  and `font-src https://fonts.gstatic.com`. A mirror uses its own two origins.

### Resolver buffers and timing

- A `data` buffer belongs to the library once it is passed to `fontFiles()` or
  returned from a resolver. Do not modify it. Documents in the same realm may
  share it. In worker mode the worker receives a copy, and your buffer is not
  detached.
- Font loading waits at most 15 s each time a document needs new fonts: once
  per DOCX document, once per PPTX slide that adds fonts or characters, and
  once per XLSX workbook plus once per sheet that adds fonts. The deadline
  covers every source in the chain together. In worker mode the main thread
  owns the deadline of a resolver call.
- The resolver is called once per family per document. For XLSX in worker mode
  the viewer and the worker share that call. Documents open at the same time
  that use the same function share its answers.
- In worker mode, the worker waits silently while the resolver runs. A slow
  resolver combined with a short `workerTimeoutMs` can make the load time out.

### Resource limits

Font sources are application configuration, so they have fixed limits. These
limits are not part of `resourceLimits`.

| Limit | Value |
| --- | --- |
| Stylesheet size | 1 MiB per stylesheet |
| Declarations | 2,048 per document |
| Requested families | 256 per document |
| One face (`data`, or the decoded body of a `data:` URL) | 30 MiB |
| Retained `data` bytes | 256 MiB per realm, across open documents |
| URL length (non-`data:` URLs and composed `cssApi` requests) | 8 KiB |

Exceeding a limit fails closed. The error is an `OoxmlError` with
`code === 'font-source-limit'`, and `error.violation` holds `{ resource:
'font-source', quota, metric, limit, observed }`. Limits that can be checked
up front make the factory throw immediately. Otherwise the error rejects
`load()`, or the slide or sheet operation that needed the fonts, and that
document's later font loads reject too. A limit breach means a misconfigured
or hostile source, not a network problem.

The library does not count the bytes of `url` faces, because the browser
downloads them. A document whose painted text has more than 65,536 distinct
characters loads every subset; that is not an error.

## Licences

The library ships no font binaries. Your application is responsible for the
licences of the fonts it serves. Google Fonts' Carlito, Caladea, Libre Franklin
and Noto families are licensed under the SIL Open Font License.

## Removal in 0.91

0.91.0 removes `useGoogleFonts` from `LoadOptions`. Passing it then throws a
`TypeError`. Replace it with `fonts` before upgrading.
