import { expect, test } from '@playwright/test';
import { build } from 'rolldown';
import { fileURLToPath } from 'node:url';

/**
 * The @font-face parser lets the host CSS parser validate font-weight,
 * font-style, font-stretch and unicode-range through probe FontFaces
 * (`packages/core/src/fonts/css-source.ts`). This checks, in Chrome, that
 * real engines mark invalid probes the way the parser assumes, in the main
 * thread and in a dedicated worker, and that probes stay invisible: no
 * request, no FontFaceSet entry, no unhandled rejection.
 */
const entry = fileURLToPath(new URL('../../packages/core/src/fonts/css-source.ts', import.meta.url));
let parserBundle: Promise<string> | undefined;
const parserScript = () => parserBundle ??= build({
  input: entry, output: { format: 'iife', name: 'cssSource' }, platform: 'browser',
}).then((bundle) => bundle.output[0].code);

const CASES: readonly (readonly [string, Record<string, string | undefined>])[] = [
  ['font-weight: 400; font-weight: calc(700);', { weight: 'calc(700)' }],
  ['font-weight: 400; font-weight: 1001;', { weight: '400' }],
  ['font-weight: 400; font-weight: heavy;', { weight: '400' }],
  ['font-weight: 300; font-weight: 100 900;', { weight: '100 900' }],
  ['font-style: normal; font-style: oblique 10deg 20deg;', { style: 'oblique 10deg 20deg' }],
  ['font-style: italic; font-style: oblique 100deg;', { style: 'italic' }],
  ['font-style: italic; font-style: slanted;', { style: 'italic' }],
  ['font-stretch: 75%; font-stretch: 50% 200%;', { stretch: '50% 200%' }],
  ['font-stretch: 75%; font-stretch: -5%;', { stretch: '75%' }],
  ['font-stretch: 75%; font-stretch: narrow;', { stretch: '75%' }],
  ['unicode-range: U+0-FF; unicode-range: U+30-20;', { unicodeRange: 'U+0-FF' }],
  ['unicode-range: U+0-FF; unicode-range: bogus;', { unicodeRange: 'U+0-FF' }],
  ['unicode-range: U+0-FF; unicode-range: U+3000-30FF, U+4E00;', { unicodeRange: 'U+3000-30FF, U+4E00' }],
];

const sheet = CASES.map(([body], index) =>
  `@font-face { font-family: Case${index}; src: url(https://fonts.example/f${index}.woff2); ${body} }`).join('\n');

interface Outcome {
  readonly host: boolean;
  readonly declarations: readonly Record<string, string | undefined>[];
  readonly fontsSize: number;
  readonly resources: number;
  readonly rejections: readonly string[];
  readonly controlStatus: string;
}

/** Runs in either realm (`self` is the window or the worker global). */
function runInRealm(css: string): Promise<Outcome> {
  const scope = self as unknown as {
    cssSource: { parseFontFaceCss(css: string, url: string): Record<string, string | undefined>[] };
    document?: { fonts: FontFaceSet };
    fonts?: FontFaceSet;
    performance: Performance;
  };
  const rejections: string[] = [];
  self.addEventListener('unhandledrejection', (event) => { rejections.push(String(event.reason)); });
  const declarations = scope.cssSource.parseFontFaceCss(css, 'https://fonts.example/kit.css');
  // Control: an invalid face whose rejected `loaded` nobody observes must be
  // reported, proving the detector works and that the engine does reject it.
  void new FontFace('control-probe', 'local(control-probe)', { weight: '1001' }).loaded;
  // Give rejected probe promises time to surface as unhandled if unobserved.
  return new Promise((resolve) => setTimeout(() => resolve({
    host: typeof FontFace === 'function',
    declarations: declarations.map(({ weight, style, stretch, unicodeRange }) => ({ weight, style, stretch, unicodeRange })),
    // document.fonts on the window, self.fonts in a worker.
    fontsSize: (scope.document?.fonts ?? scope.fonts as FontFaceSet).size,
    resources: scope.performance.getEntriesByType('resource').length,
    rejections,
    controlStatus: new FontFace('control-probe', 'local(control-probe)', { weight: '1001' }).status,
  }), 500));
}

test('host FontFace validates @font-face descriptors in the main thread and a worker', async ({ page }) => {
  test.setTimeout(60_000);
  const requests: string[] = [];
  page.context().on('request', (request) => { if (!request.url().startsWith('blob:')) requests.push(request.url()); });
  await page.goto('about:blank');
  const script = await parserScript();

  await page.addScriptTag({ content: script });
  const main = await page.evaluate(`(${runInRealm.toString()})(${JSON.stringify(sheet)})`) as Outcome;

  const worker = await page.evaluate(async ({ script, sheet, run }) => {
    const source = `${script}\nconst run = ${run};\nself.onmessage = async (event) => self.postMessage(await run(event.data));`;
    const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
    try {
      const w = new Worker(url);
      return await new Promise((resolve, reject) => {
        w.onmessage = (event) => resolve(event.data);
        w.onerror = (event) => reject(new Error(event.message));
        w.postMessage(sheet);
      });
    } finally {
      URL.revokeObjectURL(url);
    }
  }, { script, sheet, run: runInRealm.toString() }) as Outcome;

  for (const [realm, outcome] of [['main', main], ['worker', worker]] as const) {
    expect(outcome.host, realm).toBe(true);
    expect(outcome.declarations, realm).toHaveLength(CASES.length);
    CASES.forEach(([body, expected], index) => {
      expect(outcome.declarations[index], `${realm}: ${body}`).toMatchObject(expected);
    });
    expect(outcome.fontsSize, `${realm}: probes never enter the FontFaceSet`).toBe(0);
    expect(outcome.resources, `${realm}: probes fetch nothing`).toBe(0);
    // The engine marks an invalid probe the way the parser assumes.
    expect(outcome.controlStatus, realm).toBe('error');
    // Only the deliberately unobserved control surfaces.
    expect(outcome.rejections, `${realm}: no unhandled rejection from the probes`).toHaveLength(1);
    expect(outcome.rejections[0], realm).toMatch(/SyntaxError/);
  }
  expect(requests).toEqual([]);
});
