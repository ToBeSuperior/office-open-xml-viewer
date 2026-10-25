import { describe, expect, it, vi } from 'vitest';
import { fontFiles } from '@silurus/ooxml-core/internal/font-sources';
import { collectSlideFontDemand } from './font-preload-demand.js';
import { FontSourceSession, PptxInstalledFamilies, pptxFontDemand } from './font-source-load.js';
import { PresentationPreflightBuilder } from './presentation-preflight.js';
import type { Slide } from './types';
import type { PresentationBootstrap } from './worker-protocol.js';

describe('pptxFontDemand', () => {
  const facts = {
    familyCount: 4,
    tuples: [{ family: 'Calibri', weight: 700, style: 'normal' as const }],
    substituteScripts: { 'sakkal majalla': ['arabic' as const] },
  };
  const names = ['Calibri', '+mn-lt', 'Deck Sans', 'Sakkal Majalla', 'Noto Sans JP', 'serif', null];

  it('splits authored from fallback families and drops embedded, theme-token and generic names', () => {
    const demand = pptxFontDemand(names, facts, [65], new Map([['deck sans', '__ooxml_pptx_1_0']]));
    expect(demand).toEqual({
      families: ['Calibri', 'Sakkal Majalla'],
      fallbackFamilies: ['Noto Sans JP'],
      tuples: facts.tuples,
      scalars: [65],
      substituteScripts: facts.substituteScripts,
    });
  });

  it('is conservative without collected facts or scalars', () => {
    // A preflight that omitted its optional scalar demand means 'all'; without
    // facts every name is demanded as authored and no substitute scope applies.
    const demand = pptxFontDemand(['Calibri', 'Noto Sans JP'], undefined, undefined, new Map());
    expect(demand).toEqual({ families: ['Calibri', 'Noto Sans JP'], fallbackFamilies: [], tuples: [], scalars: 'all' });
  });
});

class FakeFontFace {
  status: FontFaceLoadStatus = 'unloaded';
  constructor(readonly family: string, readonly source: string | ArrayBuffer) {}
  load() { this.status = 'loaded'; return Promise.resolve(this); }
}
function fakeFontSet() {
  const faces: FakeFontFace[] = [];
  return { faces, set: { add: (face: FakeFontFace) => { faces.push(face); }, delete: () => true } as unknown as FontFaceSet };
}

const bootstrap = { slideCount: 2, slideWidth: 100, slideHeight: 100, defaultTextColor: null,
  majorFont: null, minorFont: null, hlinkColor: null, folHlinkColor: null, embeddedFonts: [],
  slides: [{ index: 0 }, { index: 1 }] } as PresentationBootstrap;
const textSlide = (index: number, text: string): Slide => ({ index, slideNumber: index + 1, background: null,
  elements: [{ type: 'shape', textBody: { vert: 'horz', paragraphs: [{ bullet: { type: 'none' },
    runs: [{ type: 'text', text, fontFamily: 'Authored' }] }] } }] } as Slide);

describe('progressive PPTX font-source loading', () => {
  it('widens a family\'s substitution scope when a later slide proves its script', async () => {
    vi.stubGlobal('FontFace', FakeFontFace);
    try {
      const sources = [fontFiles({
        faces: [{ family: 'Served', url: 'https://app.example/served.woff2' }],
        substitutes: [{ for: 'Authored', use: 'Served', script: 'arabic' }],
      })];
      const ensureAll = async (progressive: boolean) => {
        const { faces, set } = fakeFontSet();
        const session = new FontSourceSession({ sources, fontSet: set });
        const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand });
        for (const [index, text] of ['Latin', 'العربية'].entries()) {
          builder.addSlide(textSlide(index, text));
          if (progressive) {
            await session.ensure(pptxFontDemand(builder.currentFontPreloadNames, builder.currentFontPreloadFacts,
              builder.currentFontPreloadDemandDelta, new Map()));
          }
        }
        const full = builder.finish();
        const loaded = progressive ? session.loaded : await session.ensure(pptxFontDemand(
          full.fontPreloadNames, full.fontPreloadFacts, full.fontPreloadDemand, new Map()));
        session.release();
        return { families: faces.map((face) => face.family), substitution: loaded.substitutions.get('authored') };
      };
      const ordinary = await ensureAll(false);
      expect(ordinary).toEqual({ families: ['Served'], substitution: { family: 'Served', script: 'arabic' } });
      expect(await ensureAll(true)).toEqual(ordinary);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});

describe('PptxInstalledFamilies', () => {
  it('leaves installed authored families out of demand, probing each family once', async () => {
    const probe = vi.fn(async (families: readonly string[]) => families.filter((family) => family === 'Arial').map((family) => family.toLowerCase()));
    const installed = new PptxInstalledFamilies(probe);
    const first = await installed.exclude(pptxFontDemand(['Arial', 'Deck Face', 'Noto Sans JP'], {
      familyCount: 2, tuples: [], substituteScripts: {},
    }, undefined, new Map()));
    expect(first.families).toEqual(['Deck Face']);
    // Library fallback families stay demanded, as in DOCX.
    expect(first.fallbackFamilies).toEqual(['Noto Sans JP']);
    const second = await installed.exclude(pptxFontDemand(['Arial', 'Deck Face', 'Other'], undefined, undefined, new Map()));
    expect(second.families).toEqual(['Deck Face', 'Other']);
    expect(probe.mock.calls).toEqual([[['Arial', 'Deck Face']], [['Other']]]);
  });
});
