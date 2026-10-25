import { describe, expect, it } from 'vitest';
import { pptxFontDemand } from './font-source-load.js';

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
