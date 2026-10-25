/**
 * PPTX font-source loading (`LoadOptions.fonts`): the demand one presentation
 * (or one progressive slide delta) asks the configured sources for. Library
 * resource policy, not OOXML behaviour: DrawingML names faces (ECMA-376
 * §21.1.2.3) and leaves their provision to the host.
 *
 * The main thread imports this module lazily, only when sources are
 * configured, so the default offline path never loads the session or the CSS
 * adapters. The main thread and the render worker build the same demand from
 * the same preflight facts with the same descriptors, so both realms demand,
 * register and load identical faces.
 */
import { normalizeLocalFontMetricFamily, type FontPreloadDemand } from '@silurus/ooxml-core';
import {
  probeFontFamilyPresence,
  type FontDemand,
  type FontFamilyPresence,
} from '@silurus/ooxml-core/internal/font-sources';
import { excludeEmbeddedFontFamilies } from './embedded-fonts.js';
import type { PptxFontDemandFacts } from './font-requests.js';

export { FontSourceSession } from '@silurus/ooxml-core/internal/font-sources';

const CSS_GENERIC_FAMILY_NAMES: ReadonlySet<string> = new Set([
  'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui',
]);

/** A preflight name that names a face a source could serve: not a DrawingML
 * theme token (`+mj-lt`, `+mn-ea`, …, §20.1.4.1.17; its theme font is
 * demanded by name) and not a CSS generic family. */
function demandableFamily(name: string | null): name is string {
  const family = name?.trim();
  return !!family && !family.startsWith('+mj-') && !family.startsWith('+mn-')
    && !CSS_GENERIC_FAMILY_NAMES.has(family.toLowerCase());
}

/**
 * One presentation's font demand (or one progressive delta's: families
 * already handled by the session are not resolved again, and `scalars` is the
 * newly demanded scalar delta).
 *
 * - `families` / `fallbackFamilies`: the preflight's authored families and
 *   library script fallbacks (`fontPreloadNames`, split at
 *   `facts.familyCount`), minus families an embedded font part
 *   (ECMA-376 Part 1 §19.2.1.9 `p:embeddedFont`) already serves;
 * - `tuples`: the exact authored face tuples the slides request;
 * - `scalars`: the collected Unicode scalars (`'all'` when the collector
 *   declined or the preflight omitted them, the conservative answer);
 * - `substituteScripts`: per family, the scoped-substitute scripts its painted
 *   text proves, so a script-scoped substitute is never requested for text it
 *   would not paint.
 */
export function pptxFontDemand(
  names: readonly (string | null)[],
  facts: PptxFontDemandFacts | undefined,
  scalars: FontPreloadDemand | undefined,
  embeddedAliases: ReadonlyMap<string, string>,
): FontDemand {
  const familyCount = facts?.familyCount ?? names.length;
  const families = excludeEmbeddedFontFamilies(names.slice(0, familyCount), embeddedAliases).filter(demandableFamily);
  const fallbackFamilies = excludeEmbeddedFontFamilies(names.slice(familyCount), embeddedAliases).filter(demandableFamily);
  return Object.freeze({
    families: Object.freeze(families),
    fallbackFamilies: Object.freeze(fallbackFamilies),
    tuples: facts?.tuples ?? Object.freeze([]),
    scalars: scalars ?? 'all',
    ...(facts ? { substituteScripts: facts.substituteScripts } : {}),
  });
}

/** Presence of `families` (normalised keys): installed ones, and every one
 * whose probe completed. */
export type InstalledFamilyProbe = (families: readonly string[]) => Promise<FontFamilyPresence>;

/** The shared family presence probe (`local()`), batched so every family is
 * probed, against this realm's font set. */
const probeInstalledFamilies: InstalledFamilyProbe = (families) => probeFontFamilyPresence(families);

/**
 * Library policy (design §3.2, §8): the authored family names the face to use,
 * so an installed authored face is never displaced by a source face. A source
 * face registers under its real CSS family, and a document font shadows a
 * system font of the same name, so a source serving an installed family (an
 * application face named "Arial") would otherwise replace it. Every authored
 * family is probed, before any source face of the delta registers; only a
 * family whose probe completed and found no installed face is demanded and
 * cached. A family whose probe did not complete (deadline) is left out of
 * this delta, so the host keeps it, and is probed again with the next delta:
 * an unproven family never loads a source face that might shadow it. Library
 * fallback families stay demanded, as in DOCX.
 */
export class PptxInstalledFamilies {
  private readonly installed = new Map<string, Promise<boolean | undefined>>();

  constructor(private readonly probe: InstalledFamilyProbe = probeInstalledFamilies) {}

  async exclude(demand: FontDemand): Promise<FontDemand> {
    const fresh = [...new Set(demand.families.map(normalizeLocalFontMetricFamily))]
      .filter((key) => !this.installed.has(key));
    if (fresh.length) {
      const names = new Map(demand.families.map((family) => [normalizeLocalFontMetricFamily(family), family]));
      const probed = this.probe(fresh.map((key) => names.get(key) as string))
        .catch((): FontFamilyPresence => ({ installed: new Set(), checked: new Set() }));
      for (const key of fresh) {
        const outcome = probed.then(({ installed, checked }) => (checked.has(key) ? installed.has(key) : undefined));
        this.installed.set(key, outcome);
        // Unproven families are not cached: the next delta probes them again.
        void outcome.then((value) => { if (value === undefined && this.installed.get(key) === outcome) this.installed.delete(key); });
      }
    }
    // Demand only families proven missing (false); installed (true) and
    // unproven (undefined) families keep the host face.
    const keep = await Promise.all(demand.families.map(async (family) =>
      (await this.installed.get(normalizeLocalFontMetricFamily(family))) === false));
    const families = demand.families.filter((_, index) => keep[index]);
    return families.length === demand.families.length ? demand : Object.freeze({ ...demand, families: Object.freeze(families) });
  }
}
