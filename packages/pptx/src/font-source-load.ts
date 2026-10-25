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
import {
  loadOfficeFontFallbacks,
  normalizeLocalFontMetricFamily,
  type FontPreloadDemand,
} from '@silurus/ooxml-core';
import type { FontDemand } from '@silurus/ooxml-core/internal/font-sources';
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

/** Normalised families of `families` whose face is installed on the host. */
export type InstalledFamilyProbe = (families: readonly string[]) => Promise<readonly string[]>;

/** The shared family presence probe (`local()`, bounded by the office-fallback
 * preflight limits) against this realm's font set. */
const probeInstalledFamilies: InstalledFamilyProbe = async (families) =>
  (await loadOfficeFontFallbacks(families.map((family) => ({ family, presenceOnly: true })))).installed ?? [];

/**
 * Library policy (design §3.2, §8): the authored family names the face to use,
 * so an installed authored face is never displaced by a source face. A source
 * face registers under its real CSS family, and a document font shadows a
 * system font of the same name, so a source serving an installed family (an
 * application face named "Arial") would otherwise replace it. Each authored
 * family is probed once per presentation and realm, before any source face
 * of the delta registers; installed families leave the demand. Library
 * fallback families stay demanded, as in DOCX.
 */
export class PptxInstalledFamilies {
  private readonly installed = new Map<string, Promise<boolean>>();

  constructor(private readonly probe: InstalledFamilyProbe = probeInstalledFamilies) {}

  async exclude(demand: FontDemand): Promise<FontDemand> {
    const fresh = [...new Set(demand.families.map(normalizeLocalFontMetricFamily))]
      .filter((key) => !this.installed.has(key));
    if (fresh.length) {
      const names = new Map(demand.families.map((family) => [normalizeLocalFontMetricFamily(family), family]));
      const probed = this.probe(fresh.map((key) => names.get(key) as string))
        .then((found) => new Set(found), () => new Set<string>());
      for (const key of fresh) this.installed.set(key, probed.then((found) => found.has(key)));
    }
    const keep = await Promise.all(demand.families.map(async (family) =>
      !(await this.installed.get(normalizeLocalFontMetricFamily(family)))));
    const families = demand.families.filter((_, index) => keep[index]);
    return families.length === demand.families.length ? demand : Object.freeze({ ...demand, families: Object.freeze(families) });
  }
}
