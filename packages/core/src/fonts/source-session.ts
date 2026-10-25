/**
 * `FontSourceSession`: one document's font sources in one painting realm.
 *
 * The generalisation of `GoogleFontPreloadLease` (`preload.ts`) to every
 * source kind. Library resource policy, not OOXML behaviour: ECMA-376 names
 * faces but leaves their provision to the host. Each {@link
 * FontSourceSession.ensure} takes a demand delta and
 *
 * 1. **resolves** every newly demanded family (authored and library fallback)
 *    by walking the sources in option order. The first source that declares a
 *    face for the family, directly or through one of its declared
 *    substitutes, claims it. A source that attempts a family and fails (a
 *    stylesheet that cannot be fetched or yields no rule, a resolver that
 *    throws, the 15 s ceiling) also makes the family final: it keeps its host
 *    fallback and no later source is consulted, so a failing mirror never
 *    falls through to another origin (#1497 R5) and first paint is
 *    deterministic. Only a source that does not serve the family (an
 *    allow-list miss, a family absent from a stylesheet or file list) passes
 *    it on;
 * 2. **admits** the claimed declarations against the generated ceilings
 *    (families per session, declarations per session, retained bytes per
 *    realm). A breach rejects with `OoxmlError('font-source-limit')`, releases
 *    the session and is rethrown by every later ensure: it is a misconfigured
 *    or hostile source, not network weather;
 * 3. **registers** every declaration of a claimed family, in declaration
 *    order, through the refcounted per-realm registry (`font-registry.ts`).
 *    Registering excluded subsets too keeps overlapping `unicode-range`
 *    precedence independent of the order demand arrives in;
 * 4. **loads** (`FontFace.load()`) only faces whose `unicode-range`
 *    intersects the demanded scalars (every face for `'all'`, and faces
 *    without a range always). Canvas never triggers a load itself, so an
 *    unloaded face paints nothing; later deltas load further subsets;
 * 5. **snapshots** {@link LoadedFontSources}: the faces that actually loaded,
 *    the substitutions whose served face loaded, and regular-name aliases.
 *
 * Failures are reported with one `console.warn` per newly failed family set;
 * nothing else is surfaced (no strict mode, no feedback API). Ensures are
 * serialised; stylesheets and resolutions are shared across sessions of the
 * same source identity (`source-cache.ts`), and faces across sessions of the
 * same realm (`font-registry.ts`). {@link FontSourceSession.release} returns
 * every reference the session took, exactly once.
 */
import { fontSourceLimitError } from '../errors/font-source-limit.js';
import { createAbortError } from '../internal/abort.js';
import {
  HARD_MAX_FONT_DEMAND_SCALARS,
  HARD_MAX_FONT_SOURCE_DECLARATIONS,
  HARD_MAX_FONT_SOURCE_FAMILIES,
} from '../worker/resource-policy.generated.js';
import { parseUnicodeRange, resolveCssSource } from './css-source.js';
import { claimFromIndex, indexFontDeclarations, resolveFontFiles } from './files-source.js';
import { releaseFaces, retainFace } from './font-registry.js';
import { activeFontSet, FONT_LOAD_CEILING_MS } from './preload.js';
import {
  acquireSourceCache,
  chargeRetainedBytes,
  fontContentKey,
  refundRetainedBytes,
  retainedFontSourceBytes,
  SharedTask,
  type FamilyResolution,
  type SourceCache,
} from './source-cache.js';
import {
  assertFontSourceDescriptors,
  dataUrlDecodedByteLength,
  fontResolverFunction,
  validateFontFaceDeclaration,
  type FontDemand,
  type FontFaceDeclaration,
  type FontResolveContext,
  type FontResolverSource,
  type FontSource,
} from './source.js';
import { FONT_SUBSTITUTE_SCRIPTS, type FontSubstituteScript } from './substitute-script.js';
import {
  fontFamilyKey,
  loadedRegularAliases,
  scopedFamilyKey,
  type FontSubstituteScripts,
  type FontSubstitution,
} from './substitutions.js';

/** One face a source supplied that loaded in this realm. */
export interface LoadedSourceFace {
  /** CSS family the face is registered under. */
  readonly family: string;
  readonly weight: string;
  readonly style: string;
  readonly stretch: string;
  readonly unicodeRange: string;
  /** Index of the supplying source in the session's `sources`. */
  readonly sourceIndex: number;
  readonly face: FontFace;
}

/** What the renderers consume after each ensure. */
export interface LoadedFontSources {
  /** Loaded faces, in registration order. */
  readonly faces: readonly LoadedSourceFace[];
  /** Lower-cased demanded family → served family, only for substitutions at
   * least one of whose served faces loaded. */
  readonly substitutions: ReadonlyMap<string, { readonly family: string; readonly script?: FontSubstituteScript }>;
  /** `loadedRegularAliases` over the loaded faces ('lato regular' → 'Lato'). */
  readonly regularAliases: ReadonlyMap<string, string>;
}

/** Runs one `font-resolver` source. The worker bridge supplies its own
 * implementation; the default calls this realm's function directly. */
export type FontResolverRunner = (
  source: FontResolverSource,
  demand: FontDemand,
  context: FontResolveContext,
) => readonly unknown[] | Promise<readonly unknown[]>;

export interface FontSourceSessionOptions {
  /** The `fonts` option value; validated here. */
  readonly sources: readonly FontSource[];
  /** Target set; defaults to this realm's (`document.fonts` / `self.fonts`). */
  readonly fontSet?: FontFaceSet | null;
  /** Aborting releases the session and rejects pending and later ensures. */
  readonly signal?: AbortSignal;
  readonly runResolver?: FontResolverRunner;
}

const EMPTY_SNAPSHOT: LoadedFontSources = Object.freeze({
  faces: Object.freeze([]),
  substitutions: new Map(),
  regularAliases: new Map(),
});

// ---- demand normalisation -----------------------------------------------------------

interface NormalizedDemand {
  /** Authored then fallback families, trimmed, first spelling per key, with
   * the substitute scripts their text proves. */
  readonly names: readonly DemandedName[];
  readonly tuples: FontDemand['tuples'];
  /** Sorted, unique, or 'all'. */
  readonly scalars: 'all' | readonly number[];
}

interface DemandedName {
  readonly name: string;
  readonly fallback: boolean;
  readonly scripts: FontSubstituteScripts;
}

const KNOWN_SUBSTITUTE_SCRIPTS: ReadonlySet<string> = new Set<FontSubstituteScript>(FONT_SUBSTITUTE_SCRIPTS);

function substituteScriptsOf(value: unknown): (name: string) => FontSubstituteScripts {
  if (value === undefined) return () => 'all';
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('font demand substituteScripts must be an object');
  }
  const scripts = new Map<string, readonly FontSubstituteScript[]>();
  for (const [family, list] of Object.entries(value)) {
    if (!Array.isArray(list) || !list.every((script) => typeof script === 'string' && KNOWN_SUBSTITUTE_SCRIPTS.has(script))) {
      throw new TypeError(`font demand substituteScripts[${JSON.stringify(family)}] must list substitute scripts`);
    }
    scripts.set(fontFamilyKey(family), Object.freeze([...new Set(list as FontSubstituteScript[])].sort()));
  }
  return (name) => scripts.get(fontFamilyKey(name)) ?? [];
}

function familyList(value: unknown, location: string): string[] {
  if (value === undefined) return [];
  if (!Array.isArray(value)) throw new TypeError(`${location} must be an array of family names`);
  return value.map((name: unknown, index) => {
    if (typeof name !== 'string') throw new TypeError(`${location}[${index}] must be a string`);
    return name.trim();
  });
}

/** Validate a demand delta and normalise its scalars (sorted, unique; more
 * than `HARD_MAX_FONT_DEMAND_SCALARS` becomes `'all'`, the conservative
 * answer). Library-produced demand never fails this; it guards the contract. */
export function normalizeFontDemand(demand: FontDemand): NormalizedDemand {
  if (!demand || typeof demand !== 'object') throw new TypeError('font demand must be an object');
  const names: DemandedName[] = [];
  const seen = new Set<string>();
  const scriptsOf = substituteScriptsOf(demand.substituteScripts);
  const add = (list: string[], fallback: boolean) => {
    for (const name of list) {
      const key = fontFamilyKey(name);
      if (!key || seen.has(key)) continue;
      seen.add(key);
      names.push({ name, fallback, scripts: scriptsOf(name) });
    }
  };
  add(familyList(demand.families, 'font demand families'), false);
  add(familyList(demand.fallbackFamilies, 'font demand fallbackFamilies'), true);
  const rawTuples: unknown = demand.tuples ?? [];
  if (!Array.isArray(rawTuples)) throw new TypeError('font demand tuples must be an array');
  const tuples = Object.freeze(rawTuples.map((tuple: unknown, index) => {
    const fields = (tuple ?? {}) as Partial<Record<'family' | 'weight' | 'style', unknown>>;
    if (typeof fields.family !== 'string' || typeof fields.weight !== 'number' || !Number.isFinite(fields.weight)
      || (fields.style !== 'normal' && fields.style !== 'italic')) {
      throw new TypeError(`font demand tuples[${index}] must be { family, weight, style }`);
    }
    return Object.freeze({ family: fields.family, weight: fields.weight, style: fields.style });
  }));
  let scalars: 'all' | readonly number[];
  if (demand.scalars === 'all') {
    scalars = 'all';
  } else {
    if (!Array.isArray(demand.scalars)) throw new TypeError("font demand scalars must be 'all' or an array");
    const unique = new Set<number>();
    for (const scalar of demand.scalars) {
      if (!Number.isInteger(scalar) || scalar < 0 || scalar > 0x10ffff) {
        throw new TypeError('font demand scalars must be integers in U+0000..U+10FFFF');
      }
      unique.add(scalar);
    }
    scalars = unique.size > HARD_MAX_FONT_DEMAND_SCALARS ? 'all' : [...unique].sort((a, b) => a - b);
  }
  return { names, tuples, scalars };
}

// ---- helpers ---------------------------------------------------------------------------

function rangeIntersects(ranges: readonly (readonly [number, number])[], sorted: readonly number[]): boolean {
  for (const [lo, hi] of ranges) {
    let left = 0, right = sorted.length;
    while (left < right) {
      const mid = (left + right) >>> 1;
      if (sorted[mid] < lo) left = mid + 1;
      else right = mid;
    }
    if (left < sorted.length && sorted[left] <= hi) return true;
  }
  return false;
}

/** A CSS `url()` for a validated absolute URL (data: URLs are kept verbatim
 * by validation and may contain characters a CSS string must escape). */
function cssUrl(url: string): string {
  return `url("${url.replace(/[\\"\n\r\f]/g, (char) =>
    char === '\\' || char === '"' ? `\\${char}` : `\\${char.charCodeAt(0).toString(16)} `)}")`;
}

const isAbortError = (error: unknown) => error instanceof Error && error.name === 'AbortError';

/** A signal aborted by `parent` or after `ms`; `dispose()` clears the timer. */
function ceilingSignal(parent: AbortSignal, ms: number): { signal: AbortSignal; expired: () => boolean; dispose: () => void } {
  const controller = new AbortController();
  let expired = false;
  const onAbort = () => controller.abort();
  parent.addEventListener('abort', onAbort, { once: true });
  const timer = setTimeout(() => { expired = true; controller.abort(); }, ms);
  if (parent.aborted) controller.abort();
  return {
    signal: controller.signal,
    expired: () => expired,
    // Abandon whatever still waits on this phase (a sibling rejected first).
    dispose: () => { clearTimeout(timer); parent.removeEventListener('abort', onAbort); controller.abort(); },
  };
}

/** Await `promise`, or stop waiting when `signal` aborts (the work itself
 * continues; its result is then ignored). */
function until<T>(promise: Promise<T>, signal: AbortSignal): Promise<T | undefined> {
  if (signal.aborted) return Promise.resolve(undefined);
  return new Promise((resolve, reject) => {
    const onAbort = () => resolve(undefined);
    signal.addEventListener('abort', onAbort, { once: true });
    promise.then(
      (value) => { signal.removeEventListener('abort', onAbort); resolve(value); },
      (error: unknown) => { signal.removeEventListener('abort', onAbort); reject(error); },
    );
  });
}

// ---- registration byte accounting -----------------------------------------------

/** Decoded bytes of `data` faces and `data:` URL faces registered in this
 * realm, per FontFaceSet and signature (a face shared by two documents is
 * charged once; each realm's set holds its own copy). The realm total is
 * shared with the source caches (`source-cache.ts`). */
const registeredBytes = new WeakMap<FontFaceSet, Map<string, { bytes: number; refs: number }>>();

function chargeBytes(set: FontFaceSet, signature: string, bytes: number): void {
  if (bytes === 0) return;
  let bySignature = registeredBytes.get(set);
  const existing = bySignature?.get(signature);
  if (existing) { existing.refs++; return; }
  chargeRetainedBytes(bytes);
  if (!bySignature) { bySignature = new Map(); registeredBytes.set(set, bySignature); }
  bySignature.set(signature, { bytes, refs: 1 });
}

function refundBytes(set: FontFaceSet, signature: string): void {
  const bySignature = registeredBytes.get(set);
  const entry = bySignature?.get(signature);
  if (!bySignature || !entry) return;
  if (--entry.refs > 0) return;
  bySignature.delete(signature);
  refundRetainedBytes(entry.bytes);
}

export { retainedFontSourceBytes };

// ---- session ------------------------------------------------------------------------

interface OwnedFace {
  readonly signature: string;
  readonly face: FontFace;
  readonly declaration: FontFaceDeclaration;
  readonly sourceIndex: number;
  /** null: no (or an unparseable) unicode-range, so the face covers all. */
  readonly ranges: readonly (readonly [number, number])[] | null;
  readonly charged: boolean;
  /** Demanded families this face serves (for warnings). */
  readonly families: Set<string>;
  readiness?: Promise<void>;
  /** Settled load outcome; undefined while loading. */
  outcome?: boolean;
}

type FamilyState =
  | { readonly kind: 'claimed'; readonly name: string; readonly sourceIndex: number; readonly served: string;
    readonly substitution?: FontSubstitution; readonly declarations: readonly FontFaceDeclaration[];
    readonly faces: OwnedFace[] }
  | { readonly kind: 'failed'; readonly name: string; readonly reason: string }
  /** No source declared a face for the family under substitution scope
   * `scripts`. A later demand that proves a wider scope resolves it again. */
  | { readonly kind: 'absent'; readonly name: string; readonly scripts: FontSubstituteScripts };

/** The union of two substitution scopes. */
function widenScope(a: FontSubstituteScripts, b: FontSubstituteScripts): FontSubstituteScripts {
  if (a === 'all' || b === 'all') return 'all';
  return Object.freeze([...new Set([...a, ...b])].sort());
}

/** Whether scope `b` admits a script-scoped substitution `a` does not. */
function scopeWidens(a: FontSubstituteScripts, b: FontSubstituteScripts): boolean {
  if (a === 'all') return false;
  return b === 'all' || b.some((script) => !a.includes(script));
}

export class FontSourceSession {
  // Cleared on release: descriptors hold font bytes (fontFiles data,
  // inline data: URLs) and the resolver runner may capture application data.
  private sources: readonly FontSource[];
  private caches: readonly SourceCache[];
  private fontSet: FontFaceSet | null;
  private runResolver: FontResolverRunner;
  private readonly controller = new AbortController();
  private detachSignal: () => void = () => undefined;
  private closed = false;
  private abortedByCaller = false;
  private failure: unknown = undefined;
  private queue: Promise<unknown> = Promise.resolve();
  private readonly families = new Map<string, FamilyState>();
  private readonly owned = new Map<string, OwnedFace>();
  private all = false;
  private readonly points = new Set<number>();
  private readonly warned = new Set<string>();
  private snapshot: LoadedFontSources = EMPTY_SNAPSHOT;

  constructor(options: FontSourceSessionOptions) {
    this.sources = assertFontSourceDescriptors(options.sources);
    this.fontSet = options.fontSet === undefined ? activeFontSet() : options.fontSet;
    this.runResolver = options.runResolver ?? runResolverInRealm;
    this.caches = this.sources.map(acquireSourceCache);
    const signal = options.signal;
    if (signal) {
      const onAbort = () => { this.abortedByCaller = true; this.release(); };
      if (signal.aborted) onAbort();
      else {
        signal.addEventListener('abort', onAbort, { once: true });
        this.detachSignal = () => signal.removeEventListener('abort', onAbort);
      }
    }
  }

  /** The snapshot after the latest completed ensure. */
  get loaded(): LoadedFontSources { return this.snapshot; }

  /**
   * Resolve, register and load one demand delta (families already handled by
   * this session are not resolved again; new scalars may load further subsets
   * of claimed families). Resolves with the snapshot; rejects with
   * `OoxmlError('font-source-limit')` on a ceiling breach and with an
   * `AbortError` once the caller's signal aborted. After {@link release} it
   * resolves with an empty snapshot.
   */
  ensure(demand: FontDemand): Promise<LoadedFontSources> {
    let normalized: NormalizedDemand;
    try { normalized = normalizeFontDemand(demand); }
    catch (error) { return Promise.reject(error); }
    const task = this.queue.then(() => this.ensureNow(normalized));
    // The queue only orders ensures: it must never hold a settled snapshot
    // (its FontFaces) past the caller's own reference.
    this.queue = task.then(() => undefined, () => undefined);
    return task;
  }

  /** Release every face, byte charge and cache reference; idempotent. */
  release(): void {
    if (this.closed) return;
    this.closed = true;
    this.controller.abort();
    this.detachSignal();
    releaseFaces([...this.owned.values()].map((entry) => entry.face));
    if (this.fontSet) for (const entry of this.owned.values()) if (entry.charged) refundBytes(this.fontSet, entry.signature);
    this.owned.clear();
    this.families.clear();
    this.points.clear();
    for (const cache of this.caches) cache.release();
    // A released session object the application still references must not
    // keep the caches (and what they index) reachable.
    this.caches = [];
    // Nothing that holds bytes or FontFaces outlives the release: the source
    // descriptors, the target set, the resolver runner, the warned names.
    this.sources = [];
    this.fontSet = null;
    this.runResolver = releasedRunner;
    this.warned.clear();
    this.snapshot = EMPTY_SNAPSHOT;
  }

  private settledState(): LoadedFontSources {
    if (this.abortedByCaller) throw createAbortError('Font source session aborted');
    if (this.failure !== undefined) throw this.failure;
    return this.snapshot;
  }

  private async ensureNow(demand: NormalizedDemand): Promise<LoadedFontSources> {
    if (this.closed || this.failure !== undefined) return this.settledState();
    if (!this.fontSet || typeof FontFace === 'undefined' || this.sources.length === 0) return this.snapshot;
    try {
      // A family is resolved once, except that an absent family is resolved
      // again when a later delta proves a substitution scope its earlier
      // resolution excluded (progressive loading: Latin text on one slide,
      // Arabic text in the same family on a later one), under the union of
      // both scopes. A claimed or failed family is final.
      const fresh: DemandedName[] = [];
      let added = 0;
      for (const entry of demand.names) {
        const prior = this.families.get(fontFamilyKey(entry.name));
        if (!prior) { fresh.push(entry); added++; continue; }
        if (prior.kind === 'absent' && scopeWidens(prior.scripts, entry.scripts)) {
          fresh.push({ ...entry, scripts: widenScope(prior.scripts, entry.scripts) });
        }
      }
      if (this.families.size + added > HARD_MAX_FONT_SOURCE_FAMILIES) {
        throw fontSourceLimitError('families', HARD_MAX_FONT_SOURCE_FAMILIES, this.families.size + added);
      }
      const delta = this.addScalars(demand.scalars);
      const resolved = fresh.length ? await this.resolve(fresh, demand.tuples) : new Map<string, FamilyState>();
      if (this.closed) return this.settledState();
      const registered = new Set<OwnedFace>();
      const failures = new Map<string, string>();
      for (const { name, scripts } of fresh) {
        const key = fontFamilyKey(name);
        const state = resolved.get(key) ?? { kind: 'absent', name, scripts };
        this.families.set(key, state);
        if (state.kind === 'failed') failures.set(name, state.reason);
        if (state.kind === 'claimed') this.register(state, registered, failures);
      }
      await this.load(delta, registered, failures);
      if (this.closed) return this.settledState();
      this.warn(failures);
      this.snapshot = this.takeSnapshot();
      return this.snapshot;
    } catch (error) {
      if (this.closed) return this.settledState();
      this.failure = error;
      this.release();
      throw error;
    }
  }

  /** Merge demanded scalars; returns the newly demanded ones, or 'all' when
   * the session just became all-scalar (explicitly, or because the cumulative
   * set outgrew `HARD_MAX_FONT_DEMAND_SCALARS`). */
  private addScalars(scalars: 'all' | readonly number[]): 'all' | readonly number[] {
    if (this.all) return [];
    if (scalars !== 'all') {
      const delta: number[] = [];
      for (const scalar of scalars) if (!this.points.has(scalar)) { this.points.add(scalar); delta.push(scalar); }
      if (this.points.size <= HARD_MAX_FONT_DEMAND_SCALARS) return delta;
    }
    this.all = true;
    this.points.clear();
    return 'all';
  }

  /** Walk the sources in order for the fresh families (claim once). */
  private async resolve(
    fresh: NormalizedDemand['names'],
    tuples: FontDemand['tuples'],
  ): Promise<Map<string, FamilyState>> {
    const result = new Map<string, FamilyState>();
    let pending = fresh;
    const ceiling = ceilingSignal(this.controller.signal, FONT_LOAD_CEILING_MS);
    try {
      for (let sourceIndex = 0; sourceIndex < this.sources.length && pending.length; sourceIndex++) {
        if (this.closed) throw createAbortError('Font source session released');
        let answers: ReadonlyMap<string, FamilyResolution>;
        try {
          answers = await this.resolveWithSource(sourceIndex, pending, tuples, ceiling.signal);
        } catch (error) {
          if (!isAbortError(error) || !ceiling.expired()) throw error;
          for (const { name } of pending) result.set(fontFamilyKey(name), { kind: 'failed', name, reason: 'timed out' });
          pending = [];
          break;
        }
        const next: typeof pending[number][] = [];
        for (const entry of pending) {
          const key = fontFamilyKey(entry.name);
          const answer = answers.get(key) ?? { kind: 'absent' };
          if (answer.kind === 'absent') next.push(entry);
          else if (answer.kind === 'failed') result.set(key, { kind: 'failed', name: entry.name, reason: answer.reason });
          else {
            result.set(key, {
              kind: 'claimed', name: entry.name, sourceIndex, served: answer.served,
              ...(answer.substitution ? { substitution: answer.substitution } : {}),
              declarations: answer.declarations, faces: [],
            });
          }
        }
        pending = next;
      }
    } finally {
      ceiling.dispose();
    }
    return result;
  }

  private async resolveWithSource(
    sourceIndex: number,
    families: NormalizedDemand['names'],
    tuples: FontDemand['tuples'],
    signal: AbortSignal,
  ): Promise<ReadonlyMap<string, FamilyResolution>> {
    const source = this.sources[sourceIndex];
    const cache = this.caches[sourceIndex];
    const names = families.map(({ name }) => name);
    const scripts = new Map(families.map(({ name, scripts: proven }) => [fontFamilyKey(name), proven]));
    const scriptsOf = (name: string): FontSubstituteScripts => scripts.get(fontFamilyKey(name)) ?? 'all';
    switch (source.kind) {
      case 'css-api':
      case 'font-face-css':
        return resolveCssSource(source, cache, names, signal, scriptsOf);
      case 'font-files':
        return resolveFontFiles(source, cache, names, scriptsOf);
      case 'font-resolver':
        return this.resolveWithResolver(source, cache, families, tuples, signal);
    }
  }

  /** One resolver call for the families no session has resolved through this
   * source identity yet; cached families join earlier results. The batch
   * result is handed out per family (each family takes and removes its own
   * entry) and only per-family results are cached, so nothing outlives the
   * call except what each family claimed and was charged for: unclaimed or
   * rejected declarations are unreachable once the call settles. */
  private async resolveWithResolver(
    source: FontResolverSource,
    cache: SourceCache,
    families: NormalizedDemand['names'],
    tuples: FontDemand['tuples'],
    signal: AbortSignal,
  ): Promise<ReadonlyMap<string, FamilyResolution>> {
    const needed: NormalizedDemand['names'] = families.filter(({ name, scripts }) =>
      !cache.families.has(scopedFamilyKey(name, scripts)));
    let call: SharedTask<Map<string, FamilyResolution>> | undefined;
    if (needed.length) {
      const neededKeys = new Set(needed.map(({ name }) => fontFamilyKey(name)));
      const demand: FontDemand = Object.freeze({
        families: Object.freeze(needed.filter((entry) => !entry.fallback).map(({ name }) => name)),
        fallbackFamilies: Object.freeze(needed.filter((entry) => entry.fallback).map(({ name }) => name)),
        tuples: Object.freeze(tuples.filter((tuple) => neededKeys.has(fontFamilyKey(tuple.family)))),
        scalars: this.all ? 'all' : Object.freeze([...this.points].sort((a, b) => a - b)),
        ...(needed.some(({ scripts }) => scripts !== 'all') ? {
          substituteScripts: Object.freeze(Object.fromEntries(needed.flatMap(({ name, scripts }) =>
            scripts === 'all' ? [] : [[fontFamilyKey(name), scripts] as const]))),
        } : {}),
      });
      const runner = this.runResolver;
      call = new SharedTask((callSignal) => callResolver(runner, source, demand, callSignal), () => undefined);
    }
    const entries = await Promise.all(families.map(async ({ name, scripts }) => {
      const key = fontFamilyKey(name);
      const answer = await cache.families.join(scopedFamilyKey(name, scripts), async (taskSignal) => {
        if (!call) return { kind: 'absent' } satisfies FamilyResolution;
        const batch = await call.join(taskSignal);
        const own = batch.get(key) ?? { kind: 'absent' };
        batch.delete(key);
        return own;
      }, signal, cache);
      return [key, answer] as const;
    }));
    return new Map(entries);
  }

  /** Admit and register one claimed family's declarations, in order. */
  private register(state: Extract<FamilyState, { kind: 'claimed' }>, registered: Set<OwnedFace>, failures: Map<string, string>): void {
    const set = this.fontSet as FontFaceSet;
    const sourceId = this.caches[state.sourceIndex].id;
    // Registry signature (design §4.4): source identity × family × every
    // descriptor × content, so two origins never share a face and two
    // documents on one source share one face per FontFaceSet.
    const signed = state.declarations.map((declaration) => {
      const { data, url = '' } = declaration;
      const dataUrl = data === undefined && /^\s*data:/i.test(url);
      const signature = [
        `src:${sourceId}`,
        fontFamilyKey(declaration.family),
        declaration.weight ?? '',
        declaration.style ?? '',
        declaration.stretch ?? '',
        declaration.unicodeRange ?? '',
        data !== undefined ? fontContentKey(data) : dataUrl ? fontContentKey(url) : url,
      ].join('|');
      return { declaration, signature, data, url, dataUrl };
    });
    // A face registered later in a FontFaceSet takes precedence over an
    // earlier one with the same descriptors (CSS Fonts 4 §4.10 / Font Loading
    // set order), so an identical declaration repeated in the list keeps the
    // position of its LAST occurrence: `A, B, A` registers `B, A`.
    const lastIndex = new Map(signed.map(({ signature }, index) => [signature, index]));
    for (const [index, { declaration, signature, data, url, dataUrl }] of signed.entries()) {
      if (lastIndex.get(signature) !== index) continue;
      const existing = this.owned.get(signature);
      if (existing) {
        existing.families.add(state.name);
        state.faces.push(existing);
        continue;
      }
      if (this.owned.size >= HARD_MAX_FONT_SOURCE_DECLARATIONS) {
        throw fontSourceLimitError('declarations', HARD_MAX_FONT_SOURCE_DECLARATIONS, this.owned.size + 1);
      }
      const bytes = data !== undefined ? data.byteLength : dataUrl ? dataUrlDecodedByteLength(url) ?? 0 : 0;
      chargeBytes(set, signature, bytes);
      let face: FontFace;
      try {
        ({ face } = retainFace(signature, set, () => {
          const descriptors: FontFaceDescriptors = {};
          if (declaration.weight) descriptors.weight = declaration.weight;
          if (declaration.style) descriptors.style = declaration.style;
          if (declaration.stretch) descriptors.stretch = declaration.stretch;
          if (declaration.unicodeRange) descriptors.unicodeRange = declaration.unicodeRange;
          const created = new FontFace(declaration.family, data ?? cssUrl(url), descriptors);
          try { set.add(created); }
          catch (error) { set.delete(created); throw error; }
          return created;
        }));
      } catch (error) {
        refundBytes(set, signature);
        failures.set(state.name, error instanceof Error ? error.message : String(error));
        continue;
      }
      const entry: OwnedFace = {
        signature, face, declaration, sourceIndex: state.sourceIndex,
        ranges: parseUnicodeRange(declaration.unicodeRange),
        charged: bytes > 0,
        families: new Set([state.name]),
      };
      this.owned.set(signature, entry);
      state.faces.push(entry);
      registered.add(entry);
    }
  }

  /** Load faces whose ranges intersect the demand (lease semantics: a face
   * registered now is checked against all scalars demanded so far, an older
   * unloaded face only against the delta). */
  private async load(delta: 'all' | readonly number[], registered: Set<OwnedFace>, failures: Map<string, string>): Promise<void> {
    const becameAll = delta === 'all';
    const deltaIndex = delta === 'all' ? [] : [...delta].sort((a, b) => a - b);
    const cumulative = registered.size && !this.all ? [...this.points].sort((a, b) => a - b) : deltaIndex;
    const waits: OwnedFace[] = [];
    for (const entry of this.owned.values()) {
      if (entry.readiness) { waits.push(entry); continue; }
      const isNew = registered.has(entry);
      const eligible = (this.all && (becameAll || isNew)) || entry.ranges === null
        || rangeIntersects(entry.ranges, isNew ? cumulative : deltaIndex);
      if (!eligible) continue;
      // load() is idempotent and joins a shared face's in-flight load
      // (CSS Font Loading §2.2), so a reused registration reports its own result.
      entry.readiness = Promise.resolve()
        .then(() => (this.closed ? undefined : entry.face.load()))
        .then(() => { entry.outcome = true; }, () => { entry.outcome = false; });
      waits.push(entry);
    }
    if (!waits.length) return;
    const ceiling = ceilingSignal(this.controller.signal, FONT_LOAD_CEILING_MS);
    try {
      await until(Promise.all(waits.map((entry) => entry.readiness)), ceiling.signal);
      if (this.closed) return;
      for (const entry of waits) {
        if (entry.outcome === true) continue;
        const reason = entry.outcome === false ? 'font load failed' : 'timed out';
        for (const family of entry.families) if (!failures.has(family)) failures.set(family, reason);
      }
    } finally {
      ceiling.dispose();
    }
  }

  private warn(failures: Map<string, string>): void {
    const fresh = [...failures].filter(([family]) => !this.warned.has(fontFamilyKey(family)));
    if (!fresh.length) return;
    for (const [family] of fresh) this.warned.add(fontFamilyKey(family));
    console.warn(`[ooxml] failed to preload web font(s): ${fresh.map(([family, reason]) => `${family} (${reason})`)
      .join(', ')}; falling back to system fonts (text may shift or differ).`);
  }

  private takeSnapshot(): LoadedFontSources {
    const faces: LoadedSourceFace[] = [];
    for (const entry of this.owned.values()) {
      if (entry.face.status !== 'loaded') continue;
      const { declaration } = entry;
      faces.push(Object.freeze({
        family: declaration.family,
        weight: declaration.weight ?? 'normal',
        style: declaration.style ?? 'normal',
        stretch: declaration.stretch ?? 'normal',
        unicodeRange: declaration.unicodeRange ?? 'U+0-10FFFF',
        sourceIndex: entry.sourceIndex,
        face: entry.face,
      }));
    }
    const substitutions = new Map<string, { readonly family: string; readonly script?: FontSubstituteScript }>();
    for (const [key, state] of this.families) {
      if (state.kind !== 'claimed' || !state.substitution) continue;
      if (!state.faces.some((entry) => entry.face.status === 'loaded')) continue;
      const script = state.substitution.script;
      substitutions.set(key, Object.freeze({ family: state.served, ...(script ? { script } : {}) }));
    }
    return Object.freeze({
      faces: Object.freeze(faces),
      substitutions,
      regularAliases: loadedRegularAliases(faces.map((face) => ({ ...face, status: 'loaded' as const }))),
    });
  }
}

/**
 * Run one shared resolver call. Module-level and given its runner explicitly:
 * the call is shared by every session waiting on the same families, so it
 * must not read any session's state. The runner of the session that created
 * the call is kept by the call until it settles, so releasing that session
 * (which drops its own runner reference) never breaks the other waiters.
 */
async function callResolver(
runResolver: FontResolverRunner,
source: FontResolverSource,
  demand: FontDemand,
  signal: AbortSignal,
): Promise<Map<string, FamilyResolution>> {
  const failAll = (reason: string) => new Map(
    [...demand.families, ...demand.fallbackFamilies].map((name) =>
      [fontFamilyKey(name), { kind: 'failed', reason } satisfies FamilyResolution]));
  let raw: readonly unknown[];
  try {
    raw = await runResolver(source, demand, Object.freeze({ signal }));
  } catch (error) {
    if (signal.aborted) throw error;
    return failAll(error instanceof Error ? error.message : String(error));
  }
  if (signal.aborted) throw createAbortError('Font resolver abandoned');
  if (!Array.isArray(raw)) return failAll('resolver did not return an array');
  if (raw.length > HARD_MAX_FONT_SOURCE_DECLARATIONS) {
    throw fontSourceLimitError('declarations', HARD_MAX_FONT_SOURCE_DECLARATIONS, raw.length);
  }
  let declarations: FontFaceDeclaration[];
  try {
    declarations = raw.map((entry, index) =>
      validateFontFaceDeclaration(entry, `font resolver declaration ${index}`));
  } catch (error) {
    if (!(error instanceof TypeError)) throw error;
    return failAll(error.message);
  }
  const index = indexFontDeclarations(declarations);
  const scriptsOf = substituteScriptsOf(demand.substituteScripts);
  const restricted = demand.substituteScripts !== undefined;
  return new Map([...demand.families, ...demand.fallbackFamilies].map((name) =>
    [fontFamilyKey(name), claimFromIndex(index, name, source.substitutes, restricted ? scriptsOf(name) : 'all')]));
}

const releasedRunner: FontResolverRunner = () => {
  throw createAbortError('Font source session released');
};

/** Main-realm resolver execution: the application function behind a
 * descriptor this realm's `fontResolver()` issued. */
const runResolverInRealm: FontResolverRunner = (source, demand, context) => {
  const resolve = fontResolverFunction(source);
  if (!resolve) throw new Error('font resolver function is not available in this realm');
  return resolve(demand, context);
};
