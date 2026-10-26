/**
 * `fontResolver` worker bridge: main-thread half (`resolver-bridge.ts` has the
 * protocol). Loaded lazily, on a document's first bridged request, by
 * {@link FontResolverBridgeHost} (`resolver-bridge-host.ts`).
 *
 * Library API policy, not OOXML behaviour. For each `ooxmlResolveFonts`
 * request the responder looks the resolver id up among the document's own
 * `fonts` descriptors, so a worker can only run functions this document was
 * configured with, and resolves the demanded families through this realm's
 * resolution cache (`resolver-resolution.ts`, design §4.4): a family another
 * requester of this realm already resolved or is resolving (an XLSX viewer
 * realm session, another document sharing the resolver) is joined, not asked
 * again, so the application function runs once per family in both modes. A
 * new call receives a copy of the demand (`demand.scalars` bounded by
 * `HARD_MAX_FONT_DEMAND_SCALARS`, family records without a prototype) and an
 * `AbortSignal` that the worker's cancellation, document destroy and the 15 s
 * font ceiling abort (a shared call is aborted only when every requester left).
 * The reply carries each family's claiming declarations, failure or absence,
 * or the breached ceiling.
 *
 * Byte ownership (design §4.3):
 * - When the worker is the only painting realm (DOCX and PPTX worker mode),
 *   the bridge holds a cache entry only for the duration of its request.
 *   Afterwards a buffer no kept entry references (and no `fontFiles`
 *   descriptor) is transferred, so main keeps nothing; one still referenced
 *   (another document of this realm holds the family) is copied.
 * - When this realm paints as well (the XLSX viewer registers the same faces
 *   in `document.fonts`), the bridge holds its entries for the document's
 *   lifetime alongside the viewer session, main keeps the originals, and one
 *   copy per distinct buffer is transferred.
 * A cached buffer is therefore never detached. Main's cache charges what main
 * keeps and the worker's session what the worker keeps, each once; entries
 * are released with the bridge (`dispose`). A reply for a request the worker
 * cancelled, or produced after the document was destroyed, is never posted.
 */
import type { OoxmlError } from '../errors/ooxml-error.js';
import { nullRecord } from '../internal/own-key.js';
import {
  HARD_MAX_FONT_DEMAND_SCALARS,
  HARD_MAX_FONT_SOURCE_FAMILIES,
} from '../worker/resource-policy.generated.js';
import { FONT_LOAD_CEILING_MS } from './font-registry.js';
import type {
  BridgedFamilyAnswer,
  FontResolveOutcome,
  MainFontResolveReply,
  WorkerFontResolveRequest,
} from './resolver-bridge.js';
import {
  resolveFamiliesThroughCache,
  runResolverInRealm,
  substituteScriptsOf,
  type ResolverDemandedName,
} from './resolver-resolution.js';
import { acquireSourceCache, isFontBufferCached, type SourceCache } from './source-cache.js';
import {
  fontResolverFunction,
  isFontBufferHeld,
  type FontDemand,
  type FontFaceDeclaration,
  type FontResolverSource,
  type FontSource,
} from './source.js';
import { FONT_SUBSTITUTE_SCRIPTS, type FontSubstituteScript } from './substitute-script.js';

type Post = (message: unknown, transfer?: Transferable[]) => void;

export interface FontResolverResponderOptions {
  /** Posts to the document's render worker. */
  readonly post: Post;
  /** The document's validated `fonts` option (main-realm descriptors). */
  readonly sources: readonly FontSource[];
  /** Keep every original buffer and transfer copies: this realm also
   * registers the faces (XLSX viewer realm). */
  readonly copyData: boolean;
}

const KNOWN_SCRIPTS: ReadonlySet<string> = new Set<string>(FONT_SUBSTITUTE_SCRIPTS);

const isNonNegativeSafeInteger = (value: unknown): value is number =>
  typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;

function names(value: unknown): readonly string[] | undefined {
  if (!Array.isArray(value) || value.length > HARD_MAX_FONT_SOURCE_FAMILIES) return undefined;
  return value.every((name) => typeof name === 'string') ? Object.freeze([...value as string[]]) : undefined;
}

/**
 * The application's copy of a bridged demand, or undefined when the worker
 * sent something no library session produces. Record-valued fields are
 * rebuilt without a prototype (keys are document family names); more than
 * `HARD_MAX_FONT_DEMAND_SCALARS` scalars become `'all'`, as in the session.
 */
export function copyBridgedFontDemand(value: unknown): FontDemand | undefined {
  if (value === null || typeof value !== 'object') return undefined;
  const fields = value as Partial<Record<keyof FontDemand, unknown>>;
  const families = names(fields.families);
  const fallbackFamilies = names(fields.fallbackFamilies);
  if (!families || !fallbackFamilies || families.length + fallbackFamilies.length > HARD_MAX_FONT_SOURCE_FAMILIES) {
    return undefined;
  }
  if (!Array.isArray(fields.tuples)) return undefined;
  const tuples: FontDemand['tuples'][number][] = [];
  for (const tuple of fields.tuples as unknown[]) {
    const entry = (tuple ?? {}) as Partial<Record<'family' | 'weight' | 'style', unknown>>;
    if (typeof entry.family !== 'string' || typeof entry.weight !== 'number' || !Number.isFinite(entry.weight)
      || (entry.style !== 'normal' && entry.style !== 'italic')) return undefined;
    tuples.push(Object.freeze({ family: entry.family, weight: entry.weight, style: entry.style }));
  }
  let scalars: FontDemand['scalars'];
  if (fields.scalars === 'all') scalars = 'all';
  else if (!Array.isArray(fields.scalars)) return undefined;
  else if (fields.scalars.length > HARD_MAX_FONT_DEMAND_SCALARS) scalars = 'all';
  else {
    if (!fields.scalars.every((scalar) => Number.isInteger(scalar) && scalar >= 0 && scalar <= 0x10ffff)) return undefined;
    scalars = Object.freeze([...fields.scalars as number[]]);
  }
  let substituteScripts: Record<string, readonly FontSubstituteScript[]> | undefined;
  if (fields.substituteScripts !== undefined) {
    const record = fields.substituteScripts;
    if (record === null || typeof record !== 'object' || Array.isArray(record)) return undefined;
    const entries = Object.entries(record);
    if (entries.length > HARD_MAX_FONT_SOURCE_FAMILIES) return undefined;
    substituteScripts = nullRecord();
    for (const [family, list] of entries) {
      if (!Array.isArray(list) || !list.every((script) => typeof script === 'string' && KNOWN_SCRIPTS.has(script))) {
        return undefined;
      }
      substituteScripts[family] = Object.freeze([...list as FontSubstituteScript[]]);
    }
    Object.freeze(substituteScripts);
  }
  return Object.freeze({
    families,
    fallbackFamilies,
    tuples: Object.freeze(tuples),
    scalars,
    ...(substituteScripts ? { substituteScripts } : {}),
  });
}

/** Ownership of the answered buffers: originals moved, or one copy each. */
function transferable(
  answers: readonly BridgedFamilyAnswer[],
  copyData: boolean,
): { answers: readonly BridgedFamilyAnswer[]; transfer: ArrayBuffer[] } {
  const sent = new Map<ArrayBuffer, ArrayBuffer>();
  const wire = (declaration: FontFaceDeclaration): FontFaceDeclaration => {
    const { data } = declaration;
    if (!data) return declaration;
    let out = sent.get(data);
    if (!out) {
      out = copyData || isFontBufferHeld(data) || isFontBufferCached(data) ? data.slice(0) : data;
      sent.set(data, out);
    }
    return out === data ? declaration : { ...declaration, data: out };
  };
  const out = answers.map((answer) => (answer.kind === 'claimed'
    ? { ...answer, declarations: answer.declarations.map(wire) } : answer));
  return { answers: out, transfer: [...sent.values()] };
}

function isFontSourceLimit(error: unknown): error is OoxmlError & { violation: NonNullable<OoxmlError['violation']> } {
  return error instanceof Error && (error as OoxmlError).code === 'font-source-limit'
    && (error as OoxmlError).violation !== undefined;
}

/** One document's main-side bridge. */
export class FontResolverResponder {
  private readonly post: Post;
  private sources: readonly FontSource[];
  private readonly copyData: boolean;
  private disposed = false;
  /** In-flight calls by `${generation}:${requestId}` (library-generated keys). */
  private readonly calls = new Map<string, AbortController>();
  /** This document's hold on each resolver's realm cache, by resolver id. */
  private readonly caches = new Map<string, SourceCache>();

  constructor(options: FontResolverResponderOptions) {
    this.post = options.post;
    this.sources = options.sources;
    this.copyData = options.copyData;
  }

  /** Handle one worker → main bridge message (already type-guarded). */
  accept(message: { readonly kind: string }): void {
    if (this.disposed) return;
    const fields = message as Partial<WorkerFontResolveRequest>;
    if (!isNonNegativeSafeInteger(fields.generation) || !isNonNegativeSafeInteger(fields.requestId)) return;
    const key = `${fields.generation}:${fields.requestId}`;
    if (message.kind === 'ooxmlCancelResolveFonts') {
      this.calls.get(key)?.abort();
      this.calls.delete(key);
      return;
    }
    if (message.kind !== 'ooxmlResolveFonts' || this.calls.has(key)) return;
    void this.respond(key, fields.generation, fields.requestId, fields.resolverId, fields.demand);
  }

  /** The document is destroyed: abort every call and return every cache
   * entry this document held; nothing is posted after. */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    const calls = [...this.calls.values()];
    this.calls.clear();
    for (const controller of calls) controller.abort();
    for (const cache of this.caches.values()) cache.release();
    this.caches.clear();
    // The descriptors may hold application font bytes.
    this.sources = [];
  }

  /** In-flight calls (tests and leak checks). */
  get pendingCount(): number { return this.calls.size; }

  private async respond(key: string, generation: number, requestId: number, resolverId: unknown, rawDemand: unknown): Promise<void> {
    const controller = new AbortController();
    this.calls.set(key, controller);
    let expired = false;
    const timer = setTimeout(() => { expired = true; controller.abort(); }, FONT_LOAD_CEILING_MS);
    let outcome: FontResolveOutcome;
    try {
      outcome = await this.resolve(resolverId, rawDemand, controller.signal);
    } finally {
      clearTimeout(timer);
    }
    // Cancelled by the worker or the document destroyed: the worker no longer
    // waits, so the answer is dropped here (cached entries stay with their
    // holders and are released with them).
    const live = this.calls.get(key) === controller;
    this.calls.delete(key);
    if (this.disposed || !live || (controller.signal.aborted && !expired)) return;
    if (expired) outcome = { status: 'failed', reason: 'timed out' };
    let reply = outcome;
    let transfer: ArrayBuffer[] = [];
    if (outcome.status === 'families') {
      const sent = transferable(outcome.families, this.copyData);
      reply = { status: 'families', families: sent.answers };
      transfer = sent.transfer;
    }
    try {
      this.post({ kind: 'ooxmlFontsResolved', generation, requestId, outcome: reply } satisfies MainFontResolveReply, transfer);
    } catch (error) {
      // The worker is gone (nothing waits), or the reply did not clone: fail
      // the call's families instead of leaving the worker waiting.
      try {
        this.post({
          kind: 'ooxmlFontsResolved', generation, requestId,
          outcome: { status: 'failed', reason: error instanceof Error ? error.message : String(error) },
        } satisfies MainFontResolveReply);
      } catch { /* worker terminated */ }
    }
  }

  private cacheOf(source: FontResolverSource): SourceCache {
    let cache = this.caches.get(source.id);
    if (!cache) {
      cache = acquireSourceCache(source);
      this.caches.set(source.id, cache);
    }
    return cache;
  }

  private async resolve(resolverId: unknown, rawDemand: unknown, signal: AbortSignal): Promise<FontResolveOutcome> {
    const demand = copyBridgedFontDemand(rawDemand);
    if (!demand || typeof resolverId !== 'string') return { status: 'failed', reason: 'malformed font resolver request' };
    // Only this document's own resolvers: the id names the function, and the
    // same function always yields the same id (`fontResolver`).
    const source = this.sources.find((entry): entry is FontResolverSource =>
      entry.kind === 'font-resolver' && entry.id === resolverId);
    if (!source || !fontResolverFunction(source)) {
      return { status: 'failed', reason: 'font resolver function is not available in this realm' };
    }
    const scriptsOf = substituteScriptsOf(demand.substituteScripts);
    const families: ResolverDemandedName[] = [
      ...demand.families.map((name) => ({ name, fallback: false, scripts: scriptsOf(name) })),
      ...demand.fallbackFamilies.map((name) => ({ name, fallback: true, scripts: scriptsOf(name) })),
    ];
    const cache = this.cacheOf(source);
    // XLSX: held with the viewer realm's session for the document's lifetime
    // (by the cache handle). DOCX/PPTX: held for this request only.
    const holder: object = this.copyData ? cache : {};
    try {
      const answers = await resolveFamiliesThroughCache({
        source, cache, runner: runResolverInRealm, families,
        tuples: demand.tuples, scalars: demand.scalars, signal, holder,
      });
      return {
        status: 'families',
        families: [...answers].map(([family, answer]): BridgedFamilyAnswer =>
          answer.kind === 'claimed' ? { family, kind: 'claimed', declarations: answer.declarations }
            : answer.kind === 'failed' ? { family, kind: 'failed', reason: answer.reason }
              : { family, kind: 'absent' }),
      };
    } catch (error) {
      if (isFontSourceLimit(error)) return { status: 'limit', violation: error.violation };
      return { status: 'failed', reason: error instanceof Error ? error.message : String(error) };
    } finally {
      if (!this.copyData) cache.families.release(holder);
    }
  }
}
