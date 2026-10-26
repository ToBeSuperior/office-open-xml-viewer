/**
 * `fontResolver` worker bridge: main-thread half (`resolver-bridge.ts` has the
 * protocol). Loaded lazily, on a document's first bridged request, by
 * {@link FontResolverBridgeHost} (`resolver-bridge-host.ts`).
 *
 * Library API policy, not OOXML behaviour. For each `ooxmlResolveFonts`
 * request the responder looks the resolver id up among the document's own
 * `fonts` descriptors, so a worker can only run functions this document was
 * configured with, and calls the application function with a copy of the
 * demand (`demand.scalars` bounded by `HARD_MAX_FONT_DEMAND_SCALARS`, family
 * records without a prototype) and an `AbortSignal` aborted by the worker's
 * cancellation, by document destroy and by the 15 s font ceiling. The answer
 * passes the same admission a painting-realm session applies
 * (`resolver-admission.ts`), and the reply carries the admitted declarations,
 * the failure reason, or the breached ceiling.
 *
 * Byte ownership (design §4.3): when the worker is the only painting realm
 * (DOCX and PPTX worker mode) every `data` buffer is transferred, so main keeps
 * nothing. When this realm paints as well (the XLSX viewer registers the same
 * faces in `document.fonts`), or a buffer is one this realm's library still
 * references (a cached resolver answer of a main-mode document, a `fontFiles`
 * descriptor), main keeps the original and transfers one copy per distinct
 * buffer. Nothing is charged here: the worker's session charges what it keeps,
 * once. A reply for a request the worker cancelled, or produced after the
 * document was destroyed, is never posted; its buffers become unreachable.
 */
import type { OoxmlError } from '../errors/ooxml-error.js';
import { nullRecord } from '../internal/own-key.js';
import {
  HARD_MAX_FONT_DEMAND_SCALARS,
  HARD_MAX_FONT_SOURCE_FAMILIES,
} from '../worker/resource-policy.generated.js';
import { FONT_LOAD_CEILING_MS } from './font-registry.js';
import { admitFontResolverResult } from './resolver-admission.js';
import type {
  FontResolveOutcome,
  MainFontResolveReply,
  WorkerFontResolveRequest,
} from './resolver-bridge.js';
import {
  fontResolverFunction,
  isFontBufferHeld,
  type FontDemand,
  type FontFaceDeclaration,
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

/** Ownership of the admitted buffers: originals moved, or one copy each. */
function transferable(
  declarations: readonly FontFaceDeclaration[],
  copyData: boolean,
): { declarations: readonly FontFaceDeclaration[]; transfer: ArrayBuffer[] } {
  const sent = new Map<ArrayBuffer, ArrayBuffer>();
  const out = declarations.map((declaration) => {
    const { data } = declaration;
    if (!data) return declaration;
    let wire = sent.get(data);
    if (!wire) {
      wire = copyData || isFontBufferHeld(data) ? data.slice(0) : data;
      sent.set(data, wire);
    }
    return wire === data ? declaration : { ...declaration, data: wire };
  });
  return { declarations: out, transfer: [...sent.values()] };
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

  /** The document is destroyed: abort every call; nothing is posted after. */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    const calls = [...this.calls.values()];
    this.calls.clear();
    for (const controller of calls) controller.abort();
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
    // waits, so the answer (and its bytes) is dropped here.
    const live = this.calls.get(key) === controller;
    this.calls.delete(key);
    if (this.disposed || !live || (controller.signal.aborted && !expired)) return;
    if (expired) outcome = { status: 'failed', reason: 'timed out' };
    let reply = outcome;
    let transfer: ArrayBuffer[] = [];
    if (outcome.status === 'declarations') {
      const sent = transferable(outcome.declarations, this.copyData);
      reply = { status: 'declarations', declarations: sent.declarations };
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

  private async resolve(resolverId: unknown, rawDemand: unknown, signal: AbortSignal): Promise<FontResolveOutcome> {
    const demand = copyBridgedFontDemand(rawDemand);
    if (!demand || typeof resolverId !== 'string') return { status: 'failed', reason: 'malformed font resolver request' };
    // Only this document's own resolvers: the id names the function, and the
    // same function always yields the same id (`fontResolver`).
    const source = this.sources.find((entry) => entry.kind === 'font-resolver' && entry.id === resolverId);
    const resolveFn = source?.kind === 'font-resolver' ? fontResolverFunction(source) : undefined;
    if (!resolveFn) return { status: 'failed', reason: 'font resolver function is not available in this realm' };
    let raw: unknown;
    try {
      // An application that ignores its signal is not waited for past an
      // abort: its eventual answer is dropped (and becomes unreachable).
      raw = await new Promise<unknown>((resolve, reject) => {
        const onAbort = () => resolve(undefined);
        signal.addEventListener('abort', onAbort, { once: true });
        Promise.resolve().then(() => resolveFn(demand, Object.freeze({ signal }))).then(
          (value) => { signal.removeEventListener('abort', onAbort); resolve(value); },
          (error: unknown) => { signal.removeEventListener('abort', onAbort); reject(error); },
        );
      });
    } catch (error) {
      if (signal.aborted) return { status: 'failed', reason: 'Font resolver abandoned' };
      return { status: 'failed', reason: error instanceof Error ? error.message : String(error) };
    }
    if (signal.aborted) return { status: 'failed', reason: 'Font resolver abandoned' };
    try {
      const admitted = admitFontResolverResult(raw);
      return admitted.ok
        ? { status: 'declarations', declarations: admitted.declarations }
        : { status: 'failed', reason: admitted.reason };
    } catch (error) {
      if (isFontSourceLimit(error)) return { status: 'limit', violation: error.violation };
      return { status: 'failed', reason: error instanceof Error ? error.message : String(error) };
    }
  }
}
