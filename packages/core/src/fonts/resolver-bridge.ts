/**
 * `fontResolver` worker bridge: protocol and render-worker half.
 *
 * Library API policy, not OOXML behaviour. A {@link fontResolver} descriptor
 * carries only an id; its application function stays in the main realm that
 * called the factory (`source.ts`), so credentials, tokens and private
 * network clients never reach the worker (design §3.4). When the render worker
 * paints, its {@link FontSourceSession} runs every resolver call through
 * {@link WorkerFontResolverClient.run}: the call is posted to main
 * (`ooxmlResolveFonts`), main runs the function with an `AbortSignal` and
 * replies (`ooxmlFontsResolved`) with every `data` buffer in the transfer list
 * (`resolver-bridge-responder.ts`). Main resolves through its own realm
 * resolution cache, so the function runs once per family however many
 * realms ask (`resolver-resolution.ts`), and answers per family. The worker
 * session then admits and claims each family's declarations exactly as it
 * admits a direct answer, so a document resolves, registers, charges and
 * loads the same faces in both modes.
 *
 * Correlation follows the host-font measurement requests of the DOCX render
 * worker: each request carries the client's generation and a request id. A
 * worker that starts a new document calls {@link WorkerFontResolverClient.reset},
 * which abandons every pending call; a reply for an abandoned call, an older
 * generation or an unknown id is dropped and its bytes become unreachable.
 * When the worker session abandons a call (its 15 s ceiling, release, every
 * waiter gone) the client posts `ooxmlCancelResolveFonts`, which aborts the
 * application function's signal on main.
 *
 * Replies are validated fail-closed: anything but a declaration array within
 * `HARD_MAX_FONT_SOURCE_DECLARATIONS`, a failure reason or a well-formed
 * limit violation fails the call's families (host fallback), and the
 * declarations themselves are validated again by the session.
 */
import { createAbortError } from '../internal/abort.js';
import { fontSourceViolationFromWire } from '../errors/font-source-limit.js';
import type { OoxmlFontSourceViolation } from '../errors/ooxml-error.js';
import {
  HARD_MAX_FONT_SOURCE_DECLARATIONS,
  HARD_MAX_FONT_SOURCE_FAMILIES,
} from '../worker/resource-policy.generated.js';
import { bridgedFontSourceLimitError } from './resolver-admission.js';
import { BridgedFamilyResolutions } from './resolver-resolution.js';
import type { FontDemand, FontFaceDeclaration, FontResolveContext, FontResolverSource } from './source.js';

/** Worker → main: run the resolver `resolverId` for `demand`. */
export interface WorkerFontResolveRequest {
  readonly kind: 'ooxmlResolveFonts';
  readonly generation: number;
  readonly requestId: number;
  readonly resolverId: string;
  readonly demand: FontDemand;
}

/** Worker → main: the worker no longer waits for `requestId`. */
export interface WorkerFontResolveCancel {
  readonly kind: 'ooxmlCancelResolveFonts';
  readonly generation: number;
  readonly requestId: number;
}

/** Main's answer for one demanded family (lower-cased `family`). */
export type BridgedFamilyAnswer =
  /** The declarations that claim it; every `data` buffer is in the transfer list. */
  | { readonly family: string; readonly kind: 'claimed'; readonly declarations: readonly FontFaceDeclaration[] }
  | { readonly family: string; readonly kind: 'failed'; readonly reason: string }
  | { readonly family: string; readonly kind: 'absent' };

/** What main answers for one request. */
export type FontResolveOutcome =
  /** Per-family answers from main's resolution cache. */
  | { readonly status: 'families'; readonly families: readonly BridgedFamilyAnswer[] }
  /** The request could not be served (host fallback for every family). */
  | { readonly status: 'failed'; readonly reason: string }
  /** The answer breached a font-source ceiling (rejects the load). */
  | { readonly status: 'limit'; readonly violation: OoxmlFontSourceViolation };

/** Main → worker. */
export interface MainFontResolveReply {
  readonly kind: 'ooxmlFontsResolved';
  readonly generation: number;
  readonly requestId: number;
  readonly outcome: FontResolveOutcome;
}

export type WorkerFontResolveMessage = WorkerFontResolveRequest | WorkerFontResolveCancel;

type Post = (message: unknown, transfer?: Transferable[]) => void;

const kindOf = (value: unknown): unknown =>
  value !== null && typeof value === 'object' ? (value as { kind?: unknown }).kind : undefined;

/** Whether `value` is a worker → main resolver-bridge message (main side). */
export function isWorkerFontResolveMessage(value: unknown): value is WorkerFontResolveMessage {
  const kind = kindOf(value);
  return kind === 'ooxmlResolveFonts' || kind === 'ooxmlCancelResolveFonts';
}

/** Whether `value` is a main → worker resolver-bridge reply (worker side). */
export function isMainFontResolveReply(value: unknown): value is MainFontResolveReply {
  return kindOf(value) === 'ooxmlFontsResolved';
}

/** Longest failure reason kept from a reply (it only reaches console.warn). */
const MAX_REASON_LENGTH = 1024;

interface PendingCall {
  readonly resolve: (answers: BridgedFamilyResolutions) => void;
  readonly reject: (error: unknown) => void;
  readonly detach: () => void;
}

/**
 * Render-worker half of the bridge: a `FontResolverRunner` for the worker's
 * font-source session. One client per worker.
 */
export class WorkerFontResolverClient {
  private generation = 0;
  private nextRequestId = 1;
  private readonly pending = new Map<number, PendingCall>();

  constructor(private readonly post: Post) {}

  /** Pass as `FontSourceSessionOptions.runResolver`. */
  readonly run = (
    source: FontResolverSource,
    demand: FontDemand,
    context: FontResolveContext,
  ): Promise<BridgedFamilyResolutions> => {
    const { signal } = context;
    if (signal.aborted) return Promise.reject(createAbortError('Font resolver abandoned'));
    const generation = this.generation;
    const requestId = this.nextRequestId++;
    return new Promise<BridgedFamilyResolutions>((resolve, reject) => {
      const onAbort = () => {
        if (!this.settle(requestId)) return;
        reject(createAbortError('Font resolver abandoned'));
        this.postQuietly({ kind: 'ooxmlCancelResolveFonts', generation, requestId } satisfies WorkerFontResolveCancel);
      };
      signal.addEventListener('abort', onAbort, { once: true });
      this.pending.set(requestId, {
        resolve, reject, detach: () => signal.removeEventListener('abort', onAbort),
      });
      try {
        this.post({
          kind: 'ooxmlResolveFonts', generation, requestId, resolverId: source.id, demand,
        } satisfies WorkerFontResolveRequest);
      } catch (error) {
        this.settle(requestId);
        reject(error);
      }
    });
  };

  /** Route a reply; returns true when `message` belongs to this bridge (it is
   * consumed even when stale, so its bytes are dropped here). */
  accept(message: unknown): boolean {
    if (!isMainFontResolveReply(message)) return false;
    const { generation, requestId } = message as Partial<MainFontResolveReply>;
    if (generation !== this.generation || typeof requestId !== 'number') return true;
    const call = this.pending.get(requestId);
    if (!call || !this.settle(requestId)) return true;
    const outcome = readOutcome(message.outcome);
    if (outcome.status === 'families') {
      call.resolve(new BridgedFamilyResolutions(new Map(outcome.families.map((answer) => [answer.family, answer]))));
    }
    else if (outcome.status === 'limit') call.reject(bridgedFontSourceLimitError(outcome.violation));
    else call.reject(new Error(outcome.reason));
    return true;
  }

  /** A new document starts (or the worker's document is gone): abandon every
   * pending call and ignore replies of the previous generation. */
  reset(): void {
    this.generation++;
    const calls = [...this.pending.values()];
    this.pending.clear();
    for (const call of calls) {
      call.detach();
      call.reject(createAbortError('Font resolver abandoned'));
    }
  }

  /** Pending calls (tests and leak checks). */
  get pendingCount(): number { return this.pending.size; }

  private settle(requestId: number): boolean {
    const call = this.pending.get(requestId);
    if (!call) return false;
    this.pending.delete(requestId);
    call.detach();
    return true;
  }

  private postQuietly(message: WorkerFontResolveCancel): void {
    try { this.post(message); } catch { /* main is gone; nothing waits there */ }
  }
}

/** A reply outcome, validated; anything malformed is a failure. */
function readOutcome(value: unknown): FontResolveOutcome {
  const malformed: FontResolveOutcome = { status: 'failed', reason: 'malformed font resolver reply' };
  if (value === null || typeof value !== 'object') return malformed;
  const fields = value as { status?: unknown; families?: unknown; reason?: unknown; violation?: unknown };
  switch (fields.status) {
    case 'families': {
      // The session admits and claims every family's declarations again;
      // only the shape and the counts are bounded here.
      const families = fields.families;
      if (!Array.isArray(families) || families.length > HARD_MAX_FONT_SOURCE_FAMILIES) return malformed;
      const answers: BridgedFamilyAnswer[] = [];
      for (const entry of families as unknown[]) {
        const answer = readFamilyAnswer(entry);
        if (!answer) return malformed;
        answers.push(answer);
      }
      return { status: 'families', families: answers };
    }
    case 'failed':
      return typeof fields.reason === 'string'
        ? { status: 'failed', reason: fields.reason.slice(0, MAX_REASON_LENGTH) }
        : malformed;
    case 'limit': {
      const violation = fontSourceViolationFromWire(fields.violation);
      return violation ? { status: 'limit', violation } : malformed;
    }
    default:
      return malformed;
  }
}

function readFamilyAnswer(value: unknown): BridgedFamilyAnswer | undefined {
  if (value === null || typeof value !== 'object') return undefined;
  const fields = value as { family?: unknown; kind?: unknown; declarations?: unknown; reason?: unknown };
  if (typeof fields.family !== 'string') return undefined;
  const family = fields.family;
  switch (fields.kind) {
    case 'claimed':
      return Array.isArray(fields.declarations) && fields.declarations.length <= HARD_MAX_FONT_SOURCE_DECLARATIONS
        ? { family, kind: 'claimed', declarations: fields.declarations as FontFaceDeclaration[] }
        : undefined;
    case 'failed':
      return typeof fields.reason === 'string'
        ? { family, kind: 'failed', reason: fields.reason.slice(0, MAX_REASON_LENGTH) }
        : undefined;
    case 'absent':
      return { family, kind: 'absent' };
    default:
      return undefined;
  }
}
