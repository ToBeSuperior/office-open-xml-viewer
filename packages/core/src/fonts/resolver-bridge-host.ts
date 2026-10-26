/**
 * Main-thread entry of the `fontResolver` worker bridge for one document
 * (`resolver-bridge.ts`). It recognises bridge messages synchronously and
 * loads the responder only on the document's first resolver request, so a
 * document without a `fontResolver` source, or one that paints on main, loads
 * no bridge code beyond this class. Messages are handled in arrival order
 * (a cancellation never overtakes its request). Each request's deadline
 * starts when it arrives here, so the time the responder takes to load counts
 * against it, as main mode's ceiling would; a request that expires before the
 * responder takes it is failed, recorded and answered here.
 */
import { FONT_LOAD_CEILING_MS } from './font-registry.js';
import {
  isWorkerFontResolveMessage,
  type MainFontResolveReply,
  type WorkerFontResolveMessage,
  type WorkerFontResolveRequest,
} from './resolver-bridge.js';
import type { FontResolverResponder } from './resolver-bridge-responder.js';
import { recordBridgedRequestFailure, type FontResolverFailures } from './resolver-failures.js';
import type { FontSource } from './source.js';

type Post = (message: unknown, transfer?: Transferable[]) => void;

const isNonNegativeSafeInteger = (value: unknown): value is number =>
  typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;

export interface FontResolverBridgeHostOptions {
  readonly post: Post;
  /** The document's validated `fonts` option. */
  readonly sources: readonly FontSource[] | undefined;
  /** The document's resolver failure record, when its viewer-realm sessions
   * share it (XLSX); see `resolver-resolution.ts`. */
  readonly resolverFailures?: FontResolverFailures;
  /** Loads the responder chunk; defaults to the lazy import (replaced only
   * by tests that exercise a chunk that fails to load). */
  readonly loadResponder?: () => Promise<typeof import('./resolver-bridge-responder.js')>;
}

export class FontResolverBridgeHost {
  private options: FontResolverBridgeHostOptions | null;
  /** The document's outcome record, shared with the responder (and, for
   * XLSX, the viewer-realm sessions). */
  private readonly failures: FontResolverFailures;
  private responder: Promise<FontResolverResponder | null> | null = null;
  private disposed = false;
  /** Newest worker generation seen. */
  private generation = -1;
  /** Requests not yet handed to the responder, with their deadline. */
  private readonly waiting = new Map<string, {
    readonly request: WorkerFontResolveRequest;
    readonly deadline: number;
    readonly timer: ReturnType<typeof setTimeout>;
  }>();

  constructor(options: FontResolverBridgeHostOptions) {
    // Only a document configured with a resolver ever answers; any other
    // bridge message is consumed and ignored.
    this.failures = options.resolverFailures ?? new Map();
    this.options = Array.isArray(options.sources) && options.sources.some((source) => source.kind === 'font-resolver')
      ? options : null;
  }

  /** Route `message`; returns true when it belongs to the bridge. */
  accept(message: unknown): boolean {
    if (!isWorkerFontResolveMessage(message)) return false;
    const options = this.options;
    if (this.disposed || !options?.sources) return true;
    const bridged = message as WorkerFontResolveMessage;
    if (!isNonNegativeSafeInteger(bridged.generation) || !isNonNegativeSafeInteger(bridged.requestId)) return true;
    const key = `${bridged.generation}:${bridged.requestId}`;
    if (bridged.kind === 'ooxmlCancelResolveFonts') {
      // Not handed to the responder yet: the call ends here, unanswered.
      const waiting = this.waiting.get(key);
      if (waiting) { clearTimeout(waiting.timer); this.waiting.delete(key); return true; }
    } else {
      if (bridged.generation < this.generation || this.waiting.has(key)) return true;
      if (bridged.generation > this.generation) {
        // The worker moved to a new document: its older calls are abandoned.
        this.generation = bridged.generation;
        for (const [waitingKey, waiting] of [...this.waiting]) {
          if (waiting.request.generation >= bridged.generation) continue;
          clearTimeout(waiting.timer);
          this.waiting.delete(waitingKey);
        }
      }
      // The call's deadline starts on arrival: the worker's remaining ceiling
      // for its ensure, at most the font ceiling. Loading the responder counts
      // against it, and on expiry the call is failed here, recorded for the
      // document and answered, whether or not the responder has loaded.
      const { remainingMs } = bridged;
      const budget = typeof remainingMs === 'number' && Number.isFinite(remainingMs) && remainingMs >= 0
        ? Math.min(remainingMs, FONT_LOAD_CEILING_MS) : FONT_LOAD_CEILING_MS;
      const deadline = Date.now() + budget;
      const timer = setTimeout(() => {
        if (this.waiting.get(key)?.request !== bridged) return;
        this.waiting.delete(key);
        this.fail(bridged, 'timed out');
      }, budget);
      this.waiting.set(key, { request: bridged, deadline, timer });
    }
    this.responder ??= (options.loadResponder ?? (() => import('./resolver-bridge-responder.js')))().then(({ FontResolverResponder: Responder }) => {
      if (this.disposed || !this.options?.sources) return null;
      return new Responder({
        post: options.post,
        sources: options.sources ?? [],
        resolverFailures: this.failures,
      });
    }, (error: unknown) => {
      console.warn(`[ooxml] font resolver bridge failed to load: ${error instanceof Error ? error.message : String(error)}`);
      return null;
    });
    void this.responder.then((responder) => {
      if (this.disposed) return;
      if (bridged.kind === 'ooxmlCancelResolveFonts') { responder?.accept(bridged); return; }
      // Expired, cancelled or superseded while the responder loaded: not run.
      const waiting = this.waiting.get(key);
      if (waiting?.request !== bridged) return;
      clearTimeout(waiting.timer);
      this.waiting.delete(key);
      if (!responder) {
        // The responder could not load: fail the call instead of leaving the
        // worker waiting, with the same outcome for the whole document.
        this.fail(bridged, 'font resolver bridge unavailable');
        return;
      }
      // Only the time left of the call's deadline; an expired call is not run.
      const left = waiting.deadline - Date.now();
      if (left <= 0) { this.fail(bridged, 'timed out'); return; }
      const forwarded: WorkerFontResolveRequest = { ...bridged, remainingMs: left };
      responder.accept(forwarded);
    });
    return true;
  }

  /** Fail one request: recorded for the whole document, then answered. */
  private fail(request: WorkerFontResolveRequest, reason: string): void {
    const options = this.options;
    if (this.disposed || !options) return;
    recordBridgedRequestFailure(this.failures, options.sources ?? [], request, reason);
    try {
      options.post({
        kind: 'ooxmlFontsResolved', generation: request.generation, requestId: request.requestId,
        outcome: { status: 'failed', reason },
      } satisfies MainFontResolveReply);
    } catch { /* worker terminated */ }
  }

  /** The document is destroyed: abort every in-flight resolver call. */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.options = null;
    for (const { timer } of this.waiting.values()) clearTimeout(timer);
    this.waiting.clear();
    void this.responder?.then((responder) => responder?.dispose());
  }
}
