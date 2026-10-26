/**
 * Main-thread entry of the `fontResolver` worker bridge for one document
 * (`resolver-bridge.ts`). It recognises bridge messages synchronously and
 * loads the responder only on the document's first resolver request, so a
 * document without a `fontResolver` source, or one that paints on main, loads
 * no bridge code beyond this class. Messages are handled in arrival order
 * (a cancellation never overtakes its request).
 */
import { isWorkerFontResolveMessage, type MainFontResolveReply, type WorkerFontResolveMessage } from './resolver-bridge.js';
import type { FontResolverResponder } from './resolver-bridge-responder.js';
import type { FontResolverFailures } from './resolver-resolution.js';
import type { FontSource } from './source.js';

type Post = (message: unknown, transfer?: Transferable[]) => void;

export interface FontResolverBridgeHostOptions {
  readonly post: Post;
  /** The document's validated `fonts` option. */
  readonly sources: readonly FontSource[] | undefined;
  /** The document's resolver failure record, when its viewer-realm sessions
   * share it (XLSX); see `resolver-resolution.ts`. */
  readonly resolverFailures?: FontResolverFailures;
}

export class FontResolverBridgeHost {
  private options: FontResolverBridgeHostOptions | null;
  private responder: Promise<FontResolverResponder | null> | null = null;
  private disposed = false;

  constructor(options: FontResolverBridgeHostOptions) {
    // Only a document configured with a resolver ever answers; any other
    // bridge message is consumed and ignored.
    this.options = Array.isArray(options.sources) && options.sources.some((source) => source.kind === 'font-resolver')
      ? options : null;
  }

  /** Route `message`; returns true when it belongs to the bridge. */
  accept(message: unknown): boolean {
    if (!isWorkerFontResolveMessage(message)) return false;
    const options = this.options;
    if (this.disposed || !options?.sources) return true;
    this.responder ??= import('./resolver-bridge-responder.js').then(({ FontResolverResponder: Responder }) => {
      if (this.disposed || !this.options?.sources) return null;
      return new Responder({
        post: options.post,
        sources: options.sources ?? [],
        ...(options.resolverFailures ? { resolverFailures: options.resolverFailures } : {}),
      });
    }, (error: unknown) => {
      console.warn(`[ooxml] font resolver bridge failed to load: ${error instanceof Error ? error.message : String(error)}`);
      return null;
    });
    void this.responder.then((responder) => {
      const bridged = message as WorkerFontResolveMessage;
      if (responder) { responder.accept(bridged); return; }
      // The responder could not load: fail the call instead of leaving the
      // worker waiting for its ceiling.
      if (this.disposed || bridged.kind !== 'ooxmlResolveFonts') return;
      try {
        options.post({
          kind: 'ooxmlFontsResolved', generation: bridged.generation, requestId: bridged.requestId,
          outcome: { status: 'failed', reason: 'font resolver bridge unavailable' },
        } satisfies MainFontResolveReply);
      } catch { /* worker terminated */ }
    });
    return true;
  }

  /** The document is destroyed: abort every in-flight resolver call. */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.options = null;
    void this.responder?.then((responder) => responder?.dispose());
  }
}
