/**
 * Source identity and the per-realm caches font sources share across
 * documents.
 *
 * Library resource policy, not OOXML behaviour. Every cache is keyed by a
 * source's identity ({@link fontSourceId}): a stable hash of its kind and its
 * whole configuration. Two documents configured with equal sources share
 * stylesheet fetches and family resolutions; two sources that differ in
 * anything (origin, template, query, credentials, substitutions, faces) share
 * nothing, so a misbehaving or failing origin can never answer for another
 * (#1497: per-origin isolation, no retry against another origin).
 *
 * Entries live while at least one session holds the source
 * ({@link acquireSourceCache} / {@link SourceCache.release}): memory is bounded
 * by the live sessions, not by the number of documents ever opened. Sequential
 * reopening of the same document re-issues stylesheet requests, which the
 * browser HTTP cache answers.
 *
 * Work in the caches is shared through {@link SharedTask}: each waiter joins
 * with its own `AbortSignal`, the underlying fetch is aborted only when every
 * waiter has abandoned it, and a task that fails is evicted so a later session
 * retries the same source (never another).
 */
import { createAbortError } from '../internal/abort.js';
import type { FontFaceDeclaration, FontSource } from './source.js';
import type { FontSubstitution } from './substitutions.js';

// ---- identity -------------------------------------------------------------------

/** Two-lane 32-bit multiplicative hash (cyrb53 construction), 64 bits of
 * output. The input is integrator configuration, never document content, so
 * collision resistance against an adversary is not required; a collision
 * between two of one application's own sources is negligible at 64 bits. */
class IdentityHash {
  private h1 = 0xdeadbeef;
  private h2 = 0x41c6ce57;

  private unit(value: number): void {
    this.h1 = Math.imul(this.h1 ^ value, 2654435761);
    this.h2 = Math.imul(this.h2 ^ value, 1597334677);
  }

  /** Length-prefixed, so field boundaries cannot be shifted between fields. */
  text(value: string | undefined): this {
    if (value === undefined) { this.unit(0x110000); return this; }
    this.unit(value.length);
    for (let index = 0; index < value.length; index++) this.unit(value.charCodeAt(index));
    return this;
  }

  bytes(value: Uint8Array): this {
    this.unit(0x110001);
    this.unit(value.length);
    for (let index = 0; index < value.length; index++) this.unit(value[index]);
    return this;
  }

  digest(): string {
    let h1 = this.h1, h2 = this.h2;
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (h2 >>> 0).toString(16).padStart(8, '0') + (h1 >>> 0).toString(16).padStart(8, '0');
  }
}

/** Content key of one face's bytes or `data:` URL, as used in registry
 * signatures (a `data:` URL can be tens of MiB; the signature stays short). */
export function fontContentKey(value: ArrayBuffer | string): string {
  if (typeof value === 'string') return `data:${value.length}:${new IdentityHash().text(value).digest()}`;
  return `bytes:${value.byteLength}:${new IdentityHash().bytes(new Uint8Array(value)).digest()}`;
}

function hashSubstitutes(hash: IdentityHash, substitutes: readonly FontSubstitution[] | undefined): void {
  hash.text(substitutes === undefined ? undefined : String(substitutes.length));
  for (const entry of substitutes ?? []) hash.text(entry.for).text(entry.use).text(entry.script);
}

function hashDeclaration(hash: IdentityHash, face: FontFaceDeclaration): void {
  hash.text(face.family).text(face.weight).text(face.style).text(face.stretch).text(face.unicodeRange);
  if (face.data) hash.bytes(new Uint8Array(face.data));
  else hash.text(face.url);
}

const sourceIds = new WeakMap<FontSource, string>();

/**
 * Stable identity of a validated source descriptor in this realm. Descriptors
 * are frozen and their `data` buffers are owned by the library once handed
 * over, so the identity is computed once per descriptor object. A
 * `font-resolver` is identified by its function's id, so documents sharing one
 * resolver function share its resolutions.
 */
export function fontSourceId(source: FontSource): string {
  const cached = sourceIds.get(source);
  if (cached) return cached;
  const hash = new IdentityHash().text(source.protocol).text(source.kind);
  switch (source.kind) {
    case 'css-api': {
      hash.text(source.urlTemplate).text(source.query).text(source.credentials ?? 'same-origin');
      const families = source.families ? Object.entries(source.families) : undefined;
      hash.text(families === undefined ? undefined : String(families.length));
      for (const [family, value] of families ?? []) hash.text(family).text(value === true ? undefined : value);
      break;
    }
    case 'font-face-css':
      hash.text(source.url).text(source.credentials ?? 'same-origin');
      break;
    case 'font-files':
      hash.text(String(source.faces.length));
      for (const face of source.faces) hashDeclaration(hash, face);
      break;
    case 'font-resolver':
      hash.text(source.id);
      break;
  }
  hashSubstitutes(hash, source.substitutes);
  const id = `${source.kind}:${hash.digest()}`;
  sourceIds.set(source, id);
  return id;
}

// ---- shared, abandonable work ----------------------------------------------------

/** Abort `controller` when `signal` aborts; returns the detach function. */
function follow(signal: AbortSignal, onAbort: () => void): () => void {
  if (signal.aborted) { onAbort(); return () => undefined; }
  signal.addEventListener('abort', onAbort, { once: true });
  return () => signal.removeEventListener('abort', onAbort);
}

/**
 * One piece of work joined by any number of waiters. The producer receives a
 * signal that aborts only when every waiter has abandoned the task before it
 * settled; `onAbandon` then evicts it from its cache so the next waiter starts
 * afresh.
 */
export class SharedTask<T> {
  private readonly controller = new AbortController();
  private readonly promise: Promise<T>;
  private waiters = 0;
  private settled = false;

  constructor(produce: (signal: AbortSignal) => Promise<T>, private readonly onAbandon: () => void) {
    this.promise = Promise.resolve()
      .then(() => produce(this.controller.signal))
      .finally(() => { this.settled = true; });
    // The task's own rejection is observed by each joiner; never unhandled here.
    this.promise.catch(() => undefined);
  }

  /** Await the task, or reject with an `AbortError` as soon as `signal` aborts. */
  join(signal: AbortSignal): Promise<T> {
    if (signal.aborted) return Promise.reject(createAbortError('Font source request abandoned'));
    this.waiters++;
    return new Promise<T>((resolve, reject) => {
      let done = false;
      const leave = () => {
        if (done) return;
        done = true;
        this.waiters--;
      };
      const detach = follow(signal, () => {
        if (done) return;
        leave();
        if (this.waiters === 0 && !this.settled) {
          this.controller.abort();
          this.onAbandon();
        }
        reject(createAbortError('Font source request abandoned'));
      });
      this.promise.then(
        (value) => { detach(); if (!done) { leave(); resolve(value); } },
        (error: unknown) => { detach(); if (!done) { leave(); reject(error); } },
      );
    });
  }
}

/** A keyed map of shared tasks. Rejections and results `evict` accepts are
 * removed once settled, so failures are never cached. */
export class SharedTaskMap<T> {
  private readonly tasks = new Map<string, SharedTask<T>>();

  constructor(private readonly evict: (value: T) => boolean = () => false) {}

  join(key: string, produce: (signal: AbortSignal) => Promise<T>, signal: AbortSignal): Promise<T> {
    let task = this.tasks.get(key);
    if (!task) {
      const created: SharedTask<T> = new SharedTask(async (taskSignal) => {
        try {
          const value = await produce(taskSignal);
          if (this.evict(value) && this.tasks.get(key) === created) this.tasks.delete(key);
          return value;
        } catch (error) {
          if (this.tasks.get(key) === created) this.tasks.delete(key);
          throw error;
        }
      }, () => { if (this.tasks.get(key) === created) this.tasks.delete(key); });
      task = created;
      this.tasks.set(key, task);
    }
    return task.join(signal);
  }

  /** Whether a task (in flight or settled and kept) exists for `key`. */
  has(key: string): boolean { return this.tasks.has(key); }

  clear(): void { this.tasks.clear(); }
}

// ---- per-source cache ------------------------------------------------------------

/** Parsed stylesheet: its declarations and the final URL after redirects. */
export interface FontStylesheet {
  readonly ok: true;
  readonly declarations: readonly FontFaceDeclaration[];
  readonly finalUrl: string;
}

/** A transport or content failure: the families keep their host fallback. */
export interface FontSourceFailure {
  readonly ok: false;
  readonly reason: string;
}

/** How one source answered one demanded family. */
export type FamilyResolution =
  /** The source declares faces for the family; the family is now final. */
  | {
    readonly kind: 'claimed';
    /** CSS family the declarations register under. */
    readonly served: string;
    /** Present when a declared substitute serves the demanded family. */
    readonly substitution?: FontSubstitution;
    readonly declarations: readonly FontFaceDeclaration[];
  }
  /** The source attempted the family and failed; the family is also final
   * (host fallback, no other source is consulted). */
  | { readonly kind: 'failed'; readonly reason: string }
  /** The source does not serve the family; the next source is consulted. */
  | { readonly kind: 'absent' };

/** Caches one source identity shares across the sessions that hold it. */
export interface SourceCache {
  readonly id: string;
  /** Stylesheets keyed by request URL. */
  readonly stylesheets: SharedTaskMap<FontStylesheet | FontSourceFailure>;
  /** Resolutions keyed by lower-cased demanded family. */
  readonly families: SharedTaskMap<FamilyResolution>;
  /** Family index of a `font-files` source (computed once). */
  index?: ReadonlyMap<string, readonly FontFaceDeclaration[]>;
  release(): void;
}

interface CacheEntry {
  cache: SourceCache;
  refs: number;
}

const caches = new Map<string, CacheEntry>();

/** Hold the shared cache of one source identity until `release()`. */
export function acquireSourceCache(source: FontSource): SourceCache {
  const id = fontSourceId(source);
  const existing = caches.get(id);
  if (existing) {
    existing.refs++;
    return handle(existing);
  }
  const entry: CacheEntry = {
    refs: 1,
    cache: {
      id,
      stylesheets: new SharedTaskMap((value) => !value.ok),
      families: new SharedTaskMap((value) => value.kind === 'failed'),
      release: () => undefined,
    },
  };
  caches.set(id, entry);
  return handle(entry);
}

/** A per-holder view whose `release()` is idempotent. */
function handle(entry: CacheEntry): SourceCache {
  let released = false;
  const { cache } = entry;
  return {
    id: cache.id,
    stylesheets: cache.stylesheets,
    families: cache.families,
    get index() { return cache.index; },
    set index(value) { cache.index = value; },
    release() {
      if (released) return;
      released = true;
      entry.refs--;
      if (entry.refs > 0 || caches.get(cache.id) !== entry) return;
      caches.delete(cache.id);
      cache.stylesheets.clear();
      cache.families.clear();
    },
  };
}

/** Number of source identities with live holders (tests and leak checks). */
export function liveSourceCacheCount(): number {
  return caches.size;
}
