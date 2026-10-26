import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { OoxmlError } from '../errors/ooxml-error.js';
import { HARD_MAX_FONT_DEMAND_SCALARS, HARD_MAX_FONT_SOURCE_DECLARATIONS } from '../worker/resource-policy.generated.js';
import { WorkerFontResolverClient } from './resolver-bridge.js';
import { liveSourceCacheCount } from './source-cache.js';
import { FontSourceSession, retainedFontSourceBytes } from './source-session.js';
import {
  assertFontSourceDescriptors,
  type FontDemand,
  type FontFaceDeclaration,
  type FontResolveContext,
  type FontResolveFn,
  type FontSource,
} from './source.js';

// The main realm is a second instance of the module graph, so the two realms
// keep separate caches, registries and byte totals exactly as a window and its
// render worker do. The statically imported modules above are the worker realm.
interface MainRealm {
  readonly fontResolver: typeof import('./source.js').fontResolver;
  readonly FontSourceSession: typeof FontSourceSession;
  readonly FontResolverBridgeHost: typeof import('./resolver-bridge-host.js').FontResolverBridgeHost;
  readonly FontResolverResponder: typeof import('./resolver-bridge-responder.js').FontResolverResponder;
  readonly copyBridgedFontDemand: typeof import('./resolver-bridge-responder.js').copyBridgedFontDemand;
  readonly retainedFontSourceBytes: () => number;
  readonly liveSourceCacheCount: () => number;
  readonly sourceCacheEntryCount: () => number;
}
let main: MainRealm;
beforeAll(async () => {
  vi.resetModules();
  const [source, session, host, responder, cache] = await Promise.all([
    import('./source.js'), import('./source-session.js'), import('./resolver-bridge-host.js'),
    import('./resolver-bridge-responder.js'), import('./source-cache.js'),
  ]);
  main = {
    fontResolver: source.fontResolver,
    FontSourceSession: session.FontSourceSession,
    FontResolverBridgeHost: host.FontResolverBridgeHost,
    FontResolverResponder: responder.FontResolverResponder,
    copyBridgedFontDemand: responder.copyBridgedFontDemand,
    retainedFontSourceBytes: session.retainedFontSourceBytes,
    liveSourceCacheCount: cache.liveSourceCacheCount,
    sourceCacheEntryCount: cache.sourceCacheEntryCount,
  };
});
const fontResolver = (resolve: FontResolveFn) => main.fontResolver(resolve);

// The worker bridge of `fontResolver` (design §3.4, §4.3, §4.5): a worker
// session runs the resolver through the client, main answers through the
// responder, and messages cross a structured clone with real transfer lists
// (a transferred ArrayBuffer is detached on the sending side), delivered
// asynchronously like postMessage.

const G = globalThis as Record<string, unknown>;
const ORIGINAL_FONT_FACE = G.FontFace;

class FakeFontFace {
  status: FontFaceLoadStatus = 'unloaded';
  readonly unicodeRange: string;
  readonly bytes: number;
  constructor(readonly family: string, source: string | ArrayBuffer, descriptors: FontFaceDescriptors = {}) {
    this.unicodeRange = descriptors.unicodeRange ?? 'U+0-10FFFF';
    this.bytes = typeof source === 'string' ? 0 : source.byteLength;
  }
  load(): Promise<FakeFontFace> {
    this.status = 'loaded';
    return Promise.resolve(this);
  }
}

function fakeSet() {
  const faces: FakeFontFace[] = [];
  return {
    faces,
    add: (face: FakeFontFace) => { faces.push(face); },
    delete: (face: FakeFontFace) => { const at = faces.indexOf(face); if (at >= 0) faces.splice(at, 1); return at >= 0; },
  };
}
const asSet = (set: ReturnType<typeof fakeSet>) => set as unknown as FontFaceSet;

const demand = (families: string[], scalars: FontDemand['scalars'] = [0x41]): FontDemand =>
  ({ families, fallbackFamilies: [], tuples: [], scalars, substituteScripts: { [families[0]?.toLowerCase() ?? '']: [] } });

type Post = (message: unknown, transfer?: Transferable[]) => void;

/** Two realms joined by structured clone; `toMain` / `toWorker` record traffic. */
function bridge(sources: readonly FontSource[], copyData = false) {
  const toMain: unknown[] = [];
  const toWorker: unknown[] = [];
  let hold = false;
  const held: (() => void)[] = [];
  const deliver = (fn: () => void, toWorkerRealm: boolean) => {
    if (hold && toWorkerRealm) held.push(fn);
    else void Promise.resolve().then(fn);
  };
  let client: WorkerFontResolverClient;
  const mainPost: Post = (message, transfer) => {
    const copy = structuredClone(message, { transfer: transfer as Transferable[] });
    toWorker.push(copy);
    deliver(() => client.accept(copy), true);
  };
  const host = new main.FontResolverBridgeHost({ post: mainPost, sources, copyData });
  hosts.push(host);
  const workerPost: Post = (message, transfer) => {
    const copy = structuredClone(message, { transfer: transfer as Transferable[] });
    toMain.push(copy);
    deliver(() => host.accept(copy), false);
  };
  client = new WorkerFontResolverClient(workerPost);
  // The worker receives the descriptors as a structured clone and validates them.
  const workerSources = assertFontSourceDescriptors(structuredClone(sources));
  return {
    client, host, toMain, toWorker, workerSources,
    /** Stop delivering main → worker replies until `release()` (a slow realm). */
    pause() { hold = true; },
    resume() { hold = false; for (const fn of held.splice(0)) void Promise.resolve().then(fn); },
    session(set = fakeSet()) {
      const session = new FontSourceSession({ sources: workerSources, fontSet: asSet(set), runResolver: client.run });
      sessions.push(session);
      return { session, set };
    },
  };
}

const FACE = (bytes = 64, family = 'Private Face'): FontFaceDeclaration => ({ family, data: new Uint8Array(bytes).fill(7).buffer });

const sessions: { release(): void }[] = [];
const hosts: { dispose(): void }[] = [];
/** A main-realm (viewer or main-mode) session. */
function mainSession(sources: readonly FontSource[], set = fakeSet()) {
  const session = new main.FontSourceSession({ sources, fontSet: asSet(set) });
  sessions.push(session);
  return { session, set };
}
let warn: ReturnType<typeof vi.spyOn>;
beforeEach(() => {
  G.FontFace = FakeFontFace;
  warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
});
afterEach(async () => {
  vi.useRealTimers();
  for (const session of sessions.splice(0)) session.release();
  for (const host of hosts.splice(0)) host.dispose();
  // A host disposes its (lazily loaded) responder asynchronously.
  await new Promise((resolve) => setTimeout(resolve, 0));
  // Every realm returns everything it took.
  expect(liveSourceCacheCount()).toBe(0);
  expect(retainedFontSourceBytes()).toBe(0);
  expect(main.liveSourceCacheCount()).toBe(0);
  expect(main.retainedFontSourceBytes()).toBe(0);
  G.FontFace = ORIGINAL_FONT_FACE;
  vi.restoreAllMocks();
  vi.useRealTimers();
});

describe('fontResolver bridge: main and worker sessions resolve alike', () => {
  it('registers and charges the same faces as a direct session, moving the bytes once', async () => {
    const answered: ArrayBuffer[] = [];
    const seen: FontDemand[] = [];
    const resolve: FontResolveFn = (request) => {
      seen.push(request);
      const face = FACE(64);
      answered.push(face.data as ArrayBuffer);
      return request.families.map(() => face);
    };
    const source = fontResolver(resolve);

    const { session: direct } = mainSession([source]);
    const directLoaded = await direct.ensure(demand(['Private Face']));
    const directBytes = main.retainedFontSourceBytes();
    direct.release();
    expect(main.retainedFontSourceBytes()).toBe(0);

    const link = bridge([source]);
    const { session, set } = link.session();
    const loaded = await session.ensure(demand(['Private Face']));
    expect(loaded.faces.map(({ family }) => family)).toEqual(directLoaded.faces.map(({ family }) => family));
    expect(set.faces.map((face) => face.bytes)).toEqual([64]);
    // Charged once, by the realm that keeps the bytes; main kept nothing
    // (the bridge held main's cache entry only for its request).
    expect(retainedFontSourceBytes()).toBe(directBytes);
    expect(main.retainedFontSourceBytes()).toBe(0);
    expect(main.sourceCacheEntryCount()).toBe(0);
    // The worker was the only painting realm: the original was transferred.
    expect(answered[1]?.byteLength).toBe(0);
    // The application saw the same demand in both modes, scalars included,
    // with a prototype-free family record.
    expect(seen[1]).toEqual(seen[0]);
    for (const request of seen) expect(Object.getPrototypeOf(request.substituteScripts)).toBeNull();
    session.release();
    expect(retainedFontSourceBytes()).toBe(0);
  });

  it('keeps the original and sends one copy per buffer when main also paints', async () => {
    const face = FACE(32);
    const source = fontResolver(() => [face, { ...face, weight: '700' }]);
    const link = bridge([source], true);
    const { set, session } = link.session();
    await session.ensure(demand(['Private Face']));
    expect((face.data as ArrayBuffer).byteLength).toBe(32);
    expect(set.faces.map((entry) => entry.bytes)).toEqual([32, 32]);
    const reply = link.toWorker[0] as { outcome: { families: { declarations: FontFaceDeclaration[] }[] } };
    const [first, second] = reply.outcome.families[0]?.declarations ?? [];
    // One copy shared by both declarations, not one per declaration.
    expect(first?.data).toBe(second?.data);
  });

  for (const order of ['worker first', 'viewer first', 'concurrently'] as const) {
    it(`runs the resolver once per family when the viewer realm also paints (${order})`, async () => {
      let calls = 0;
      const resolve: FontResolveFn = (request) => {
        calls++;
        return request.families.map((family) => FACE(16, family));
      };
      const source = fontResolver(resolve);
      const link = bridge([source], true);
      const viewer = mainSession([source]);
      const worker = link.session();
      const ensureViewer = () => viewer.session.ensure(demand(['Private Face']));
      const ensureWorker = () => worker.session.ensure(demand(['Private Face']));
      if (order === 'worker first') { await ensureWorker(); await ensureViewer(); }
      else if (order === 'viewer first') { await ensureViewer(); await ensureWorker(); }
      else await Promise.all([ensureWorker(), ensureViewer()]);
      // As in main mode: one call for the family, shared by both realms.
      expect(calls).toBe(1);
      expect(viewer.set.faces.map((face) => face.bytes)).toEqual([16]);
      expect(worker.set.faces.map((face) => face.bytes)).toEqual([16]);
      // A later family of the same workbook is again one call.
      await Promise.all([
        worker.session.ensure(demand(['Second Face'])),
        viewer.session.ensure(demand(['Second Face'])),
      ]);
      expect(calls).toBe(2);
      // Main's cache keeps one charged copy for the workbook while either
      // the viewer session or the bridge holds it.
      expect(main.retainedFontSourceBytes()).toBeGreaterThan(0);
      viewer.session.release();
      expect(main.liveSourceCacheCount()).toBe(1);
      link.host.dispose();
      await vi.waitFor(() => expect(main.liveSourceCacheCount()).toBe(0));
      expect(main.retainedFontSourceBytes()).toBe(0);
    });
  }

  it('joins a main-mode document\'s resolution and never detaches its cached buffer', async () => {
    let calls = 0;
    const face = FACE(16);
    const source = fontResolver(() => { calls++; return [face]; });
    // A main-mode document of the same application holds the answer.
    await mainSession([source]).session.ensure(demand(['Private Face']));
    // A worker-mode document (the worker is its only painting realm) joins it.
    const link = bridge([source]);
    const { set, session } = link.session();
    await session.ensure(demand(['Private Face']));
    expect(calls).toBe(1);
    expect(set.faces.map((entry) => entry.bytes)).toEqual([16]);
    // Without the guard the transfer would have detached it (byteLength 0).
    expect((face.data as ArrayBuffer).byteLength).toBe(16);
  });
});

describe('fontResolver bridge: abort, staleness and malformed replies', () => {
  it('aborts the application call when the worker abandons it and posts no late answer', async () => {
    let signal: AbortSignal | undefined;
    let finish!: (faces: readonly FontFaceDeclaration[]) => void;
    const source = fontResolver((_demand, context: FontResolveContext) => {
      signal = context.signal;
      return new Promise((resolve) => { finish = resolve; });
    });
    const link = bridge([source]);
    const { session, set } = link.session();
    const pending = session.ensure(demand(['Private Face']));
    await vi.waitFor(() => expect(signal).toBeDefined());
    session.release();
    await vi.waitFor(() => expect(signal?.aborted).toBe(true));
    expect(link.toMain.map((message) => (message as { kind: string }).kind))
      .toEqual(['ooxmlResolveFonts', 'ooxmlCancelResolveFonts']);
    const late = FACE(8);
    finish([late]);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(link.toWorker).toEqual([]);
    expect((late.data as ArrayBuffer).byteLength).toBe(8);
    expect((await pending).faces).toEqual([]);
    expect(set.faces).toHaveLength(0);
    expect(link.client.pendingCount).toBe(0);
  });

  it('aborts every in-flight call when the document is destroyed', async () => {
    let signal: AbortSignal | undefined;
    const source = fontResolver((_demand, context) => {
      signal = context.signal;
      return new Promise(() => undefined);
    });
    const link = bridge([source]);
    const { session } = link.session();
    void session.ensure(demand(['Private Face']));
    await vi.waitFor(() => expect(signal).toBeDefined());
    link.host.dispose();
    await vi.waitFor(() => expect(signal?.aborted).toBe(true));
    expect(link.toWorker).toEqual([]);
  });

  it('drops a reply of an earlier generation and never registers it', async () => {
    const source = fontResolver(() => [FACE(8)]);
    const link = bridge([source]);
    link.pause();
    const { session, set } = link.session();
    const pending = session.ensure(demand(['Private Face']));
    await vi.waitFor(() => expect(link.toWorker).toHaveLength(1));
    // The worker starts another document before the reply arrives.
    link.client.reset();
    link.resume();
    const loaded = await pending;
    expect(loaded.faces).toEqual([]);
    expect(set.faces).toHaveLength(0);
    expect(link.client.pendingCount).toBe(0);
    session.release();
  });

  it('fails closed on a malformed reply: host fallback, nothing registered', async () => {
    const source = fontResolver(() => [FACE(8)]);
    const sources = [source];
    const workerSources = assertFontSourceDescriptors(structuredClone(sources));
    for (const outcome of [
      undefined,
      { status: 'families', families: 'nope' },
      { status: 'families', families: [{ family: 'private face', kind: 'claimed', declarations: new Array(HARD_MAX_FONT_SOURCE_DECLARATIONS + 1).fill(null) }] },
      { status: 'families', families: [{ family: 'private face', kind: 'claimed', declarations: [{ family: 'Private Face', data: new Uint8Array(4) }] }] },
      { status: 'families', families: [{ family: 1, kind: 'absent' }] },
      { status: 'families', families: [{ family: 'private face', kind: 'failed' }] },
      { status: 'limit', violation: { resource: 'font-source', quota: 'face', metric: 'bytes', limit: 2, observed: 1 } },
      { status: 'unknown' },
    ]) {
      const client: WorkerFontResolverClient = new WorkerFontResolverClient((message) => {
        const { generation, requestId } = message as { generation: number; requestId: number };
        void Promise.resolve().then(() => client.accept({ kind: 'ooxmlFontsResolved', generation, requestId, outcome }));
      });
      const set = fakeSet();
      const session = new FontSourceSession({ sources: workerSources, fontSet: asSet(set), runResolver: client.run });
      expect((await session.ensure(demand(['Private Face']))).faces).toEqual([]);
      expect(set.faces).toHaveLength(0);
      session.release();
    }
    expect(warn).toHaveBeenCalled();
  });

  it('rejects the load for a ceiling breach found by the main-side admission', async () => {
    const source = fontResolver(() => Array.from({ length: HARD_MAX_FONT_SOURCE_DECLARATIONS + 1 }, () => FACE(1)));
    const link = bridge([source]);
    const { session } = link.session();
    const error = await session.ensure(demand(['Private Face'])).catch((caught: unknown) => caught);
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).violation).toMatchObject({ quota: 'declarations', observed: HARD_MAX_FONT_SOURCE_DECLARATIONS + 1 });
  });

  it('keeps an application error a family failure, with the same reason as main mode', async () => {
    const source = fontResolver(() => { throw new Error('token expired'); });
    const link = bridge([source]);
    const { session } = link.session();
    expect((await session.ensure(demand(['Private Face']))).faces).toEqual([]);
    expect(String(warn.mock.calls.at(-1)?.[0])).toContain('Private Face (token expired)');
  });

  it('runs only resolvers the document was configured with', async () => {
    const configured = fontResolver(() => [FACE(4)]);
    const other = fontResolver(() => [FACE(4, 'Other')]);
    const link = bridge([configured]);
    const session = new FontSourceSession({
      sources: assertFontSourceDescriptors(structuredClone([other])),
      fontSet: asSet(fakeSet()),
      runResolver: link.client.run,
    });
    sessions.push(session);
    expect((await session.ensure(demand(['Other']))).faces).toEqual([]);
    expect(String(warn.mock.calls.at(-1)?.[0])).toContain('not available in this realm');
  });

  it('times the application call out after the font ceiling', async () => {
    vi.useFakeTimers();
    let signal: AbortSignal | undefined;
    const source = fontResolver((_demand, context) => {
      signal = context.signal;
      return new Promise(() => undefined);
    });
    const posts: unknown[] = [];
    const responder = new main.FontResolverResponder({ post: (message) => posts.push(message), sources: [source], copyData: false });
    responder.accept({ kind: 'ooxmlResolveFonts', generation: 0, requestId: 1, resolverId: source.id, demand: demand(['A']) } as never);
    await vi.advanceTimersByTimeAsync(15_000);
    expect(signal?.aborted).toBe(true);
    expect(posts).toEqual([{ kind: 'ooxmlFontsResolved', generation: 0, requestId: 1, outcome: { status: 'failed', reason: 'timed out' } }]);
    expect(responder.pendingCount).toBe(0);
    responder.dispose();
  });
});

describe('fontResolver bridge: demand copy', () => {
  it('bounds scalars, rejects malformed demand and keeps family records prototype-free', () => {
    const many = Array.from({ length: HARD_MAX_FONT_DEMAND_SCALARS + 1 }, (_, index) => index);
    expect(main.copyBridgedFontDemand({ ...demand(['A']), scalars: many })?.scalars).toBe('all');
    expect(main.copyBridgedFontDemand({ ...demand(['A']), scalars: [0x110000] })).toBeUndefined();
    expect(main.copyBridgedFontDemand({ ...demand(['A']), families: [1] })).toBeUndefined();
    expect(main.copyBridgedFontDemand({ ...demand(['A']), substituteScripts: { a: ['latin'] } })).toBeUndefined();
    const copied = main.copyBridgedFontDemand({ ...demand(['A']), substituteScripts: JSON.parse('{"__proto__":["arabic"]}') });
    expect(Object.keys(copied?.substituteScripts ?? {})).toEqual(['__proto__']);
  });
});
