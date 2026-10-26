/**
 * One document's resolver outcomes that the realm cache does not keep.
 * Library resource policy, not OOXML behaviour.
 *
 * The realm resolution cache never keeps a failure (a later document retries
 * the same source), but every requester of one document must report one
 * outcome per family: an XLSX workbook's viewer-realm sessions measure with
 * the faces its worker paints. Every terminal non-success of a family within
 * the document (a failed answer, a malformed answer, a ceiling breach, the
 * font ceiling, a bridge that cannot answer) is recorded here, through
 * {@link recordFontResolverFailures}, keyed by source identity and scoped
 * family; every requester consults it before and after resolving.
 *
 * Light on purpose: the main-thread bridge host records a bootstrap failure
 * without loading the responder.
 *
 * The record lives on main only. A bridged call's deadline is main's (the
 * worker sends the time left of its ceiling with each request), so main sees
 * every terminal outcome of the worker's resolver calls and records it here.
 */
import { fontSourceId } from './source-cache.js';
import type { FontSource } from './source.js';
import type { FontSubstituteScript } from './substitute-script.js';
import { fontFamilyKey, scopedFamilyKey, type FontSubstituteScripts } from './substitutions.js';

/** Failure reason by {@link fontResolverFailureKey}. */
export type FontResolverFailures = Map<string, string>;

/** Key of one scoped family of one source identity (`fontSourceId`). */
export function fontResolverFailureKey(sourceId: string, name: string, scripts: FontSubstituteScripts): string {
  return `${sourceId}\u0000${scopedFamilyKey(name, scripts)}`;
}

/** Record `reason` for `families` of the source `sourceId`; a family keeps
 * the first terminal outcome recorded for it. */
export function recordFontResolverFailures(
  failures: FontResolverFailures | undefined,
  sourceId: string,
  families: readonly { readonly name: string; readonly scripts: FontSubstituteScripts }[],
  reason: string,
): void {
  if (!failures) return;
  for (const { name, scripts } of families) {
    const key = fontResolverFailureKey(sourceId, name, scripts);
    if (!failures.has(key)) failures.set(key, reason);
  }
}

/**
 * The scoped families a bridged request names, read leniently (the request
 * may be malformed): authored and fallback names with the substitution scope
 * the demand proves, as the responder derives them. Empty when nothing
 * readable is named.
 */
export function bridgedRequestFamilies(
  demand: unknown,
): { readonly name: string; readonly scripts: FontSubstituteScripts }[] {
  if (demand === null || typeof demand !== 'object') return [];
  const fields = demand as { families?: unknown; fallbackFamilies?: unknown; substituteScripts?: unknown };
  const names = [fields.families, fields.fallbackFamilies]
    .flatMap((list) => (Array.isArray(list) ? list.filter((name): name is string => typeof name === 'string') : []));
  const record = fields.substituteScripts;
  const scoped = record !== null && typeof record === 'object' && !Array.isArray(record)
    ? new Map(Object.entries(record).map(([family, list]) =>
      [fontFamilyKey(family), Array.isArray(list)
        ? [...new Set(list.filter((script): script is FontSubstituteScript => typeof script === 'string'))].sort()
        : []]))
    : undefined;
  return names.map((name) => ({
    name,
    scripts: scoped ? Object.freeze(scoped.get(fontFamilyKey(name)) ?? []) : 'all',
  }));
}

/**
 * Record a failed reply to a bridged request (`resolver-bridge.ts`) for every
 * family it names, when it names one of the document's own resolver sources
 * (`sources[sourceIndex]` with the request's resolver id). Every path that
 * answers a bridged request with a failure goes through here, so the
 * document's other requesters (the XLSX viewer realm) report the same.
 */
export function recordBridgedRequestFailure(
  failures: FontResolverFailures | undefined,
  sources: readonly FontSource[],
  request: { readonly sourceIndex?: unknown; readonly resolverId?: unknown; readonly demand?: unknown },
  reason: string,
): void {
  const { sourceIndex, resolverId } = request;
  if (!failures || typeof sourceIndex !== 'number' || !Number.isSafeInteger(sourceIndex) || sourceIndex < 0) return;
  const source = sources[sourceIndex];
  if (source?.kind !== 'font-resolver' || source.id !== resolverId) return;
  recordFontResolverFailures(failures, fontSourceId(source), bridgedRequestFamilies(request.demand), reason);
}
