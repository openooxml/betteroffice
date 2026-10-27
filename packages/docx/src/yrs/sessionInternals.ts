/** Engine operations the package runs on sessions it creates itself; not public API. */

import type { editorSaveKeys, noteSeparatorKeys } from './editorSaveKeys';
import type { YrsSession } from './index';
import type { Document } from '../types/document';

export interface YrsSessionInternals {
  headerFooterAliases(): ReadonlyMap<string, string>;
  setHeaderFooterAliases(json: string): void;
  sourcePackage(): { buffer: ArrayBuffer; keys: ReturnType<typeof editorSaveKeys> } | null;
  materializedSource(): Readonly<Document> | null;
  /** The opened source's note separators, keyed as `noteSeparatorKeys` does. */
  sourceNoteSeparators(): ReturnType<typeof noteSeparatorKeys> | null;
  /** Compares two DOCX packages into the empty session; bridge JSON. */
  compareDocx(original: Uint8Array, revised: Uint8Array, options: string): string;
  /** The final comparison result for the saved bytes; bridge JSON. */
  finishComparedDocx(bytes: Uint8Array): string;
  /** The final comparison result when saving failed; bridge JSON. */
  failComparedDocx(message: string): string;
}

const registry = new WeakMap<object, YrsSessionInternals>();

/** Returns the session's opened source package when available. @internal */
export function sessionSourcePackage(
  session: YrsSession
): { buffer: ArrayBuffer; keys: ReturnType<typeof editorSaveKeys> } | null {
  return registry.get(session)?.sourcePackage() ?? null;
}

export function sessionMaterializedSource(session: YrsSession): Readonly<Document> | null {
  return registry.get(session)?.materializedSource() ?? session.materializeDocx?.() ?? null;
}

/** The opened source's note separator keys when available. @internal */
export function sessionSourceNoteSeparators(
  session: YrsSession
): ReturnType<typeof noteSeparatorKeys> | null {
  return registry.get(session)?.sourceNoteSeparators() ?? null;
}

export function registerSessionInternals(session: object, internals: YrsSessionInternals): void {
  registry.set(session, internals);
}

export function sessionInternals(session: object): YrsSessionInternals {
  const internals = registry.get(session);
  if (!internals) throw new Error('the session was not created by createYrsSession');
  return internals;
}

/** @internal */
export function headerFooterStory(session: object | null | undefined, rId: string): string {
  const canonical = session && registry.get(session)?.headerFooterAliases().get(rId);
  return `hf:${canonical || rId}`;
}

/** The relationship ids whose bands show `rId`'s story. @internal */
export function headerFooterDisplayRIds(session: object, rId: string): string[] {
  return routedHeaderFooterRIds(registry.get(session)?.headerFooterAliases(), rId);
}

/** `rId`'s canonical relationship id and every alias `aliases` routes to it. @internal */
export function routedHeaderFooterRIds(
  aliases: ReadonlyMap<string, string> | undefined,
  rId: string
): string[] {
  const canonical = aliases?.get(rId) ?? rId;
  return [canonical, ...[...(aliases ?? [])].flatMap(([alias, owner]) =>
    owner === canonical ? [alias] : []
  )];
}
