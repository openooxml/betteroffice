import { LEGACY_SYNC_API } from './runtimeOptions';

export type XlsxEngineWarning = 'deprecated' | 'no-worker' | 'no-offscreen-canvas' | 'legacy-sync-api';

const messages: Record<XlsxEngineWarning, string> = {
  deprecated: `[XlsxEditor] \`experimentalWorkerOpen={false}\` selects the deprecated in-thread engine. Omit it and \`${LEGACY_SYNC_API}\` to use the worker-owned async editor.`,
  'no-worker': '[XlsxEditor] `Worker` is not available in this environment; using the deprecated in-thread engine.',
  'no-offscreen-canvas': '[XlsxEditor] `OffscreenCanvas` is not available in this environment; using the deprecated in-thread engine.',
  'legacy-sync-api': `[XlsxEditor] \`${LEGACY_SYNC_API}\` selects the deprecated hybrid runtime and sync API, removed in 0.6.0. Omit it to use the async editor.`,
};
const warningKey = Symbol.for('@betteroffice/xlsx-react/engine-warnings');
const registry = globalThis as typeof globalThis & { [warningKey]?: Set<string> };
const warned = registry[warningKey] ??= new Set<string>();

export function warnXlsxEngine(reason: XlsxEngineWarning): void {
  const message = messages[reason];
  if (warned.has(message)) return;
  warned.add(message);
  console.warn(message);
}

export function resetXlsxEngineWarningsForTests(): void {
  warned.clear();
}
