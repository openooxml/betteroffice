import { LEGACY_SYNC_API } from './runtimeOptions';

export type PptxEngineWarning = 'deprecated' | 'collaboration' | 'no-worker' | 'no-offscreen-canvas' | 'legacy-sync-api';

const messages: Record<PptxEngineWarning, string> = {
  deprecated: '[PptxEditor] `experimentalWorkerOpen={false}` selects the deprecated in-thread engine. Omit `experimentalWorkerOpen` to use the worker-owned editor.',
  collaboration: '[PptxEditor] the legacy hybrid runtime does not support `collaboration`; using the deprecated in-thread engine.',
  'no-worker': '[PptxEditor] `Worker` is not available in this environment; using the deprecated in-thread engine.',
  'no-offscreen-canvas': '[PptxEditor] `OffscreenCanvas` is not available in this environment; using the deprecated in-thread engine.',
  'legacy-sync-api': `[PptxEditor] \`${LEGACY_SYNC_API}\` selects the deprecated hybrid runtime and sync API, removed in 0.6.0. Omit it to use the async editor.`,
};
const warningKey = Symbol.for('@betteroffice/pptx-react/engine-warnings');
const registry = globalThis as typeof globalThis & { [warningKey]?: Set<string> };
const warned = registry[warningKey] ??= new Set<string>();

export function warnPptxEngine(reason: PptxEngineWarning): void {
  const message = messages[reason];
  if (warned.has(message)) return;
  warned.add(message);
  console.warn(message);
}

export function resetPptxEngineWarningsForTests(): void {
  warned.clear();
}
