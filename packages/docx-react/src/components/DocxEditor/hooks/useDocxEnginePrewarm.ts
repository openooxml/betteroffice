import { useEffect } from 'react';
import { preloadDocxEngine, retainPreloadedResidentEngineWorker } from '@betteroffice/docx/yrs';

function prewarm(): () => void {
  void preloadDocxEngine().catch(() => {});
  return retainPreloadedResidentEngineWorker();
}

export function useDocxEnginePrewarm(enabled: boolean): void {
  useEffect(() => {
    if (enabled) return prewarm();
  }, [enabled]);
}

export function useDocxEnginePrewarmOnBytes(enabled: boolean, bytes: Uint8Array | null): void {
  useEffect(() => {
    if (enabled && bytes) return prewarm();
  }, [enabled, bytes]);
}
