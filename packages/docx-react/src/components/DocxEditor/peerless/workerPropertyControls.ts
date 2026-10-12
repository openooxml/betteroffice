import { useCallback, useState } from 'react';
import type { SectionProperties, Watermark } from '@betteroffice/docx/types/document';
import { applyWorkerProperty, readWorkerSection, type WorkerPropertyQuery, type WorkerShellController } from './workerProperties';

export function useWorkerPageSetupControls(controller: WorkerShellController, onError?: (error: Error) => void) {
  const [query, setQuery] = useState<WorkerPropertyQuery | null>(null);
  const handleOpenPageSetup = useCallback(async () => {
    try {
      const result = await readWorkerSection(controller);
      if (result.value?.kind !== 'section') throw new Error('Section properties are unavailable');
      setQuery(result);
      return { ok: true, status: 'opened' } as const;
    } catch (error) { onError?.(error instanceof Error ? error : new Error(String(error)));
      return { ok: false, failure: { code: 'command-failed', message: 'Section properties are unavailable.' } } as const; }
  }, [controller, onError]);
  const handlePageSetupApply = useCallback(async (patch: Partial<SectionProperties>) => {
    if (!query || query.value?.kind !== 'section') return;
    if (await applyWorkerProperty(controller, query, { kind: 'setSectionProperties',
      payload: { section: query.value.section, patch, expectVersion: query.version } }, onError)) setQuery(null);
  }, [controller, query, onError]);
  return { showPageSetup: query !== null, close: () => setQuery(null), currentProps: query?.value?.kind === 'section' ? query.value.properties : undefined,
    handleOpenPageSetup, handlePageSetupApply };
}

export function useWorkerWatermarkControls(controller: WorkerShellController, onError?: (error: Error) => void) {
  const [query, setQuery] = useState<WorkerPropertyQuery | null>(null);
  const handleOpenWatermark = useCallback(async () => {
    try {
      const result = await controller.query({ kind: 'properties', target: { kind: 'watermark' } });
      if (result.value?.kind !== 'watermark') throw new Error('Watermark properties are unavailable');
      setQuery(result);
      return { ok: true, status: 'opened' } as const;
    } catch (error) { onError?.(error instanceof Error ? error : new Error(String(error)));
      return { ok: false, failure: { code: 'command-failed', message: 'Watermark properties are unavailable.' } } as const; }
  }, [controller, onError]);
  const handleWatermarkApply = useCallback(async (watermark: Watermark | null) => {
    if (!query) return;
    if (await applyWorkerProperty(controller, query, { kind: 'setWatermark', payload: { watermark, expectVersion: query.version } }, onError)) setQuery(null);
  }, [controller, query, onError]);
  return { showWatermark: query !== null, close: () => setQuery(null), currentWatermark: query?.value?.kind === 'watermark' ? query.value.watermark ?? undefined : undefined,
    handleOpenWatermark, handleWatermarkApply };
}
