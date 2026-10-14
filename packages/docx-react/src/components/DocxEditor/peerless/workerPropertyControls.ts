import { useCallback, useRef, useState } from 'react';
import type { YrsParagraphAttrs } from '@betteroffice/docx/yrs';
import type { SectionProperties, TabStop, Watermark } from '@betteroffice/docx/types/document';
import type { PeerlessInputIntent } from './inputEvents';
import type { PeerlessWorkerQuery } from './shellController';
import { useWorkerQuery } from './useWorkerQuery';
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

export function useWorkerRulerControls(controller: WorkerShellController, onError?: (error: Error) => void, enabled = true) {
  const query = useWorkerQuery<'sectionProperties'>(controller, enabled ? { kind: 'sectionProperties' } : null);
  const state = controller.getSnapshot();
  const page = state.view?.geometry?.pages.find(page => page.page.pageIndex === controller.ref.getCurrentPage() - 1);
  const section = query?.value.sections[page?.displayPage?.sectionIndex ?? query.value.sections.length - 1] ?? query?.value.sections.at(-1);
  const paragraph = query?.value.paragraphs.find(paragraph => paragraph.story === query.selection?.focus.story && paragraph.paraId === query.selection.focus.paraId);
  const attrs = state.view?.editState?.toolbar?.paragraphProperties;
  const env = controller.environment();
  const editable = !!query && !env.readOnly && env.mode === 'editing';
  const paragraphEditable = editable && !!query?.selection && !!paragraph;
  const drag = useRef<{ query: PeerlessWorkerQuery<'sectionProperties'>; intent?: PeerlessInputIntent } | null>(null);
  const apply = async (intent: PeerlessInputIntent) => {
    const origin = drag.current?.query ?? query;
    if (!origin || !editable) return;
    if (drag.current) { drag.current.intent = intent; return; }
    await applyWorkerProperty(controller, origin, intent, onError);
  };
  const margin = (property: 'marginLeft' | 'marginRight' | 'marginTop' | 'marginBottom') => (twips: number) => {
    if (!section || !query) return Promise.resolve();
    return apply({ kind: 'setSectionMargins', payload: { section: section.section, margins: { [property]: twips }, expectVersion: query.version } });
  };
  const indent = (attrs: YrsParagraphAttrs) => paragraphEditable && query?.selection ? apply({ kind: 'setParagraphAttrs', payload: { attrs, selection: query.selection } }) : Promise.resolve();
  return { sectionProps: section?.properties, editable, paragraphEditable, indentLeft: Number(attrs?.indentLeft ?? 0), indentRight: Number(attrs?.indentRight ?? 0),
    firstLineIndent: Number(attrs?.indentFirstLine ?? 0), hangingIndent: attrs?.hangingIndent === true, tabStops: paragraph ? [...paragraph.tabs] : null,
    startDrag: () => { if (editable && query) drag.current = { query }; },
    endDrag: async () => { const pending = drag.current; drag.current = null;
      if (pending?.intent) await applyWorkerProperty(controller, pending.query, pending.intent, onError); },
    handleLeftMarginChange: margin('marginLeft'), handleRightMarginChange: margin('marginRight'),
    handleTopMarginChange: margin('marginTop'), handleBottomMarginChange: margin('marginBottom'),
    handleIndentLeftChange: (twips: number) => indent({ indentLeft: twips > 0 ? twips : null }),
    handleIndentRightChange: (twips: number) => indent({ indentRight: twips > 0 ? twips : null }),
    handleFirstLineIndentChange: (twips: number) => indent({ indentFirstLine: Math.abs(twips) || null, hangingIndent: twips < 0 }),
    handleTabStopsChange: (tabs: TabStop[]) => paragraphEditable && query?.selection ? apply({ kind: 'setTabStops', payload: {
      selection: query.selection, tabs,
      expectVersion: query.version } }) : Promise.resolve(),
    handleTabStopRemove: (positionTwips: number) => paragraphEditable && query?.selection ? apply({ kind: 'removeTabStop', payload: {
      selection: query.selection, positionTwips, expectVersion: query.version } }) : Promise.resolve(),
  };
}
