import { useCallback, useLayoutEffect, useRef, useState } from 'react';
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

type RulerQuery = PeerlessWorkerQuery<'sectionProperties'>;

function rulerDocumentKey(state: ReturnType<WorkerShellController['getSnapshot']>) {
  const identity = state.status === 'ready' ? state.view?.identity : undefined;
  return identity ? `${identity.session}:${identity.incarnation}:${identity.epoch}` : null;
}

export function useWorkerRulerControls(controller: WorkerShellController, onError?: (error: Error) => void, enabled = true) {
  const live = useWorkerQuery<'sectionProperties'>(controller, enabled ? { kind: 'sectionProperties' } : null);
  const state = controller.getSnapshot();
  const documentKey = enabled ? rulerDocumentKey(state) : null;
  // the markers show the last read while an edit's re-read is in flight; writes always use a fresh one
  const settled = useRef<{ documentKey: string; query: RulerQuery } | null>(null);
  useLayoutEffect(() => { if (live && documentKey) settled.current = { documentKey, query: live }; }, [live, documentKey]);
  const query = live ?? (documentKey && settled.current?.documentKey === documentKey ? settled.current.query : null);
  const pageIndex = query ? controller.ref.getCurrentPage() - 1 : -1;
  const page = state.view?.geometry?.pages.find(page => page.page.pageIndex === pageIndex);
  const sectionOf = (read: RulerQuery) => read.value.sections[page?.displayPage?.sectionIndex ?? read.value.sections.length - 1] ?? read.value.sections.at(-1);
  const paragraphOf = (read: RulerQuery) => read.value.paragraphs.find(paragraph => paragraph.story === read.selection?.focus.story && paragraph.paraId === read.selection.focus.paraId);
  const section = query ? sectionOf(query) : undefined;
  const paragraph = query ? paragraphOf(query) : undefined;
  const attrs = state.view?.editState?.toolbar?.paragraphProperties;
  const env = controller.environment();
  const editable = !!query && !env.readOnly && env.mode === 'editing';
  const paragraphEditable = editable && !!query?.selection && !!paragraph;
  type Build = (read: RulerQuery) => PeerlessInputIntent | null;
  const drag = useRef<{ query: RulerQuery | null; documentKey: string | null; shown: typeof paragraph; build?: Build } | null>(null);
  const shownParagraph = () => drag.current ? drag.current.shown : paragraph;
  const commit = async (origin: RulerQuery | null, authoredIn: string | null, build: Build) => {
    let read: RulerQuery;
    try { read = origin ?? await controller.query<'sectionProperties'>({ kind: 'sectionProperties' }); } catch (error) {
      onError?.(error instanceof Error ? error : new Error(String(error))); return;
    }
    if (!authoredIn || rulerDocumentKey(controller.getSnapshot()) !== authoredIn) return;
    const intent = build(read);
    if (intent) await applyWorkerProperty(controller, read, intent, onError);
  };
  const apply = async (build: Build) => {
    if (!editable) return;
    if (drag.current) { drag.current.build = build; return; }
    await commit(live, documentKey, build);
  };
  const margin = (property: 'marginLeft' | 'marginRight' | 'marginTop' | 'marginBottom') => (twips: number) => apply(read => {
    const target = sectionOf(read);
    return target ? { kind: 'setSectionMargins', payload: { section: target.section, margins: { [property]: twips }, expectVersion: read.version } } : null;
  });
  type Paragraph = NonNullable<ReturnType<typeof paragraphOf>>;
  const paragraphWrite = (write: (read: RulerQuery, selection: NonNullable<RulerQuery['selection']>, target: Paragraph) => PeerlessInputIntent | null, shownOnly = false) => {
    const shown = shownParagraph();
    return paragraphEditable ? apply(read => {
      const target = paragraphOf(read);
      return read.selection && target && (!shownOnly || (target.story === shown?.story && target.paraId === shown.paraId)) ? write(read, read.selection, target) : null;
    }) : Promise.resolve();
  };
  const indent = (attrs: YrsParagraphAttrs) => paragraphWrite((_read, selection) => ({ kind: 'setParagraphAttrs', payload: { attrs, selection } }));
  return { sectionProps: section?.properties, editable, paragraphEditable, indentLeft: Number(attrs?.indentLeft ?? 0), indentRight: Number(attrs?.indentRight ?? 0),
    firstLineIndent: Number(attrs?.indentFirstLine ?? 0), hangingIndent: attrs?.hangingIndent === true, tabStops: paragraph ? [...paragraph.tabs] : null,
    startDrag: () => { if (editable) drag.current = { query: live, documentKey, shown: paragraph }; },
    endDrag: async () => { const pending = drag.current; drag.current = null;
      if (pending?.build) await commit(pending.query, pending.documentKey, pending.build); },
    handleLeftMarginChange: margin('marginLeft'), handleRightMarginChange: margin('marginRight'),
    handleTopMarginChange: margin('marginTop'), handleBottomMarginChange: margin('marginBottom'),
    handleIndentLeftChange: (twips: number) => indent({ indentLeft: Math.max(0, twips) }),
    handleIndentRightChange: (twips: number) => indent({ indentRight: Math.max(0, twips) }),
    handleFirstLineIndentChange: (twips: number) => indent({ indentFirstLine: Math.abs(twips) || null, hangingIndent: twips < 0 }),
    handleTabStopsChange: (tabs: TabStop[]) => {
      const shownTabs = JSON.stringify(shownParagraph()?.tabs);
      return paragraphWrite((read, selection, target) => JSON.stringify(target.tabs) === shownTabs
        ? { kind: 'setTabStops', payload: { selection, tabs, expectVersion: read.version } } : null, true);
    },
    handleTabStopRemove: (positionTwips: number) => paragraphWrite((read, selection) => ({ kind: 'removeTabStop', payload: {
      selection, positionTwips, expectVersion: read.version } }), true),
  };
}
