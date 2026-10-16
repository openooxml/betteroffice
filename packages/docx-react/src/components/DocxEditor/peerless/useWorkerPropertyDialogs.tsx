import { useCallback, useEffect, useState, useSyncExternalStore } from 'react';
import type { EndnoteProperties, FootnoteProperties } from '@betteroffice/docx/types/document';
import type { DocxCommandSelection } from '@betteroffice/docx/yrs';
import type { DocxEditorProps } from '../../DocxEditor';
import { HyperlinkDialog, type HyperlinkData } from '../../dialogs/HyperlinkDialog';
import { FootnotePropertiesDialog } from '../../dialogs/FootnotePropertiesDialog';
import { ImagePropertiesDialog, type ImagePropertiesData } from '../../dialogs/ImagePropertiesDialog';
import { ImagePositionDialog, type ImagePositionData } from '../../dialogs/ImagePositionDialog';
import { PageSetupDialog } from '../../dialogs/PageSetupDialog';
import { SplitCellDialog } from '../../dialogs/SplitCellDialog';
import type { DocxCommandResult } from '../../../commands/types';
import { WatermarkDialog } from '../../dialogs/WatermarkDialog';
import { useWorkerQuery } from './useWorkerQuery';
import { useWorkerPageSetupControls, useWorkerWatermarkControls } from './workerPropertyControls';
import { applyWorkerProperty, readWorkerSection, type WorkerPropertyQuery, type WorkerShellController } from './workerProperties';

type WorkerLink = { query: WorkerPropertyQuery; initialData: HyperlinkData; targetSelection: DocxCommandSelection; replaceText: boolean };

export function useWorkerPropertyDialogs(controller: WorkerShellController, props: Pick<DocxEditorProps, 'onError' | 'watermarkPresets'>) {
  const state = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  const page = useWorkerPageSetupControls(controller, props.onError);
  const watermark = useWorkerWatermarkControls(controller, props.onError);
  const selection = state.view?.workerSelection;
  const object = useWorkerQuery<'properties'>(controller, selection ? { kind: 'properties', target: { kind: 'objectSelection', selection } } : null);
  const [image, setImage] = useState<{ query: WorkerPropertyQuery; position: boolean } | null>(null);
  const [notes, setNotes] = useState<WorkerPropertyQuery | null>(null);
  const [link, setLink] = useState<WorkerLink | null>(null);
  const [split, setSplit] = useState<{ defaults: { initialRows: number; initialCols: number; minRows: number; minCols: number };
    apply(rows: number, columns: number): Promise<DocxCommandResult> } | null>(null);
  const openSplitCell = useCallback((defaults: { initialRows: number; initialCols: number; minRows: number; minCols: number },
    apply: (rows: number, columns: number) => Promise<DocxCommandResult>) => { setSplit({ defaults, apply }); }, []);
  const report = (error: unknown) => props.onError?.(error instanceof Error ? error : new Error(String(error)));
  useEffect(() => { setImage(null); setNotes(null); setLink(null); setSplit(null); },
    [state.view?.identity.session, state.view?.identity.incarnation, state.view?.identity.epoch]);
  const openImage = async (position = false) => {
    try {
      await controller.ref.flushPendingInput();
      const selected = controller.getSnapshot().view?.workerSelection;
      if (!selected) throw new Error('Image selection is unavailable');
      const query = await controller.query({ kind: 'properties', target: { kind: 'objectSelection', selection: selected } });
      if (query.value?.kind !== 'object') throw new Error('Image properties are unavailable');
      setImage({ query, position });
      return { ok: true, status: 'opened' } as const;
    } catch (error) { report(error); return { ok: false, failure: { code: 'command-failed', message: 'Image properties are unavailable.' } } as const; }
  };
  const readLink = async (): Promise<WorkerLink> => {
    await controller.ref.flushPendingInput();
    const selection = controller.getSnapshot().view?.workerSelection;
    if (!selection) throw new Error('Link selection is unavailable');
    const query = await controller.query({ kind: 'properties', target: { kind: 'hyperlink', selection } });
    if (query.value?.kind !== 'hyperlink') throw new Error('Link properties are unavailable');
    const { range, hyperlink } = query.value;
    const targetSelection = hyperlink ? { anchor: { story: range.story, ...range.start }, focus: { story: range.story, ...range.end } } : selection;
    const text = await controller.readAt({ kind: 'rawSelectionInfo', selection: targetSelection }, query.version);
    const collapsed = selection.anchor.paraId === selection.focus.paraId && selection.anchor.offset === selection.focus.offset;
    return { query, initialData: { url: hyperlink?.href, tooltip: hyperlink?.tooltip ?? undefined, displayText: text.selectedText },
      targetSelection, replaceText: range.start.paraId === range.end.paraId && (collapsed || hyperlink !== null) };
  };
  const openLink = async () => {
    try {
      setLink(await readLink());
      return { ok: true, status: 'opened' } as const;
    } catch (error) { report(error); return { ok: false, failure: { code: 'command-failed', message: 'Link properties are unavailable.' } } as const; }
  };
  const openNotes = async () => {
    try {
      const query = await readWorkerSection(controller);
      if (query.value?.kind !== 'section') throw new Error('Note properties are unavailable');
      setNotes(query);
    } catch (error) { report(error); }
  };
  const writeLink = async (opened: WorkerLink, data: HyperlinkData | null) => {
    if (opened.query.value?.kind !== 'hyperlink') return false;
    const hyperlink = data ? { href: data.bookmark ? `#${data.bookmark}` : data.url ?? '', ...(data.tooltip ? { tooltip: data.tooltip } : {}) } : null;
    const displayText = data && opened.replaceText && data.displayText ? data.displayText : undefined;
    if (!await applyWorkerProperty(controller, opened.query, { kind: 'setHyperlink', payload: { selection: opened.targetSelection,
      hyperlink, displayText, expectVersion: opened.query.version } }, props.onError)) return false;
    if (data) {
      const range = opened.query.value.range;
      const end = { story: range.story, ...(displayText ? { ...range.start, offset: range.start.offset + displayText.length } : range.end) };
      await controller.submit({ kind: 'select', payload: { selection: { anchor: end, focus: end } } });
    }
    return true;
  };
  const applyLink = async (data: HyperlinkData | null) => {
    if (link && await writeLink(link, data)) setLink(null);
  };
  const changeLinkAtCaret = async (data: HyperlinkData | null) => {
    try {
      const opened = await readLink();
      const story = opened.targetSelection.focus.story;
      if (opened.query.value?.kind !== 'hyperlink' || !opened.query.value.hyperlink || (story !== 'body' && !story.startsWith('body:'))) return false;
      return await writeLink(opened, data);
    } catch (error) { report(error); return false; }
  };
  const applyNotes = async (footnotePr: FootnoteProperties, endnotePr: EndnoteProperties) => {
    if (notes?.value?.kind !== 'section') return;
    if (await applyWorkerProperty(controller, notes, { kind: 'setNoteProperties', payload: { section: notes.value.section,
      footnotePr, endnotePr, expectVersion: notes.version } }, props.onError)) setNotes(null);
  };
  const applyImage = async (patch: Record<string, unknown>) => {
    if (image?.query.value?.kind !== 'object') return;
    if (await applyWorkerProperty(controller, image.query, { kind: 'setObjectProperties', payload: {
      handle: image.query.value.handle, patch, expectVersion: image.query.version,
    } }, props.onError)) setImage(null);
  };
  const applyImageProperties = (data: ImagePropertiesData) => applyImage({ ...data, alt: data.alt ?? null,
    borderWidth: data.borderWidth ?? 0, borderColor: data.borderColor ?? null, borderStyle: data.borderStyle ?? null });
  const applyImagePosition = (data: ImagePositionData) => applyImage({ position: { horizontal: data.horizontal, vertical: data.vertical },
    distTop: data.distTop, distBottom: data.distBottom, distLeft: data.distLeft, distRight: data.distRight,
    ...(image?.query.value?.kind === 'object' && (image.query.value.properties.wrapType === undefined || image.query.value.properties.wrapType === 'inline')
      ? { wrap: 'square' } : {}) });
  const hyperlink = link?.query.value?.kind === 'hyperlink' ? link.query.value.hyperlink : null;
  const imageData = image?.query.value?.kind === 'object' ? image.query.value.properties : null;
  const imageProperties = imageData ? { ...imageData,
    width: typeof imageData.width === 'number' ? imageData.width : typeof imageData.widthEmu === 'number' ? imageData.widthEmu / 9525 : undefined,
    height: typeof imageData.height === 'number' ? imageData.height : typeof imageData.heightEmu === 'number' ? imageData.heightEmu / 9525 : undefined,
  } : undefined;
  return { openLink, applyPopupLink: (displayText: string, href: string, tooltip?: string) => changeLinkAtCaret({ displayText, url: href, tooltip }),
    removeLinkAtCaret: () => changeLinkAtCaret(null), openNotes, openImage, openImagePosition: () => openImage(true), imagePropertiesAvailable: object?.value?.kind === 'object',
    openPageSetup: page.handleOpenPageSetup, openWatermark: watermark.handleOpenWatermark, openSplitCell,
    element: <>
      {page.showPageSetup && <PageSetupDialog isOpen onClose={page.close} currentProps={page.currentProps} onApply={page.handlePageSetupApply} />}
      {watermark.showWatermark && <WatermarkDialog isOpen onClose={watermark.close} current={watermark.currentWatermark}
        presets={props.watermarkPresets} onApply={watermark.handleWatermarkApply} />}
      {link && <HyperlinkDialog isOpen onClose={() => setLink(null)} isEditing={hyperlink !== null} selectedText={link.initialData.displayText}
        initialData={link.initialData} onSubmit={applyLink}
        onRemove={hyperlink ? () => { void applyLink(null); } : undefined} />}
      {notes?.value?.kind === 'section' && <FootnotePropertiesDialog isOpen onClose={() => setNotes(null)} onApply={applyNotes}
        footnotePr={notes.value.properties.footnotePr} endnotePr={notes.value.properties.endnotePr} />}
      {split && <SplitCellDialog isOpen onClose={() => setSplit(null)} initialRows={split.defaults.initialRows}
        initialCols={split.defaults.initialCols} minRows={split.defaults.minRows} minCols={split.defaults.minCols}
        onApply={(rows, columns) => { void split.apply(rows, columns); setSplit(null); }} />}
      {image && imageData && (image.position ? <ImagePositionDialog isOpen onClose={() => setImage(null)} onApply={applyImagePosition}
        currentData={{ ...imageData, ...(imageData.position && typeof imageData.position === 'object' ? imageData.position : {}) }} /> :
        <ImagePropertiesDialog isOpen onClose={() => setImage(null)} currentData={imageProperties} onApply={applyImageProperties}
          onPosition={() => setImage(current => current ? { ...current, position: true } : null)} />)}
    </> };
}
