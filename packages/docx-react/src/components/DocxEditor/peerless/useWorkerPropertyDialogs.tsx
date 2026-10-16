import { useCallback, useEffect, useState, useSyncExternalStore } from 'react';
import type { EndnoteProperties, FootnoteProperties } from '@betteroffice/docx/types/document';
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

export function useWorkerPropertyDialogs(controller: WorkerShellController, props: Pick<DocxEditorProps, 'onError' | 'watermarkPresets'>) {
  const state = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  const page = useWorkerPageSetupControls(controller, props.onError);
  const watermark = useWorkerWatermarkControls(controller, props.onError);
  const selection = state.view?.workerSelection;
  const object = useWorkerQuery<'properties'>(controller, selection ? { kind: 'properties', target: { kind: 'objectSelection', selection } } : null);
  const [image, setImage] = useState<{ query: WorkerPropertyQuery; position: boolean } | null>(null);
  const [notes, setNotes] = useState<WorkerPropertyQuery | null>(null);
  const [link, setLink] = useState<{ query: WorkerPropertyQuery; text: string } | null>(null);
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
  const openLink = async () => {
    try {
      await controller.ref.flushPendingInput();
      const selection = controller.getSnapshot().view?.workerSelection;
      if (!selection) throw new Error('Link selection is unavailable');
      const query = await controller.query({ kind: 'properties', target: { kind: 'hyperlink', selection } });
      if (query.value?.kind !== 'hyperlink') throw new Error('Link properties are unavailable');
      const text = await controller.readAt({ kind: 'rawSelectionInfo', selection }, query.version);
      setLink({ query, text: text.selectedText });
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
  const applyLink = async (data: HyperlinkData | null) => {
    if (!link?.query.selection) return;
    const hyperlink = data ? { href: data.bookmark ? `#${data.bookmark}` : data.url ?? '', ...(data.tooltip ? { tooltip: data.tooltip } : {}) } : null;
    if (await applyWorkerProperty(controller, link.query, { kind: 'setHyperlink', payload: { selection: link.query.selection,
      hyperlink, expectVersion: link.query.version } }, props.onError)) setLink(null);
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
  return { openLink, openNotes, openImage, openImagePosition: () => openImage(true), imagePropertiesAvailable: object?.value?.kind === 'object',
    openPageSetup: page.handleOpenPageSetup, openWatermark: watermark.handleOpenWatermark, openSplitCell,
    element: <>
      {page.showPageSetup && <PageSetupDialog isOpen onClose={page.close} currentProps={page.currentProps} onApply={page.handlePageSetupApply} />}
      {watermark.showWatermark && <WatermarkDialog isOpen onClose={watermark.close} current={watermark.currentWatermark}
        presets={props.watermarkPresets} onApply={watermark.handleWatermarkApply} />}
      {link && <HyperlinkDialog isOpen onClose={() => setLink(null)} isEditing={hyperlink !== null} selectedText={link.text} displayTextReadOnly
        initialData={{ url: hyperlink?.href, tooltip: hyperlink?.tooltip ?? undefined, displayText: link.text }} onSubmit={applyLink}
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
