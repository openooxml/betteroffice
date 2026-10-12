import { useState } from 'react';
import type { DocxEditorProps } from '../../DocxEditor';
import { HyperlinkDialog, type HyperlinkData } from '../../dialogs/HyperlinkDialog';
import { PageSetupDialog } from '../../dialogs/PageSetupDialog';
import { WatermarkDialog } from '../../dialogs/WatermarkDialog';
import { useWorkerPageSetupControls, useWorkerWatermarkControls } from './workerPropertyControls';
import { applyWorkerProperty, type WorkerPropertyQuery, type WorkerShellController } from './workerProperties';

export function useWorkerPropertyDialogs(controller: WorkerShellController, props: Pick<DocxEditorProps, 'onError' | 'watermarkPresets'>) {
  const page = useWorkerPageSetupControls(controller, props.onError);
  const watermark = useWorkerWatermarkControls(controller, props.onError);
  const [link, setLink] = useState<{ query: WorkerPropertyQuery; text: string } | null>(null);
  const report = (error: unknown) => props.onError?.(error instanceof Error ? error : new Error(String(error)));
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
  const applyLink = async (data: HyperlinkData | null) => {
    if (!link?.query.selection) return;
    const hyperlink = data ? { href: data.bookmark ? `#${data.bookmark}` : data.url ?? '', ...(data.tooltip ? { tooltip: data.tooltip } : {}) } : null;
    if (await applyWorkerProperty(controller, link.query, { kind: 'setHyperlink', payload: { selection: link.query.selection,
      hyperlink, expectVersion: link.query.version } }, props.onError)) setLink(null);
  };
  const hyperlink = link?.query.value?.kind === 'hyperlink' ? link.query.value.hyperlink : null;
  return { openLink,
    openPageSetup: page.handleOpenPageSetup, openWatermark: watermark.handleOpenWatermark,
    element: <>
      {page.showPageSetup && <PageSetupDialog isOpen onClose={page.close} currentProps={page.currentProps} onApply={page.handlePageSetupApply} />}
      {watermark.showWatermark && <WatermarkDialog isOpen onClose={watermark.close} current={watermark.currentWatermark}
        presets={props.watermarkPresets} onApply={watermark.handleWatermarkApply} />}
      {link && <HyperlinkDialog isOpen onClose={() => setLink(null)} isEditing={hyperlink !== null} selectedText={link.text} displayTextReadOnly
        initialData={{ url: hyperlink?.href, tooltip: hyperlink?.tooltip ?? undefined, displayText: link.text }} onSubmit={applyLink}
        onRemove={hyperlink ? () => { void applyLink(null); } : undefined} />}
    </> };
}
