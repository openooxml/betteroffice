import type { DocxCommandFrame } from '@betteroffice/docx/yrs';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';

export function createWorkerPrintCapture(binding: PeerlessEditSessionBinding,
  render: (bitmap: ImageBitmap) => string, isCurrent: () => boolean) {
  let version: number | undefined;
  let failure: Error | null = null;
  const images = new Map<number, string>();
  const assertCurrent = () => {
    if (!isCurrent() || binding.viewStore.getSnapshot().version !== version) throw new Error('Document changed while printing');
    if (failure) throw failure;
  };
  return {
    assertCurrent,
    accept(frame: DocxCommandFrame) {
      if (version === undefined) return;
      try {
        assertCurrent();
        for (const { page, bitmap } of frame.payload.stagedPresentation?.pages ?? []) images.set(page.pageIndex, render(bitmap));
      } catch (cause) { failure = cause instanceof Error ? cause : new Error(String(cause)); }
    },
    async collect(buildPages: (pages: number[], frameEpoch: number) => Promise<unknown>) {
      await binding.admission.whenCompositionSettled?.();
      await binding.session.flush();
      if (!isCurrent()) throw new Error('Document changed while printing');
      const snapshot = binding.viewStore.getSnapshot();
      version = snapshot.version;
      const pages = snapshot.geometry?.pages;
      if (!pages?.length) throw new Error('Worker print layout is unavailable');
      for (const { page } of pages) {
        assertCurrent();
        await buildPages([page.pageIndex], binding.viewStore.getSnapshot().geometry?.frameEpoch ?? 0);
        assertCurrent();
        if (!images.has(page.pageIndex)) throw new Error(`Worker print pixels are unavailable for page ${page.pageIndex + 1}`);
      }
      return pages.map(({ page }) => images.get(page.pageIndex)!);
    },
  };
}

export function printBitmap(bitmap: ImageBitmap): string {
  const canvas = document.createElement('canvas');
  canvas.width = bitmap.width; canvas.height = bitmap.height;
  try {
    const context = canvas.getContext('2d');
    if (!context) throw new Error('Print canvas is unavailable');
    context.drawImage(bitmap, 0, 0);
    return canvas.toDataURL('image/png');
  } finally { canvas.width = 0; canvas.height = 0; }
}
