import { useCallback, useRef } from 'react';
import type { Comment } from '@betteroffice/docx/types/content';
import {
  createDocx,
  injectReplyRangeMarkers,
  injectTCReplyRangeMarkers,
  repackDocx,
} from '@betteroffice/docx/docx';
import { readDocxFileFromInput, type DocxInput } from '@betteroffice/docx/utils';
import { openPrintWindow } from '@betteroffice/docx';
import {
  rasterizeDisplayListPages,
  type DisplayList,
  type ImageResolver,
} from '@betteroffice/docx/layout/render';
import type { PagedEditorRef } from '../PagedEditor';

const INSERT_IMAGE_MAX_WIDTH_PX = 612;

function toFileIOError(error: unknown, fallbackMessage: string): Error {
  return error instanceof Error ? error : new Error(fallbackMessage);
}

// Page-break CSS for the print popup. The core `openPrintWindow` already zeros
// the page margins; these rules put one canvas raster per printed sheet.
// — never `document.write` / `innerHTML` — per the print security contract in
// the repo security guidelines.
const PRINT_CANVAS_CSS =
  '* { margin: 0; padding: 0; }\n' +
  'body { background: #fff; }\n' +
  'img.print-page { display: block; width: 100%; break-after: page; }\n' +
  'img.print-page:last-child { break-after: auto; }\n' +
  '@page { margin: 0; size: auto; }';

// Print once fonts + images have settled, then close. Prints as soon as
// everything is ready (usually well under the cap) with a hard timeout so a
// browser that never resolves `fonts.ready`/`decode()` still prints.
function finishPrint(w: Window, images: HTMLImageElement[] = []): void {
  let done = false;
  const runPrint = () => {
    if (done || w.closed) return;
    done = true;
    w.focus();
    w.print();
    w.close();
  };
  Promise.all([
    w.document.fonts?.ready ?? Promise.resolve(),
    ...images.map((img) => img.decode().catch(() => undefined)),
  ]).then(runPrint, runPrint);
  setTimeout(runPrint, 2000);
}

/**
 * Print from the experimental canvas renderer: raster each visible page canvas
 * to a PNG and print one image per sheet. Fully painter-independent, so print
 * keeps working once the DOM painter is retired. Built through the hardened
 * `openPrintWindow` (no `document.write`) with every node created via DOM APIs;
 * the PNG data URLs are machine-generated base64 assigned as `img.src`
 * properties, never interpolated into markup.
 */
function printDisplayListPages(
  displayList: DisplayList,
  resolveImage: ImageResolver,
  onPrint?: () => void
): void {
  const w = openPrintWindow('Print', '');
  if (!w) {
    window.print();
    onPrint?.();
    return;
  }
  const style = w.document.createElement('style');
  style.textContent = PRINT_CANVAS_CSS;
  w.document.head.appendChild(style);

  void rasterizeDisplayListPages(displayList, { resolveImage }).then(
    (canvases) => {
      const images: HTMLImageElement[] = [];
      for (const canvas of canvases) {
        try {
          const img = w.document.createElement('img');
          img.className = 'print-page';
          img.src = canvas.toDataURL('image/png');
          w.document.body.appendChild(img);
          images.push(img);
        } catch {
          // Skip an unexpectedly tainted page without exposing markup.
        }
      }
      finishPrint(w, images);
      onPrint?.();
    },
    () => {
      w.close();
      window.print();
      onPrint?.();
    }
  );
}

/**
 * File-IO surface of the editor: save (to buffer), download, print, open
 * a DOCX from disk, insert an image from disk. The two file <input> refs
 * live here too because they're hidden inputs whose `click()` is wrapped
 * by the trigger callbacks.
 *
 * Image insertion targets the authoritative body Yrs selection.
 */
export function useFileIO({
  pagedEditorRef,
  displayList,
  resolveImage,
  comments,
  documentName,
  onSave,
  onOpen,
  onError,
  onPrint,
  onDocumentNameChange,
  loadBuffer,
  focusActiveEditor,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  displayList: DisplayList | null;
  resolveImage: ImageResolver;
  comments: Comment[];
  documentName: string | undefined;
  onSave: ((buffer: ArrayBuffer) => void) | undefined;
  onOpen: ((file: File) => void | Promise<void>) | undefined;
  onError: ((error: Error) => void) | undefined;
  onPrint: (() => void) | undefined;
  onDocumentNameChange: ((name: string) => void) | undefined;
  loadBuffer: (buffer: DocxInput) => Promise<void>;
  focusActiveEditor: () => void;
}) {
  const imageInputRef = useRef<HTMLInputElement>(null);
  const docxInputRef = useRef<HTMLInputElement>(null);

  const handleSave = useCallback(
    async (): Promise<ArrayBuffer | null> => {
      try {
        const document = pagedEditorRef.current?.getDocument();
        if (!document) return null;

        // Sync React comments state (including new replies) back to the document model
        document.package.document.comments = comments;

        // Inject commentRangeStart/End for reply comments that share the parent's range.
        // Pages/Word require every comment (including replies) to have range markers in document.xml.
        injectReplyRangeMarkers(document.package.document.content, comments);
        // Also inject range markers for comments that reply to tracked changes.
        injectTCReplyRangeMarkers(document.package.document.content, comments);

        const buffer = document.originalBuffer
          ? await repackDocx(document)
          : await createDocx(document);
        document.originalBuffer = buffer;

        onSave?.(buffer);
        return buffer;
      } catch (error) {
        onError?.(toFileIOError(error, 'Failed to save document'));
        return null;
      }
    },
    [pagedEditorRef, comments, onSave, onError]
  );

  const handleDirectPrint = useCallback(() => {
    if (!displayList) {
      window.print();
      onPrint?.();
      return;
    }
    printDisplayListPages(displayList, resolveImage, onPrint);
  }, [displayList, resolveImage, onPrint]);

  const handleDownloadDocument = useCallback(async () => {
    const buffer = await handleSave();
    if (!buffer) return;
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const url = URL.createObjectURL(blob);
    const a = window.document.createElement('a');
    a.href = url;
    a.download = `${(documentName?.trim() || 'document').replace(/\.docx$/i, '')}.docx`;
    a.click();
    // Defer revoke so Safari has time to start the download.
    setTimeout(() => URL.revokeObjectURL(url), 0);
  }, [handleSave, documentName]);

  const handleOpenDocument = useCallback(() => {
    docxInputRef.current?.click();
  }, []);

  const handleDocxFileChange = useCallback(
    async (event: React.ChangeEvent<HTMLInputElement>) => {
      if (onOpen) {
        const input = event.currentTarget;
        const file = input.files?.[0];
        input.value = '';
        if (!file) return;

        try {
          await onOpen(file);
        } catch (error) {
          onError?.(toFileIOError(error, 'Failed to open document'));
        }
        return;
      }

      try {
        const result = await readDocxFileFromInput(event.nativeEvent);
        if (!result) return;
        await loadBuffer(result.buffer);
        onDocumentNameChange?.(result.name);
      } catch (error) {
        onError?.(toFileIOError(error, 'Failed to open document'));
      }
    },
    [loadBuffer, onDocumentNameChange, onError, onOpen]
  );

  const handleInsertImageClick = useCallback(() => {
    imageInputRef.current?.click();
  }, []);

  const handleImageFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      // Reset the input so the same file can be selected again
      e.target.value = '';
      if (!file) return;

      const reader = new FileReader();
      reader.onload = () => {
        const dataUrl = reader.result;
        if (typeof dataUrl !== 'string') return;
        const image = new Image();
        image.onload = () => {
          let width = image.naturalWidth;
          let height = image.naturalHeight;
          if (width > INSERT_IMAGE_MAX_WIDTH_PX) {
            height = Math.round(height * (INSERT_IMAGE_MAX_WIDTH_PX / width));
            width = INSERT_IMAGE_MAX_WIDTH_PX;
          }
          const rId = `rId_img_${Date.now()}_${Math.round(Math.random() * 1e9)}`;
          const inserted = pagedEditorRef.current?.applyYrsCommand({
            type: 'insertImage',
            image: {
              src: dataUrl,
              alt: file.name,
              width,
              height,
              rId,
              wrapType: 'inline',
              displayMode: 'inline',
            },
          });
          if (inserted) focusActiveEditor();
        };
        image.onerror = () => onError?.(new Error('Failed to decode image'));
        image.src = dataUrl;
      };
      reader.onerror = () => onError?.(reader.error ?? new Error('Failed to read image'));
      reader.readAsDataURL(file);
    },
    [focusActiveEditor, onError, pagedEditorRef]
  );

  return {
    imageInputRef,
    docxInputRef,
    handleSave,
    handleDirectPrint,
    handleDownloadDocument,
    handleOpenDocument,
    handleDocxFileChange,
    handleInsertImageClick,
    handleImageFileChange,
  };
}
