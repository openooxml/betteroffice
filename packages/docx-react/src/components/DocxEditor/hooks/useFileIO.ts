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
import type { DocxEditorProps } from '../../DocxEditor';
import type { DocxImageInsert, DocxSaveOutcome } from './useDocxCommands';

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

/** Prints display-list pages as PNG images without interpolating data into markup. */
async function printDisplayListPages(
  w: Window,
  displayList: DisplayList,
  resolveImage: ImageResolver
): Promise<void> {
  const style = w.document.createElement('style');
  style.textContent = PRINT_CANVAS_CSS;
  w.document.head.appendChild(style);
  const canvases = await rasterizeDisplayListPages(displayList, { resolveImage });
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
}

/** A print window opened during a user gesture, filled once the document has settled. */
export interface DocxPrintJob {
  /** Prints `displayList`; false when the reserved window was closed first. */
  finish(displayList: DisplayList): Promise<boolean>;
  cancel(): void;
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
  resolveImage,
  comments,
  documentName,
  onSave,
  onSaveRequest,
  downloadOnSave = true,
  onOpen,
  onError,
  onPrint,
  onDocumentNameChange,
  loadBuffer,
  focusActiveEditor,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  resolveImage: ImageResolver;
  comments: Comment[];
  documentName: string | undefined;
  onSave: ((buffer: ArrayBuffer) => void) | undefined;
  onSaveRequest?: DocxEditorProps['onSaveRequest'];
  downloadOnSave?: boolean;
  onOpen: ((file: File) => void | Promise<void>) | undefined;
  onError: ((error: Error) => void) | undefined;
  onPrint: (() => void) | undefined;
  onDocumentNameChange: ((name: string) => void) | undefined;
  loadBuffer: (buffer: DocxInput) => Promise<void>;
  focusActiveEditor: () => void;
}) {
  const imageInputRef = useRef<HTMLInputElement>(null);
  const imageInsertRef = useRef<DocxImageInsert | null>(null);
  const docxInputRef = useRef<HTMLInputElement>(null);
  const saveRequestRef = useRef<Promise<DocxSaveOutcome> | null>(null);

  const handleSave = useCallback(
    async (): Promise<ArrayBuffer | null> => {
      try {
        const editor = pagedEditorRef.current;
        if (!editor) return null;
        const session = editor.getYrsSession();
        await editor.flushPendingInput();
        const assertCurrent = () => {
          if (editor !== pagedEditorRef.current || session !== editor.getYrsSession()) {
            throw new Error('The document changed while saving');
          }
        };
        assertCurrent();
        const document = editor.getDocument();
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
        assertCurrent();
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

  const reservePrint = useCallback((): DocxPrintJob => {
    const w = openPrintWindow('Print', '');
    return {
      async finish(displayList) {
        if (!w) {
          window.print();
        } else {
          if (w.closed) return false;
          await printDisplayListPages(w, displayList, resolveImage);
        }
        onPrint?.();
        return true;
      },
      cancel() {
        if (w && !w.closed) w.close();
      },
    };
  }, [resolveImage, onPrint]);

  const handleDownloadDocument = useCallback((): Promise<DocxSaveOutcome> => {
    if (saveRequestRef.current) return saveRequestRef.current;
    const pending = Promise.resolve().then(async (): Promise<DocxSaveOutcome> => {
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      if (onSaveRequest && (await onSaveRequest()) !== true) return 'requested';
      if (editor !== pagedEditorRef.current || session !== editor?.getYrsSession()) {
        throw new Error('The document changed during the save request');
      }
      const buffer = await handleSave();
      if (!buffer) return 'failed';
      if (!downloadOnSave) return 'saved';
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      const url = URL.createObjectURL(blob);
      const a = window.document.createElement('a');
      a.href = url;
      a.download = `${(documentName?.trim() || 'document').replace(/\.docx$/i, '')}.docx`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 0);
      return 'saved';
    }).catch((error): DocxSaveOutcome => {
      onError?.(toFileIOError(error, 'Failed to save document'));
      return 'failed';
    }).finally(() => {
      if (saveRequestRef.current === pending) saveRequestRef.current = null;
    });
    saveRequestRef.current = pending;
    return pending;
  }, [handleSave, documentName, downloadOnSave, onSaveRequest, onError, pagedEditorRef]);

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

  /** Opens the picker; `insert` stays with the file chosen in it through decoding. */
  const handleInsertImageClick = useCallback((insert?: DocxImageInsert) => {
    imageInsertRef.current = insert ?? null;
    imageInputRef.current?.click();
  }, []);

  const handleImageFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      const insert = imageInsertRef.current;
      imageInsertRef.current = null;
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
          const picture = {
            src: dataUrl,
            alt: file.name,
            width,
            height,
            rId,
            wrapType: 'inline',
            displayMode: 'inline',
          };
          if (insert) {
            void insert(picture).then((result) => {
              if (result.ok && result.status === 'executed') focusActiveEditor();
            });
            return;
          }
          const inserted = pagedEditorRef.current?.applyYrsCommand({
            type: 'insertImage',
            image: picture,
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
    reservePrint,
    handleDownloadDocument,
    handleOpenDocument,
    handleDocxFileChange,
    handleInsertImageClick,
    handleImageFileChange,
  };
}
