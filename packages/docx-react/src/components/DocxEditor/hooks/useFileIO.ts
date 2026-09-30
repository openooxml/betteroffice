import { useCallback, useRef } from 'react';
import type { Comment } from '@betteroffice/docx/types/content';
import type { Document } from '@betteroffice/docx/types/document';
import {
  createDocx,
  injectReplyRangeMarkers,
  injectTCReplyRangeMarkers,
  repackDocx,
} from '@betteroffice/docx/docx';
import { readDocxFileFromInput, type DocxInput } from '@betteroffice/docx/utils';
import {
  captureSessionSave,
  writeSessionSave,
  yrsToDocument,
  type DocxSessionSave,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { openPrintWindow } from '@betteroffice/docx';
import {
  rasterizeDisplayListPages,
  type DisplayList,
  type ImageResolver,
} from '@betteroffice/docx/layout/render';
import type { PagedEditorRef } from '../PagedEditor';
import { flushedSession } from '../editorBatches';
import { dirtyProjectionStory } from './useYrsCoreSession';
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

// Resolves once fonts + images have settled (usually well under the cap), with
// a hard timeout so a browser that never resolves `fonts.ready`/`decode()`
// still prints.
function settled(w: Window, images: HTMLImageElement[]): Promise<void> {
  const loaded = Promise.all([
    w.document.fonts?.ready ?? Promise.resolve(),
    ...images.map((img) => img.decode().catch(() => undefined)),
  ]).then(() => undefined, () => undefined);
  return Promise.race([loaded, new Promise<void>((resolve) => setTimeout(resolve, 2000))]);
}

/** Fills `w` with display-list pages as PNG images without interpolating data into markup. */
async function renderDisplayListPages(
  w: Window,
  displayList: DisplayList,
  resolveImage: ImageResolver,
  fontFamilies: ReadonlyMap<string, string> | undefined
): Promise<void> {
  const style = w.document.createElement('style');
  style.textContent = PRINT_CANVAS_CSS;
  w.document.head.appendChild(style);
  const canvases = await rasterizeDisplayListPages(displayList, { resolveImage, fontFamilies });
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
  await settled(w, images);
}

/** A print window opened during a user gesture, filled once the document has settled. */
export interface DocxPrintJob {
  /** Renders `displayList` into the reserved window and waits for its pages to load. */
  prepare(displayList: DisplayList): Promise<void>;
  /** Prints the prepared pages; false when the reserved window was closed first. */
  print(): boolean;
  cancel(): void;
}

/** Writes the editor's document, through the session save when it has one. */
async function writeEditorDocument(
  document: Document,
  session: YrsSession | null,
  capture: DocxSessionSave | null
): Promise<ArrayBuffer> {
  const original = document.originalBuffer;
  if (!original) return createDocx(document);
  if (!session || !capture) return repackDocx(document);
  // The original buffer can be the last save rather than the session source, so none is patched.
  const { bytes } = await writeSessionSave(session, document, capture, original, {}, () => false);
  return bytes.buffer as ArrayBuffer;
}

/**
 * `document` with the comments a save writes, the stories they are anchored
 * in projected again with them: the editor projects its host's comments.
 */
function withSavedComments(document: Document, session: YrsSession, comments: Comment[]): Document {
  const base: Document = {
    ...document,
    package: { ...document.package, document: { ...document.package.document, comments } },
  };
  const storyIds = new Set<string>();
  for (const comment of comments) {
    try {
      for (const anchor of session.resolveComment(String(comment.id))) {
        storyIds.add(dirtyProjectionStory(anchor.story));
      }
    } catch {
      // Replies and comments whose anchors are gone hold no range.
    }
  }
  return storyIds.size > 0 ? yrsToDocument(session, base, { storyIds }) : base;
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
  shownImageResolver,
  fontFamilies,
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
  /** The resolver of the frame published last; print reads it once its display list settles. */
  shownImageResolver?: () => ImageResolver;
  /** The CSS family each document font family paints browser text with, where they differ. */
  fontFamilies?: ReadonlyMap<string, string>;
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
        if (!pagedEditorRef.current) return null;
        const { editor, session } = await flushedSession(pagedEditorRef);
        if (session.isDisplayOnly?.()) throw new Error('The document is still opening');
        const projected = editor.getDocument();
        if (!projected) return null;
        const capture = projected.originalBuffer ? captureSessionSave(session) : null;
        const document = withSavedComments(projected, session, comments);

        // Inject commentRangeStart/End for reply comments that share the parent's range.
        // Pages/Word require every comment (including replies) to have range markers in document.xml.
        injectReplyRangeMarkers(document.package.document.content, comments);
        // Also inject range markers for comments that reply to tracked changes.
        injectTCReplyRangeMarkers(document.package.document.content, comments);

        const buffer = await writeEditorDocument(document, session, capture);
        if (pagedEditorRef.current?.getYrsSession() !== session) {
          throw new Error('The document changed while saving');
        }
        projected.originalBuffer = buffer;

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
      async prepare(displayList) {
        if (w && !w.closed) {
          await renderDisplayListPages(
            w,
            displayList,
            shownImageResolver?.() ?? resolveImage,
            fontFamilies
          );
        }
      },
      print() {
        if (!w) {
          window.print();
        } else {
          if (w.closed) return false;
          w.focus();
          w.print();
          w.close();
        }
        onPrint?.();
        return true;
      },
      cancel() {
        if (w && !w.closed) w.close();
      },
    };
  }, [resolveImage, shownImageResolver, fontFamilies, onPrint]);

  const handleDownloadDocument = useCallback((): Promise<DocxSaveOutcome> => {
    if (saveRequestRef.current) return saveRequestRef.current;
    const pending = Promise.resolve().then(async (): Promise<DocxSaveOutcome> => {
      const session = pagedEditorRef.current?.getYrsSession();
      if (onSaveRequest && (await onSaveRequest()) !== true) return 'requested';
      if (session !== pagedEditorRef.current?.getYrsSession()) {
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
