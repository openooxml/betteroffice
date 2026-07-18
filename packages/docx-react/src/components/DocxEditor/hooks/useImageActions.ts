import { useCallback, useState } from 'react';
import type {
  Document,
  FootnoteProperties,
  EndnoteProperties,
} from '@betteroffice/docx/types/document';
import { toolbarValueToLayoutTarget } from '@betteroffice/docx/docx';
import {
  captureInlinePositionEmuFromDisplayList,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import type { ImagePositionData } from '../../dialogs/ImagePositionDialog';
import type { ImagePropertiesData } from '../../dialogs/ImagePropertiesDialog';
import type { PagedEditorRef } from '../PagedEditor';

/** Minimal shape the hook needs from the parent's selection-tracker state. */
interface ImageContext {
  pos: number;
  wrapType?: string;
}

/**
 * Image-related dialogs and toolbar actions:
 *  - wrap type (inline ↔ float-wrap variants) via setImageWrapType
 *  - 90° rotate + horizontal/vertical flip via transform attr
 *  - position dialog (horizontal/vertical anchor + distFrom* offsets)
 *  - properties dialog (alt text, border, width/height)
 *  - footnote/endnote properties dialog (footnote numbering/format)
 *
 * Owns the open/closed state for each dialog; the JSX consumer reads the
 * `*Open` flags + the apply/cancel callbacks. `pmImageContext` comes
 * from the parent's selection-tracker state because it's set by the
 * image right-click flow.
 */
export function useImageActions({
  document,
  pmImageContext,
  displayListQueries,
  pagedEditorRef,
  focusActiveEditor,
  pushDocument,
}: {
  document: Document | null;
  pmImageContext: ImageContext | null | undefined;
  displayListQueries: DisplayListQueries | null;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  focusActiveEditor: () => void;
  pushDocument: (doc: Document) => void;
}) {
  const [imagePositionOpen, setImagePositionOpen] = useState(false);
  const [imagePropsOpen, setImagePropsOpen] = useState(false);
  const [footnotePropsOpen, setFootnotePropsOpen] = useState(false);

  const handleImageWrapType = useCallback(
    (toolbarValue: string) => {
      if (!pmImageContext) return;
      const pos = pmImageContext.pos;

      // The toolbar and the right-click menu share `setImageWrapType` and its
      // `resolveAnchorAttrs` taxonomy. `toolbarValueToLayoutTarget` lives in
      // core so the Vue adapter doesn't have to duplicate it.
      const target = toolbarValueToLayoutTarget(toolbarValue);
      if (!target) return;

      // For inline → anchor, capture the inline glyph's rendered offset so the
      // new float lands at the same X/Y (Word's behavior). The core helper
      // handles the zoom + EMU conversion uniformly.
      let opts: { initialPositionEmu?: { horizontalEmu: number; verticalEmu: number } } | undefined;
      if (pmImageContext.wrapType === 'inline' && target !== 'inline') {
        const captured = displayListQueries
          ? captureInlinePositionEmuFromDisplayList(displayListQueries, pos)
          : undefined;
        if (captured) opts = { initialPositionEmu: captured };
      }

      if (pagedEditorRef.current?.applyYrsCommand({
        type: 'imageWrap',
        pmPos: pos,
        target,
        options: opts,
      })) {
        focusActiveEditor();
      }
    },
    [displayListQueries, focusActiveEditor, pagedEditorRef, pmImageContext]
  );

  const handleImageTransform = useCallback(
    (action: 'rotateCW' | 'rotateCCW' | 'flipH' | 'flipV') => {
      if (!pmImageContext) return;
      if (pagedEditorRef.current?.applyYrsCommand({
        type: 'imageTransform',
        pmPos: pmImageContext.pos,
        action,
      })) {
        focusActiveEditor();
      }
    },
    [focusActiveEditor, pagedEditorRef, pmImageContext]
  );

  const handleApplyImagePosition = useCallback(
    (data: ImagePositionData) => {
      if (!pmImageContext) return;
      const pos = pmImageContext.pos;
      const patch = {
        position: {
          horizontal: data.horizontal,
          vertical: data.vertical,
        },
        ...(data.distTop != null ? { distTop: data.distTop } : {}),
        ...(data.distBottom != null ? { distBottom: data.distBottom } : {}),
        ...(data.distLeft != null ? { distLeft: data.distLeft } : {}),
        ...(data.distRight != null ? { distRight: data.distRight } : {}),
      };
      if (pagedEditorRef.current?.applyYrsCommand({ type: 'imageGeometry', pmPos: pos, patch })) {
        focusActiveEditor();
      }
    },
    [focusActiveEditor, pagedEditorRef, pmImageContext]
  );

  const handleOpenImageProperties = useCallback(() => {
    setImagePropsOpen(true);
  }, []);

  const handleApplyImageProperties = useCallback(
    (data: ImagePropertiesData) => {
      if (!pmImageContext) return;
      const pos = pmImageContext.pos;
      const patch = {
        alt: data.alt ?? null,
        borderWidth: data.borderWidth ?? null,
        borderColor: data.borderColor ?? null,
        borderStyle: data.borderStyle ?? null,
        width: data.width ?? null,
        height: data.height ?? null,
      };
      if (pagedEditorRef.current?.applyYrsCommand({ type: 'imageGeometry', pmPos: pos, patch })) {
        focusActiveEditor();
      }
    },
    [focusActiveEditor, pagedEditorRef, pmImageContext]
  );

  const handleApplyFootnoteProperties = useCallback(
    (footnotePr: FootnoteProperties, endnotePr: EndnoteProperties) => {
      if (!document?.package) return;
      const newDoc = {
        ...document.package.document,
        finalSectionProperties: {
          ...document.package.document.finalSectionProperties,
          footnotePr,
          endnotePr,
        },
      };
      pushDocument({
        ...document,
        package: {
          ...document.package,
          document: newDoc,
        },
      });
    },
    [document, pushDocument]
  );

  return {
    imagePositionOpen,
    setImagePositionOpen,
    imagePropsOpen,
    setImagePropsOpen,
    footnotePropsOpen,
    setFootnotePropsOpen,
    handleImageWrapType,
    handleImageTransform,
    handleApplyImagePosition,
    handleOpenImageProperties,
    handleApplyImageProperties,
    handleApplyFootnoteProperties,
  };
}
