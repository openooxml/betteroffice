import { useCallback, useState } from 'react';
import type {
  Document,
  FootnoteProperties,
  EndnoteProperties,
} from '@betteroffice/docx/types/document';
import { updateFinalSectionProperties } from '@betteroffice/docx/editor';
import type { ImagePositionData } from '../../dialogs/ImagePositionDialog';
import type { ImagePropertiesData } from '../../dialogs/ImagePropertiesDialog';

/** Minimal shape the hook needs from the parent's selection-tracker state. */
interface ImageContext {
  pos: number;
  wrapType?: string;
}

/**
 * Image and note dialogs:
 *  - position dialog (horizontal/vertical anchor + distFrom* offsets)
 *  - properties dialog (alt text, border, width/height)
 *  - footnote/endnote properties dialog (footnote numbering/format)
 *
 * Owns the open/closed state for each dialog; the JSX consumer reads the
 * `*Open` flags + the apply/cancel callbacks. `applyGeometry` writes to the
 * image the dialog opened for, resolved again by its identity.
 */
export function useImageActions({
  document,
  pmImageContext,
  applyGeometry,
  pushDocument,
}: {
  document: Document | null;
  pmImageContext: ImageContext | null | undefined;
  applyGeometry: (patch: Readonly<Record<string, unknown>>) => void;
  pushDocument: (doc: Document) => void;
}) {
  const [imagePositionOpen, setImagePositionOpen] = useState(false);
  const [imagePropsOpen, setImagePropsOpen] = useState(false);
  const [footnotePropsOpen, setFootnotePropsOpen] = useState(false);

  const handleApplyImagePosition = useCallback(
    (data: ImagePositionData) => {
      if (!pmImageContext) return;
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
      applyGeometry(patch);
    },
    [applyGeometry, pmImageContext]
  );

  const handleOpenImageProperties = useCallback(() => {
    setImagePropsOpen(true);
  }, []);

  const handleApplyImageProperties = useCallback(
    (data: ImagePropertiesData) => {
      if (!pmImageContext) return;
      const patch = {
        alt: data.alt ?? null,
        borderWidth: data.borderWidth ?? null,
        borderColor: data.borderColor ?? null,
        borderStyle: data.borderStyle ?? null,
        width: data.width ?? null,
        height: data.height ?? null,
      };
      applyGeometry(patch);
    },
    [applyGeometry, pmImageContext]
  );

  const handleApplyFootnoteProperties = useCallback(
    (footnotePr: FootnoteProperties, endnotePr: EndnoteProperties) => {
      if (!document?.package) return;
      pushDocument({
        ...document,
        package: {
          ...document.package,
          document: updateFinalSectionProperties(document.package.document, (properties) => ({
            ...properties,
            footnotePr,
            endnotePr,
          })),
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
    handleApplyImagePosition,
    handleOpenImageProperties,
    handleApplyImageProperties,
    handleApplyFootnoteProperties,
  };
}
