import { useCallback, useState } from 'react';
import type { Document, Watermark } from '@betteroffice/docx/types/document';
import { getDocumentWatermark, setDocumentWatermark } from '@betteroffice/docx/docx';

/**
 * Watermark dialog controls. Watermarks are immutable header-part metadata,
 * outside the editable Yrs stories, and therefore use the Document API.
 */
export function useWatermarkControls({
  readOnly,
  document,
  pushDocument,
}: {
  readOnly: boolean;
  document: Document | null;
  pushDocument: (document: Document) => void;
}) {
  const [showWatermark, setShowWatermark] = useState(false);
  const handleOpenWatermark = useCallback(() => setShowWatermark(true), []);

  const currentWatermark = getDocumentWatermark(document);

  const handleWatermarkApply = useCallback(
    (watermark: Watermark | null) => {
      if (readOnly) return;
      if (document) pushDocument(setDocumentWatermark(document, watermark));
    },
    [document, pushDocument, readOnly]
  );

  return {
    showWatermark,
    setShowWatermark,
    handleOpenWatermark,
    currentWatermark,
    handleWatermarkApply,
  };
}
