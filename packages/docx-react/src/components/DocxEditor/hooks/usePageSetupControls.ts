import { useCallback, useMemo, useState } from 'react';
import type { Document, SectionProperties } from '@betteroffice/docx/types/document';
import type { PagedEditorRef } from '../PagedEditor';

/**
 * Page setup + ruler controls: page-level margin handlers (header/footer
 * page setup dialog + drag in the rulers), paragraph indent handlers, and
 * tab-stop removal. Margin changes go through `handleDocumentChange` so
 * they land in the undo/redo history; indent and tab-stop changes
 * apply through the authoritative Yrs session.
 */
export function usePageSetupControls({
  document,
  readOnly,
  handleDocumentChange,
  pagedEditorRef,
}: {
  document: Document | null;
  readOnly: boolean;
  handleDocumentChange: (doc: Document) => void;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
}) {
  const [showPageSetup, setShowPageSetup] = useState(false);
  const handleOpenPageSetup = useCallback(() => setShowPageSetup(true), []);

  const createMarginHandler = useCallback(
    (property: 'marginLeft' | 'marginRight' | 'marginTop' | 'marginBottom') =>
      (marginTwips: number) => {
        if (!document || readOnly) return;
        const newDoc = {
          ...document,
          package: {
            ...document.package,
            document: {
              ...document.package.document,
              finalSectionProperties: {
                ...document.package.document.finalSectionProperties,
                [property]: marginTwips,
              },
            },
          },
        };
        handleDocumentChange(newDoc);
      },
    [document, readOnly, handleDocumentChange]
  );

  const handleLeftMarginChange = useMemo(
    () => createMarginHandler('marginLeft'),
    [createMarginHandler]
  );
  const handleRightMarginChange = useMemo(
    () => createMarginHandler('marginRight'),
    [createMarginHandler]
  );
  const handleTopMarginChange = useMemo(
    () => createMarginHandler('marginTop'),
    [createMarginHandler]
  );
  const handleBottomMarginChange = useMemo(
    () => createMarginHandler('marginBottom'),
    [createMarginHandler]
  );

  const handlePageSetupApply = useCallback(
    (props: Partial<SectionProperties>) => {
      if (!document || readOnly) return;
      const newDoc = {
        ...document,
        package: {
          ...document.package,
          document: {
            ...document.package.document,
            finalSectionProperties: {
              ...document.package.document.finalSectionProperties,
              ...props,
            },
          },
        },
      };
      handleDocumentChange(newDoc);
    },
    [document, readOnly, handleDocumentChange]
  );

  const handleIndentLeftChange = useCallback(
    (twips: number) => {
      pagedEditorRef.current?.applyYrsCommand({
        type: 'paragraphAttrs',
        attrs: { indentLeft: twips > 0 ? twips : null },
      });
    },
    [pagedEditorRef]
  );

  const handleIndentRightChange = useCallback(
    (twips: number) => {
      pagedEditorRef.current?.applyYrsCommand({
        type: 'paragraphAttrs',
        attrs: { indentRight: twips > 0 ? twips : null },
      });
    },
    [pagedEditorRef]
  );

  const handleFirstLineIndentChange = useCallback(
    (twips: number) => {
      pagedEditorRef.current?.applyYrsCommand({
        type: 'paragraphAttrs',
        attrs: {
          indentFirstLine: Math.abs(twips) > 0 ? Math.abs(twips) : null,
          hangingIndent: twips < 0,
        },
      });
    },
    [pagedEditorRef]
  );

  const handleTabStopRemove = useCallback(
    (positionTwips: number) => {
      pagedEditorRef.current?.applyYrsCommand({ type: 'removeTabStop', positionTwips });
    },
    [pagedEditorRef]
  );

  return {
    showPageSetup,
    setShowPageSetup,
    handleOpenPageSetup,
    handleLeftMarginChange,
    handleRightMarginChange,
    handleTopMarginChange,
    handleBottomMarginChange,
    handlePageSetupApply,
    handleIndentLeftChange,
    handleIndentRightChange,
    handleFirstLineIndentChange,
    handleTabStopRemove,
  };
}
