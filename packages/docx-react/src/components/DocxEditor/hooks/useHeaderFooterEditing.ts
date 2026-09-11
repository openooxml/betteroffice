import { useCallback, useMemo } from 'react';
import type { Document, HeaderFooter } from '@betteroffice/docx/types/document';
import {
  resolvedFinalSectionProperties,
  updateFinalSectionProperties,
} from '@betteroffice/docx/editor';
import { resolveHeaderFooter } from '@betteroffice/docx/layout';

import type { PartEditTarget } from '../partEdit';

/**
 * Owns the inline header/footer editing mode: which slot is being edited and
 * which page it was opened from (both carried by `partEditTarget`, the one
 * value naming whatever non-body part is open), the resolved header/footer
 * content for the current section, plus the double-click → edit, save,
 * remove, and "click out" workflows.
 *
 * References and `titlePg` are read from the resolved last section, which
 * inherits them from earlier sections; the authored body sectPr may carry
 * neither. Edits are written to both views.
 *
 * Empty headers/footers are materialised on first double-click so the
 * user can start typing — the helper writes the new HeaderFooter into
 * `package.headers` / `package.footers` and registers the relationship
 * so the serializer picks it up (#274).
 */
export function useHeaderFooterEditing({
  document,
  pushDocument,
  partEditTarget,
  setPartEditTarget,
}: {
  document: Document | null;
  pushDocument: (doc: Document) => void;
  // State + setter live in the parent so selection and canvas overlays can
  // route to the open part's story.
  partEditTarget: PartEditTarget | null;
  setPartEditTarget: React.Dispatch<React.SetStateAction<PartEditTarget | null>>;
}) {
  const finalSectionProperties = useMemo(
    () => resolvedFinalSectionProperties(document?.package.document),
    [document]
  );

  const { headerContent, footerContent, firstPageHeaderContent, firstPageFooterContent } =
    useMemo(() => {
      const { header, footer, firstHeader, firstFooter } = resolveHeaderFooter(
        document ?? null,
        finalSectionProperties
      );
      return {
        headerContent: header,
        footerContent: footer,
        firstPageHeaderContent: firstHeader,
        firstPageFooterContent: firstFooter,
      };
    }, [document, finalSectionProperties]);

  const handleHeaderFooterDoubleClick = useCallback(
    (position: 'header' | 'footer', pageNumber?: number) => {
      // No scroll-to-page-1 — the HF content is shared across all pages by
      // `r:id`, so the painter renders the same edits on every page in real
      // time. Whichever page the user double-clicked, the chrome bar floats
      // over THAT page's header and edits propagate visually to all others.
      const isFirstPage = finalSectionProperties?.titlePg === true && (pageNumber ?? 1) === 1;
      const opened: PartEditTarget = {
        kind: position,
        isFirstPage,
        pageIndex: Math.max(0, (pageNumber ?? 1) - 1),
      };
      const hf = isFirstPage
        ? position === 'header'
          ? firstPageHeaderContent
          : firstPageFooterContent
        : position === 'header'
          ? headerContent
          : footerContent;
      if (hf) {
        setPartEditTarget(opened);
        return;
      }

      // Materialise an empty header/footer so the user can start typing.
      if (!document?.package || !finalSectionProperties) return;
      const pkg = document.package;

      const hdrFtrType = isFirstPage ? 'first' : 'default';
      const rId = `rId_new_${position}_${hdrFtrType}`;
      const emptyHf: HeaderFooter = {
        type: position === 'header' ? 'header' : 'footer',
        hdrFtrType,
        content: [{ type: 'paragraph', content: [] }],
      };

      const mapKey = position === 'header' ? 'headers' : 'footers';
      const newMap = new Map(pkg[mapKey] ?? []);
      newMap.set(rId, emptyHf);

      const refKey = position === 'header' ? 'headerReferences' : 'footerReferences';
      const newRef = { type: hdrFtrType as 'default' | 'first', rId };

      // Register the rel so the serializer wires up content types + doc rels (#274).
      const existingRels = pkg.relationships;
      const usedTargets = new Set<string>();
      for (const rel of existingRels?.values() ?? []) {
        if (rel.target) usedTargets.add(rel.target);
      }
      let targetNum = 1;
      while (usedTargets.has(`${position}${targetNum}.xml`)) targetNum++;
      const relType =
        position === 'header'
          ? 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
          : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
      const newRelationships = new Map(existingRels);
      newRelationships.set(rId, {
        id: rId,
        type: relType,
        target: `${position}${targetNum}.xml`,
      });

      const newDoc: Document = {
        ...document,
        package: {
          ...pkg,
          [mapKey]: newMap,
          relationships: newRelationships,
          document: updateFinalSectionProperties(pkg.document, (properties) => ({
            ...properties,
            [refKey]: [...(properties[refKey] ?? []), newRef],
          })),
        },
      };
      pushDocument(newDoc);
      setPartEditTarget(opened);
    },
    [
      headerContent,
      footerContent,
      firstPageHeaderContent,
      firstPageFooterContent,
      finalSectionProperties,
      document,
      pushDocument,
      setPartEditTarget,
    ]
  );

  const handleBodyClick = useCallback(() => {
    setPartEditTarget(null);
  }, [setPartEditTarget]);

  const handleRemoveHeaderFooter = useCallback(() => {
    const band =
      partEditTarget?.kind === 'header' || partEditTarget?.kind === 'footer'
        ? partEditTarget
        : null;
    if (!band || !document?.package) {
      setPartEditTarget(null);
      return;
    }

    const pkg = document.package;
    const refKey = band.kind === 'header' ? 'headerReferences' : 'footerReferences';
    const mapKey = band.kind === 'header' ? 'headers' : 'footers';
    const refs = finalSectionProperties?.[refKey];
    const delTargetType = band.isFirstPage ? 'first' : 'default';
    const removedId = (
      refs?.find((r) => r.type === delTargetType) ??
      refs?.find((r) => r.type === 'default') ??
      refs?.find((r) => r.type === 'first') ??
      refs?.[0]
    )?.rId;

    if (removedId) {
      const newMap = new Map(pkg[mapKey] ?? []);
      newMap.delete(removedId);

      const newDoc: Document = {
        ...document,
        package: {
          ...pkg,
          [mapKey]: newMap,
          document: updateFinalSectionProperties(pkg.document, (properties) => {
            const own = properties[refKey];
            return own
              ? { ...properties, [refKey]: own.filter((r) => r.rId !== removedId) }
              : properties;
          }),
        },
      };
      pushDocument(newDoc);
    }

    setPartEditTarget(null);
  }, [partEditTarget, document, finalSectionProperties, pushDocument, setPartEditTarget]);

  return {
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    finalSectionProperties,
    handleHeaderFooterDoubleClick,
    handleBodyClick,
    handleRemoveHeaderFooter,
  };
}
