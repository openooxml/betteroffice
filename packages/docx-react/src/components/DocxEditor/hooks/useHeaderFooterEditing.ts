import { useCallback, useMemo } from 'react';
import type {
  Document,
  HeaderFooter,
  SectionProperties,
} from '@betteroffice/docx/types/document';
import { resolveHeaderFooter } from '@betteroffice/docx/layout';

import type { PartEditTarget } from '../partEdit';

/**
 * Owns the inline header/footer editing mode: which slot is being edited and
 * which page it was opened from (both carried by `partEditTarget`, the one
 * value naming whatever non-body part is open), the resolved header/footer
 * content for the current section, plus the double-click → edit, save,
 * remove, and "click out" workflows.
 *
 * Empty headers/footers are materialised on first double-click so the
 * user can start typing — the helper writes the new HeaderFooter into
 * `package.headers` / `package.footers` and registers the relationship
 * so the serializer picks it up (#274).
 */
export function useHeaderFooterEditing({
  document,
  pushDocument,
  initialSectionProperties,
  finalSectionProperties,
  partEditTarget,
  setPartEditTarget,
}: {
  document: Document | null;
  pushDocument: (doc: Document) => void;
  initialSectionProperties: SectionProperties | undefined;
  finalSectionProperties: SectionProperties | undefined;
  // State + setter live in the parent so selection and canvas overlays can
  // route to the open part's story.
  partEditTarget: PartEditTarget | null;
  setPartEditTarget: React.Dispatch<React.SetStateAction<PartEditTarget | null>>;
}) {
  const { headerContent, footerContent, firstPageHeaderContent, firstPageFooterContent } =
    useMemo(() => {
      const { header, footer, firstHeader, firstFooter } = resolveHeaderFooter(
        document ?? null,
        finalSectionProperties ?? initialSectionProperties
      );
      return {
        headerContent: header,
        footerContent: footer,
        firstPageHeaderContent: firstHeader,
        firstPageFooterContent: firstFooter,
      };
    }, [document, initialSectionProperties, finalSectionProperties]);

  const handleHeaderFooterDoubleClick = useCallback(
    (position: 'header' | 'footer', pageNumber?: number) => {
      // No scroll-to-page-1 — the HF content is shared across all pages by
      // `r:id`, so the painter renders the same edits on every page in real
      // time. Whichever page the user double-clicked, the chrome bar floats
      // over THAT page's header and edits propagate visually to all others.
      const sectProps = document?.package?.document?.finalSectionProperties;
      const isFirstPage = sectProps?.titlePg === true && (pageNumber ?? 1) === 1;
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
      if (!document?.package) return;
      const pkg = document.package;
      const sectionProps = pkg.document?.finalSectionProperties;
      if (!sectionProps) return;

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
      const existingRefs = sectionProps[refKey] ?? [];
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
          document: pkg.document
            ? {
                ...pkg.document,
                finalSectionProperties: {
                  ...sectionProps,
                  [refKey]: [...existingRefs, newRef],
                },
              }
            : pkg.document,
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
    const sectionProps = pkg.document?.finalSectionProperties;
    const refKey = band.kind === 'header' ? 'headerReferences' : 'footerReferences';
    const mapKey = band.kind === 'header' ? 'headers' : 'footers';
    const refs = sectionProps?.[refKey];
    const delTargetType = band.isFirstPage ? 'first' : 'default';
    const activeRef =
      refs?.find((r) => r.type === delTargetType) ??
      refs?.find((r) => r.type === 'default') ??
      refs?.find((r) => r.type === 'first') ??
      refs?.[0];

    if (activeRef?.rId) {
      const newMap = new Map(pkg[mapKey] ?? []);
      newMap.delete(activeRef.rId);

      const newRefs = (refs ?? []).filter((r) => r.rId !== activeRef.rId);

      const newDoc: Document = {
        ...document,
        package: {
          ...pkg,
          [mapKey]: newMap,
          document: pkg.document
            ? {
                ...pkg.document,
                finalSectionProperties: {
                  ...sectionProps,
                  [refKey]: newRefs,
                },
              }
            : pkg.document,
        },
      };
      pushDocument(newDoc);
    }

    setPartEditTarget(null);
  }, [partEditTarget, document, pushDocument, setPartEditTarget]);

  return {
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    handleHeaderFooterDoubleClick,
    handleBodyClick,
    handleRemoveHeaderFooter,
  };
}
