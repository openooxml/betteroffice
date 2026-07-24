import { useEffect, useMemo, useState } from 'react';
import { createPortal } from 'react-dom';
import type { ReactNode } from 'react';
import type {
  Document,
  Theme,
  SectionProperties,
  HeaderFooter,
} from '@betteroffice/docx/types/document';
import type { Comment } from '@betteroffice/docx/types/content';
import { type BundledFontProvider } from '@betteroffice/docx/layout';
import { PagedEditor, type PagedEditorRef } from './PagedEditor';
import type { RustFontChainsProvider } from './hooks/useRustMeasurement';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { DisplayList, DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { YrsResidentCaretSnapshot } from '@betteroffice/docx/yrs';
import type { ResidentFrameApplyResult } from './hooks/useDisplayList';
import type { ResolveDisplayListQueries } from './hooks/displayListQueryEpochGate';
import {
  InlineHeaderFooterEditor,
} from '../InlineHeaderFooterEditor';
import { UnifiedSidebar } from '../UnifiedSidebar';
import { CommentMarginMarkers } from '../CommentMarginMarkers';
import { useCanvasOverlayTarget } from './internals/useCanvasOverlayTarget';
import { CanvasCellSelectionOverlay } from './overlays/CanvasCellSelectionOverlay';
import { CanvasHfSelectionOverlay } from './overlays/CanvasHfSelectionOverlay';
import { projectPageLocalRect } from './internals/canvasProjection';
import { Tooltip } from '../ui/Tooltip';
import { MaterialSymbol } from '../ui/Icons';
import type { HyperlinkPopupData } from '../ui/HyperlinkPopup';
import type { WrapType } from '@betteroffice/docx/docx/wrapTypes';
import type { ReactSidebarItem } from '../../plugin-api/types';
import type { RenderedDomContext } from '../../plugin-api/types';
import type { YrsToolbarSelection } from './yrsToolbar';
import type { TrackedChangesResult } from '@betteroffice/docx/layout/render';
import type { DocxEditorCollaborationOptions } from './types';
import type { YrsCoreSession } from './hooks/useYrsCoreSession';

/**
 * Body of the editor: the paged ProseMirror host, its sidebar overlay
 * (UnifiedSidebar + comment margin markers), the floating "Add comment"
 * button anchored to a non-empty selection, and the inline header/footer
 * editor that appears when a user double-clicks an H/F slot.
 *
 * The floating button dispatches a pending comment mark inline rather
 * than going through onAddComment — same shape as the right-click menu's
 * addComment branch.
 */
export function DocxEditorPagedArea({
  // PagedEditor refs + state
  pagedEditorRef,
  scrollContainerRef,
  editorContentRef,
  // Document + section
  document,
  yrsCore,
  collaboration,
  theme,
  initialSectionProperties,
  finalSectionProperties,
  // Header/footer
  headerContent,
  footerContent,
  firstPageHeaderContent,
  firstPageFooterContent,
  hfEditPosition,
  setHfEditPosition,
  hfEditIsFirstPage,
  hfEditPageIndex,
  onHeaderFooterDoubleClick,
  onRemoveHeaderFooter,
  onBodyClick,
  // Editor
  zoom,
  readOnly,
  onYrsContentChange,
  onYrsHistoryChange,
  onPagedSelectionChange,
  onYrsSelectionChange,
  onRenderedDomContextReady,
  pluginOverlays,
  onHyperlinkClick,
  hyperlinkPopupData,
  onHyperlinkPopupNavigate,
  onHyperlinkPopupCopy,
  onHyperlinkPopupEdit,
  onHyperlinkPopupRemove,
  onHyperlinkPopupClose,
  onContextMenu,
  // Sidebar
  sidebarOpen,
  sidebarItems,
  anchorPositions,
  onAnchorPositionsChange,
  onYrsTrackedChangesChange,
  pluginRenderedDomContext,
  pageWidthPx,
  expandedSidebarItem,
  setExpandedSidebarItem,
  comments,
  resolvedCommentIds,
  resolvedIdsForRender,
  setShowCommentsSidebar,
  // Scroll page indicator
  onTotalPagesChange,
  onLayoutComputed,
  applyResidentInput,
  applyResidentDelete,
  displayListQueries,
  resolveDisplayListQueries,
  canvasDisplayList,
  displayListFrameEpoch,
  residentCaret,
  residentCaretAuthoritative,
  paintedCaretActive,
  onCaretInput,
  onCaretInputDispatched,
  onCaretInterrupt,
  canvasHostRef,
  // Floating comment button
  floatingCommentBtn,
  isAddingComment,
  setCommentSelectionRange,
  setAddCommentYPosition,
  setIsAddingComment,
  setFloatingCommentBtn,
  isSuggesting = false,
  author = 'User',
  measurementFontProvider,
  rustFontChainsProviderRef,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  scrollContainerRef: React.RefObject<HTMLDivElement | null>;
  editorContentRef: React.RefObject<HTMLDivElement | null>;
  document: Document | null;
  yrsCore: YrsCoreSession;
  collaboration?: DocxEditorCollaborationOptions;
  theme: Theme | null | undefined;
  initialSectionProperties: SectionProperties | undefined;
  finalSectionProperties: SectionProperties | undefined;
  headerContent: HeaderFooter | null | undefined;
  footerContent: HeaderFooter | null | undefined;
  firstPageHeaderContent: HeaderFooter | null | undefined;
  firstPageFooterContent: HeaderFooter | null | undefined;
  hfEditPosition: 'header' | 'footer' | null;
  setHfEditPosition: React.Dispatch<React.SetStateAction<'header' | 'footer' | null>>;
  hfEditIsFirstPage: boolean;
  hfEditPageIndex: number;
  onHeaderFooterDoubleClick: (position: 'header' | 'footer', pageNumber?: number) => void;
  onRemoveHeaderFooter: () => void;
  onBodyClick: () => void;
  zoom: number;
  readOnly: boolean;
  onYrsContentChange: () => void;
  onYrsHistoryChange: (canUndo: boolean, canRedo: boolean) => void;
  onPagedSelectionChange: () => void;
  onYrsSelectionChange: (selection: YrsToolbarSelection) => void;
  onRenderedDomContextReady: ((ctx: RenderedDomContext) => void) | undefined;
  pluginOverlays: ReactNode;
  onHyperlinkClick: (data: HyperlinkPopupData) => void;
  hyperlinkPopupData: HyperlinkPopupData | null;
  onHyperlinkPopupNavigate: (href: string) => void;
  onHyperlinkPopupCopy: (href: string) => void;
  onHyperlinkPopupEdit: (displayText: string, href: string) => void;
  onHyperlinkPopupRemove: () => void;
  onHyperlinkPopupClose: () => void;
  onContextMenu: (data: {
    x: number;
    y: number;
    hasSelection: boolean;
    image?: {
      pos: number;
      wrapType: WrapType;
      cssFloat?: 'left' | 'right' | 'none' | null;
      inlinePositionEmu?: { horizontalEmu: number; verticalEmu: number };
    } | null;
  }) => void;
  sidebarOpen: boolean;
  sidebarItems: ReactSidebarItem[];
  anchorPositions: Map<string, number>;
  onAnchorPositionsChange: (positions: Map<string, number>) => void;
  onYrsTrackedChangesChange: (result: TrackedChangesResult) => void;
  pluginRenderedDomContext: RenderedDomContext | null | undefined;
  pageWidthPx: number;
  expandedSidebarItem: string | null;
  setExpandedSidebarItem: React.Dispatch<React.SetStateAction<string | null>>;
  comments: Comment[];
  resolvedCommentIds: Set<number>;
  resolvedIdsForRender: Set<number>;
  setShowCommentsSidebar: React.Dispatch<React.SetStateAction<boolean>>;
  onTotalPagesChange: (totalPages: number) => void;
  /** Layout of each pass, surfaced for the experimental canvas renderer. */
  onLayoutComputed?: (layout: Layout | null) => void;
  applyResidentInput?: (text: string) => Promise<ResidentFrameApplyResult | null>;
  applyResidentDelete?: (
    direction: 'backward' | 'forward'
  ) => Promise<ResidentFrameApplyResult | null>;
  /** Display-list query source while the canvas renderer paints (null on the DOM-painter path). */
  displayListQueries?: DisplayListQueries | null;
  resolveDisplayListQueries?: ResolveDisplayListQueries;
  canvasDisplayList?: DisplayList | null;
  displayListFrameEpoch?: number | null;
  residentCaret?: YrsResidentCaretSnapshot | null;
  residentCaretAuthoritative?: boolean;
  /** Worker-painted caret line is on screen; hide the DOM blink caret. */
  paintedCaretActive?: boolean;
  onCaretInput?: () => void;
  onCaretInputDispatched?: () => void;
  onCaretInterrupt?: () => void;
  /** `.canvas-pages` host element — canvas-path pointer events attach here. */
  canvasHostRef?: React.RefObject<HTMLDivElement | null>;
  floatingCommentBtn: { top: number; left: number } | null;
  isAddingComment: boolean;
  setCommentSelectionRange: React.Dispatch<
    React.SetStateAction<{ from: number; to: number } | null>
  >;
  setAddCommentYPosition: React.Dispatch<React.SetStateAction<number | null>>;
  setIsAddingComment: React.Dispatch<React.SetStateAction<boolean>>;
  setFloatingCommentBtn: React.Dispatch<React.SetStateAction<{ top: number; left: number } | null>>;
  isSuggesting?: boolean;
  author?: string;
  /** Bundled metric-compatible font provider for Rust measurement. */
  measurementFontProvider?: BundledFontProvider;
  /**
   * Host slot the Rust measure source fills with the merged doc-wide font
   * chains — forwarded to PagedEditor and read by the canvas display-list
   * build to gate GlyphRun emission.
   */
  rustFontChainsProviderRef?: React.RefObject<RustFontChainsProvider | null>;
}) {
  const sidebarCommentIds = useMemo(() => comments.map((comment) => comment.id), [comments]);

  // Resolve the active HF block for the inline editor — first-page variant
  // wins when `titlePg` is set and the user double-clicked page 1.
  const activeHf = hfEditPosition
    ? hfEditIsFirstPage
      ? hfEditPosition === 'header'
        ? firstPageHeaderContent
        : firstPageFooterContent
      : hfEditPosition === 'header'
        ? headerContent
        : footerContent
    : null;

  // Relationship id of the active HF part, resolved the way `getHfView` and the
  // display-list HF variant builder (`buildDisplayListHeadersFooters`) both do:
  // off the section's header/footer references. The first-page variant wins when
  // the user is editing page 1 of a titlePg section. This keys the region-aware
  // range-rect query so a first-page vs default variant on another page never
  // contributes stray rects.
  const activeHfRid = hfEditPosition
    ? (() => {
        const refs =
          hfEditPosition === 'header'
            ? finalSectionProperties?.headerReferences
            : finalSectionProperties?.footerReferences;
        if (!refs) return null;
        const wantType = hfEditIsFirstPage ? 'first' : 'default';
        const entry =
          refs.find((r) => r.type === wantType) ??
          refs.find((r) => r.type === 'default') ??
          refs.find((r) => r.type === 'first') ??
          null;
        return entry?.rId ?? null;
      })()
    : null;

  // Live HF Yrs selection, captured on every HF selection change.
  const [hfSelection, setHfSelection] = useState<{ from: number; to: number } | null>(null);
  useEffect(() => {
    setHfSelection(null);
  }, [hfEditPosition, activeHfRid]);

  // UI chrome is independent of renderer readiness and always portals onto
  // the positioned editor-content host.
  const canvasOverlayTarget = useCanvasOverlayTarget(true, editorContentRef);

  const activeHfPage = displayListQueries
    ? (displayListQueries.displayList.pages[hfEditPageIndex] ??
      displayListQueries.displayList.pages.find((page) => {
        const band = hfEditPosition ? page[hfEditPosition] : null;
        return band?.rId === activeHfRid;
      }))
    : null;
  const activeHfBand = activeHfPage && hfEditPosition ? activeHfPage[hfEditPosition] : null;
  const hfChromeRect =
    activeHfPage &&
    activeHfBand &&
    displayListQueries &&
    canvasOverlayTarget &&
    canvasHostRef?.current
      ? projectPageLocalRect(
          canvasHostRef.current,
          canvasOverlayTarget,
          displayListQueries,
          activeHfPage.pageIndex,
          0,
          activeHfBand.y,
          activeHfPage.width,
          activeHfBand.height
        )
      : null;

  const sidebarOverlayNode = (
    <>
      {sidebarItems.length > 0 && (
        <UnifiedSidebar
          items={sidebarItems}
          anchorPositions={anchorPositions}
          renderedDomContext={pluginRenderedDomContext ?? null}
          pageWidth={pageWidthPx}
          zoom={zoom}
          editorContainerRef={scrollContainerRef}
          onExpandedItemChange={setExpandedSidebarItem}
          activeItemId={expandedSidebarItem}
        />
      )}
      <CommentMarginMarkers
        comments={comments}
        anchorPositions={anchorPositions}
        zoom={zoom}
        pageWidth={pageWidthPx}
        sidebarOpen={sidebarOpen}
        resolvedCommentIds={resolvedCommentIds}
        onMarkerClick={() => setShowCommentsSidebar(true)}
      />
    </>
  );

  const floatingCommentButton =
    floatingCommentBtn != null && !isAddingComment && !readOnly ? (
      <Tooltip content="Add comment" side="bottom" delayMs={300}>
        <button
          type="button"
          onMouseDown={(e) => {
            e.preventDefault();
            e.stopPropagation();
            const selection = pagedEditorRef.current?.getSelectionRange();
            if (selection && selection.from !== selection.to) {
              setCommentSelectionRange(selection);
              pagedEditorRef.current?.setSelection(selection.to);
            }
            setAddCommentYPosition(floatingCommentBtn.top);
            setShowCommentsSidebar(true);
            setIsAddingComment(true);
            setFloatingCommentBtn(null);
          }}
          style={{
            position: 'absolute',
            top: floatingCommentBtn.top,
            left: floatingCommentBtn.left,
            transform: 'translate(-50%, -50%)',
            zIndex: 50,
            width: 28,
            height: 28,
            borderRadius: 6,
            border: '1px solid var(--doc-focus-ring)',
            backgroundColor: 'var(--doc-surface)',
            color: 'var(--doc-primary)',
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            boxShadow: '0 1px 3px var(--doc-shadow)',
            transition: 'background-color 0.15s ease, box-shadow 0.15s ease',
          }}
          onMouseOver={(e) => {
            (e.currentTarget as HTMLButtonElement).style.backgroundColor =
              'var(--doc-primary-light)';
            (e.currentTarget as HTMLButtonElement).style.boxShadow =
              '0 1px 4px var(--doc-focus-ring)';
          }}
          onMouseOut={(e) => {
            (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'var(--doc-surface)';
            (e.currentTarget as HTMLButtonElement).style.boxShadow = '0 1px 3px var(--doc-shadow)';
          }}
        >
          <MaterialSymbol name="add_comment" size={16} />
        </button>
      </Tooltip>
    ) : null;

  return (
    <>
      <PagedEditor
        ref={pagedEditorRef}
        document={document}
        yrsCore={yrsCore}
        collaboration={collaboration}
        styles={document?.package.styles}
        theme={document?.package.theme || theme}
        sectionProperties={initialSectionProperties}
        finalSectionProperties={finalSectionProperties}
        headerContent={headerContent}
        footerContent={footerContent}
        firstPageHeaderContent={firstPageHeaderContent}
        firstPageFooterContent={firstPageFooterContent}
        onHeaderFooterDoubleClick={onHeaderFooterDoubleClick}
        hfEditMode={hfEditPosition}
        hfEditRId={activeHfRid}
        onBodyClick={onBodyClick}
        isSuggesting={isSuggesting}
        author={author}
        measurementFontProvider={measurementFontProvider}
        rustFontChainsProviderRef={rustFontChainsProviderRef}
        // Click routing through `onHfPagesMouseDown` was retired; usePagesPointer
        // now routes every HF gesture (click, drag, dblclick, image, hyperlink,
        // context menu) through the active-surface helper directly.
        zoom={zoom}
        readOnly={readOnly}
        onYrsContentChange={onYrsContentChange}
        onYrsHistoryChange={onYrsHistoryChange}
        onSelectionChange={onPagedSelectionChange}
        onYrsSelectionChange={onYrsSelectionChange}
        onYrsHfSelectionChange={(rId, selection) => {
          if (rId === activeHfRid) setHfSelection(selection);
        }}
        onRenderedDomContextReady={onRenderedDomContextReady}
        pluginOverlays={pluginOverlays}
        onHyperlinkClick={onHyperlinkClick}
        hyperlinkPopupData={hyperlinkPopupData}
        onHyperlinkPopupNavigate={onHyperlinkPopupNavigate}
        onHyperlinkPopupCopy={onHyperlinkPopupCopy}
        onHyperlinkPopupEdit={onHyperlinkPopupEdit}
        onHyperlinkPopupRemove={onHyperlinkPopupRemove}
        onHyperlinkPopupClose={onHyperlinkPopupClose}
        onContextMenu={onContextMenu}
        commentsSidebarOpen={sidebarOpen}
        onAnchorPositionsChange={onAnchorPositionsChange}
        sidebarCommentIds={sidebarCommentIds}
        onYrsTrackedChangesChange={onYrsTrackedChangesChange}
        onTotalPagesChange={onTotalPagesChange}
        onLayoutComputed={onLayoutComputed}
        applyResidentInput={applyResidentInput}
        applyResidentDelete={applyResidentDelete}
        displayListQueries={displayListQueries}
        resolveDisplayListQueries={resolveDisplayListQueries}
        canvasDisplayList={canvasDisplayList}
        displayListFrameEpoch={displayListFrameEpoch}
        residentCaret={residentCaret}
        residentCaretAuthoritative={residentCaretAuthoritative}
        paintedCaretActive={paintedCaretActive}
        onCaretInput={onCaretInput}
        onCaretInputDispatched={onCaretInputDispatched}
        onCaretInterrupt={onCaretInterrupt}
        canvasHostRef={canvasHostRef}
        canvasOverlayTarget={canvasOverlayTarget}
        resolvedCommentIds={resolvedIdsForRender}
        scrollContainerRef={scrollContainerRef}
        sidebarOverlay={canvasOverlayTarget ? null : sidebarOverlayNode}
      />

      {/* Canvas mode: the sidebar overlay + floating button portal onto the
          visible canvas pages (their offset origin is `editorContentRef`, which
          shares the `.canvas-pages` host's top-left). The wrapper mirrors the
          one PagedEditor renders in default mode. */}
      {canvasOverlayTarget &&
        createPortal(
          <div
            style={{
              position: 'absolute',
              top: 0,
              left: 0,
              right: 0,
              height: '100%',
              pointerEvents: 'none',
              overflow: 'visible',
            }}
          >
            <div
              style={{ pointerEvents: 'auto' }}
              // In default mode the sidebar lives inside `.paged-editor`, whose
              // mousedown handling keeps the hidden PM focused when its chrome is
              // clicked. Portalled out for canvas mode it loses that, so a bare
              // click on a card would blur the PM and move the caret. Re-assert
              // it here in the capture phase (before the card's own
              // stopPropagation): preventDefault keeps the PM's selection for
              // non-interactive clicks, while inputs/buttons still focus.
              onMouseDownCapture={(e) => {
                const target = e.target as HTMLElement;
                if (!target.closest('input, textarea, select, button, a, [contenteditable]')) {
                  e.preventDefault();
                }
              }}
            >
              {sidebarOverlayNode}
              {floatingCommentButton}
            </div>
          </div>,
          canvasOverlayTarget
        )}

      {!canvasOverlayTarget && floatingCommentButton}

      {canvasOverlayTarget &&
        displayListQueries &&
        hfEditPosition &&
        activeHfRid &&
        canvasHostRef && (
          <>
            <CanvasCellSelectionOverlay
              session={pagedEditorRef.current?.getYrsSession() ?? null}
              positionProjection={null}
              overlayTarget={canvasOverlayTarget}
              canvasHostRef={canvasHostRef}
              displayListQueries={displayListQueries}
              region={{ kind: hfEditPosition, rId: activeHfRid }}
              sidebarOpen={sidebarOpen}
              zoom={zoom}
            />
            <CanvasHfSelectionOverlay
              region={hfEditPosition}
              rId={activeHfRid}
              selection={hfSelection}
              overlayTarget={canvasOverlayTarget}
              canvasHostRef={canvasHostRef}
              displayListQueries={displayListQueries}
              activePageIndex={activeHfPage?.pageIndex}
              sidebarOpen={sidebarOpen}
              zoom={zoom}
            />
          </>
        )}

      {hfEditPosition &&
        activeHf &&
        hfChromeRect &&
        (() => {
          const editor = (
            <InlineHeaderFooterEditor
              position={hfEditPosition}
              targetRect={hfChromeRect}
              onClose={() => {
                setHfEditPosition(null);
              }}
              onRemove={onRemoveHeaderFooter}
            />
          );
          return canvasOverlayTarget ? createPortal(editor, canvasOverlayTarget) : editor;
        })()}
    </>
  );
}
