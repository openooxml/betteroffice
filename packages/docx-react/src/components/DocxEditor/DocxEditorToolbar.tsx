import type { ReactNode } from 'react';
import type { Theme, Document } from '@betteroffice/docx/types/document';
import { EditorToolbar } from '../EditorToolbar';
import { ToolbarSeparator, type SelectionFormatting, type FormattingAction } from '../Toolbar';
import type { FontOption } from '../ui/FontPicker';
import type { TableAction } from '../ui/TableToolbar';
import type { TableContextInfo } from './types';
import { CommentsSidebarToggle } from './CommentsSidebarToggle';
import { EditingModeDropdown } from './EditingModeDropdown';
import type { EditorMode } from './internals/editing-modes';

interface ImageContext {
  pos: number;
  wrapType: string;
  displayMode: string;
  cssFloat: string | null;
  transform: string | null;
  alt: string | null;
  borderWidth: number | null;
  borderColor: string | null;
  borderStyle: string | null;
  width: number | null;
  height: number | null;
}

/**
 * Top-of-editor toolbar — the EditorToolbar compound component wired up
 * with the document state (selection formatting, table/image context,
 * undo/redo depth), plus the title bar slots (logo, document name,
 * right-side actions, menu bar) and the trailing toolbar extras (comments
 * sidebar toggle and editing mode dropdown).
 *
 * Undo/redo availability comes from the authoritative Yrs history.
 */
export function DocxEditorToolbar({
  toolbarRefCallback,
  // Doc state
  document,
  theme,
  authoritativeCanUndo,
  authoritativeCanRedo,
  selectionFormatting,
  tableContext,
  imageContext,
  // Editor modes + flags
  readOnly,
  editingMode,
  setEditingMode,
  setShowCommentsSidebar,
  setExpandedSidebarItem,
  showCommentsSidebar,
  // Customisation slots
  renderLogo,
  documentName,
  onDocumentNameChange,
  documentNameEditable,
  renderTitleBarRight,
  toolbarExtra,
  fontFamilies,
  documentFonts,
  zoom,
  showZoomControl,
  // Handlers
  onFormat,
  onUndo,
  onRedo,
  onPrint,
  showFileOpen,
  showHelpMenu,
  onOpen,
  onSave,
  onZoomChange,
  onRefocusEditor,
  onInsertTable,
  onInsertImage,
  onInsertPageBreak,
  onInsertSectionBreakNextPage,
  onInsertSectionBreakContinuous,
  onInsertTOC,
  onImageWrapType,
  onImageTransform,
  onOpenImageProperties,
  onPageSetup,
  onWatermark,
  onTableAction,
}: {
  toolbarRefCallback: (el: HTMLDivElement | null) => void;
  document: Document | null;
  theme: Theme | null | undefined;
  authoritativeCanUndo: boolean;
  authoritativeCanRedo: boolean;
  selectionFormatting: SelectionFormatting;
  tableContext: TableContextInfo | null;
  imageContext: ImageContext | null;
  readOnly: boolean;
  editingMode: EditorMode;
  setEditingMode: (mode: EditorMode) => void;
  setShowCommentsSidebar: React.Dispatch<React.SetStateAction<boolean>>;
  setExpandedSidebarItem: React.Dispatch<React.SetStateAction<string | null>>;
  showCommentsSidebar: boolean;
  renderLogo: (() => ReactNode) | undefined;
  documentName: string | undefined;
  onDocumentNameChange: ((name: string) => void) | undefined;
  documentNameEditable: boolean | undefined;
  renderTitleBarRight: (() => ReactNode) | undefined;
  toolbarExtra: ReactNode;
  fontFamilies: ReadonlyArray<string | FontOption> | undefined;
  documentFonts?: readonly FontOption[];
  zoom: number;
  showZoomControl: boolean;
  onFormat: (action: FormattingAction) => void;
  onUndo: () => void;
  onRedo: () => void;
  onPrint: () => void;
  showFileOpen: boolean;
  showHelpMenu: boolean;
  onOpen: () => void;
  onSave: () => void | Promise<void>;
  onZoomChange: (zoom: number) => void;
  onRefocusEditor: () => void;
  onInsertTable: (rows: number, columns: number) => void;
  onInsertImage: () => void;
  onInsertPageBreak: () => void;
  onInsertSectionBreakNextPage: () => void;
  onInsertSectionBreakContinuous: () => void;
  onInsertTOC: () => void;
  onImageWrapType: (value: string) => void;
  onImageTransform: (action: 'rotateCW' | 'rotateCCW' | 'flipH' | 'flipV') => void;
  onOpenImageProperties: () => void;
  onPageSetup: () => void;
  onWatermark: () => void;
  onTableAction: (action: TableAction) => void;
}) {
  return (
    <div ref={toolbarRefCallback} className="z-50 flex flex-col gap-0 flex-shrink-0">
      <EditorToolbar
        currentFormatting={selectionFormatting}
        onFormat={onFormat}
        onUndo={onUndo}
        onRedo={onRedo}
        canUndo={authoritativeCanUndo}
        canRedo={authoritativeCanRedo}
        disabled={readOnly}
        documentStyles={document?.package.styles?.styles}
        theme={document?.package.theme || theme}
        fontFamilies={fontFamilies}
        documentFonts={documentFonts}
        onPrint={onPrint}
        onOpen={showFileOpen ? onOpen : undefined}
        onSave={onSave}
        showZoomControl={showZoomControl}
        zoom={zoom}
        onZoomChange={onZoomChange}
        onRefocusEditor={onRefocusEditor}
        onInsertTable={onInsertTable}
        showTableInsert={true}
        showHelpMenu={showHelpMenu}
        onInsertImage={onInsertImage}
        onInsertPageBreak={onInsertPageBreak}
        onInsertSectionBreakNextPage={onInsertSectionBreakNextPage}
        onInsertSectionBreakContinuous={onInsertSectionBreakContinuous}
        onInsertTOC={onInsertTOC}
        imageContext={imageContext}
        onImageWrapType={onImageWrapType}
        onImageTransform={onImageTransform}
        onOpenImageProperties={onOpenImageProperties}
        onPageSetup={onPageSetup}
        onWatermark={onWatermark}
        tableContext={tableContext}
        onTableAction={onTableAction}
      >
        <EditorToolbar.TitleBar>
          {renderLogo && <EditorToolbar.Logo>{renderLogo()}</EditorToolbar.Logo>}
          {documentName !== undefined && (
            <EditorToolbar.DocumentName
              value={documentName}
              onChange={onDocumentNameChange}
              editable={documentNameEditable}
            />
          )}
          {renderTitleBarRight && (
            <EditorToolbar.TitleBarRight>{renderTitleBarRight()}</EditorToolbar.TitleBarRight>
          )}
          <EditorToolbar.MenuBar />
        </EditorToolbar.TitleBar>
        <EditorToolbar.Toolbar>
          <ToolbarSeparator />
          <CommentsSidebarToggle
            active={showCommentsSidebar}
            onClick={() => {
              // Reset expansion so reshowing the sidebar lands on the default
              // collapsed state — resolved threads stay as checkmarks, not opened.
              setShowCommentsSidebar((v) => !v);
              setExpandedSidebarItem(null);
            }}
          />
          <ToolbarSeparator />
          <EditingModeDropdown
            mode={editingMode}
            onModeChange={(mode) => {
              setEditingMode(mode);
              if (mode === 'suggesting') setShowCommentsSidebar(true);
            }}
          />
          {toolbarExtra}
        </EditorToolbar.Toolbar>
      </EditorToolbar>
    </div>
  );
}
