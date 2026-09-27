import type { ReactNode } from 'react';
import { EditorToolbar } from '../EditorToolbar';

/**
 * The default editor chrome: the title bar with its host slots (logo,
 * document name, right-side actions) above the menu bar, and the default
 * formatting rail. Controls read and run the editor's commands.
 */
export function DocxEditorToolbar({
  renderLogo,
  documentName,
  onDocumentNameChange,
  documentNameEditable,
  renderTitleBarRight,
}: {
  renderLogo: (() => ReactNode) | undefined;
  documentName: string | undefined;
  onDocumentNameChange: ((name: string) => void) | undefined;
  documentNameEditable: boolean | undefined;
  renderTitleBarRight: (() => ReactNode) | undefined;
}) {
  return (
    <div className="z-50 flex flex-col gap-0 flex-shrink-0">
      <EditorToolbar>
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
        <EditorToolbar.Toolbar />
      </EditorToolbar>
    </div>
  );
}
