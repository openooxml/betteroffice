/**
 * EditorToolbar — Google Docs-style 2-level compound component, bound to the
 * nearest editor's commands.
 *
 * Usage:
 *   <EditorToolbar>
 *     <EditorToolbar.TitleBar>
 *       <EditorToolbar.Logo><MyIcon /></EditorToolbar.Logo>
 *       <EditorToolbar.DocumentName value={name} onChange={setName} />
 *       <EditorToolbar.MenuBar />
 *       <EditorToolbar.TitleBarRight>
 *         <button>Save</button>
 *       </EditorToolbar.TitleBarRight>
 *     </EditorToolbar.TitleBar>
 *     <EditorToolbar.Toolbar />
 *   </EditorToolbar>
 *
 * Rendered outside `DocxEditor` (under a `DocxCommandProvider`), it supplies
 * its own styling root and locale, and its keyboard shortcuts reach the editor.
 */

import { useLayoutEffect, useRef } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type { Translations } from '@betteroffice/docx-i18n';
import { docxCommandController } from '../commands/createDocxCommandStore';
import { useDocxChrome, useDocxCommands } from '../commands/hooks';
import { LocaleProvider } from '../i18n';
import { useEditorChrome } from './EditorToolbarContext';
import { TitleBar, Logo, DocumentName, MenuBar, TitleBarRight } from './TitleBar';
import type { TitleBarProps, LogoProps, DocumentNameProps, TitleBarRightProps } from './TitleBar';
import { Toolbar, type ToolbarProps } from './Toolbar';
import { ToolbarReviewControls, type ToolbarReviewControlsProps } from './toolbar/ToolbarReviewControls';
import { useIsDark } from './DocxEditor/hooks/useIsDark';
import { cn } from '../lib/utils';
import { Z_INDEX } from '../styles/zIndex';

export interface EditorToolbarProps {
  /** Compound parts; defaults to the formatting rail. */
  children?: ReactNode;
  className?: string;
  style?: CSSProperties;
  /** Locale when rendered outside the editor; defaults to the editor's. */
  i18n?: Translations;
  /** Color mode when rendered outside the editor; defaults to the editor's. */
  colorMode?: 'light' | 'dark' | 'system';
}

interface EditorToolbarComponent {
  (props: EditorToolbarProps): React.JSX.Element;
  TitleBar: typeof TitleBar;
  Logo: typeof Logo;
  DocumentName: typeof DocumentName;
  MenuBar: typeof MenuBar;
  TitleBarRight: typeof TitleBarRight;
  Toolbar: typeof Toolbar;
  Review: typeof ToolbarReviewControls;
}

function ExternalChrome({
  i18n,
  colorMode,
  children,
}: {
  i18n: Translations | undefined;
  colorMode: EditorToolbarProps['colorMode'];
  children: ReactNode;
}) {
  const controller = docxCommandController(useDocxCommands());
  const chrome = useDocxChrome();
  const rootRef = useRef<HTMLDivElement>(null);
  const systemDark = useIsDark(colorMode ?? 'light');
  const isDark = colorMode ? systemDark : (chrome?.isDark ?? false);
  useLayoutEffect(() => {
    const root = rootRef.current;
    if (!root || !controller) return;
    return controller.registerChrome(root);
  }, [controller]);
  return (
    <div ref={rootRef} className={cn('oox-root', isDark && 'dark')}>
      <LocaleProvider i18n={i18n ?? chrome?.i18n}>{children}</LocaleProvider>
    </div>
  );
}

function EditorToolbarBase({ children, className, style, i18n, colorMode }: EditorToolbarProps) {
  const chrome = useEditorChrome();
  const toolbar = (
    <div
      className={cn('flex flex-col bg-doc-surface shadow-sm flex-shrink-0', className)}
      style={{ position: 'relative', zIndex: Z_INDEX.toolbar, ...style }}
      data-testid="editor-toolbar"
    >
      {children ?? <Toolbar />}
    </div>
  );
  return chrome ? (
    toolbar
  ) : (
    <ExternalChrome i18n={i18n} colorMode={colorMode}>
      {toolbar}
    </ExternalChrome>
  );
}

const EditorToolbar = EditorToolbarBase as EditorToolbarComponent;
EditorToolbar.TitleBar = TitleBar;
EditorToolbar.Logo = Logo;
EditorToolbar.DocumentName = DocumentName;
EditorToolbar.MenuBar = MenuBar;
EditorToolbar.TitleBarRight = TitleBarRight;
EditorToolbar.Toolbar = Toolbar;
EditorToolbar.Review = ToolbarReviewControls;

export { EditorToolbar };
export type {
  TitleBarProps,
  LogoProps,
  DocumentNameProps,
  TitleBarRightProps,
  ToolbarProps,
  ToolbarReviewControlsProps,
};
