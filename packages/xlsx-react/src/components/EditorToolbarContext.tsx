import { createContext, useContext } from 'react';
import type { Translations } from '@betteroffice/xlsx-i18n';
import type { ToolbarMode, ToolbarProps } from './Toolbar';

export interface EditorToolbarProps extends ToolbarProps {
  /** Locale of a command-mode toolbar rendered outside the editor; defaults to the editor's. */
  i18n?: Translations;
}

export const EditorToolbarContext = createContext<ToolbarProps | null>(null);

/** The mode an `EditorToolbar` root chose explicitly, inherited by its parts. */
export const ToolbarModeContext = createContext<ToolbarMode | null>(null);

/** Whether chrome renders inside an `XlsxEditor`, which supplies locale and shortcuts. */
export const EditorChromeContext = createContext(false);

export function useEditorToolbar(): ToolbarProps {
  const context = useContext(EditorToolbarContext);
  if (!context)
    throw new Error('useEditorToolbar must be used within an <EditorToolbar> component');
  return context;
}
