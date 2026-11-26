import { createContext, useContext, useMemo } from 'react';
import type { PptxCommandId } from '../commands/types';
import type { ToolbarProps } from './Toolbar';

export type EditorToolbarProps = ToolbarProps;

export const EditorToolbarContext = createContext<ToolbarProps | null>(null);

export function useEditorToolbar(): ToolbarProps {
  const context = useContext(EditorToolbarContext);
  if (!context) {
    throw new Error('useEditorToolbar must be used within an <EditorToolbar> component');
  }
  return context;
}

/** The binding style of the surrounding `EditorToolbar`, inherited by `Toolbar`. */
export const ToolbarModeContext = createContext<'legacy' | 'commands' | null>(null);

/** Set by the editor around its toolbar region, whose locale and shortcuts it already owns. */
export const EditorChromeContext = createContext(false);

/** Built-in commands whose controls the default toolbar leaves out. */
export const HiddenToolbarCommandsContext = createContext<ReadonlySet<PptxCommandId>>(new Set());

export function useHiddenToolbarCommands(ids: readonly PptxCommandId[] | undefined): ReadonlySet<PptxCommandId> {
  const key = ids?.join('\0') ?? '';
  return useMemo(() => new Set(ids), [key]);
}
