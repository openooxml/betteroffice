import { createContext, useContext } from 'react';
import type { ToolbarProps } from './Toolbar';

export type EditorToolbarProps = ToolbarProps;

export const EditorToolbarContext = createContext<ToolbarProps | null>(null);

export function useEditorToolbar(): ToolbarProps {
  const context = useContext(EditorToolbarContext);
  if (!context)
    throw new Error('useEditorToolbar must be used within an <EditorToolbar> component');
  return context;
}
