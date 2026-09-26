import { createContext } from 'react';
import type { ReactNode } from 'react';
import { UNAVAILABLE_DOCX_COMMANDS } from './createDocxCommandStore';
import type { DocxCommandStore } from './types';

export const DocxCommandContext = createContext<DocxCommandStore | null>(null);

/** @experimental */
export interface DocxCommandProviderProps {
  /** An editor's `ref.commands`; `null` until the editor is attached. */
  commands: DocxCommandStore | null;
  children?: ReactNode;
}

/**
 * Provides one editor's commands to toolbar parts rendered anywhere in the tree.
 * @experimental
 */
export function DocxCommandProvider({ commands, children }: DocxCommandProviderProps) {
  return (
    <DocxCommandContext.Provider value={commands ?? UNAVAILABLE_DOCX_COMMANDS}>
      {children}
    </DocxCommandContext.Provider>
  );
}
