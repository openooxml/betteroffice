import { createContext } from 'react';
import type { ReactNode } from 'react';
import { UNAVAILABLE_PPTX_COMMANDS } from './createPptxCommandStore';
import type { PptxCommandStore } from './types';

export const PptxCommandContext = createContext<PptxCommandStore | null>(null);

/** @experimental */
export interface PptxCommandProviderProps {
  /** An editor's `api.commands`; `null` until the editor is ready. */
  commands: PptxCommandStore | null;
  children?: ReactNode;
}

/**
 * Provides one editor's commands to toolbar parts rendered anywhere in the tree.
 * @experimental
 */
export function PptxCommandProvider({ commands, children }: PptxCommandProviderProps) {
  return (
    <PptxCommandContext.Provider value={commands ?? UNAVAILABLE_PPTX_COMMANDS}>
      {children}
    </PptxCommandContext.Provider>
  );
}
