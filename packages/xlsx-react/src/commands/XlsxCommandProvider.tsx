import { createContext } from 'react';
import type { ReactNode } from 'react';
import { UNAVAILABLE_XLSX_COMMANDS } from './createXlsxCommandStore';
import type { XlsxCommandStore } from './types';

export const XlsxCommandContext = createContext<XlsxCommandStore | null>(null);

export interface XlsxCommandProviderProps {
  /** An editor's `api.commands`; `null` until the editor is ready. */
  commands: XlsxCommandStore | null;
  children?: ReactNode;
}

/** Provides one editor's commands to toolbar parts rendered anywhere in the tree. */
export function XlsxCommandProvider({ commands, children }: XlsxCommandProviderProps) {
  return (
    <XlsxCommandContext.Provider value={commands ?? UNAVAILABLE_XLSX_COMMANDS}>
      {children}
    </XlsxCommandContext.Provider>
  );
}
