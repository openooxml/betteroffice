import { createContext, useContext } from 'react';
import type { ReactNode } from 'react';
import type { SectionProperties } from '@betteroffice/docx/types/document';

/** Configuration of the default chrome, supplied by the editor that renders it. */
export interface EditorChrome {
  showZoomControl: boolean;
  prepareSave?(): void;
  pageLayout?: {
    currentProps?: SectionProperties;
    unit: 'inch' | 'cm';
    readCurrent?(): Promise<void>;
    onApply(properties: Partial<SectionProperties>): void | Promise<void>;
  };
  /** Host controls appended to the default toolbar row. */
  toolbarExtra: ReactNode;
}

export const EditorChromeContext = createContext<EditorChrome | null>(null);

/** The surrounding editor's chrome configuration, or `null` outside an editor. */
export function useEditorChrome(): EditorChrome | null {
  return useContext(EditorChromeContext);
}
