import type { ZipContainerReader } from './zipContainer';

/**
 * Legacy raw-package shape retained for the public rezip compatibility
 * signature. Live parsing no longer constructs it; Rust S9 owns extraction.
 */
export interface RawDocxContent {
  documentXml: string | null;
  stylesXml: string | null;
  themeXml: string | null;
  numberingXml: string | null;
  fontTableXml: string | null;
  settingsXml: string | null;
  webSettingsXml: string | null;
  headers: Map<string, string>;
  footers: Map<string, string>;
  footnotesXml: string | null;
  endnotesXml: string | null;
  commentsXml: string | null;
  commentsExtensibleXml: string | null;
  commentsExtendedXml: string | null;
  documentRels: string | null;
  packageRels: string | null;
  contentTypesXml: string | null;
  corePropsXml: string | null;
  appPropsXml: string | null;
  customPropsXml: string | null;
  media: Map<string, ArrayBuffer>;
  fonts: Map<string, ArrayBuffer>;
  allXml: Map<string, string>;
  container: ZipContainerReader;
  originalBuffer: ArrayBuffer;
}
