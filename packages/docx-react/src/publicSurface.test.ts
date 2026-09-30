import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { RenderedDomContext } from '@betteroffice/docx/plugin-api';
import type { DocxEditorProps, DocxEditorRef } from './components/DocxEditor';
import { DOCX_REF_ASYNC_TWINS } from './components/DocxEditor/hooks/useDocxEditorRefApi';
import type {
  DocxPluginCommandClient,
  DocxPluginContext,
  DocxPluginEditClient,
  DocxPluginGeometry,
  DocxPluginNavigation,
  DocxPluginReadClient,
} from './plugins/types';

type Access = 'value' | 'callback' | 'document-callback-deprecated';
type Access2 = 'async' | 'projection' | 'deprecated';

const PROPS = {
  documentBuffer: 'value',
  experimentalPrewarm: 'value',
  document: 'value',
  onSave: 'callback',
  onSaveRequest: 'callback',
  downloadOnSave: 'value',
  collaboration: 'value',
  experimentalWorkerOpen: 'value',
  mediaTokens: 'value',
  onOpen: 'callback',
  author: 'value',
  onChange: 'document-callback-deprecated',
  onDocumentChange: 'callback',
  onSelectionChange: 'callback',
  onError: 'callback',
  onMemoryPressure: 'callback',
  memoryBudget: 'value',
  onFontsLoaded: 'callback',
  colorMode: 'value',
  theme: 'value',
  showToolbar: 'value',
  showFileOpen: 'value',
  showHelpMenu: 'value',
  showZoomControl: 'value',
  showMarginGuides: 'value',
  marginGuideColor: 'value',
  showRuler: 'value',
  rulerUnit: 'value',
  initialZoom: 'value',
  showHiddenText: 'value',
  readOnly: 'value',
  previewFirstPage: 'value',
  allowHostProposals: 'value',
  showHostProposalsInSidebar: 'value',
  disableFindReplaceShortcuts: 'value',
  toolbarExtra: 'value',
  toolbar: 'value',
  className: 'value',
  style: 'value',
  placeholder: 'value',
  loadingIndicator: 'value',
  showOutline: 'value',
  showOutlineButton: 'value',
  fontFamilies: 'value',
  fonts: 'value',
  watermarkPresets: 'value',
  printOptions: 'value',
  onPrint: 'callback',
  onCopy: 'callback',
  onCut: 'callback',
  onPaste: 'callback',
  mode: 'value',
  onModeChange: 'callback',
  onCommentAdd: 'callback',
  onCommentResolve: 'callback',
  onCommentDelete: 'callback',
  onCommentReply: 'callback',
  comments: 'value',
  onCommentsChange: 'callback',
  commentsSidebarOpen: 'value',
  onCommentsSidebarOpenChange: 'callback',
  onRenderedDomContextReady: 'callback',
  onFirstPagePainted: 'callback',
  pluginOverlays: 'value',
  pluginSidebarItems: 'value',
  pluginRenderedDomContext: 'value',
  renderLogo: 'callback',
  documentName: 'value',
  onDocumentNameChange: 'callback',
  documentNameEditable: 'value',
  renderTitleBarRight: 'callback',
  i18n: 'value',
  measurementFontProvider: 'value',
  plugins: 'value',
  pluginGrants: 'value',
  onPluginError: 'callback',
} as const satisfies Record<keyof DocxEditorProps, Access>;

const PLUGIN_CONTEXT = {
  pluginId: 'projection',
  snapshot: 'projection',
  state: 'projection',
  signal: 'projection',
  lifetimeSignal: 'projection',
  read: 'projection',
  commands: 'projection',
  edits: 'projection',
  geometry: 'projection',
  navigation: 'projection',
  setState: 'projection',
  onCleanup: 'projection',
  run: 'async',
} as const satisfies Record<keyof DocxPluginContext<unknown>, Access2>;

const PLUGIN_GEOMETRY = {
  layout: 'projection',
  dom: 'projection',
  toOverlayRect: 'projection',
  getPositionAtPoint: 'deprecated',
  readPositionAtPoint: 'async',
  getAnchorGeometry: 'projection',
  readAnchorGeometry: 'async',
} as const satisfies Record<keyof DocxPluginGeometry, Access2>;

const PLUGIN_READ = {
  version: 'async',
  readParagraphs: 'async',
  findText: 'async',
  validateEdits: 'async',
} as const satisfies Record<keyof DocxPluginReadClient, Access2>;

const PLUGIN_EDITS = {
  applyEdits: 'async',
} as const satisfies Record<keyof DocxPluginEditClient, Access2>;

const PLUGIN_COMMANDS = {
  getDescriptor: 'projection',
  getState: 'projection',
  subscribe: 'projection',
  execute: 'async',
} as const satisfies Record<keyof DocxPluginCommandClient, Access2>;

const PLUGIN_NAVIGATION = {
  scrollToParagraph: 'async',
} as const satisfies Record<keyof DocxPluginNavigation, Access2>;

const RENDERED_DOM = {
  pagesContainer: 'projection',
  getCoordinatesForPosition: 'projection',
  getPositionAtPoint: 'projection',
  findElementsForRange: 'projection',
  getRectsForRange: 'projection',
  getCoordinatesForHfPosition: 'projection',
  getRectsForHfRange: 'projection',
  getPageBounds: 'projection',
  zoom: 'projection',
  getContainerOffset: 'projection',
} as const satisfies Record<keyof RenderedDomContext, Access2>;

void RENDERED_DOM;

const EDITOR_SOURCE = readFileSync(resolve(import.meta.dir, 'components/DocxEditor.tsx'), 'utf8');
const PLUGIN_SOURCE = readFileSync(resolve(import.meta.dir, 'plugins/types.ts'), 'utf8');
const REF_DEPRECATIONS = [
  'getEditorRef',
  'onContentChange',
] as const satisfies readonly (keyof DocxEditorRef)[];

function memberDocs(source: string, name: string): Map<string, string> {
  const body = source.match(
    new RegExp(`^export interface ${name}\\b[^\\n]*\\{\\r?\\n([\\s\\S]*?)^}`, 'm')
  )?.[1];
  if (body === undefined) throw new Error(`Missing interface ${name}`);
  const docs = new Map<string, string>();
  for (const member of body.matchAll(/^  (?:readonly )?([\w$]+)\??\s*(?=[:(<])/gm)) {
    const before = body.slice(0, member.index);
    const comment = before.slice(before.lastIndexOf('/**'));
    const end = comment.indexOf('*/') + 2;
    docs.set(member[1]!, end > 1 && comment.slice(end).trim() === '' ? comment.slice(0, end) : '');
  }
  return docs;
}

test('document callbacks are deprecated on the public props', () => {
  const docs = memberDocs(EDITOR_SOURCE, 'DocxEditorProps');
  for (const [member, access] of Object.entries(PROPS)) {
    if (access === 'document-callback-deprecated') {
      expect({ member, deprecated: docs.get(member)?.includes('@deprecated') }).toEqual({
        member,
        deprecated: true,
      });
    }
  }
});

test('deprecated plugin members retain their JSDoc annotation', () => {
  const surfaces = {
    DocxPluginContext: PLUGIN_CONTEXT,
    DocxPluginGeometry: PLUGIN_GEOMETRY,
    DocxPluginReadClient: PLUGIN_READ,
    DocxPluginEditClient: PLUGIN_EDITS,
    DocxPluginCommandClient: PLUGIN_COMMANDS,
    DocxPluginNavigation: PLUGIN_NAVIGATION,
  };
  for (const [name, surface] of Object.entries(surfaces)) {
    const docs = memberDocs(PLUGIN_SOURCE, name);
    for (const [member, access] of Object.entries(surface)) {
      if (access === 'deprecated') {
        expect({ name, member, deprecated: docs.get(member)?.includes('@deprecated') }).toEqual({
          name,
          member,
          deprecated: true,
        });
      }
    }
  }
});

test('synchronous ref members are deprecated while their async twins remain public', () => {
  const docs = memberDocs(EDITOR_SOURCE, 'DocxEditorRef');
  for (const [member, twin] of Object.entries(DOCX_REF_ASYNC_TWINS)) {
    expect({ member, deprecated: docs.get(member)?.includes('@deprecated') }).toEqual({
      member,
      deprecated: true,
    });
    expect({ twin, deprecated: docs.get(twin)?.includes('@deprecated') }).toEqual({
      twin,
      deprecated: false,
    });
  }
});

test('every deprecated ref member has an async twin or an explicit exemption', () => {
  const allowed = new Set<string>([...Object.keys(DOCX_REF_ASYNC_TWINS), ...REF_DEPRECATIONS]);
  for (const [member, doc] of memberDocs(EDITOR_SOURCE, 'DocxEditorRef')) {
    if (doc.includes('@deprecated')) {
      expect({ member, classified: allowed.has(member) }).toEqual({ member, classified: true });
    }
  }
});
