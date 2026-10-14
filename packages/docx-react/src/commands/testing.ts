import type { YrsSelectionContext } from '@betteroffice/docx/yrs';
import type { DocxChromeContext, DocxCommandBinding } from './createDocxCommandStore';
import type { DocxCommandEnvironment } from './evaluate';
import type { DocxCommandArgs, DocxCommandId, DocxCommandResult } from './types';

/** A selection context with every mark off, for tests. */
export const PLAIN_CONTEXT: YrsSelectionContext = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  superscript: false,
  subscript: false,
  fontFamily: 'Arial',
  fontSize: 22,
  color: null,
  highlight: null,
  paraId: 'p1',
  styleId: 'Normal',
  alignment: 'left',
  paragraphProperties: { indentLeft: 720 },
  hasSelection: true,
  isMultiParagraph: false,
  inTable: false,
  isSingleEmbed: false,
  embedKind: null,
  isImage: false,
  inInsertion: false,
  inDeletion: false,
};

/** An editing environment where every command has what it needs, for tests. */
export function testEnvironment(
  overrides: Partial<DocxCommandEnvironment> = {}
): DocxCommandEnvironment {
  const env: DocxCommandEnvironment = {
    status: 'ready',
    readOnly: false,
    mode: 'editing',
    modeControl: 'internal',
    sidebarOpen: false,
    sidebarControl: 'internal',
    zoom: 1,
    canOpen: true,
    canReportIssue: true,
    bodyStory: true,
    pendingInput: false,
    canUndo: true,
    canRedo: true,
    selection: { context: PLAIN_CONTEXT, fontFamily: 'Arial', fontSize: 11 },
    table: {
      isInTable: true,
      rowCount: 3,
      columnCount: 3,
      hasMultiCellSelection: true,
      canSplitCell: true,
    },
    image: { wrap: 'inline' },
    revisions: [
      { revisionId: 'r1', position: 4 },
      { revisionId: 'r2', position: 12 },
    ],
    currentRevisionId: 'r1',
    styles: [
      { args: { styleId: 'Normal' }, label: 'Normal text' },
      { args: { styleId: 'Heading1' }, label: 'Heading 1', preview: { fontSize: 20, bold: true } },
    ],
    fonts: [
      { args: { family: 'Arial' }, label: 'Arial', preview: { fontFamily: 'Arial', group: 'sans-serif' } },
      { args: { family: 'Georgia' }, label: 'Georgia', preview: { fontFamily: 'Georgia', group: 'serif' } },
    ],
    revisionIds: new Set(),
    translate: (key) => key,
    ...overrides,
  };
  if (!overrides.revisionIds) {
    env.revisionIds = new Set(env.revisions.map((revision) => revision.revisionId));
  }
  return env;
}

export interface TestBindingCall {
  id: DocxCommandId;
  args: unknown;
  ordered: boolean;
}

/**
 * A binding over a mutable test environment. Ordered commands wait on
 * `admission`, which tests may replace to hold or reject them.
 */
export function testBinding(initial: Partial<DocxCommandEnvironment> = {}) {
  const calls: TestBindingCall[] = [];
  const state = {
    env: testEnvironment(initial),
    admission: (): Promise<void> => Promise.resolve(),
    result: (): DocxCommandResult => ({ ok: true, status: 'executed' }),
    focused: 0,
    /** Replace to simulate loading another document. */
    document: {} as object,
    /** Set to simulate the target of a deferred action changing. */
    targetChanged: false,
    chrome: { i18n: undefined, isDark: false, theme: null } as DocxChromeContext,
  };
  const immediate: ReadonlySet<DocxCommandId> = new Set<DocxCommandId>([
    'editingMode',
    'commentsSidebar',
    'open',
    'save',
    'print',
    'find',
    'replace',
    'reportIssue',
    'zoom',
    'insertImage',
    'imageProperties',
    'pageSetup',
    'watermark',
  ]);
  const binding: DocxCommandBinding = {
    environment: (executing) => (executing ? { ...state.env, pendingInput: false } : state.env),
    ordered: (id) => !immediate.has(id),
    admit: (operation) => state.admission().then(operation),
    perform<K extends DocxCommandId>(id: K, args: DocxCommandArgs[K]) {
      calls.push({ id, args, ordered: !immediate.has(id) });
      return state.result();
    },
    capture: () => ({ document: state.document }),
    resume: (origin) =>
      origin.document !== state.document
        ? 'document-replaced'
        : state.targetChanged
          ? 'target-changed'
          : null,
    chrome: () => state.chrome,
    focusEditor: () => {
      state.focused += 1;
    },
  };
  return {
    binding,
    calls,
    state,
    update(overrides: Partial<DocxCommandEnvironment>) {
      state.env = { ...state.env, ...overrides };
    },
  };
}
