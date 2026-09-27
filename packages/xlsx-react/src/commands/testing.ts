import type { SelectionFormatting } from '@betteroffice/xlsx';
import type { Translations } from '@betteroffice/xlsx-i18n';
import type { XlsxCommandBinding } from './createXlsxCommandStore';
import {
  DEFAULT_FONT_FAMILIES,
  DEFAULT_FONT_SIZES,
  isCellCommand,
  type XlsxCommandEnvironment,
} from './evaluate';
import type { XlsxCommandArgs, XlsxCommandId, XlsxCommandResult } from './types';

/** Uniform formatting of a plain selection, for tests. */
export const PLAIN_FORMATTING: SelectionFormatting = {
  numberFormat: 'automatic',
  fontFamily: 'Calibri',
  fontSize: 11,
  bold: false,
  italic: false,
  strikethrough: false,
  textColor: '#000000',
  fillColor: '#ffffff',
  horizontalAlignment: 'left',
  verticalAlignment: 'bottom',
  textWrapping: 'overflow',
};

/** An editing environment where every command has what it needs, for tests. */
export function testEnvironment(
  overrides: Partial<XlsxCommandEnvironment> = {}
): XlsxCommandEnvironment {
  return {
    status: 'ready',
    readOnly: false,
    collaborative: false,
    selection: {
      sheet: 0,
      top: 0,
      left: 0,
      bottom: 2,
      right: 1,
      merged: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }],
      formatting: PLAIN_FORMATTING,
    },
    canUndo: true,
    canRedo: true,
    pendingInput: false,
    zoom: 1,
    paintFormat: false,
    borderStyle: null,
    borderColor: null,
    pngExport: true,
    proposals: [{ id: 'p1', label: 'Audit agent' }],
    proposalsPanelOpen: false,
    fontFamilies: DEFAULT_FONT_FAMILIES,
    fontSizes: DEFAULT_FONT_SIZES,
    translate: (key) => key,
    ...overrides,
  };
}

export interface TestBindingCall {
  id: XlsxCommandId;
  args: unknown;
  ordered: boolean;
}

/**
 * A binding over a mutable test environment. Ordered commands wait on
 * `admission`, which tests may replace to hold or reject them.
 */
export function testBinding(initial: Partial<XlsxCommandEnvironment> = {}) {
  const calls: TestBindingCall[] = [];
  const state = {
    env: testEnvironment(initial),
    admission: (): Promise<void> => Promise.resolve(),
    result: (): XlsxCommandResult => ({ ok: true, status: 'executed' }),
    focused: 0,
    generation: 1,
    target: 'A1:B3',
    i18n: undefined as Translations | undefined,
  };
  const immediate: ReadonlySet<XlsxCommandId> = new Set<XlsxCommandId>([
    'zoom',
    'searchMenus',
    'proposalsPanel',
  ]);
  const binding: XlsxCommandBinding = {
    environment: (executing) => (executing ? { ...state.env, pendingInput: false } : state.env),
    ordered: (id) => !immediate.has(id),
    admit: (operation) => state.admission().then(operation),
    perform<K extends XlsxCommandId>(id: K, args: XlsxCommandArgs[K]) {
      calls.push({ id, args, ordered: !immediate.has(id) });
      return state.result();
    },
    capture: (id) => ({
      generation: state.generation,
      target: isCellCommand(id) ? state.target : 'document',
    }),
    resume: (origin, id) =>
      origin.generation !== state.generation
        ? 'document-replaced'
        : origin.target !== (isCellCommand(id) ? state.target : 'document')
          ? 'target-changed'
          : null,
    chrome: () => ({ i18n: state.i18n }),
    focusEditor: () => {
      state.focused += 1;
    },
  };
  return {
    binding,
    calls,
    state,
    update(overrides: Partial<XlsxCommandEnvironment>) {
      state.env = { ...state.env, ...overrides };
    },
  };
}
