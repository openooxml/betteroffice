import { createT, en } from '@betteroffice/pptx-i18n';
import type { LocaleStrings } from '@betteroffice/pptx-i18n';
import type { PptxChromeContext, PptxCommandBinding } from './createPptxCommandStore';
import { DEFAULT_FONT_FAMILIES, DEFAULT_FONT_SIZES, type PptxCommandEnvironment } from './evaluate';
import type { PptxCommandArgs, PptxCommandId, PptxCommandResult } from './types';

/** An editing environment where every command has what it needs, for tests. */
export function testEnvironment(
  overrides: Partial<PptxCommandEnvironment> = {}
): PptxCommandEnvironment {
  return {
    status: 'ready',
    readOnly: false,
    reviewing: false,
    pendingInput: false,
    canUndo: true,
    canRedo: true,
    slide: { id: 'slide:0:256', layoutPartPath: 'ppt/slideLayouts/slideLayout1.xml' },
    layouts: [
      { args: { layoutPartPath: 'ppt/slideLayouts/slideLayout1.xml' }, label: 'Layout 1' },
      { args: { layoutPartPath: 'ppt/slideLayouts/slideLayout2.xml' }, label: 'Layout 2' },
    ],
    text: {
      kind: 'range',
      bold: false,
      italic: false,
      underline: false,
      fontFamily: 'Arial',
      fontSize: 24,
      color: '#111827',
      alignment: 'l',
    },
    shape: {
      formattable: true,
      geometry: 'roundRect',
      fill: '#d9eaf7',
      stroke: '#202124',
      strokeWidth: 1,
      adjustments: { adj: 0.17 },
      order: { index: 1, count: 3 },
    },
    tool: 'select',
    zoom: 'fit',
    fontFamilies: DEFAULT_FONT_FAMILIES,
    fontSizes: DEFAULT_FONT_SIZES,
    proposals: {
      open: false,
      pending: [
        { id: 'p1', label: '1. Agent', stale: false },
        { id: 'p2', label: '2. Agent', stale: true },
      ],
      canvas: {
        available: [{ id: 'p1', label: '1. Agent', stale: false }],
        selectedId: 'p1',
        diff: false,
      },
    },
    translate: createT(en as LocaleStrings, 'en'),
    ...overrides,
  };
}

export interface TestBindingCall {
  id: PptxCommandId;
  args: unknown;
}

/**
 * A binding over a mutable test environment. Ordered commands wait on
 * `admission`, which tests may replace to hold or reject them.
 */
export function testBinding(initial: Partial<PptxCommandEnvironment> = {}) {
  const calls: TestBindingCall[] = [];
  const state = {
    env: testEnvironment(initial),
    admission: (): Promise<void> => Promise.resolve(),
    result: (): PptxCommandResult => ({ ok: true, status: 'executed' }),
    focused: 0,
    generation: 0,
    /** Set to simulate the target of an admitted command changing. */
    targetChanged: false,
    chrome: { i18n: undefined } as PptxChromeContext,
  };
  const immediate: ReadonlySet<PptxCommandId> = new Set<PptxCommandId>([
    'save',
    'zoom',
    'slideshow',
    'insertImage',
    'proposalsPanel',
    'proposalSelect',
    'proposalDiff',
  ]);
  const binding: PptxCommandBinding = {
    environment: (executing) => (executing ? { ...state.env, pendingInput: false } : state.env),
    ordered: (id) => !immediate.has(id),
    admit: (operation) => state.admission().then(operation),
    perform<K extends PptxCommandId>(id: K, args: PptxCommandArgs[K]) {
      calls.push({ id, args });
      return state.result();
    },
    capture: () => ({ generation: state.generation }),
    resume: (origin) =>
      origin.generation !== state.generation
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
    update(overrides: Partial<PptxCommandEnvironment>) {
      state.env = { ...state.env, ...overrides };
    },
  };
}
