import { describe, expect, test } from 'bun:test';
import { DOCX_COMMAND_DESCRIPTORS, DOCX_COMMAND_IDS } from './descriptors';
import { evaluateDocxCommand, type DocxCommandEnvironment } from './evaluate';
import { createDocxCommandController } from './createDocxCommandStore';
import { PLAIN_CONTEXT, testBinding, testEnvironment } from './testing';
import type {
  DocxCommandArgs,
  DocxCommandDisabledCode,
  DocxCommandId,
  DocxCommandResult,
  DocxTableAction,
} from './types';

function codeOf<K extends DocxCommandId>(
  id: K,
  env: DocxCommandEnvironment | null,
  args?: DocxCommandArgs[K]
): DocxCommandDisabledCode | null {
  const state = evaluateDocxCommand(id, args, env);
  return state.enabled ? null : state.disabledReason.code;
}

const READING_COMMANDS = DOCX_COMMAND_IDS.filter(
  (id) => !DOCX_COMMAND_DESCRIPTORS[id].mutatesDocument
);
const WRITING_COMMANDS = DOCX_COMMAND_IDS.filter(
  (id) => DOCX_COMMAND_DESCRIPTORS[id].mutatesDocument
);

const TABLE_ACTIONS: readonly DocxTableAction[] = [
  'addRowAbove',
  'addRowBelow',
  'addColumnLeft',
  'addColumnRight',
  'deleteRow',
  'deleteColumn',
  'mergeCells',
  'splitCell',
  'deleteTable',
  'selectTable',
  'selectRow',
  'selectColumn',
  'borderAll',
  'borderOutside',
  'borderInside',
  'borderNone',
  'borderTop',
  'borderBottom',
  'borderLeft',
  'borderRight',
  { type: 'cellFillColor', color: 'FFCC00' },
  { type: 'borderColor', color: '000000' },
  { type: 'borderWidth', size: 8 },
  { type: 'cellBorder', side: 'all', style: 'single', size: 4, color: '000000' },
  { type: 'tableProperties', props: { justification: 'center' } },
  { type: 'openTableProperties' },
  { type: 'applyTableStyle', styleId: 'TableGrid' },
  { type: 'cellVerticalAlign', align: 'center' },
  { type: 'cellMargins', margins: { top: 100 } },
  { type: 'cellTextDirection', direction: 'btLr' },
  { type: 'toggleNoWrap' },
  { type: 'rowHeight', height: 400, rule: 'exact' },
  { type: 'toggleHeaderRow' },
  { type: 'distributeColumns' },
  { type: 'autoFitContents' },
];

const UNSUPPORTED_TABLE_ACTIONS = new Set([
  'cellVerticalAlign',
  'cellMargins',
  'cellTextDirection',
  'toggleNoWrap',
  'rowHeight',
  'toggleHeaderRow',
  'distributeColumns',
  'autoFitContents',
]);

const SUGGESTING_TABLE_ACTIONS = new Set([
  'addRowAbove',
  'addRowBelow',
  'deleteRow',
  'selectTable',
  'selectRow',
  'selectColumn',
]);

const SELECTION_TABLE_ACTIONS = new Set(['selectTable', 'selectRow', 'selectColumn']);

function tableKey(action: DocxTableAction): string {
  return typeof action === 'string' ? action : action.type;
}

describe('command gate', () => {
  test('every command but the table of contents is available while editing', () => {
    const env = testEnvironment();
    const unavailable = DOCX_COMMAND_IDS.filter((id) => codeOf(id, env) !== null);
    expect(unavailable).toEqual(['insertTOC']);
    expect(codeOf('insertTOC', env)).toBe('unsupported-command');
    const toc = evaluateDocxCommand('insertTOC', undefined, env);
    expect(toc.enabled ? null : toc.disabledReason.message).toBe(
      'commands.reasons.tableOfContentsUnsupported'
    );
  });

  test('viewing mode blocks document writes and keeps reading, navigation and output', () => {
    const env = testEnvironment({ mode: 'viewing' });
    for (const id of WRITING_COMMANDS) expect([id, codeOf(id, env)]).toEqual([id, 'viewing-mode']);
    for (const id of READING_COMMANDS) expect([id, codeOf(id, env)]).toEqual([id, null]);
  });

  test('host read-only blocks writes and cannot be escaped through the mode', () => {
    const env = testEnvironment({ readOnly: true });
    for (const id of WRITING_COMMANDS) expect([id, codeOf(id, env)]).toEqual([id, 'read-only']);
    expect(codeOf('editingMode', env)).toBe('read-only');
    expect(codeOf('editingMode', env, { mode: 'editing' })).toBe('read-only');
    expect(codeOf('editingMode', env, { mode: 'suggesting' })).toBe('read-only');
    expect(codeOf('editingMode', env, { mode: 'viewing' })).toBeNull();
    expect(evaluateDocxCommand('editingMode', undefined, env).value).toBe('viewing');
    for (const id of ['save', 'print', 'find', 'zoom', 'reviewNext', 'commentsSidebar'] as const) {
      expect(codeOf(id, env)).toBeNull();
    }
  });

  test('suggesting keeps direct formatting and tracked operations, and refuses untracked structure', () => {
    const env = testEnvironment({ mode: 'suggesting' });
    for (const id of [
      'bold',
      'fontSize',
      'paragraphStyle',
      'alignment',
      'bulletList',
      'insertLink',
      'insertImage',
      'undo',
      'reviewAccept',
    ] as const) {
      expect([id, codeOf(id, env)]).toEqual([id, null]);
    }
    for (const id of [
      'insertTable',
      'insertPageBreak',
      'insertSectionBreakNextPage',
      'insertSectionBreakContinuous',
      'imageWrap',
      'imageTransform',
      'imageProperties',
      'pageSetup',
      'watermark',
    ] as const) {
      expect([id, codeOf(id, env)]).toEqual([id, 'suggesting-unsupported']);
    }
  });

  test('table actions keep their specific reasons in every mode', () => {
    const modes: [string, Partial<DocxCommandEnvironment>][] = [
      ['editing', {}],
      ['suggesting', { mode: 'suggesting' }],
      ['viewing', { mode: 'viewing' }],
      ['read-only', { readOnly: true }],
    ];
    for (const [mode, overrides] of modes) {
      const env = testEnvironment(overrides);
      for (const action of TABLE_ACTIONS) {
        const key = tableKey(action);
        let expected: DocxCommandDisabledCode | null = null;
        if (mode === 'viewing' && !SELECTION_TABLE_ACTIONS.has(key)) expected = 'viewing-mode';
        else if (mode === 'read-only' && !SELECTION_TABLE_ACTIONS.has(key)) expected = 'read-only';
        else if (UNSUPPORTED_TABLE_ACTIONS.has(key)) expected = 'unsupported-command';
        else if (mode === 'suggesting' && !SUGGESTING_TABLE_ACTIONS.has(key)) {
          expected = 'suggesting-unsupported';
        }
        expect([mode, key, codeOf('tableAction', env, action)]).toEqual([mode, key, expected]);
      }
    }
  });

  test('table structure limits are argument specific', () => {
    const env = testEnvironment({
      table: {
        isInTable: true,
        rowCount: 1,
        columnCount: 1,
        hasMultiCellSelection: false,
        canSplitCell: false,
      },
    });
    expect(codeOf('tableAction', env)).toBeNull();
    expect(codeOf('tableAction', env, 'deleteRow')).toBe('last-row');
    expect(codeOf('tableAction', env, 'deleteColumn')).toBe('last-column');
    expect(codeOf('tableAction', env, 'mergeCells')).toBe('multiple-cells-required');
    expect(codeOf('tableAction', env, 'splitCell')).toBe('cannot-split-cell');
    expect(codeOf('tableAction', env, 'addRowBelow')).toBeNull();
    expect(codeOf('tableAction', testEnvironment({ table: null }), 'addRowBelow')).toBe(
      'table-required'
    );
    expect(codeOf('tableAction', testEnvironment({ bodyStory: false }))).toBe('unsupported-story');
  });

  test('document, selection and story gates come first', () => {
    expect(codeOf('bold', null)).toBe('editor-unavailable');
    expect(codeOf('bold', testEnvironment({ status: 'loading' }))).toBe('document-loading');
    expect(codeOf('save', testEnvironment({ status: 'empty' }))).toBe('no-document');
    expect(codeOf('open', testEnvironment({ status: 'empty' }))).toBeNull();
    expect(codeOf('bold', testEnvironment({ selection: null }))).toBe('selection-required');
    expect(codeOf('alignment', testEnvironment({ selection: 'unsupported' }))).toBe(
      'unsupported-selection'
    );
    const part = testEnvironment({ bodyStory: false });
    expect(codeOf('bold', part)).toBeNull();
    expect(codeOf('insertLink', part)).toBe('unsupported-story');
    expect(codeOf('insertTable', part)).toBe('unsupported-story');
    expect(codeOf('imageWrap', part)).toBe('unsupported-story');
    expect(codeOf('imageWrap', testEnvironment({ image: null }))).toBe('image-required');
  });

  test('history availability counts input that has not been applied yet', () => {
    const empty = testEnvironment({ canUndo: false, canRedo: false });
    expect(codeOf('undo', empty)).toBe('nothing-to-undo');
    expect(codeOf('redo', empty)).toBe('nothing-to-redo');
    const typing = testEnvironment({ canUndo: false, canRedo: false, pendingInput: true });
    expect(codeOf('undo', typing)).toBeNull();
    expect(codeOf('redo', typing)).toBeNull();
  });

  test('host-controlled mode and sidebar without setters report why', () => {
    const fixed = testEnvironment({ modeControl: 'fixed', sidebarControl: 'fixed', sidebarOpen: true });
    expect(codeOf('editingMode', fixed)).toBe('controlled-mode');
    expect(codeOf('editingMode', fixed, { mode: 'suggesting' })).toBe('controlled-mode');
    expect(codeOf('editingMode', fixed, { mode: 'editing' })).toBeNull();
    expect(codeOf('commentsSidebar', fixed)).toBe('controlled-sidebar');
    expect(evaluateDocxCommand('commentsSidebar', undefined, fixed).active).toBe(true);
    const host = testEnvironment({ modeControl: 'host', sidebarControl: 'host' });
    expect(codeOf('editingMode', host, { mode: 'suggesting' })).toBeNull();
    expect(codeOf('commentsSidebar', host)).toBeNull();
  });

  test('review resolution needs a live revision and navigation needs any', () => {
    expect(codeOf('reviewAccept', testEnvironment({ currentRevisionId: null }))).toBe(
      'revision-required'
    );
    expect(codeOf('reviewReject', testEnvironment(), { revisionId: 'gone' })).toBe(
      'revision-not-found'
    );
    expect(codeOf('reviewAccept', testEnvironment(), { revisionId: 'r2' })).toBeNull();
    const header = testEnvironment({ revisionIds: new Set(['r1', 'r2', 'header-r3']) });
    expect(codeOf('reviewReject', header, { revisionId: 'header-r3' })).toBeNull();
    expect(codeOf('reviewNext', testEnvironment({ revisions: [] }))).toBe('no-revisions');
    expect(codeOf('reviewPrevious', testEnvironment({ mode: 'viewing' }))).toBeNull();
  });

  test('host feature configuration disables open and issue reporting', () => {
    const env = testEnvironment({ canOpen: false, canReportIssue: false });
    expect(codeOf('open', env)).toBe('host-disabled');
    expect(codeOf('reportIssue', env)).toBe('host-disabled');
  });

  test('outdent needs indentation or a list', () => {
    const flat = testEnvironment({
      selection: {
        context: { ...PLAIN_CONTEXT, paragraphProperties: {} },
        fontFamily: null,
        fontSize: null,
      },
    });
    expect(codeOf('outdent', flat)).toBe('cannot-outdent');
    const listed = testEnvironment({
      selection: {
        context: { ...PLAIN_CONTEXT, paragraphProperties: { numPr: { numId: 1, ilvl: 0 } } },
        fontFamily: null,
        fontSize: null,
      },
    });
    expect(codeOf('outdent', listed)).toBeNull();
  });

  test('arguments are validated before anything runs', () => {
    const env = testEnvironment();
    expect(codeOf('fontSize', env, { points: 0 })).toBe('invalid-arguments');
    expect(codeOf('fontSize', env, { points: 10.5 })).toBeNull();
    expect(codeOf('alignment', env, { value: 'sideways' } as never)).toBe('invalid-arguments');
    expect(codeOf('insertTable', env, { rows: 0, columns: 2 })).toBe('invalid-arguments');
    expect(codeOf('zoom', env, { scale: 9 })).toBe('invalid-arguments');
    expect(codeOf('paragraphStyle', env, { styleId: 'Missing' })).toBe('invalid-arguments');
    expect(codeOf('paragraphStyle', env, { styleId: 'Heading1' })).toBeNull();
    expect(codeOf('textColor', env, { color: { rgb: 'red' } })).toBe('invalid-arguments');
    expect(codeOf('textColor', env, { color: 'auto' })).toBeNull();
    expect(codeOf('tableAction', env, 'explode' as never)).toBe('invalid-arguments');
    expect(codeOf('bold', env, { value: true } as never)).toBe('invalid-arguments');
  });

  test('toggle, value and option state mirror the authoritative selection', () => {
    const env = testEnvironment({
      selection: {
        context: {
          ...PLAIN_CONTEXT,
          bold: 'mixed',
          superscript: true,
          alignment: 'center',
          styleId: 'Heading1',
          highlight: 'yellow',
          paragraphProperties: { numPr: { numId: 1, ilvl: 0 }, lineSpacing: 360 },
        },
        fontFamily: 'Georgia',
        fontSize: 12,
      },
    });
    expect(evaluateDocxCommand('bold', undefined, env).active).toBe('mixed');
    expect(evaluateDocxCommand('italic', undefined, env).active).toBe(false);
    expect(evaluateDocxCommand('superscript', undefined, env).active).toBe(true);
    expect(evaluateDocxCommand('subscript', undefined, env).active).toBe(false);
    expect(evaluateDocxCommand('bulletList', undefined, env).active).toBe(true);
    expect(evaluateDocxCommand('numberedList', undefined, env).active).toBe(false);
    expect(evaluateDocxCommand('fontFamily', undefined, env).value).toBe('Georgia');
    expect(evaluateDocxCommand('fontSize', undefined, env).value).toBe(12);
    expect(evaluateDocxCommand('highlightColor', undefined, env).value).toBe('yellow');
    expect(evaluateDocxCommand('lineSpacing', undefined, env).value).toBe(360);
    expect(evaluateDocxCommand('lineSpacing', { value: 360 }, env).active).toBe(true);
    expect(evaluateDocxCommand('paragraphStyle', undefined, env).value).toBe('Heading1');
    expect(evaluateDocxCommand('paragraphStyle', { styleId: 'Heading1' }, env).active).toBe(true);
    expect(evaluateDocxCommand('alignment', { value: 'center' }, env).active).toBe(true);
    expect(evaluateDocxCommand('alignment', { value: 'right' }, env).active).toBe(false);
    expect(evaluateDocxCommand('editingMode', { mode: 'editing' }, env).active).toBe(true);
  });

  test('states and reasons are plain JSON', () => {
    const envs = [testEnvironment(), testEnvironment({ mode: 'viewing' }), null];
    for (const env of envs) {
      for (const id of DOCX_COMMAND_IDS) {
        const state = evaluateDocxCommand(id, undefined, env);
        expect(JSON.parse(JSON.stringify(state))).toEqual(state);
      }
    }
  });
});

describe('execution re-checks the gate', () => {
  test('a stale enabled snapshot is refused when the editor changed', async () => {
    const harness = testBinding();
    const controller = createDocxCommandController();
    controller.attach(harness.binding);
    expect(controller.store.getState('bold').enabled).toBe(true);
    harness.update({ mode: 'viewing' });
    const result = await controller.store.execute('bold', null);
    expect(result).toEqual({
      ok: false,
      failure: { code: 'viewing-mode', message: 'commands.reasons.viewingMode' },
    });
    expect(harness.calls).toEqual([]);
  });

  test('a mode change while the command waits for input is honored', async () => {
    const harness = testBinding();
    let release!: () => void;
    harness.state.admission = () =>
      new Promise<void>((resolve) => {
        release = resolve;
      });
    const controller = createDocxCommandController();
    controller.attach(harness.binding);
    const pending = controller.store.execute('italic', null);
    harness.update({ readOnly: true });
    release();
    const result = await pending;
    expect(result.ok ? null : result.failure.code).toBe('read-only');
    expect(harness.calls).toEqual([]);
  });

  test('undo refused only after preceding input fails to create history', async () => {
    const harness = testBinding({ canUndo: false, pendingInput: true });
    const controller = createDocxCommandController();
    controller.attach(harness.binding);
    expect(controller.store.getState('undo').enabled).toBe(true);
    const refused = await controller.store.execute('undo', null);
    expect(refused.ok ? null : refused.failure.code).toBe('nothing-to-undo');
    harness.update({ canUndo: true });
    expect(await controller.store.execute('undo', null)).toEqual({ ok: true, status: 'executed' });
  });

  test('deferred actions are gated again and bound to the document and target they opened for', async () => {
    const harness = testBinding();
    const controller = createDocxCommandController();
    controller.attach(harness.binding);
    let writes = 0;
    const write = (): DocxCommandResult => {
      writes += 1;
      return { ok: true, status: 'executed' };
    };
    const code = (result: DocxCommandResult) => (result.ok ? null : result.failure.code);

    const pageSetup = controller.defer('pageSetup', null, 'document');
    harness.update({ mode: 'suggesting' });
    expect(code(await pageSetup.complete(write))).toBe('suggesting-unsupported');
    harness.update({ mode: 'editing' });

    const opened = controller.defer('pageSetup', null, 'document');
    harness.state.document = {};
    expect(code(await opened.complete(write))).toBe('document-replaced');

    const link = controller.defer('insertLink', null, 'selection');
    harness.state.targetChanged = true;
    expect(code(await link.complete(write))).toBe('target-changed');
    harness.state.targetChanged = false;

    let release!: () => void;
    harness.state.admission = () =>
      new Promise<void>((resolve) => {
        release = resolve;
      });
    const pending = link.complete(write);
    await Promise.resolve();
    expect(writes).toBe(0);
    release();
    expect(await pending).toEqual({ ok: true, status: 'executed' });
    expect(writes).toBe(1);
  });
});
