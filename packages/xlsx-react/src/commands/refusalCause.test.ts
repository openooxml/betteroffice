import { describe, expect, test } from 'bun:test';
import { createT, en } from '@betteroffice/xlsx-i18n';
import type { CommandStatus } from '@betteroffice/xlsx/worker';
import { commandReason } from './evaluate';
import { PeerlessWorkerCommandRefusal } from './peerlessWorkerInputCoordinator';
import { REFUSAL_CAUSE_KEYS, refusalCause } from './refusalCause';

const t = createT(en);
const draft = { sheet: 0, row: 0, col: 0, value: 'x', source: 'cell' } as never;
const insert = { kind: 'applyOps', payload: { args: [[{ type: 'insertRows', sheet: 0, at: 1, count: 1 }]] } } as never;
const bold = { kind: 'applyOps', payload: { toolbar: { id: 'bold', args: null } } } as never;
const editCell = { kind: 'editCell', payload: { args: [0, 0, 0, 'x'] } } as never;

function refused(reason: string, result?: unknown, status: CommandStatus['status'] = 'refused'): CommandStatus {
  return { seq: 1, status, reason, committed: false, ...(result === undefined ? {} : { result }) };
}

function structural(cause: string): CommandStatus {
  return refused('unsupported', {
    failure: { code: 'unsupported', cause, message: `xl/${cause}s/part1.xml references cells this edit would move, and it cannot be rewritten` },
  });
}

describe('refusal banner mapping', () => {
  test('each structural cause names what the edit would have broken, and says nothing changed', () => {
    for (const [cause, text] of [
      ['chart', "This edit can't be applied: a chart uses cells it would move. The workbook is unchanged."],
      ['table', "This edit can't be applied: a table uses cells it would move. The workbook is unchanged."],
      ['pivot', "This edit can't be applied: a pivot table uses cells it would move. The workbook is unchanged."],
      ['other', "This edit can't be applied here. The workbook is unchanged."],
    ] as const) {
      const error = new PeerlessWorkerCommandRefusal(structural(cause), null, null, insert);
      expect(error.code).toBe('edit-unsupported');
      expect(error.cause).toBe(cause);
      expect(t(REFUSAL_CAUSE_KEYS[error.cause!])).toBe(text);
    }
  });

  test('an unsupported refusal without a known cause gets the generic message, never the cell entry one', () => {
    for (const status of [
      refused('unsupported'),
      refused('unsupported', { failure: { cause: 'sparkline' } }),
      refused('unsupported', { refusedUpdate: 'x' }),
    ]) {
      expect(refusalCause(status)).toBe('other');
      const error = new PeerlessWorkerCommandRefusal(status, null, null, insert);
      expect(error.code).toBe('edit-unsupported');
      expect(t(REFUSAL_CAUSE_KEYS[error.cause!])).toBe("This edit can't be applied here. The workbook is unchanged.");
    }
    expect(new PeerlessWorkerCommandRefusal(refused('unsupported'), null, draft, editCell).code).toBe('edit-unsupported');
    expect(new PeerlessWorkerCommandRefusal(refused('unsupported'), null, draft).code).toBe('edit-unsupported');
  });

  test('any other refusal of an edit command without a cell draft is a declined edit', () => {
    for (const input of [bold, insert, { kind: 'undo', payload: { args: [] } } as never]) {
      for (const reason of ['invalid-step', 'limit-exceeded', 'missing-target', 'locked-target']) {
        const error = new PeerlessWorkerCommandRefusal(refused(reason), null, null, input);
        expect([reason, error.code, error.cause]).toEqual([reason, 'edit-unsupported', 'other']);
      }
    }
  });

  test('navigation, read-only toggles and background reads are not reported as declined edits', () => {
    for (const input of [
      { kind: 'select', payload: { selection: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } } } },
      { kind: 'setActiveSheet', payload: { args: [1] } },
      { kind: 'setReadOnly', payload: { readOnly: true } },
      { kind: 'applyOps', payload: { toolbar: { id: 'refreshProposals', args: null } } },
      { kind: 'applyOps', payload: { toolbar: { id: 'paintFormat', args: null } } },
    ] as never[]) {
      for (const reason of ['target-sheet-missing', 'stale-basis', 'unsupported', 'missing-target']) {
        const error = new PeerlessWorkerCommandRefusal(refused(reason), null, null, input);
        expect([(input as { kind: string }).kind, reason, error.code, error.cause])
          .toEqual([(input as { kind: string }).kind, reason, 'input-failed', null]);
      }
    }
    expect(new PeerlessWorkerCommandRefusal(refused('invalid-step'), null, null).code).toBe('input-failed');
  });

  test('cell input that could not be written stays an input failure', () => {
    for (const reason of ['invalid-step', 'locked-target', 'limit-exceeded']) {
      const error = new PeerlessWorkerCommandRefusal(refused(reason), null, draft, editCell);
      expect([reason, error.code, error.cause]).toEqual([reason, 'input-failed', null]);
    }
    const cancelled = new PeerlessWorkerCommandRefusal(refused('dependency-refused', undefined, 'cancelled'), null, null, bold);
    expect([cancelled.code, cancelled.cause]).toEqual(['input-failed', null]);
    expect(commandReason('input-failed', { translate: t }).message).toBe('The cell entry before this could not be saved.');
  });

  test('protection keeps its own message', () => {
    const error = new PeerlessWorkerCommandRefusal(refused('locked-target', {
      ok: false, failure: { code: 'locked-target', message: 'A1 is locked on protected sheet Sheet1' },
    }), null, null);
    expect([error.code, error.cause, error.message]).toEqual(['sheet-protected', null, 'This cell is on a protected sheet.']);
  });

  test('the generic code reads the same in every editor context', () => {
    expect(commandReason('edit-unsupported', { translate: t }).message)
      .toBe("This edit can't be applied here. The workbook is unchanged.");
    expect(commandReason('edit-unsupported', null).message)
      .toBe("This edit can't be applied here. The workbook is unchanged.");
  });
});
