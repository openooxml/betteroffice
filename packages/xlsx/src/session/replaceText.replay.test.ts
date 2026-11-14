import { expect, test } from 'bun:test';
import { validateWorkbookReplayEnvelope } from './replay';

function envelope(step: unknown) {
  return {
    sequence: 1,
    calculation: { nowSerial: 46000, randSeed: 1 },
    op: {
      method: 'applyEdits',
      args: [{
        expectVersion: 'version',
        steps: [{
          op: 'replaceText',
          target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1:A2' } },
          ...step as Record<string, unknown>,
        }],
      }],
    },
  };
}

test('replay accepts a replace text step', () => {
  expect(() => validateWorkbookReplayEnvelope(envelope({ query: 'alpha', replacement: 'beta' })))
    .not.toThrow();
  expect(() => validateWorkbookReplayEnvelope(envelope({
    query: 'alpha', replacement: 'beta', entireCell: true,
  }))).not.toThrow();
  expect(() => validateWorkbookReplayEnvelope(envelope({
    query: 'alpha', replacement: 'beta', lookIn: 'formulas',
  }))).not.toThrow();
});

test('replay rejects malformed replace text steps before mutation', () => {
  for (const step of [
    { query: '', replacement: 'beta' },
    { query: 'alpha', replacement: 1 },
    { query: 'alpha', replacement: 'beta', entireCell: 'yes' },
    { query: 'alpha', replacement: 'beta', lookIn: 'comments' },
    {
      query: 'alpha',
      replacement: 'beta',
      expect: { cells: [[{ value: { kind: 'text', value: 'alpha' } }]] },
    },
  ]) expect(() => validateWorkbookReplayEnvelope(envelope(step))).toThrow();
});
