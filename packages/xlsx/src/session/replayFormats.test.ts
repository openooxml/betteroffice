import { expect, test } from 'bun:test';
import { validateWorkbookReplayEnvelope } from './replay';
import type { CapturedCellFormat, CapturedFormat } from '../wasm/loader';

const format: CapturedCellFormat = {
  font: {}, fill: { Pattern: { pattern: 'lightUp', fg: { Theme: { idx: 4, tint: 0.25 } }, bg: 'Auto' } },
  border: { top: { style: 'MediumDashed', color: { Indexed: 5 } } },
  numberFormat: { kind: 'builtin', id: 0 }, alignment: {},
};

function envelope(value: unknown) {
  const captured: CapturedFormat = { rows: 1, columns: 1, formats: [value] };
  return { sequence: 1, calculation: { nowSerial: 46000, randSeed: 1 },
    op: { method: 'applyFormat', args: [0, 'A1', captured] } };
}

test('replay accepts preserved pattern fills and extended borders', () => {
  expect(() => validateWorkbookReplayEnvelope(envelope(format))).not.toThrow();
  for (const fill of ['None', { Solid: { Rgb: '#123456' } }, { Pattern: { pattern: 'gray125', fg: null, bg: null } }]) {
    expect(() => validateWorkbookReplayEnvelope(envelope({ ...format, fill }))).not.toThrow();
  }
});

test('replay rejects malformed pattern facets before mutation', () => {
  for (const fill of [
    { Pattern: { pattern: 'unknown', fg: null, bg: null } },
    { Pattern: { pattern: 'lightUp', fg: { Rgb: 42 }, bg: null } },
    { Pattern: { pattern: 'lightUp', fg: null, bg: { Theme: { idx: 4, tint: Infinity } } } },
    { Solid: { Indexed: -1 } },
  ]) expect(() => validateWorkbookReplayEnvelope(envelope({ ...format, fill }))).toThrow();
  expect(() => validateWorkbookReplayEnvelope(envelope({ ...format, border: { top: { style: 'unknown', color: null } } }))).toThrow();
});
