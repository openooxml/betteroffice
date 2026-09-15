import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');

let session: YrsSession;

beforeAll(async () => {
  preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  session = await createYrsSession({ clientId: 5150 });
  session.createStory('body', 'AXZ');
  session.applyRawOps('body', [{ op: 'format', index: 1, len: 1, attrs: { hidden: true } }]);
});

interface LoweredRun {
  text?: string;
}

function runTexts(env: Record<string, unknown>): (string | undefined)[] {
  const blocks = session.yrsBlocksForStory('body', env) as Array<{ runs?: LoweredRun[] }>;
  return (blocks[0]?.runs ?? []).map((run) => run.text);
}

describe('yrsBlocksForStory showHiddenText', () => {
  test('omits hidden runs by default', () => {
    expect(runTexts({})).toEqual(['A', 'Z']);
  });

  test('reveals hidden runs without touching the document', () => {
    const before = session.encodeState();
    expect(runTexts({ showHiddenText: true })).toEqual(['AXZ']);
    expect(runTexts({})).toEqual(['A', 'Z']);
    expect(session.encodeState()).toEqual(before);
  });
});
