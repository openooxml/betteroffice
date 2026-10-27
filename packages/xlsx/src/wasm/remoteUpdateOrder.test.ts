import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { initWasm, openWorkbook } from './loader';
import type { WorkbookHandle } from './loader';

const FIXTURE = resolve(import.meta.dir, '../../test-fixtures/sample.xlsx');
const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');
const CHILD = 'XLSX_REMOTE_UPDATE_ORDER_TEST';

async function isolated(name: string): Promise<boolean> {
  if (process.env[CHILD] === name) return false;
  const child = Bun.spawn([process.execPath, 'test', import.meta.path, '-t', name], {
    env: { ...process.env, [CHILD]: name }, stdout: 'pipe', stderr: 'pipe',
  });
  const deadline = setTimeout(() => child.kill(), 20_000);
  const [status, stdout, stderr] = await Promise.all([
    child.exited, new Response(child.stdout).text(), new Response(child.stderr).text(),
  ]);
  clearTimeout(deadline);
  expect({ status, output: status === 0 ? '' : stdout + stderr }).toEqual({ status: 0, output: '' });
  return true;
}

function collaborative(clientId: number): WorkbookHandle {
  return openWorkbook(new Uint8Array(readFileSync(FIXTURE)), { collaborative: true, clientId });
}

function edit(source: WorkbookHandle, edits: Array<[number, number, string]>): Uint8Array[] {
  const updates: Uint8Array[] = [];
  const unsubscribe = source.onUpdate((update) => updates.push(update));
  try {
    for (const [row, col, input] of edits) source.editCell(0, row, col, input);
  } finally {
    unsubscribe();
  }
  expect(updates).toHaveLength(edits.length);
  return updates;
}

describe('remote update order', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  const overtaking = 'an overwrite that overtakes a cleared edit settles and converges';
  it(overtaking, async () => {
    if (await isolated(overtaking)) return;
    const source = collaborative(921);
    const target = collaborative(922);
    try {
      const updates = edit(source, [[0, 3, 'one'], [19, 7, 'two'], [19, 7, ''], [0, 3, 'three']]);
      target.applyUpdate(updates[0]);
      target.applyUpdate(updates[3]);
      target.applyUpdate(updates[1]);
      target.applyUpdate(updates[2]);
      expect(target.cell(0, 0, 3).input).toBe('three');
      expect(target.cell(0, 19, 7).input).toBe('');
      expect(target.encodeStateVector()).toEqual(source.encodeStateVector());
    } finally {
      source.dispose();
      target.dispose();
    }
  }, 30_000);
});
