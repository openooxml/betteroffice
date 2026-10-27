import { beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { CollaborationProvider } from '../collaboration/provider';
import type { CollaborationTransport } from '../collaboration/types';
import { initWasm, openPresentation, type PresentationHandle } from '../wasm/loader';
import { createPresentationCollaboration } from './collaboration';

let source: Uint8Array;
beforeAll(async () => {
  await initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  const zip = await JSZip.loadAsync(await readFile(new URL('../../../../crates/pptx-render/tests/fixtures/table-basic.pptx', import.meta.url)));
  const path = 'ppt/slides/slide1.xml';
  zip.file(path, (await zip.file(path)!.async('string')).replace(/ (gridSpan|rowSpan|hMerge|vMerge)="[^"]*"/gu, ''));
  source = await zip.generateAsync({ type: 'uint8array' });
});

function table(handle: PresentationHandle): [string, string] {
  const slide = handle.snapshot().slides[0]!;
  const shape = slide.shapes.find((shape) => (shape.graphic as { type?: string } | null)?.type === 'table')!;
  return [slide.id, shape.id];
}

const edits: Array<(handle: PresentationHandle) => unknown> = [
  (handle) => handle.insertTableRow(...table(handle), 0, 'after', ['New']),
  (handle) => handle.insertTableColumn(...table(handle), 0, 'before'),
  (handle) => handle.deleteTableRow(...table(handle), 0),
  (handle) => handle.deleteTableColumn(...table(handle), 0),
];

const idle: CollaborationTransport = {
  connect() {},
  disconnect() {},
  send() { return true; },
  onEvent() { return () => {}; },
};

test('table structure edits work alone and are refused once the deck is collaborative', () => {
  const solo = openPresentation(source);
  const shared = openPresentation(source);
  const pipe = createPresentationCollaboration(shared);
  const fork = shared.fork();
  try {
    const original = solo.snapshot();
    for (const edit of edits) {
      expect(() => edit(solo)).not.toThrow();
      solo.addUndoBoundary();
    }
    const edited = solo.snapshot();
    for (let step = 0; step < 8 && solo.undo().applied; step++);
    expect(solo.snapshot()).toEqual(original);
    for (let step = 0; step < 8 && solo.redo().applied; step++);
    expect(solo.snapshot()).toEqual(edited);
    const before = shared.encodeStateAsUpdate();
    for (const handle of [shared, fork]) {
      for (const edit of edits) expect(() => edit(handle)).toThrow('unavailable in collaborative mode');
    }
    expect(shared.encodeStateAsUpdate()).toEqual(before);
    expect(shared.isCollaborative()).toBe(true);
    expect(fork.isCollaborative()).toBe(true);
  } finally { fork.dispose(); pipe.dispose(); shared.dispose(); solo.dispose(); }
});

test('a provider marks its replica collaborative and drops local table structure undo', () => {
  const handle = openPresentation(source);
  try {
    handle.insertTableRow(...table(handle), 0, 'after');
    expect(handle.canUndo()).toBe(true);
    const provider = new CollaborationProvider(handle, idle);
    try {
      expect(handle.canUndo()).toBe(false);
      expect(() => edits[0]!(handle)).toThrow('unavailable in collaborative mode');
      expect(handle.isCollaborative()).toBe(true);
    } finally { provider.destroy(); }
  } finally { handle.dispose(); }
});
