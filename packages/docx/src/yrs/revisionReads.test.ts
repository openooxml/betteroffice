import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm, type EditSession } from '../wasm/edit';
import { EditSession as WasmEditSession } from '../wasm/generated/edit/docx_edit';
import { createYrsSession, type YrsRevisionInfo } from './index';

beforeAll(() => preloadEditWasm(
  new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
));
const AUTHOR = { name: 'Ada', date: '2026-09-29T00:00:00Z' };

function revisionReads() {
  const list = WasmEditSession.prototype.list_revisions;
  let raw: EditSession;
  const spy = spyOn(WasmEditSession.prototype, 'list_revisions');
  spy.mockImplementation(function (this: EditSession) {
    raw = this;
    return list.call(this);
  });
  return { spy, uncached: (): YrsRevisionInfo[] => JSON.parse(list.call(raw)) };
}

test('listRevisions shares a wasm read and returns independently owned results', async () => {
  const live = await createYrsSession({ clientId: 80031 });
  const reads = revisionReads();
  try {
    const { paraId } = live.createStory('body', 'alpha');
    live.insertText({ story: 'body', paraId, offset: 5 }, ' beta', AUTHOR);
    const first = live.listRevisions();
    const second = live.listRevisions();
    expect(first).toHaveLength(1);
    expect(first).toEqual(second);
    first[0]!.preview = 'changed';
    first[0]!.range.start.offset = 99;
    first[0]!.range.end.offset = 99;
    first.length = 0;
    expect(live.listRevisions()).toEqual(second);
    expect(second).toEqual(reads.uncached());
    expect(reads.spy).toHaveBeenCalledTimes(1);
  } finally {
    reads.spy.mockRestore();
    live.destroy();
  }
});

test('listRevisions refreshes after edits, undo, revision decisions, and mirrored updates', async () => {
  const live = await createYrsSession({ clientId: 80032 });
  const peer = await createYrsSession({ clientId: 80033 });
  const reads = revisionReads();
  try {
    const { paraId } = live.createStory('body', 'alpha');
    live.insertText({ story: 'body', paraId, offset: 5 }, ' beta', AUTHOR);
    live.listRevisions();
    const fresh = (change: () => unknown, changesVersion = true): YrsRevisionInfo[] => {
      const version = live.version();
      const count = reads.spy.mock.calls.length;
      change();
      expect(live.version() !== version).toBe(changesVersion);
      const result = live.listRevisions();
      expect(result).toEqual(reads.uncached());
      expect(live.listRevisions()).toEqual(result);
      expect(reads.spy).toHaveBeenCalledTimes(count + 1);
      return result;
    };
    live.addUndoBoundary();
    fresh(() => live.insertText({ story: 'body', paraId, offset: 10 }, ' local', AUTHOR));
    fresh(() => expect(live.undo()).toBe(true));
    peer.loadState(live.encodeState());
    peer.insertText({ story: 'body', paraId, offset: 0 }, 'remote ', { ...AUTHOR, name: 'Bob' });
    const remote = fresh(() => live.applyUpdate(peer.encodeStateAsUpdate(live.encodeStateVector())));
    const accepted = fresh(() => live.acceptChange({ revisionId: remote[0]!.revisionId }));
    expect(accepted.length).toBeLessThan(remote.length);
    expect(fresh(() => live.rejectChange({ revisionId: accepted[0]!.revisionId }))).toEqual([]);
    const mirror = { version: 'worker:1', proposals: { previewVersion: 0, entries: [] } };
    fresh(() => live.mirrorWorkerDocument(mirror));
    peer.loadState(live.encodeState());
    peer.insertText({ story: 'body', paraId, offset: 0 }, 'host ', AUTHOR);
    fresh(() => live.applyHostUpdate(peer.encodeStateAsUpdate(live.encodeStateVector())), false);
    fresh(() => live.mirrorWorkerDocument({ ...mirror, version: 'worker:2' }));
  } finally {
    reads.spy.mockRestore();
    peer.destroy();
    live.destroy();
  }
});
