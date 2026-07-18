import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import {
  CollaborationProvider,
  createDirectCollaborationReplica,
  createWorkerCollaborationReplica,
} from '../collaboration';
import type {
  CollaborationReplica,
  CollaborationTransport,
  CollaborationTransportEvent,
} from '../collaboration';
import { createYrsSession, type YrsSession } from '../yrs';
import { createEditSession, preloadEditWasm } from './edit';

const WASM = resolve(import.meta.dir, './generated/edit/docx_edit_bg.wasm');

class LoopbackTransport implements CollaborationTransport {
  peer: LoopbackTransport | null = null;
  accept = true;
  private readonly listeners = new Set<(event: CollaborationTransportEvent) => void>();

  connect(): void {}

  disconnect(): void {}

  send(data: Uint8Array): boolean {
    if (!this.accept) return false;
    this.peer?.emit({ type: 'message', data: data.slice() });
    return true;
  }

  onEvent(listener: (event: CollaborationTransportEvent) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  open(): void {
    this.emit({ type: 'open' });
  }

  close(): void {
    this.emit({ type: 'close' });
  }

  drain(): void {
    this.emit({ type: 'drain' });
  }

  private emit(event: CollaborationTransportEvent): void {
    for (const listener of [...this.listeners]) listener(event);
  }
}

function loopbackPair(): [LoopbackTransport, LoopbackTransport] {
  const left = new LoopbackTransport();
  const right = new LoopbackTransport();
  left.peer = right;
  right.peer = left;
  return [left, right];
}

function connectPair(
  left: CollaborationReplica,
  right: CollaborationReplica
): {
  leftProvider: CollaborationProvider;
  rightProvider: CollaborationProvider;
  leftTransport: LoopbackTransport;
  rightTransport: LoopbackTransport;
} {
  const [leftTransport, rightTransport] = loopbackPair();
  const leftProvider = new CollaborationProvider(left, leftTransport);
  const rightProvider = new CollaborationProvider(right, rightTransport);
  leftProvider.connect();
  rightProvider.connect();
  leftTransport.open();
  rightTransport.open();
  leftTransport.close();
  leftProvider.connect();
  leftTransport.open();
  expect(leftProvider.synced).toBe(true);
  expect(rightProvider.synced).toBe(true);
  return { leftProvider, rightProvider, leftTransport, rightTransport };
}

async function seededPair(): Promise<[YrsSession, YrsSession]> {
  const left = await createYrsSession({ clientId: 1001 });
  left.loadStories([
    {
      storyId: 'body',
      paragraphs: [{ text: 'alpha beta' }, { text: 'second paragraph' }],
    },
  ]);
  const right = await createYrsSession({ clientId: 1002 });
  right.loadState(left.encodeState());
  return [left, right];
}

function renderedText(session: YrsSession): string {
  return session
    .storyIds()
    .sort()
    .flatMap((story) => session.paragraphs(story).map((paragraph) => paragraph.text))
    .join('\n');
}

describe('docx wasm collaboration', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('exports vectors, diffs, and origin-prefixed polling at the wasm boundary', () => {
    const left = createEditSession(501);
    const right = createEditSession(502);
    try {
      const receipt = JSON.parse(
        left.load_json(
          JSON.stringify([{ storyId: 'body', paragraphs: [{ text: 'before' }] }])
        )
      ) as { body: string[] };
      right.load(left.encode_state());
      expect([...left.encode_state_vector()]).toEqual([...right.encode_state_vector()]);
      expect([...left.encode_diff(right.encode_state_vector())]).toEqual([0, 0]);
      expect(() => left.encode_diff(Uint8Array.of(0xff))).toThrow('invalid yrs state vector');

      left.start_update_event_observation();
      right.start_update_event_observation();
      left.insert_text('body', receipt.body[0], 6, ' local');
      const local = left.drain_update_event();
      expect(local[0]).toBe(0);
      expect(local.byteLength).toBeGreaterThan(1);
      right.apply_update(local.slice(1));
      const remote = right.drain_update_event();
      expect(remote[0]).toBe(1);
      expect([...remote.slice(1)]).toEqual([...local.slice(1)]);
      expect([...left.encode_state_vector()]).toEqual([...right.encode_state_vector()]);
      expect(left.drain_update_event()).toEqual(new Uint8Array());
      left.clear_update_event_observation();
      right.clear_update_event_observation();
    } finally {
      left.free();
      right.free();
    }
  });

  it('reports owned local and remote events and isolates duplicate listeners', async () => {
    const [left, right] = await seededPair();
    const received: Array<{ update: Uint8Array; origin: string }> = [];
    let duplicateCalls = 0;
    const duplicate = () => {
      duplicateCalls += 1;
    };
    const offFirst = left.onUpdate(duplicate);
    const offSecond = left.onUpdate(duplicate);
    left.onUpdate((update, origin) => {
      received.push({ update, origin });
      update.fill(0);
      throw new Error('listener failure');
    });
    const safe: Uint8Array[] = [];
    left.onUpdate((update) => safe.push(update));
    try {
      offFirst();
      const paraId = left.paragraphs('body')[0].paraId;
      left.insertText({ story: 'body', paraId, offset: 0 }, 'local ');
      expect(duplicateCalls).toBe(1);
      expect(received[0].origin).toBe('local');
      expect([...safe[0]].some((byte) => byte !== 0)).toBe(true);

      const remoteParaId = right.paragraphs('body')[1].paraId;
      right.insertText({ story: 'body', paraId: remoteParaId, offset: 0 }, 'remote ');
      left.applyUpdate(right.encodeStateAsUpdate(left.encodeStateVector()));
      expect(received.at(-1)?.origin).toBe('remote');
      offSecond();
    } finally {
      left.destroy();
      right.destroy();
    }
  });

  it('converges text, paragraph structure, and a table edit through providers', async () => {
    const [left, right] = await seededPair();
    const connection = connectPair(left, right);
    try {
      connection.leftTransport.accept = false;
      connection.rightTransport.accept = false;
      const leftFirst = left.paragraphs('body')[0];
      const rightSecond = right.paragraphs('body')[1];
      left.insertText({ story: 'body', paraId: leftFirst.paraId, offset: 5 }, ' LEFT');
      right.insertText({ story: 'body', paraId: rightSecond.paraId, offset: 0 }, 'RIGHT ');
      expect(connection.leftProvider.pendingBytes).toBeGreaterThan(0);
      expect(connection.rightProvider.pendingBytes).toBeGreaterThan(0);
      connection.leftTransport.accept = true;
      connection.rightTransport.accept = true;
      connection.leftTransport.drain();
      connection.rightTransport.drain();

      const splitAt = left.paragraphs('body')[0];
      left.splitParagraph({ story: 'body', paraId: splitAt.paraId, offset: 5 });
      expect(right.paragraphs('body')).toEqual(left.paragraphs('body'));
      right.mergeParagraphs('body', right.paragraphs('body')[0].paraId);
      expect(right.paragraphs('body')).toEqual(left.paragraphs('body'));

      const tableAnchor = left.paragraphs('body')[0];
      const table = left.insertTable(
        { story: 'body', paraId: tableAnchor.paraId, offset: 0 },
        2,
        2
      );
      const cellStory = table.createdStoryIds[0];
      const cellParagraph = right.paragraphs(cellStory)[0];
      right.insertText(
        { story: cellStory, paraId: cellParagraph.paraId, offset: 0 },
        'table edit'
      );

      expect([...left.encodeStateVector()]).toEqual([...right.encodeStateVector()]);
      expect(renderedText(left)).toBe(renderedText(right));
      expect(left.yrsBlocksForStory('body')).toEqual(right.yrsBlocksForStory('body'));
      expect(left.paragraphs(cellStory)[0].text).toBe('table edit');
      expect(right.paragraphs(cellStory)[0].text).toBe('table edit');
    } finally {
      connection.leftProvider.destroy();
      connection.rightProvider.destroy();
      left.destroy();
      right.destroy();
    }
  });

  it('keeps local undo isolated and preserves tracked-change attribution', async () => {
    const [left, right] = await seededPair();
    const connection = connectPair(left, right);
    try {
      left.beginUndoCapture('body');
      const rightPara = right.paragraphs('body')[0];
      right.insertText({ story: 'body', paraId: rightPara.paraId, offset: 0 }, 'REMOTE ');
      const leftPara = left.paragraphs('body')[0];
      left.insertText({ story: 'body', paraId: leftPara.paraId, offset: 0 }, 'LOCAL ');
      expect(left.undo()).toBe(true);
      expect(left.paragraphs('body')[0].text).toContain('REMOTE ');
      expect(left.paragraphs('body')[0].text).not.toContain('LOCAL ');

      const trackedPara = right.paragraphs('body')[1];
      right.insertText(
        { story: 'body', paraId: trackedPara.paraId, offset: 0 },
        'suggested ',
        { name: 'Alice', date: '2026-07-18T20:00:00Z' }
      );
      expect(left.listRevisions()).toEqual(right.listRevisions());
      expect(left.listRevisions()).toEqual([
        expect.objectContaining({
          author: 'Alice',
          date: '2026-07-18T20:00:00Z',
          kind: 'insertion',
        }),
      ]);
      expect([...left.encodeStateVector()]).toEqual([...right.encodeStateVector()]);
      expect(renderedText(left)).toBe(renderedText(right));
    } finally {
      connection.leftProvider.destroy();
      connection.rightProvider.destroy();
      left.destroy();
      right.destroy();
    }
  });

  it('resynchronizes a missed edit through a reconnect state-vector handshake', async () => {
    const [left, right] = await seededPair();
    const connection = connectPair(left, right);
    try {
      connection.leftTransport.close();
      const rightPara = right.paragraphs('body')[0];
      right.insertText({ story: 'body', paraId: rightPara.paraId, offset: 0 }, 'offline ');
      expect(renderedText(left)).not.toBe(renderedText(right));

      connection.leftProvider.connect();
      connection.leftTransport.open();
      expect(connection.leftProvider.synced).toBe(true);
      expect([...left.encodeStateVector()]).toEqual([...right.encodeStateVector()]);
      expect(renderedText(left)).toBe(renderedText(right));
    } finally {
      connection.leftProvider.destroy();
      connection.rightProvider.destroy();
      left.destroy();
      right.destroy();
    }
  });

  it('adapts remote provider updates into the resident worker protocol', async () => {
    const [host, source] = await seededPair();
    const invalidations: Array<{ update: Uint8Array; selection: unknown }> = [];
    const workerReplica = createWorkerCollaborationReplica(host, {
      invalidate(update, selection) {
        invalidations.push({ update: update.slice(), selection });
      },
    });
    try {
      expect(createDirectCollaborationReplica(host)).toBe(host);
      const sourcePara = source.paragraphs('body')[0];
      source.insertText({ story: 'body', paraId: sourcePara.paraId, offset: 0 }, 'worker ');
      const update = source.encodeStateAsUpdate(workerReplica.encodeStateVector());
      workerReplica.applyUpdate(update);

      expect(invalidations).toHaveLength(1);
      expect([...invalidations[0].update]).toEqual([...update]);
      expect(invalidations[0].selection).toBeNull();
      expect([...workerReplica.encodeStateVector()]).toEqual([...source.encodeStateVector()]);
      expect(renderedText(host)).toBe(renderedText(source));
    } finally {
      host.destroy();
      source.destroy();
    }
  });
});
