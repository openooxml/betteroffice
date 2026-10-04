import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { SessionFailure, type SessionTransport } from '../../../../shared/office-session';
import { openWorkbook, type WorkbookCalculationContext, type WorkbookHandle } from '../wasm/loader';
import { hydratePeer } from './client';
import { createWorkbookEditPeer, WorkbookEditPeerFailedError } from './editPeer';
import type { WorkbookReplayOp } from './replay';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from './testHelpers';

let fixture: Uint8Array;
let wasmBytes: Uint8Array<ArrayBuffer>;
const calculation: WorkbookCalculationContext = { nowSerial: 46_000.5, randSeed: 123456789 };

beforeAll(async () => {
  ({ fixture, wasmBytes } = await loadWorkbookSessionFixtures());
});

function crashableHost(): {
  wrap(transport: SessionTransport): SessionTransport;
  crash(): void;
} {
  let crash: ((error: unknown) => void) | undefined;
  return {
    wrap: (transport) => ({
      ...transport,
      onError(listener) { crash = listener; return transport.onError(listener); },
    }),
    crash() {
      if (!crash) throw new Error('Missing worker crash callback');
      crash(new SessionFailure('crash', 'Worker stopped'));
    },
  };
}

describe('workbook peer hydration', () => {
  test('keeps hydration opt-in and normal sessions usable without retained state', async () => {
    const session = await createTestWorkbookSession(fixture, undefined, { calculation });
    const peer = openWorkbook(fixture, { calculation });
    try {
      await expect(hydratePeer(session)).rejects.toThrow('does not retain peer hydration state');
      expect(await session.call.version()).toBe(peer.version());
      expect(await session.save()).toEqual(peer.save());
      expect(session.state.stage).toBe('ready');
      expect(session.failure).toBeUndefined();
    } finally {
      peer.dispose();
      await session.dispose();
    }
  });

  test('refuses retained hydration when the worker supplies no compiled module', async () => {
    await expect(createTestWorkbookSession(fixture, undefined, {
      calculation, retainPeerHydration: true,
    })).rejects.toThrow('did not retain a compiled module');
  });

  test('hydrates retained bytes after pre-peer worker failure and saves queued recovery edits once', async () => {
    const host = crashableHost();
    const bytes = fixture.slice();
    const openingCalculation = { ...calculation };
    const session = await createTestWorkbookSession(bytes, host.wrap, {
      calculation: openingCalculation, wasm: wasmBytes.buffer, retainPeerHydration: true,
    });
    const errors: Error[] = [];
    let peer: WorkbookHandle | undefined;
    const original = openWorkbook(fixture, { calculation });
    try {
      bytes.fill(0);
      openingCalculation.nowSerial = 0;
      const waiting = Promise.allSettled([session.call.readCells({ ranges: [] }), session.save()]);
      host.crash();
      for (const result of await waiting) {
        if (result.status !== 'rejected') throw new Error('Normal request survived worker failure');
        expect(result.reason).toBeInstanceOf(SessionFailure);
      }
      expect(session.state.stage).toBe('failed');
      const compiling = spyOn(WebAssembly, 'compile');
      const streaming = spyOn(WebAssembly, 'compileStreaming');
      try {
        const hydration = hydratePeer(session);
        expect(hydratePeer(session)).toBe(hydration);
        peer = await hydration;
        expect(compiling).not.toHaveBeenCalled();
        expect(streaming).not.toHaveBeenCalled();
      } finally {
        compiling.mockRestore();
        streaming.mockRestore();
      }
      if (!peer) throw new Error('Missing hydrated workbook peer');
      expect(peer.save()).toEqual(original.save());
      const edits = createWorkbookEditPeer({
        session, peer, now: () => 0, randomSeed: () => 1, onError: (error) => { errors.push(error); },
      });
      const input: WorkbookReplayOp = { method: 'editCell', args: [0, 2, 1, 'queued before hydration'] };
      const rows: WorkbookReplayOp = {
        method: 'applyOps', args: [[{ type: 'insertRows', sheet: 0, at: 2, count: 1 }]],
      };
      try {
        expect(edits.state).toBe('failed');
        expect(() => edits.editCell(0, 2, 1, 'blocked')).toThrow(WorkbookEditPeerFailedError);
        await expect(edits.flush()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
        await expect(edits.save()).rejects.toBeInstanceOf(WorkbookEditPeerFailedError);
        expect(edits.applyRecoveryOp(input)).toMatchObject({ applied: true });
        expect(edits.applyRecoveryOp(rows)).toMatchObject({ applied: true });
        const version = peer.version();
        const expected = peer.save();
        expect(edits.applyRecoveryOp(input)).toMatchObject({ applied: true });
        expect(edits.applyRecoveryOp(rows)).toMatchObject({ applied: true });
        expect(peer.version()).toBe(version);
        expect(peer.save()).toEqual(expected);
        expect(peer.cell(0, 3, 1).input).toBe('queued before hydration');
        expect(edits.sentSequence).toBe(0);
        expect(edits.acknowledgedSequence).toBe(0);
        const recovery = edits.recoverySave();
        expect(recovery.recovery).toBe(true);
        expect(new Uint8Array(recovery.bytes)).toEqual(expected);
        expect(errors).toEqual([session.failure!]);
        const reopened = openWorkbook(new Uint8Array(recovery.bytes), { calculation });
        try {
          expect(reopened.cell(0, 3, 1).input).toBe('queued before hydration');
        } finally { reopened.dispose(); }
      } finally { edits.dispose(); }
    } finally {
      peer?.dispose();
      original.dispose();
      await session.dispose();
    }
  });

  test('finishes in-flight hydration when the worker fails', async () => {
    const host = crashableHost();
    const session = await createTestWorkbookSession(fixture, host.wrap, {
      calculation, wasm: wasmBytes.buffer, retainPeerHydration: true,
    });
    let peer: WorkbookHandle | undefined;
    try {
      const hydration = hydratePeer(session);
      const waiting = session.call.version();
      host.crash();
      await expect(waiting).rejects.toBeInstanceOf(SessionFailure);
      peer = await hydration;
      expect(await hydratePeer(session)).toBe(peer);
      const errors: Error[] = [];
      const edits = createWorkbookEditPeer({ session, peer, onError: (error) => { errors.push(error); } });
      try {
        expect(new Uint8Array(edits.recoverySave().bytes)).toEqual(peer.save());
        expect(errors).toEqual([session.failure!]);
      } finally { edits.dispose(); }
    } finally {
      peer?.dispose();
      await session.dispose();
    }
  });

  test('rejects hydration after disposal and cancels an in-flight hydration', async () => {
    const session = await createTestWorkbookSession(fixture, undefined, {
      calculation, wasm: wasmBytes.buffer, retainPeerHydration: true,
    });
    const hydration = hydratePeer(session);
    const disposal = session.dispose();
    await expect(hydration).rejects.toMatchObject({ code: 'disposed' });
    await disposal;
    await expect(hydratePeer(session)).rejects.toMatchObject({ code: 'disposed' });
  });
});
