import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { openDiagram, type DiagramHandle } from '../index';
import { initWasm } from '../wasm/loader';
import { VsdxDocument } from '../wasm/generated/vsdx_wasm.js';
import { CollaborationProvider } from './provider';
import type {
  CollaborationTransport,
  CollaborationTransportEvent,
} from './types';

const root = resolve(import.meta.dir, '../../../..');
let fixture: Uint8Array;
let left: DiagramHandle;
let right: DiagramHandle;

beforeAll(async () => {
  const [wasm, vsdx] = await Promise.all([
    readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/foundation.vsdx')),
  ]);
  await initWasm(wasm);
  fixture = vsdx;
});

afterAll(() => {
  left?.dispose();
  right?.dispose();
});

describe('VSDX collaboration replica', () => {
  test('two seeded providers converge a real WASM cell edit through the protocol', async () => {
    const source = openDiagram(fixture, { clientId: 4100 });
    const seed = source.encodeStateAsUpdate();
    source.dispose();
    left = openDiagram(fixture, { clientId: 4101, initialUpdate: seed });
    right = openDiagram(fixture, { clientId: 4102, initialUpdate: seed });
    expect([...left.encodeStateAsUpdate()]).toEqual([...right.encodeStateAsUpdate()]);
    const hub = new LoopbackHub();
    const leftTransport = hub.createTransport();
    const rightTransport = hub.createTransport();
    const leftProvider = new CollaborationProvider(left, leftTransport);
    const rightProvider = new CollaborationProvider(right, rightTransport);
    const leftOrigins: string[] = [];
    const rightOrigins: string[] = [];
    const stopLeft = left.onUpdate((_update, origin) => leftOrigins.push(origin));
    const stopRight = right.onUpdate((_update, origin) => rightOrigins.push(origin));

    leftProvider.connect();
    rightProvider.connect();
    await hub.open();
    expect(leftProvider.synced).toBe(true);
    expect(rightProvider.synced).toBe(true);

    const baseline = left.snapshot();
    const pageId = baseline.pages[0].id;
    const shapeId = baseline.pages[0].shapes[0].id;
    hub.pause();
    left.setCellFormula(pageId, shapeId, { cellName: 'Both' }, '3');
    hub.resume();

    expect([...left.encodeStateVector()]).toEqual([...right.encodeStateVector()]);
    expect(left.snapshot()).toEqual(right.snapshot());
    expect(left.snapshot().pages[0].shapes[0].cells.find((cell) => cell.name === 'Both')?.formula).toBe('3');
    expect(leftOrigins).toContain('local');
    expect(rightOrigins).toContain('remote');

    stopLeft();
    stopRight();
    leftProvider.destroy();
    rightProvider.destroy();
  });

  test('refuses a guarded cell edit after the guard arrives through collaboration', async () => {
    const source = openDiagram(fixture, { clientId: 4200 });
    const seed = source.encodeStateAsUpdate();
    source.dispose();
    const guarded = openDiagram(fixture, { clientId: 4201, initialUpdate: seed });
    const peer = openDiagram(fixture, { clientId: 4202, initialUpdate: seed });
    const hub = new LoopbackHub();
    const guardedProvider = new CollaborationProvider(guarded, hub.createTransport());
    const peerProvider = new CollaborationProvider(peer, hub.createTransport());
    guardedProvider.connect();
    peerProvider.connect();
    await hub.open();
    const pageId = guarded.snapshot().pages[0].id;
    const shapeId = guarded.snapshot().pages[0].shapes[0].id;
    guarded.setCellFormula(pageId, shapeId, { cellName: 'Both' }, 'GUARD(1)');
    expect(() => peer.setCellFormula(pageId, shapeId, { cellName: 'Both' }, '2')).toThrow('GUARD protects the requested cell');
    guardedProvider.destroy();
    peerProvider.destroy();
    guarded.dispose();
    peer.dispose();
  });

  test('resyncs a peer after the WASM reports dropped observations', async () => {
    const source = openDiagram(fixture, { clientId: 4301 });
    const peer = openDiagram(fixture, { clientId: 4302 });
    const hub = new LoopbackHub();
    const sourceProvider = new CollaborationProvider(source, hub.createTransport());
    const peerProvider = new CollaborationProvider(peer, hub.createTransport());
    sourceProvider.connect();
    peerProvider.connect();
    await hub.open();
    const pageId = source.snapshot().pages[0].id;
    const shapeId = source.snapshot().pages[0].shapes[0].id;
    const drain = VsdxDocument.prototype.drainUpdateEvent;
    VsdxDocument.prototype.drainUpdateEvent = () => new Uint8Array();
    try {
      for (let index = 0; index <= 1024; index++) source.setCellFormula(pageId, shapeId, { cellName: 'Both' }, String(index));
    } finally {
      VsdxDocument.prototype.drainUpdateEvent = drain;
    }
    source.setCellFormula(pageId, shapeId, { cellName: 'Both' }, '1025');
    expect(peer.snapshot().pages[0].shapes[0].cells.find((cell) => cell.name === 'Both')?.formula).toBe('1025');
    sourceProvider.destroy();
    peerProvider.destroy();
    source.dispose();
    peer.dispose();
  });
});

class LoopbackHub {
  private transports: LoopbackTransport[] = [];
  private queued: Array<{ target: LoopbackTransport; data: Uint8Array }> = [];
  private paused = false;

  createTransport(): LoopbackTransport {
    const transport = new LoopbackTransport(this);
    this.transports.push(transport);
    return transport;
  }

  async open(): Promise<void> {
    await Promise.resolve();
    for (const transport of this.transports) transport.open();
    this.flush();
  }

  pause(): void {
    this.paused = true;
  }

  resume(): void {
    this.paused = false;
    this.flush();
  }

  route(source: LoopbackTransport, data: Uint8Array): void {
    for (const target of this.transports) {
      if (target !== source) this.queued.push({ target, data: data.slice() });
    }
    this.flush();
  }

  private flush(): void {
    if (this.paused) return;
    while (this.queued.length > 0) {
      const message = this.queued.shift();
      if (message) message.target.receive(message.data);
    }
  }
}

class LoopbackTransport implements CollaborationTransport {
  private listener: ((event: CollaborationTransportEvent) => void) | null = null;
  private isOpen = false;
  private pending: Uint8Array[] = [];

  constructor(private readonly hub: LoopbackHub) {}

  connect(): void {}

  disconnect(): void {
    this.isOpen = false;
  }

  send(data: Uint8Array): boolean {
    this.hub.route(this, data);
    return true;
  }

  onEvent(listener: (event: CollaborationTransportEvent) => void): () => void {
    this.listener = listener;
    return () => {
      if (this.listener === listener) this.listener = null;
    };
  }

  open(): void {
    this.isOpen = true;
    this.listener?.({ type: 'open' });
    for (const data of this.pending.splice(0)) this.receive(data);
  }

  receive(data: Uint8Array): void {
    if (!this.isOpen) {
      this.pending.push(data);
      return;
    }
    this.listener?.({ type: 'message', data });
  }
}
