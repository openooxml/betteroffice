import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { openPresentation, type PresentationHandle } from '../index';
import { initWasm } from '../wasm/loader';
import { CollaborationProvider } from './provider';
import type {
  CollaborationTransport,
  CollaborationTransportEvent,
} from './types';

const root = resolve(import.meta.dir, '../../../..');
let fixture: Uint8Array;
let left: PresentationHandle;
let right: PresentationHandle;

beforeAll(async () => {
  const [wasm, pptx] = await Promise.all([
    readFile(resolve(root, 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
  ]);
  await initWasm(wasm);
  fixture = pptx;
});

afterAll(() => {
  left?.dispose();
  right?.dispose();
});

describe('PPTX collaboration replica', () => {
  test('two seeded providers converge text and slide edits through the protocol', async () => {
    const source = openPresentation(fixture, { clientId: 4100 });
    const seed = source.encodeStateAsUpdate();
    source.dispose();
    left = openPresentation(fixture, { clientId: 4101, initialUpdate: seed });
    right = openPresentation(fixture, { clientId: 4102, initialUpdate: seed });
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
    const story = baseline.slides[0].shapes.find((shape) => shape.textStories.length > 0)!
      .textStories[0];
    const slideId = baseline.slides[0].id;
    hub.pause();
    left.insertText(story.id, story.length - 1, ' LEFT', { bold: true });
    right.insertSlide(1, baseline.slides[0].layoutPartPath ?? undefined);
    right.addTextBox(slideId, {
      name: 'Remote note',
      rect: { x: 500_000, y: 4_500_000, width: 3_000_000, height: 700_000 },
      text: 'RIGHT',
      style: { italic: true, fontSizePt: 18 },
    });
    hub.resume();

    expect([...left.encodeStateVector()]).toEqual([...right.encodeStateVector()]);
    expect(left.snapshot()).toEqual(right.snapshot());
    expect(left.story(story.id).paragraphs.some((paragraph) =>
      paragraph.runs.some((run) => run.text.includes('LEFT'))
    )).toBe(true);
    expect(right.story(story.id)).toEqual(left.story(story.id));
    expect(left.snapshot().slides).toHaveLength(baseline.slides.length + 1);
    expect(right.snapshot().slides).toHaveLength(baseline.slides.length + 1);
    expect(left.snapshot().slides.find((slide) => slide.id === slideId)?.shapes.some(
      (shape) => shape.name === 'Remote note'
    )).toBe(true);
    expect(leftOrigins).toContain('local');
    expect(leftOrigins).toContain('remote');
    expect(rightOrigins).toContain('local');
    expect(rightOrigins).toContain('remote');

    stopLeft();
    stopRight();
    leftProvider.destroy();
    rightProvider.destroy();
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
