import { beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { openPresentation, type PresentationHandle } from '../index';
import { initWasm } from '../wasm/loader';

const root = resolve(import.meta.dir, '../../../..');
let fixture: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx] = await Promise.all([
    readFile(resolve(root, 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
  ]);
  await initWasm(wasm);
  fixture = pptx;
});

function replicas(count: number, text: string) {
  const source = openPresentation(fixture, { clientId: 4200 });
  const slideId = source.snapshot().slides[0].id;
  const { shapeId } = source.addTextBox(slideId, {
    name: 'Cleanup',
    rect: { x: 500_000, y: 500_000, width: 3_000_000, height: 700_000 },
    text,
    style: {},
  });
  const storyId = source.snapshot().slides[0].shapes.find((shape) => shape.id === shapeId)!
    .textStories[0].id;
  const seed = source.encodeStateAsUpdate();
  source.dispose();
  const handles = Array.from({ length: count }, (_, index) =>
    openPresentation(fixture, { clientId: 4201 + index, initialUpdate: seed })
  );
  return { handles, storyId };
}

function mesh(handles: PresentationHandle[]) {
  const inboxes = handles.map(() => [] as Uint8Array[]);
  const sent: Uint8Array[] = [];
  const stops = handles.map((handle, from) => handle.onUpdate((update, origin) => {
    if (origin !== 'local') return;
    sent.push(update.slice());
    inboxes.forEach((inbox, to) => { if (to !== from) inbox.push(update.slice()); });
  }));
  const settle = (): number => {
    let rounds = 0;
    while (inboxes.some((inbox) => inbox.length > 0)) {
      rounds += 1;
      if (rounds > 8) throw new Error('replicas kept publishing updates');
      handles.forEach((handle, to) => {
        for (const update of inboxes[to].splice(0)) handle.applyUpdate(update);
      });
    }
    return rounds;
  };
  return { settle, sent, stop: () => stops.forEach((stop) => stop()) };
}

function text(handle: PresentationHandle, storyId: string): string {
  return handle.story(storyId).paragraphs[0].runs.map((run) => run.text).join('');
}

function expectConverged(handles: PresentationHandle[], storyId: string) {
  const [first, ...rest] = handles;
  for (const handle of rest) {
    expect([...handle.encodeStateVector()]).toEqual([...first.encodeStateVector()]);
    expect(handle.story(storyId)).toEqual(first.story(storyId));
  }
}

test('seeded concurrent deck text edits converge through incremental updates alone', () => {
  const { handles, storyId } = replicas(2, 'abcdefghij');
  const network = mesh(handles);
  let seed = 0x5eed;
  const random = () => {
    seed = (seed * 1_103_515_245 + 12_345) % 2_147_483_648;
    return seed / 2_147_483_648;
  };
  const pick = (limit: number) => Math.floor(random() * limit);
  try {
    for (let edit = 0; edit < 300; edit += 1) {
      const handle = handles[pick(handles.length)];
      const length = text(handle, storyId).length;
      const roll = random();
      if (roll < 0.15 || length < 3) {
        handle.insertText(storyId, pick(length + 1), 'xy');
      } else if (roll < 0.25) {
        const start = pick(length - 1);
        handle.deleteText(storyId, start, start + 1);
      } else {
        const start = pick(length - 1);
        const end = start + 1 + pick(length - start);
        const key = random() < 0.6 ? 'bold' : 'italic';
        handle.formatText(storyId, start, end, { [key]: random() < 0.5 });
      }
      if (edit % 3 === 2) {
        expect(network.settle()).toBeLessThanOrEqual(4);
        expectConverged(handles, storyId);
      }
    }
    const published = network.sent.length;
    for (const update of network.sent.slice()) {
      for (const handle of handles) handle.applyUpdate(update);
    }
    expect(network.settle()).toBe(0);
    expect(network.sent.length).toBe(published);
    expectConverged(handles, storyId);
  } finally {
    network.stop();
    for (const handle of handles) handle.dispose();
  }
});
