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

function replicas(): [PresentationHandle, PresentationHandle] {
  const source = openPresentation(fixture, { clientId: 4200 });
  const seed = source.encodeStateAsUpdate();
  source.dispose();
  return [
    openPresentation(fixture, { clientId: 4201, initialUpdate: seed }),
    openPresentation(fixture, { clientId: 4202, initialUpdate: seed }),
  ];
}

function exchange(left: PresentationHandle, right: PresentationHandle): void {
  const fromLeft = left.encodeStateAsUpdate();
  const fromRight = right.encodeStateAsUpdate();
  left.applyUpdate(fromRight);
  right.applyUpdate(fromLeft);
}

test('a slide moved on one replica and deleted on another is gone everywhere', () => {
  const [mover, deleter] = replicas();
  try {
    const ids = mover.snapshot().slides.map((slide) => slide.id);
    mover.moveSlide(ids[0], 1);
    deleter.deleteSlide(ids[0]);
    expect(() => exchange(mover, deleter)).not.toThrow();
    expect(mover.snapshot()).toEqual(deleter.snapshot());
    expect(mover.snapshot().slides.map((slide) => slide.id)).toEqual(ids.slice(1));
  } finally {
    mover.dispose();
    deleter.dispose();
  }
});

test('undoing a delete merged with a concurrent move restores the slide on both replicas', () => {
  const [mover, deleter] = replicas();
  try {
    const ids = mover.snapshot().slides.map((slide) => slide.id);
    mover.moveSlide(ids[0], 1);
    deleter.deleteSlide(ids[0]);
    exchange(mover, deleter);
    deleter.undo();
    expect(() => exchange(mover, deleter)).not.toThrow();
    expect(mover.snapshot()).toEqual(deleter.snapshot());
    expect(deleter.snapshot().slides.map((slide) => slide.id)).toEqual(ids);
  } finally {
    mover.dispose();
    deleter.dispose();
  }
});
