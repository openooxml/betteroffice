import { afterEach, expect, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { DocxDisplaySelectionText, ResidentDocumentRead, ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import { ViewerSelectionController } from './viewerSelectionController';

const controllers: ViewerSelectionController[] = [];
afterEach(() => { for (const controller of controllers.splice(0)) controller.reset(); });

function captured(text = 'Alpha'): DocxDisplaySelectionText {
  return { text, range: null, sticky: {
    anchor: { story: 'body', encoded: new Uint8Array([1]) },
    head: { story: 'body', encoded: new Uint8Array([2]) },
  } };
}

async function drain(): Promise<void> {
  for (let turn = 0; turn < 8; turn += 1) await Promise.resolve();
}

function setup(preview = false) {
  const pending: Array<{
    request: ResidentDocumentRead;
    resolve(reply: { version: string; value: unknown }): void;
    reject(error: Error): void;
  }> = [];
  const issued: ResidentDocumentRead[] = [];
  const read = ((request: ResidentDocumentRead) => new Promise((resolve, reject) => {
    issued.push(request);
    pending.push({ request, resolve, reject });
  })) as ResidentEngineWorkerClient['documentRead'];
  const queries = {
    displayList: { pages: [{ primitives: [{ docStart: 1, docEnd: 100 }] }] },
    isReady: () => true,
    paragraphRects: () => [],
    visualLineAtPosition: (position: number) => ({ from: 1, to: 100, position }),
  } as unknown as DisplayListQueries;
  const controller = new ViewerSelectionController({ read, story: 'body', queries: () => queries });
  controllers.push(controller);
  controller.onFrame({ version: 'A', preview });
  const take = (kind: ResidentDocumentRead['kind'], version: string) => {
    const index = pending.findIndex((entry) => entry.request.kind === kind &&
      'expectVersion' in entry.request && entry.request.expectVersion === version);
    expect(index).toBeGreaterThanOrEqual(0);
    return pending.splice(index, 1)[0]!;
  };
  const answer = async (kind: ResidentDocumentRead['kind'], version: string, value: unknown, replyVersion = version) => {
    take(kind, version).resolve({ version: replyVersion, value });
    await drain();
  };
  return { controller, pending, issued, answer, take };
}

test('R1: capture and unit channels keep only the newest queued read', async () => {
  const { controller, pending, answer, issued } = setup();
  controller.expand(1, 'word');
  controller.expand(5, 'word');
  controller.expand(9, 'word');
  expect(pending.filter((entry) => entry.request.kind === 'selectionUnit')).toHaveLength(1);
  expect(pending.filter((entry) => entry.request.kind === 'selectionText')).toHaveLength(1);
  await answer('selectionUnit', 'A', { anchor: 1, head: 4 });
  expect(issued.filter((request) => request.kind === 'selectionUnit')).toEqual([
    { kind: 'selectionUnit', story: 'body', position: 1, unit: 'word', expectVersion: 'A' },
    { kind: 'selectionUnit', story: 'body', position: 9, unit: 'word', expectVersion: 'A' },
  ]);
  await answer('selectionUnit', 'A', { anchor: 8, head: 12 });
  await answer('selectionText', 'A', null);
  expect(controller.displaySelection()).toEqual({ anchor: 8, head: 12 });
  expect(pending.filter((entry) => entry.request.kind === 'selectionText')).toHaveLength(1);
  await answer('selectionText', 'A', captured('Beta'));
  expect(controller.settledText()).toBe('Beta');
});

test('R2/R3: a superseded mapping changes nothing and only onFrame re-issues it', async () => {
  const { controller, answer, issued } = setup();
  controller.select(1, 6);
  await answer('selectionText', 'A', captured());
  controller.onFrame({ version: 'B', preview: false });
  expect(controller.displaySelection()).toBeNull();
  await answer('stickyPosition', 'B', null, 'C');
  expect(issued.filter((request) => request.kind === 'stickyPosition')).toHaveLength(1);
  controller.onFrame({ version: 'C', preview: false });
  await answer('stickyPosition', 'C', { anchor: 3, head: 8 });
  expect(controller.displaySelection()).toEqual({ anchor: 3, head: 8 });
  await answer('selectionText', 'C', captured());
  expect(controller.settledText()).toBe('Alpha');
});

test('R1/R3: mapping coalesces intermediate frames to the newest version', async () => {
  const { controller, answer, issued } = setup();
  controller.select(1, 6);
  await answer('selectionText', 'A', captured());
  for (const version of ['B', 'C', 'D']) controller.onFrame({ version, preview: false });
  await answer('stickyPosition', 'B', null, 'D');
  expect(issued.filter((request) => request.kind === 'stickyPosition').map((request) =>
    'expectVersion' in request ? request.expectVersion : null)).toEqual(['B', 'D']);
  await answer('stickyPosition', 'D', { anchor: 5, head: 10 });
  expect(controller.displaySelection()).toEqual({ anchor: 5, head: 10 });
});

test('R2: superseded and rejected captures preserve the live selection', async () => {
  const { controller, answer, take } = setup();
  controller.select(1, 6);
  await answer('selectionText', 'A', null, 'B');
  expect(controller.displaySelection()).toEqual({ anchor: 1, head: 6 });
  controller.select(8, 12);
  take('selectionText', 'A').reject(new Error('read failed'));
  await drain();
  expect(controller.displaySelection()).toEqual({ anchor: 8, head: 12 });
});

test('R2: a null capture for an older revision cannot drop a corrected word', async () => {
  const { controller, answer } = setup();
  controller.expand(4, 'word');
  await answer('selectionUnit', 'A', { anchor: 1, head: 6 });
  await answer('selectionText', 'A', null);
  expect(controller.displaySelection()).toEqual({ anchor: 1, head: 6 });
  await answer('selectionText', 'A', captured());
  expect(controller.settledText()).toBe('Alpha');
});

test('R3: a frame change waits for an in-flight capture before mapping its sticky ends', async () => {
  const { controller, answer, pending } = setup();
  controller.select(1, 6);
  const gesture = controller.currentGesture();
  const copy = controller.whenSettled(gesture);
  controller.onFrame({ version: 'B', preview: false });
  expect(pending.some((entry) => entry.request.kind === 'stickyPosition')).toBe(false);
  await answer('selectionText', 'A', captured());
  await answer('stickyPosition', 'B', { anchor: 3, head: 8 });
  await answer('selectionText', 'B', captured());
  expect(await copy).toBe('Alpha');
  expect(controller.isCurrent(gesture)).toBe(true);
});

test('R3: a superseded uncaptured selection drops only when a new frame is presented', async () => {
  const { controller, answer } = setup();
  controller.select(1, 6);
  const copy = controller.whenSettled(controller.currentGesture()).catch((error: Error) => error);
  await answer('selectionText', 'A', null, 'B');
  expect(controller.displaySelection()).toEqual({ anchor: 1, head: 6 });
  controller.onFrame({ version: 'B', preview: false });
  await drain();
  expect(controller.displaySelection()).toBeNull();
  expect((await copy as Error).message).toBe('Selection dropped');
});

test('R3/R9: preview handover re-issues pending unit and capture under the same gesture', async () => {
  const { controller, answer } = setup(true);
  controller.expand(4, 'word');
  const gesture = controller.currentGesture();
  const copy = controller.whenSettled(gesture);
  controller.onFrame({ version: 'B', preview: false });
  expect(controller.displaySelection()).toEqual({ anchor: 4, head: 4 });
  await answer('selectionUnit', 'A', null, 'B');
  await answer('selectionText', 'A', null, 'B');
  await answer('selectionUnit', 'B', { anchor: 1, head: 6 });
  await answer('selectionText', 'B', captured(''));
  await answer('selectionText', 'B', captured());
  expect(await copy).toBe('Alpha');
  expect(controller.isCurrent(gesture)).toBe(true);
});

test('R4/R7: a preview story stays pending and every new version re-reads the story', async () => {
  const { controller, answer, issued } = setup(true);
  controller.selectAll();
  await answer('selectionUnit', 'A', { anchor: 0, head: 10 });
  await answer('selectionText', 'A', captured('prefix'));
  await answer('selectionText', 'A', captured('prefix'));
  expect(controller.settledText()).toBeNull();
  const gesture = controller.currentGesture();
  const copy = controller.whenSettled(gesture);
  controller.onFrame({ version: 'B', preview: false });
  await answer('selectionUnit', 'B', { anchor: 0, head: 100 });
  await answer('selectionText', 'B', captured('whole story'));
  expect(await copy).toBe('whole story');
  controller.onFrame({ version: 'C', preview: false });
  await answer('selectionUnit', 'C', { anchor: 0, head: 110 });
  expect(controller.displaySelection()).toEqual({ anchor: 0, head: 110 });
  expect(issued.some((request) => request.kind === 'stickyPosition')).toBe(false);
});

test('R3: a new frame at the same version only refreshes visibility', async () => {
  const { controller, answer, issued } = setup();
  controller.select(1, 6);
  await answer('selectionText', 'A', captured());
  const count = issued.length;
  let changes = 0;
  const unsubscribe = controller.subscribe(() => { changes += 1; });
  controller.onFrame({ version: 'A', preview: false });
  expect(issued).toHaveLength(count);
  expect(controller.displaySelection()).toEqual({ anchor: 1, head: 6 });
  expect(changes).toBe(1);
  unsubscribe();
});

test('R6: copy rejects when a new gesture starts', async () => {
  const { controller } = setup();
  controller.select(1, 6);
  const copy = controller.whenSettled(controller.currentGesture()).catch((error: Error) => error);
  controller.beginGesture();
  expect((await copy as Error).message).toBe('Selection gesture changed');
});

test('R6: copy rejects when a valid null answer drops the selection', async () => {
  const { controller, answer } = setup();
  controller.select(1, 6);
  await answer('selectionText', 'A', captured());
  controller.onFrame({ version: 'B', preview: false });
  const copy = controller.whenSettled(controller.currentGesture()).catch((error: Error) => error);
  await answer('stickyPosition', 'B', null);
  expect((await copy as Error).message).toBe('Selection dropped');
  expect(controller.displaySelection()).toBeNull();
});

test('a selection the worker cannot capture stays shown and copies nothing', async () => {
  const { controller, answer } = setup();
  controller.select(1, 6);
  const copy = controller.whenSettled(controller.currentGesture());
  await answer('selectionText', 'A', null);
  expect(await copy).toBe('');
  expect(controller.displaySelection()).toEqual({ anchor: 1, head: 6 });
});

test('a word unit the worker cannot resolve keeps the caret and settles', async () => {
  const { controller, answer } = setup();
  controller.expand(5, 'word');
  await answer('selectionUnit', 'A', null);
  expect(controller.displaySelection()).not.toBeNull();
});

test('R6: copy times out without polling or issuing more reads', async () => {
  const { controller, issued } = setup();
  controller.select(1, 6);
  const copy = controller.whenSettled(controller.currentGesture(), 5).catch((error: Error) => error);
  expect((await copy as Error).message).toBe('Selection did not settle');
  expect(issued).toHaveLength(1);
});

test('R8: reset cancels copies and ignores old replies without freeing an occupied channel', async () => {
  const { controller, answer, pending } = setup();
  controller.select(1, 6);
  const gesture = controller.currentGesture();
  const copy = controller.whenSettled(gesture).catch((error: Error) => error);
  controller.reset();
  expect(controller.isCurrent(gesture)).toBe(false);
  expect(controller.displaySelection()).toBeNull();
  expect((await copy as Error).message).toBe('Selection gesture changed');
  controller.onFrame({ version: 'B', preview: false });
  controller.select(8, 12);
  expect(pending).toHaveLength(1);
  await answer('selectionText', 'A', captured('old'));
  expect(controller.displaySelection()).toEqual({ anchor: 8, head: 12 });
  await answer('selectionText', 'B', captured('new'));
  expect(controller.settledText()).toBe('new');
});

test('R2/R6: drag revisions share one gesture and copy waits for the newest capture', async () => {
  const { controller, answer } = setup();
  const gesture = controller.beginGesture();
  controller.select(1, 6, gesture);
  const copy = controller.whenSettled(gesture);
  controller.select(1, 12, gesture);
  await answer('selectionText', 'A', captured('old'));
  expect(controller.settledText()).toBeNull();
  await answer('selectionText', 'A', captured('latest'));
  expect(await copy).toBe('latest');
  expect(controller.isCurrent(gesture)).toBe(true);
});

test('R3: restoring a temporarily absent frame re-issues reads ignored while it was absent', async () => {
  const { controller, answer } = setup();
  controller.expand(4, 'word');
  const copy = controller.whenSettled(controller.currentGesture());
  controller.onFrame(null);
  await answer('selectionUnit', 'A', { anchor: 1, head: 6 });
  await answer('selectionText', 'A', captured(''));
  controller.onFrame({ version: 'A', preview: false });
  await answer('selectionUnit', 'A', { anchor: 1, head: 6 });
  await answer('selectionText', 'A', captured(''));
  await answer('selectionText', 'A', captured());
  expect(await copy).toBe('Alpha');
});
