import { afterEach, expect, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { DocxDisplaySelectionText, ResidentDocumentRead, ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import { ViewerSelectionController } from './viewerSelectionController';

const controllers: ViewerSelectionController[] = [];
afterEach(() => { for (const controller of controllers.splice(0)) controller.reset(); });

function captured(text = 'Alpha'): DocxDisplaySelectionText {
  return { text, range: null };
}

async function drain(): Promise<void> {
  for (let turn = 0; turn < 8; turn += 1) await Promise.resolve();
}

function setup(preview = false, lines = [{ pageIndex: 0, from: 1, to: 100 }]) {
  const pending: Array<{
    request: ResidentDocumentRead;
    resolve(reply: { version: string; value: unknown }): void;
    reject(error: Error): void;
  }> = [];
  const issued: ResidentDocumentRead[] = [];
  const read = ((request: ResidentDocumentRead) => new Promise<{ version: string; value: unknown }>((resolve, reject) => {
    issued.push(request);
    pending.push({ request, resolve, reject });
  })) as unknown as ResidentEngineWorkerClient['documentRead'];
  const queries = {
    displayList: { pages: Array.from(new Set(lines.map((line) => line.pageIndex)), (pageIndex) => ({
      pageIndex,
      primitives: lines.filter((line) => line.pageIndex === pageIndex).map((line) => ({
        docStart: line.from, docEnd: line.to,
      })),
    })) },
    isReady: () => true,
    paragraphRects: () => [],
    visualLinesOnPage: (pageIndex: number) => lines.filter((line) => line.pageIndex === pageIndex),
    visualLineAtPosition: (position: number) => lines.find((line) => position >= line.from && position <= line.to) ?? null,
    verticalMove: () => null,
  } as unknown as DisplayListQueries;
  const controller = new ViewerSelectionController({ read, story: 'body', queries: () => queries });
  controllers.push(controller);
  controller.onFrame({ version: 'A', preview, asOpened: false });
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
  return { controller, queries, pending, issued, answer, take };
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

test('a settled capture retains its worker range and emits with unchanged endpoints', async () => {
  const { controller, answer, issued } = setup();
  const changes: Array<{ selection: ReturnType<ViewerSelectionController['displaySelection']>;
    capture: DocxDisplaySelectionText | null }> = [];
  controller.subscribe(() => changes.push({
    selection: controller.displaySelection(), capture: controller.settledCapture(),
  }));
  controller.select(12, 1);
  const range: NonNullable<DocxDisplaySelectionText['range']> = {
    story: 'body', view: 'accepted',
    start: { paraId: '00000001', offset: 0 },
    end: { paraId: '00000002', offset: 4 },
  };
  expect(changes).toEqual([{ selection: { anchor: 12, head: 1 }, capture: null }]);
  await answer('selectionText', 'A', { text: 'Alpha\nBeta', range });
  expect(controller.settledCapture()).toEqual({ text: 'Alpha\nBeta', range });
  expect(changes).toEqual([
    { selection: { anchor: 12, head: 1 }, capture: null },
    { selection: { anchor: 12, head: 1 }, capture: { text: 'Alpha\nBeta', range } },
  ]);
  expect(issued).toHaveLength(1);
});

test('a superseded capture cannot expose its range for the current revision', async () => {
  const { controller, answer } = setup();
  const gesture = controller.beginGesture();
  controller.select(1, 6, gesture);
  controller.select(1, 12, gesture);
  await answer('selectionText', 'A', {
    text: 'Alpha', range: { story: 'body', view: 'accepted',
      start: { paraId: '00000001', offset: 0 }, end: { paraId: '00000001', offset: 5 } },
  });
  expect(controller.settledCapture()).toBeNull();
  await answer('selectionText', 'A', captured('Alpha\nBeta'));
  expect(controller.settledCapture()).toEqual(captured('Alpha\nBeta'));
  controller.beginGesture();
  expect(controller.settledCapture()).toBeNull();
});

test('a version change clears the selection, rejects its copy and ends its gesture', async () => {
  const { controller, answer, issued } = setup();
  controller.select(1, 6);
  const gesture = controller.currentGesture();
  const copy = controller.readSelectedText()!.catch((error: Error) => error);
  const settled = controller.whenSettled(gesture).catch((error: Error) => error);
  controller.onFrame({ version: 'B', preview: false, asOpened: false });
  expect(controller.displaySelection()).toBeNull();
  expect(controller.isCurrent(gesture)).toBe(false);
  expect((await copy as Error).message).toBe('Selection cleared');
  expect((await settled as Error).message).toBe('Selection cleared');
  controller.select(1, 12, gesture);
  await answer('selectionText', 'A', captured());
  expect(controller.displaySelection()).toBeNull();
  expect(controller.settledText()).toBeNull();
  expect(controller.readSelectedText()).toBeNull();
  expect(issued).toHaveLength(1);
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

test('the hand-over from a preview to the document as opened keeps the selection and its gesture', async () => {
  const { controller, answer, issued } = setup(true);
  controller.expand(5, 'word');
  const selection = controller.displaySelection();
  const gesture = controller.currentGesture();
  const copy = controller.readSelectedText()!;
  controller.onFrame({ version: 'F', preview: false, asOpened: true });
  expect(controller.displaySelection()).toEqual(selection);
  expect(controller.isCurrent(gesture)).toBe(true);
  await answer('selectionUnit', 'A', { anchor: 1, head: 6 });
  await answer('selectionText', 'A', captured('preview'));
  expect(controller.displaySelection()).toEqual(selection);
  expect(issued).toContainEqual({
    kind: 'selectionUnit', story: 'body', position: 5, unit: 'word', expectVersion: 'F',
  });
  expect(issued).toContainEqual({
    kind: 'selectionText', story: 'body', anchor: 5, head: 5, expectVersion: 'F',
  });
  await answer('selectionText', 'F', captured(''));
  expect(controller.settledText()).toBeNull();
  await answer('selectionUnit', 'F', { anchor: 1, head: 6 });
  await answer('selectionText', 'F', captured());
  expect(await copy).toBe('Alpha');
  expect(controller.isCurrent(gesture)).toBe(true);
});

test('a preview hand-over to a changed document clears the selection', async () => {
  const { controller, answer, issued } = setup(true);
  controller.expand(5, 'word');
  const gesture = controller.currentGesture();
  const copy = controller.readSelectedText()!.catch((error: Error) => error);
  controller.onFrame({ version: 'F', preview: false, asOpened: false });
  expect(controller.displaySelection()).toBeNull();
  expect(controller.isCurrent(gesture)).toBe(false);
  expect((await copy as Error).message).toBe('Selection cleared');
  await answer('selectionUnit', 'A', { anchor: 1, head: 6 });
  await answer('selectionText', 'A', captured());
  expect(controller.displaySelection()).toBeNull();
  expect(controller.settledText()).toBeNull();
  expect(issued).toHaveLength(2);
});

test('select-all on a preview reads the whole story on the document as opened', async () => {
  const { controller, answer, issued } = setup(true);
  controller.selectAll();
  await answer('selectionUnit', 'A', { anchor: 0, head: 10 });
  await answer('selectionText', 'A', captured('prefix'));
  await answer('selectionText', 'A', captured('prefix'));
  expect(controller.settledText()).toBeNull();
  const gesture = controller.currentGesture();
  let finished = false;
  const copy = controller.readSelectedText()!.then((text) => { finished = true; return text; });
  await drain();
  expect(finished).toBe(false);
  controller.onFrame({ version: 'F', preview: false, asOpened: true });
  expect(controller.isCurrent(gesture)).toBe(true);
  expect(issued).toContainEqual({
    kind: 'selectionUnit', story: 'body', position: 0, unit: 'story', expectVersion: 'F',
  });
  await answer('selectionText', 'F', captured('prefix'));
  expect(finished).toBe(false);
  await answer('selectionUnit', 'F', { anchor: 0, head: 100 });
  await answer('selectionText', 'F', captured('whole story'));
  expect(await copy).toBe('whole story');
  expect(controller.displaySelection()).toEqual({ anchor: 0, head: 100 });
  const count = issued.length;
  controller.onFrame({ version: 'C', preview: false, asOpened: false });
  expect(controller.displaySelection()).toBeNull();
  expect(controller.isCurrent(gesture)).toBe(false);
  expect(controller.settledText()).toBeNull();
  expect(issued).toHaveLength(count);
});

test('keyboard extension stops at the last built page', async () => {
  const { controller, answer, issued } = setup(false, [
    { pageIndex: 0, from: 1, to: 40 },
    { pageIndex: 1, from: 41, to: 80 },
  ]);
  controller.selectAll();
  await answer('selectionUnit', 'A', { anchor: 0, head: 500 });
  expect(controller.move('ArrowRight', false)).toBe(true);
  expect(controller.displaySelection()).toEqual({ anchor: 500, head: 500 });
  const gesture = controller.currentGesture();
  const count = issued.length;
  expect(controller.move('ArrowLeft', true)).toBe(true);
  expect(controller.displaySelection()).toEqual({ anchor: 500, head: 500 });
  expect(controller.isCurrent(gesture)).toBe(true);
  expect(issued).toHaveLength(count);
  controller.select(80);
  const lastPageGesture = controller.currentGesture();
  expect(controller.move('ArrowRight', true)).toBe(true);
  expect(controller.displaySelection()).toEqual({ anchor: 80, head: 80 });
  expect(controller.isCurrent(lastPageGesture)).toBe(true);
  controller.select(40);
  expect(controller.move('ArrowRight', true)).toBe(true);
  expect(controller.displaySelection()).toEqual({ anchor: 40, head: 41 });
});

test('a horizontal step in a split table row stays on the head page and the next one', () => {
  const { controller } = setup(false, [
    { pageIndex: 0, from: 4, to: 40 },
    { pageIndex: 0, from: 508, to: 540 },
    { pageIndex: 3, from: 124, to: 160 },
  ]);
  controller.select(40);
  expect(controller.move('ArrowRight', true)).toBe(true);
  expect(controller.displaySelection()).toEqual({ anchor: 40, head: 508 });
  controller.select(508);
  expect(controller.move('ArrowLeft', true)).toBe(true);
  expect(controller.displaySelection()).toEqual({ anchor: 508, head: 40 });
});

test('a vertical move does not jump across unbuilt pages', () => {
  const { controller, queries, issued } = setup(false, [
    { pageIndex: 0, from: 1, to: 40 },
    { pageIndex: 5, from: 201, to: 240 },
  ]);
  queries.verticalMove = () => ({ position: 210, goalX: 10 });
  controller.select(20);
  const gesture = controller.currentGesture();
  const count = issued.length;
  expect(controller.move('ArrowDown', true)).toBe(true);
  expect(controller.displaySelection()).toEqual({ anchor: 20, head: 20 });
  expect(controller.isCurrent(gesture)).toBe(true);
  expect(issued).toHaveLength(count);
});

test('R3: a new frame at the same version only refreshes visibility', async () => {
  const { controller, answer, issued } = setup();
  controller.select(1, 6);
  await answer('selectionText', 'A', captured());
  const count = issued.length;
  let changes = 0;
  const unsubscribe = controller.subscribe(() => { changes += 1; });
  controller.onFrame({ version: 'A', preview: false, asOpened: false });
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

test('a gesture that selects nothing hides the selection it replaced', async () => {
  const { controller, answer } = setup();
  let changes = 0;
  controller.subscribe(() => { changes += 1; });
  controller.select(1, 6);
  await answer('selectionText', 'A', captured());
  expect(controller.readSelectedText()).not.toBeNull();
  changes = 0;
  controller.beginGesture();
  expect(changes).toBe(1);
  expect(controller.displaySelection()).toBeNull();
  expect(controller.readSelectedText()).toBeNull();
  controller.select(3, 8);
  expect(controller.displaySelection()).toEqual({ anchor: 3, head: 8 });
  await answer('selectionText', 'A', captured('Beta'));
  expect(await controller.readSelectedText()).toBe('Beta');
});

test('copy rejects when a valid null story answer drops the selection', async () => {
  const { controller, answer } = setup();
  controller.selectAll();
  const copy = controller.whenSettled(controller.currentGesture()).catch((error: Error) => error);
  await answer('selectionUnit', 'A', null);
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
  expect((await copy as Error).message).toBe('Selection cleared');
  controller.onFrame({ version: 'B', preview: false, asOpened: false });
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
  controller.onFrame({ version: 'A', preview: false, asOpened: false });
  await answer('selectionUnit', 'A', { anchor: 1, head: 6 });
  await answer('selectionText', 'A', captured(''));
  await answer('selectionText', 'A', captured());
  expect(await copy).toBe('Alpha');
});
