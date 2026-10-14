import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, spyOn, test } from 'bun:test';
import { useRef } from 'react';
import {
  applyFrameDelta,
  buildInteractiveOverlayPage,
  buildMirrorPage,
  displayPageNoteAnchorRevision,
  displayPageRevision,
  FRAME_DELTA_VERSION,
  type DecodedFrameDelta,
  type DisplayPage,
  type DisplayPrimitive,
  type FramePageOperation,
  type RetainedFrame,
} from '@betteroffice/docx/layout/render';
import type { TFunction } from '@betteroffice/docx-i18n';
import { usePageChrome } from './usePageChrome';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render } = await import('@testing-library/react');
const t = ((key: string) => key) as TFunction;
const pending = new Map<number, IdleRequestCallback>();
const restorers: (() => void)[] = [];

beforeEach(() => {
  let id = 0;
  const request = globalThis.requestIdleCallback;
  const cancel = globalThis.cancelIdleCallback;
  globalThis.requestIdleCallback = (work) => {
    pending.set(++id, work);
    return id;
  };
  globalThis.cancelIdleCallback = (id) => { pending.delete(id); };
  restorers.push(() => { globalThis.requestIdleCallback = request; globalThis.cancelIdleCallback = cancel; });
});
afterEach(() => {
  cleanup();
  pending.clear();
  for (const restore of restorers.splice(0).reverse()) restore();
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

async function idle(): Promise<void> {
  await act(async () => {
    for (const [id, work] of [...pending]) {
      pending.delete(id);
      work({ didTimeout: false, timeRemaining: () => 50 });
    }
  });
}

function page(pageIndex: number): DisplayPage {
  const start = 10 + pageIndex * 100;
  const text = (offset: number, blockKey: string): DisplayPrimitive => ({
    kind: 'text', text: 'words', x: 10, baselineY: 20, width: 30, font: '12px Arial', color: '#000',
    blockKey, docStart: start + offset, docEnd: start + offset + 5,
    fragmentDocStart: start + offset - 1, fragmentDocEnd: start + offset + 6,
  });
  return {
    pageIndex, width: 100, height: 100, positionSpan: [start - 1, start + 6],
    primitives: [text(0, 'body'), {
      kind: 'rect', x: 10, y: 30, w: 10, h: 10, docStart: start, docEnd: start + 1,
      inlineSdtWidget: { kind: 'checkbox', groupId: 'control', controlId: 1, pos: start, checked: false },
    }],
    noteAreas: [{ kind: 'footnote', y: 80, height: 20, noteIds: [1],
      notes: [{ id: 1, anchorDocStart: start, anchorDocEnd: start + 1 }],
      separatorPrimitives: [text(20, 'separator')], primitives: [{ ...text(30, 'note'), groupId: 'footnote-1' }] }],
    header: { kind: 'header', rId: 'header', y: 0, height: 10, primitives: [text(40, 'header')] },
    footer: { kind: 'footer', rId: 'footer', y: 90, height: 10, primitives: [text(50, 'footer')] },
  } as DisplayPage;
}

function delta(epoch: number, operations: FramePageOperation[]): DecodedFrameDelta {
  return { protocolVersion: FRAME_DELTA_VERSION, full: epoch === 1, docEpoch: epoch, layoutEpoch: epoch,
    frameEpoch: epoch, baseFrameEpoch: epoch - 1, pageCount: 3, operations, bytes: new Uint8Array() };
}

function upsert(page: DisplayPage, fingerprint: bigint): FramePageOperation {
  return { kind: 'upsert', pageId: BigInt(page.pageIndex + 1), pageIndex: page.pageIndex, fingerprint,
    page, primitiveIds: new BigUint64Array([1n, 2n, 3n, 4n, 5n, 6n]) };
}

function initial(): RetainedFrame {
  return applyFrameDelta(null, delta(1, [0, 1, 2].map((index) => upsert(page(index), 1n))));
}

function shift(epoch: number, anchors = false): DecodedFrameDelta {
  return delta(epoch, [{ kind: 'shift-positions', pageIndex: 1, pageId: 2n, fingerprint: BigInt(epoch),
    runs: [{ start: 0, count: 6, changedMask: 1 | 2 | 4 | 8 | 16 | 128, delta: 5 }], spanDelta: 5,
    anchors: anchors ? [{ area: 0, note: 0, start: 115, end: 116 }] : [] },
  { kind: 'shift-range', pageIndex: 2, pageId: 3n, count: 1, delta: 5, salt: 5n }]);
}

function Chrome({ page, make, overlay = false, translate = t }: {
  page: DisplayPage; make(page: DisplayPage): HTMLElement; overlay?: boolean; translate?: TFunction;
}) {
  const ref = useRef<HTMLDivElement>(null);
  usePageChrome(ref, { page, t: translate, active: true, defer: false, rebuildAtOnce: true,
    urgentRevision: overlay ? displayPageRevision(page) : displayPageNoteAnchorRevision(page), make });
  return <div ref={ref} data-page={page.pageIndex} />;
}

function mirrors(first: RetainedFrame) {
  const calls: number[] = [];
  const make = (page: DisplayPage) => { calls.push(page.pageIndex); return buildMirrorPage(page); };
  const view = (frame: RetainedFrame) => <>{frame.displayList.pages.map((page) =>
    <Chrome key={page.pageIndex} page={page} make={make} />)}</>;
  const rendered = render(view(first));
  const hosts = [...rendered.container.querySelectorAll<HTMLElement>('[data-page]')];
  const replacements = hosts.map((host) => spyOn(host, 'replaceChildren'));
  restorers.push(...replacements.map((replacement) => () => replacement.mockRestore()));
  calls.length = 0;
  return { ...rendered, hosts, calls, replacements, update: (frame: RetainedFrame) => rendered.rerender(view(frame)) };
}

function exactMirrors(hosts: HTMLElement[], frame: RetainedFrame): void {
  hosts.forEach((host, index) => {
    expect(host.firstElementChild?.outerHTML).toBe(buildMirrorPage(frame.displayList.pages[index]!).outerHTML);
  });
}

test('copied shifts do not build or replace visible mirrors synchronously, then match fresh builds', async () => {
  const first = initial();
  const before = structuredClone(first.displayList);
  const { hosts, calls, replacements, update } = mirrors(first);
  const edited = page(0);
  edited.primitives = [{ ...edited.primitives[0]!, text: 'edited' } as DisplayPrimitive, edited.primitives[1]!];
  const changes = shift(2);
  const second = applyFrameDelta(first, { ...changes, operations: [upsert(edited, 2n), ...changes.operations] });
  update(second);
  expect(calls).toEqual([0]);
  expect(replacements.map((replace) => replace.mock.calls.length)).toEqual([1, 0, 0]);
  expect(first.displayList).toEqual(before);
  await idle();
  exactMirrors(hosts, second);
});

test('copied shift lineage composes beyond the shift log while a mirror waits for idle', async () => {
  const first = initial();
  const { hosts, calls, update } = mirrors(first);
  let latest = first;
  for (let epoch = 2; epoch <= 67; epoch++) {
    latest = applyFrameDelta(latest, shift(epoch));
    update(latest);
  }
  expect(calls).toEqual([]);
  await idle();
  expect(calls).toEqual([1, 2]);
  exactMirrors(hosts, latest);
});

test('content after a copied shift rebuilds the visible mirror at once', async () => {
  const first = initial();
  const { hosts, calls, update } = mirrors(first);
  const second = applyFrameDelta(first, shift(2));
  update(second);
  calls.length = 0;
  const shifted = second.displayList.pages[1]!;
  const edited = { ...shifted, primitives: shifted.primitives.map((primitive) =>
    primitive.kind === 'text' ? { ...primitive, text: 'changed' } : primitive) };
  const third = applyFrameDelta(second, delta(3, [upsert(edited, 3n)]));
  update(third);
  expect(calls).toEqual([1]);
  expect(hosts[1]!.firstElementChild?.outerHTML).toBe(buildMirrorPage(edited).outerHTML);
});

test('note anchors after a copied shift rebuild the visible mirror at once', async () => {
  const first = initial();
  const { hosts, calls, update } = mirrors(first);
  const second = applyFrameDelta(first, shift(2));
  update(second);
  calls.length = 0;
  const third = applyFrameDelta(second, shift(3, true));
  update(third);
  expect(calls).toEqual([1]);
  expect(hosts[1]!.querySelector<HTMLElement>('aside[data-note-id="1"]')?.dataset.anchorDocStart).toBe('115');
});

test('translation after a copied shift rebuilds the visible mirror at once', () => {
  const first = initial();
  const make = spyOn({ build: buildMirrorPage }, 'build');
  const rendered = render(<Chrome page={first.displayList.pages[1]!} make={make} />);
  const second = applyFrameDelta(first, shift(2));
  rendered.rerender(<Chrome page={second.displayList.pages[1]!} make={make} />);
  make.mockClear();
  rendered.rerender(<Chrome page={second.displayList.pages[1]!} make={make} translate={((key: string) => `new ${key}`) as TFunction} />);
  expect(make).toHaveBeenCalledTimes(1);
});

test('copied shifts advance overlay revisions and immediately match fresh hit targets', () => {
  const first = initial();
  const base = first.displayList.pages[1]!;
  const make = spyOn({ build: buildInteractiveOverlayPage }, 'build');
  const { container, rerender } = render(<Chrome page={base} make={make} overlay />);
  const old = container.firstElementChild!.firstElementChild;
  make.mockClear();
  const second = applyFrameDelta(first, shift(2));
  const shifted = second.displayList.pages[1]!;
  rerender(<Chrome page={shifted} make={make} overlay />);
  expect(displayPageRevision(shifted)).toBe(displayPageRevision(base) + 1);
  expect(make).toHaveBeenCalledTimes(1);
  expect(container.firstElementChild!.firstElementChild).not.toBe(old);
  expect(container.firstElementChild!.firstElementChild?.outerHTML).toBe(buildInteractiveOverlayPage(shifted).outerHTML);
  expect(container.querySelector<HTMLElement>('button[data-sdt-pos]')?.dataset.sdtPos).toBe('115');
  expect(base.primitives[1]!.inlineSdtWidget?.pos).toBe(110);
});
