import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from '@betteroffice/docx/layout/render';
import { CanvasPageMirror } from './CanvasPageMirror';
import { describeNode, sameNode } from './__fixtures__/nodeIdentity';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render } = await import('@testing-library/react');

afterEach(() => {
  getSelection()?.removeAllRanges();
  cleanup();
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

function page(pageIndex: number, texts: Record<string, string>): DisplayPage {
  const base = 10 + pageIndex * 1000;
  const primitives = Object.entries(texts).map(([blockKey, text], index): DisplayPrimitive => ({
    kind: 'text', text, x: 10, baselineY: 20 + index * 20, width: 80, font: '12px Arial', color: '#000',
    blockKey, docStart: base + index * 100, docEnd: base + index * 100 + text.length,
    fragmentDocStart: base + index * 100 - 1, fragmentDocEnd: base + index * 100 + text.length + 1,
  } as DisplayPrimitive));
  return { pageIndex, width: 100, height: 100, positionSpan: [base - 1, base + 300], primitives } as DisplayPage;
}

const first = { p1: 'First paragraph', p2: 'Second paragraph' };
const edited = { p1: 'First paragraph', p2: 'Second paragraph, edited' };

function Pages({ pages }: { pages: DisplayPage[] }) {
  return <>{pages.map((page) => (
    <div key={page.pageIndex} data-page-index={page.pageIndex}><CanvasPageMirror page={page} /></div>
  ))}</>;
}

function mount(texts: Record<string, string>[]) {
  const rendered = render(<><Pages pages={texts.map((text, index) => page(index, text))} /><button>outside</button></>);
  const host = (index: number) => rendered.container.querySelector<HTMLElement>(`[data-page-index="${index}"] .canvas-page-mirror`)!;
  const block = (index: number, id: string) => host(index).querySelector<HTMLElement>(`[data-block-id="${id}"]`)!;
  const update = (next: Record<string, string>[]) =>
    rendered.rerender(<><Pages pages={next.map((text, index) => page(index, text))} /><button>outside</button></>);
  return { ...rendered, host, block, update };
}

function textPoint(block: HTMLElement, offset: number): [Node, number] {
  const walker = block.ownerDocument.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  let seen = 0;
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const length = node.textContent?.length ?? 0;
    if (offset <= seen + length) return [node, offset - seen];
    seen += length;
  }
  throw new Error(`offset ${offset} is past the block's text`);
}

function select(from: [HTMLElement, number], to: [HTMLElement, number]): void {
  const selection = getSelection()!;
  const range = document.createRange();
  range.setStart(...textPoint(...from));
  selection.removeAllRanges();
  selection.addRange(range);
  selection.extend(...textPoint(...to));
}

test('a rebuilt page keeps focus on the focused mirror paragraph that is not a tab stop', () => {
  const view = mount([first]);
  const before = view.block(0, 'p2');
  before.tabIndex = -1;
  before.focus();
  expect(sameNode(document.activeElement, before)).toBe('same');
  view.update([edited]);
  const after = view.block(0, 'p2');
  expect(after === before).toBe(false);
  expect(after.textContent).toContain('edited');
  expect(sameNode(document.activeElement, after)).toBe('same');
});

test('a rebuilt page keeps a selection inside it by paragraph and offset', () => {
  const view = mount([first]);
  select([view.block(0, 'p2'), 0], [view.block(0, 'p2'), 6]);
  expect(getSelection()!.toString()).toBe('Second');
  const before = view.block(0, 'p2');
  view.update([edited]);
  const after = view.block(0, 'p2');
  expect(after === before).toBe(false);
  const selection = getSelection()!;
  expect(selection.toString()).toBe('Second');
  expect(after.contains(selection.anchorNode)).toBe(true);
  expect(after.contains(selection.focusNode)).toBe(true);
});

test('a rebuilt page does not pull back focus that moved elsewhere', () => {
  const view = mount([first]);
  const block = view.block(0, 'p2');
  block.tabIndex = -1;
  block.focus();
  const outside = view.getByText('outside');
  outside.focus();
  view.update([edited]);
  expect(sameNode(document.activeElement, outside)).toBe('same');
});

test('a rebuilt page does not collapse or move a selection that extends to another page', () => {
  const view = mount([first, first]);
  select([view.block(0, 'p2'), 3], [view.block(1, 'p1'), 5]);
  const focusNode = getSelection()!.focusNode;
  view.update([edited, first]);
  const selection = getSelection()!;
  expect(selection.isCollapsed).toBe(false);
  expect(sameNode(selection.focusNode, focusNode)).toBe('same');
  expect(selection.focusOffset).toBe(5);
  expect(view.host(0).contains(selection.focusNode)).toBe(false);
});

test('a rebuilt page whose focused paragraph is gone restores neither focus nor selection', () => {
  const view = mount([first]);
  const block = view.block(0, 'p2');
  block.tabIndex = -1;
  block.focus();
  select([block, 0], [block, 6]);
  view.update([{ p1: 'First paragraph' }]);
  expect(describeNode(view.host(0).querySelector('[data-block-id="p2"]'))).toBe('null');
  expect(view.host(0).contains(document.activeElement)).toBe(false);
  expect(view.host(0).contains(getSelection()!.anchorNode)).toBe(false);
  expect(view.host(0).contains(getSelection()!.focusNode)).toBe(false);
});

test('a rebuilt page restores the selection the user made last inside it, not an earlier one', () => {
  const view = mount([first]);
  select([view.block(0, 'p2'), 0], [view.block(0, 'p2'), 6]);
  select([view.block(0, 'p1'), 6], [view.block(0, 'p1'), 15]);
  expect(getSelection()!.toString()).toBe('paragraph');
  view.update([edited]);
  const selection = getSelection()!;
  expect(selection.toString()).toBe('paragraph');
  expect(view.block(0, 'p1').contains(selection.anchorNode)).toBe(true);
});

test('a rebuilt page leaves a selection whose text survived as it was', () => {
  const view = mount([first]);
  const p1 = view.block(0, 'p1');
  select([p1, 0], [p1, 5]);
  const { anchorNode, focusNode } = getSelection()!;
  view.update([edited]);
  const selection = getSelection()!;
  expect(sameNode(view.block(0, 'p1'), p1)).toBe('same');
  expect(sameNode(selection.anchorNode, anchorNode)).toBe('same');
  expect(sameNode(selection.focusNode, focusNode)).toBe('same');
  expect(selection.toString()).toBe('First');
});
