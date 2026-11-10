import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, test } from 'bun:test';
import { useRef } from 'react';
import type { DisplayPage, DisplayPrimitive } from '@betteroffice/docx/layout/render';
import type { TFunction } from '@betteroffice/docx-i18n';
import { CanvasPageMirror } from '../CanvasPageMirror';
import { fakeIdleScheduler } from '../__fixtures__/fakeIdleScheduler';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { useUnbuiltPageMirrors } from './UnbuiltPageMirrors';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render } = await import('@testing-library/react');
const t = ((key: string) => key) as TFunction;

let scheduler: ReturnType<typeof fakeIdleScheduler>;
beforeEach(() => { scheduler = fakeIdleScheduler(); });
afterEach(() => {
  getSelection()?.removeAllRanges();
  cleanup();
  scheduler.restore();
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const idle = () =>
  act(async () => {
    await new Promise<void>((resolve) => setTimeout(resolve, 200));
    while (scheduler.idleWork.size > 0) {
      scheduler.flushIdle({ didTimeout: false, timeRemaining: () => 50 });
      await Promise.resolve();
    }
  });

const settle = () => act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, 0)); });

const TEXT = 'Second paragraph';

type Block = [blockKey: string, text: string, paraId?: string, marker?: 'numbered' | 'bullet'];

function page(blocks: Block[] = [['p1', 'First paragraph'], ['p2', TEXT]]): DisplayPage {
  const primitives = blocks.flatMap(([blockKey, text, paraId, marker], index): DisplayPrimitive[] => {
    const at = { baselineY: 20 + index * 20, font: '12px Arial', color: '#000', blockKey, paraId,
      fragmentDocStart: 1009 + index * 100, fragmentDocEnd: 1011 + index * 100 + text.length };
    const body = { ...at, kind: 'text', text, x: 30, width: 60, docStart: 1010 + index * 100, docEnd: 1010 + index * 100 + text.length } as DisplayPrimitive;
    if (!marker) return [body];
    return [{ ...at, kind: 'text', text: marker === 'numbered' ? '1.' : '•', x: 10, width: 10, listMarker: true,
      listMarkerNumbered: marker === 'numbered', listMarkerBullet: marker === 'bullet' } as DisplayPrimitive, body];
  });
  return { pageIndex: 1, width: 100, height: 100, positionSpan: [1009, 1300], primitives } as DisplayPage;
}

type Unbuilt = [paraId: string | undefined, text: string][];

type Phase = 'unbuilt' | 'text' | 'full';
type Mirrors = ReturnType<typeof useUnbuiltPageMirrors>;

function Harness({ phase, displayPage, expose, fullMirror = true, unbuilt }: { phase: Phase; displayPage: DisplayPage; expose(mirrors: Mirrors): void; fullMirror?: boolean; unbuilt: Unbuilt }) {
  const column = useRef<HTMLDivElement>(null);
  const mirrors = useUnbuiltPageMirrors({
    binding: {} as PeerlessEditSessionBinding, geometry: null, mountedPages: [], center: 1, enabled: true,
    navigate: () => {}, column, t, fullMirror,
  });
  expose(mirrors);
  return <div ref={column}><div data-docx-pages-host="">
    {phase === 'unbuilt'
      ? <section data-docx-unbuilt-page="1">{unbuilt.map(([paraId, text], index) => <p key={index} tabIndex={-1} data-para-id={paraId}>{text}</p>)}</section>
      : <div data-page-index="1"><CanvasPageMirror page={displayPage} active={phase === 'full'} /></div>}
  </div></div>;
}

function mount(fullMirror = true, unbuilt: Unbuilt = [[undefined, TEXT]]) {
  let displayPage = page();
  let mirrors!: Mirrors;
  const expose = (value: Mirrors) => { mirrors = value; };
  const view = render(<Harness phase="unbuilt" displayPage={displayPage} expose={expose} fullMirror={fullMirror} unbuilt={unbuilt} />);
  const to = async (phase: Phase, next: DisplayPage = displayPage) => {
    displayPage = next;
    view.rerender(<Harness phase={phase} displayPage={displayPage} expose={expose} fullMirror={fullMirror} unbuilt={unbuilt} />);
    await settle();
  };
  const section = () => view.container.querySelector<HTMLElement>('[data-docx-unbuilt-page="1"]')!;
  const textBlock = () => Array.from(view.container.querySelectorAll<HTMLElement>('[data-page-index="1"] .canvas-page-mirror [role="paragraph"], [data-page-index="1"] .canvas-page-mirror p'))
    .find((element) => element.textContent?.trim() === TEXT) ?? null;
  const fullBlock = () =>
    view.container.querySelector<HTMLElement>('[data-page-index="1"] .canvas-page-mirror [data-block-id="p2"]');
  return { ...view, to, section, fullBlock, textBlock, remember: (...args: Parameters<Mirrors['remember']>) => mirrors.remember(...args) };
}

test('focus handed from an unbuilt paragraph lands in the full mirror, not in its text stand-in', async () => {
  const view = mount();
  const paragraph = view.section().querySelector('p')!;
  await act(async () => { paragraph.focus(); });
  await view.to('text');
  await idle();
  await settle();
  await view.to('full');
  await settle();
  const block = view.fullBlock();
  expect(block?.textContent).toBe(TEXT);
  expect(document.activeElement).toBe(block);
});

test('a selection handed from an unbuilt paragraph lands in the full mirror, not in its text stand-in', async () => {
  const view = mount();
  const section = view.section();
  const paragraph = section.querySelector('p')!;
  view.remember(section, paragraph, true);
  await view.to('text');
  await idle();
  await settle();
  await view.to('full');
  await settle();
  const block = view.fullBlock();
  expect(block?.textContent).toBe(TEXT);
  const selection = getSelection()!;
  expect(selection.toString()).toBe(TEXT);
  expect(block!.contains(selection.anchorNode)).toBe(true);
});

test('with text page mirrors, focus handed from an unbuilt paragraph lands in the text mirror', async () => {
  const view = mount(false);
  const paragraph = view.section().querySelector('p')!;
  await act(async () => { paragraph.focus(); });
  await view.to('text');
  await idle();
  await settle();
  const block = view.textBlock();
  expect(block).not.toBeNull();
  expect(document.activeElement).toBe(block);
});

const shell = (): DisplayPage => ({ ...page([]), unbuilt: true });
const mirrorRoot = (view: ReturnType<typeof mount>) =>
  view.container.querySelector<HTMLElement>('[data-page-index="1"] .canvas-page-mirror')!.firstElementChild;
const blockOf = (view: ReturnType<typeof mount>, key: string) =>
  view.container.querySelector<HTMLElement>(`[data-page-index="1"] .canvas-page-mirror [data-block-id="${key}"]`);
const focusUnbuilt = (view: ReturnType<typeof mount>, index: number) =>
  act(async () => { view.section().querySelectorAll<HTMLElement>('p')[index]!.focus(); });

test('focus handed to a page whose full mirror fills in place waits for the paragraph', async () => {
  const view = mount();
  const paragraph = view.section().querySelector('p')!;
  await act(async () => { paragraph.focus(); });
  await view.to('full', shell());
  const root = mirrorRoot(view);
  expect(root).not.toBeNull();
  expect(view.fullBlock()).toBeNull();
  await view.to('full', page());
  expect(mirrorRoot(view)).toBe(root);
  const block = view.fullBlock();
  expect(block?.textContent).toBe(TEXT);
  expect(document.activeElement).toBe(block);
});

test('a selection handed to a page whose full mirror fills in place waits for the paragraph', async () => {
  const view = mount();
  const section = view.section();
  view.remember(section, section.querySelector('p')!, true);
  await view.to('full', shell());
  const root = mirrorRoot(view);
  expect(root).not.toBeNull();
  await view.to('full', page());
  expect(mirrorRoot(view)).toBe(root);
  const block = view.fullBlock();
  const selection = getSelection()!;
  expect(selection.toString()).toBe(TEXT);
  expect(block!.contains(selection.anchorNode)).toBe(true);
});

test('while the page is an empty shell, focus waits on the named page and then lands on the paragraph', async () => {
  const view = mount();
  await focusUnbuilt(view, 0);
  await view.to('full', shell());
  const root = mirrorRoot(view);
  expect(root?.childNodes.length).toBeGreaterThan(0);
  expect(root?.querySelector('[role="paragraph"], p')).toBeNull();
  expect(document.activeElement).toBe(root);
  expect(root?.getAttribute('role')).toBe('document');
  expect(root?.getAttribute('aria-label')).toBe('a11y.pageLabel');
  await view.to('full', page());
  expect(mirrorRoot(view)).toBe(root);
  expect(document.activeElement).toBe(view.fullBlock());
});

test('a hand-off to a page built without paragraphs ends on the page and is not taken back', async () => {
  const view = mount();
  await focusUnbuilt(view, 0);
  await view.to('full', page([]));
  const root = mirrorRoot(view);
  expect(document.activeElement).toBe(root);
  await view.to('full', page());
  expect(view.fullBlock()).not.toBeNull();
  expect(document.activeElement).toBe(root);
});

test('a paragraph that left the page ends the hand-off at its remembered place, not where it reappears', async () => {
  const view = mount(true, [['A', 'First paragraph'], ['B', 'Moved paragraph'], ['C', 'Third paragraph']]);
  await focusUnbuilt(view, 1);
  await view.to('full', shell());
  await view.to('full', page([['p1', 'First paragraph', 'A'], ['p3', 'Third paragraph', 'C']]));
  expect(document.activeElement).toBe(blockOf(view, 'p3'));
  await view.to('full', page([['p1', 'First paragraph', 'A'], ['p2', 'Moved paragraph', 'B'], ['p3', 'Third paragraph', 'C']]));
  expect(document.activeElement?.getAttribute('data-block-id')).toBe('p3');
});

test('a selection whose paragraph left the page ends on the paragraph at its remembered place', async () => {
  const view = mount(true, [['A', 'First paragraph'], ['B', 'Moved paragraph']]);
  const section = view.section();
  view.remember(section, section.querySelectorAll('p')[1]!, true);
  await view.to('full', shell());
  await view.to('full', page([['p1', 'First paragraph', 'A'], ['p3', 'Third paragraph', 'C']]));
  const selection = getSelection()!;
  expect(selection.toString()).toBe('Third paragraph');
  expect(blockOf(view, 'p3')!.contains(selection.anchorNode)).toBe(true);
});

test('of two paragraphs with the same text, the hand-off lands on the one at the remembered place', async () => {
  const view = mount(true, [[undefined, 'Same text'], [undefined, 'Same text']]);
  await focusUnbuilt(view, 1);
  await view.to('full', shell());
  await view.to('full', page([['p1', 'Same text'], ['p2', 'Same text']]));
  expect(document.activeElement).toBe(blockOf(view, 'p2'));
});

for (const marker of ['numbered', 'bullet'] as const) {
  test(`focus handed from an unbuilt ${marker} list item lands on that item`, async () => {
    const view = mount(true, [['A', 'First paragraph'], ['B', 'List item']]);
    await focusUnbuilt(view, 1);
    await view.to('full', shell());
    await view.to('full', page([['p0', 'Inserted paragraph', 'Z'], ['p1', 'First paragraph', 'A'], ['p2', 'List item', 'B', marker]]));
    const item = blockOf(view, 'p2');
    expect(item?.textContent?.trim()).not.toBe('List item');
    expect(document.activeElement).toBe(item);
  });
}

test('without paragraph ids, focus handed from an unbuilt list item ends on the item at its remembered place', async () => {
  const view = mount(true, [[undefined, 'First paragraph'], [undefined, 'List item']]);
  await focusUnbuilt(view, 1);
  await view.to('full', shell());
  await view.to('full', page([['p1', 'First paragraph'], ['p2', 'List item', undefined, 'numbered'], ['p3', 'Third paragraph']]));
  const item = blockOf(view, 'p2');
  expect(item?.textContent?.trim()).not.toBe('List item');
  expect(document.activeElement).toBe(item);
});

test('focus handed from an unbuilt paragraph lands on it after its text was edited', async () => {
  const view = mount(true, [['A', 'First paragraph'], ['B', TEXT]]);
  await focusUnbuilt(view, 1);
  await view.to('full', shell());
  await view.to('full', page([['p0', 'Inserted paragraph', 'Z'], ['p1', 'First paragraph', 'A'], ['p2', `${TEXT}, edited`, 'B']]));
  expect(document.activeElement).toBe(blockOf(view, 'p2'));
});

test('focus moved elsewhere while a hand-off waits drops it', async () => {
  const view = mount();
  await focusUnbuilt(view, 0);
  await view.to('full', shell());
  const outside = document.createElement('button');
  document.body.appendChild(outside);
  await act(async () => { outside.focus(); });
  await view.to('full', page());
  expect(view.fullBlock()).not.toBeNull();
  expect(document.activeElement).toBe(outside);
  outside.remove();
});
