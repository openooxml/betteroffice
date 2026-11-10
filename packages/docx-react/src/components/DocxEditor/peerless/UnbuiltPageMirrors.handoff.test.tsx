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

function page(blocks: string[][] = [['p1', 'First paragraph'], ['p2', TEXT]]): DisplayPage {
  const primitives = blocks.map(([blockKey, text], index): DisplayPrimitive => ({
    kind: 'text', text, x: 10, baselineY: 20 + index * 20, width: 80, font: '12px Arial', color: '#000',
    blockKey, docStart: 1010 + index * 100, docEnd: 1010 + index * 100 + text!.length,
    fragmentDocStart: 1009 + index * 100, fragmentDocEnd: 1011 + index * 100 + text!.length,
  } as DisplayPrimitive));
  return { pageIndex: 1, width: 100, height: 100, positionSpan: [1009, 1300], primitives } as DisplayPage;
}

type Phase = 'unbuilt' | 'text' | 'full';
type Mirrors = ReturnType<typeof useUnbuiltPageMirrors>;

function Harness({ phase, displayPage, expose, fullMirror = true }: { phase: Phase; displayPage: DisplayPage; expose(mirrors: Mirrors): void; fullMirror?: boolean }) {
  const column = useRef<HTMLDivElement>(null);
  const mirrors = useUnbuiltPageMirrors({
    binding: {} as PeerlessEditSessionBinding, geometry: null, mountedPages: [], center: 1, enabled: true,
    navigate: () => {}, column, t, fullMirror,
  });
  expose(mirrors);
  return <div ref={column}><div data-docx-pages-host="">
    {phase === 'unbuilt'
      ? <section data-docx-unbuilt-page="1"><p tabIndex={-1}>{TEXT}</p></section>
      : <div data-page-index="1"><CanvasPageMirror page={displayPage} active={phase === 'full'} /></div>}
  </div></div>;
}

function mount(fullMirror = true) {
  let displayPage = page();
  let mirrors!: Mirrors;
  const expose = (value: Mirrors) => { mirrors = value; };
  const view = render(<Harness phase="unbuilt" displayPage={displayPage} expose={expose} fullMirror={fullMirror} />);
  const to = async (phase: Phase, next: DisplayPage = displayPage) => {
    displayPage = next;
    view.rerender(<Harness phase={phase} displayPage={displayPage} expose={expose} fullMirror={fullMirror} />);
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

const shell = () => page([['p1', 'First paragraph']]);
const mirrorRoot = (view: ReturnType<typeof mount>) =>
  view.container.querySelector<HTMLElement>('[data-page-index="1"] .canvas-page-mirror')!.firstElementChild;

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

test('focus handed to a page whose full mirror starts as an empty shell waits for the paragraph', async () => {
  const view = mount();
  const paragraph = view.section().querySelector('p')!;
  await act(async () => { paragraph.focus(); });
  await view.to('full', page([]));
  const root = mirrorRoot(view);
  expect(root?.childNodes.length).toBeGreaterThan(0);
  expect(root?.querySelector('[role="paragraph"], p')).toBeNull();
  await view.to('full', page());
  expect(mirrorRoot(view)).toBe(root);
  const block = view.fullBlock();
  expect(block?.textContent).toBe(TEXT);
  expect(document.activeElement).toBe(block);
});
