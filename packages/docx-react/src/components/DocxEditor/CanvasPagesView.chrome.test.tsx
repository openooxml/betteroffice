import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef } from 'react';
import type { DisplayList } from '@betteroffice/docx/layout/render';
import { CanvasPagesView } from './CanvasPagesView';

const { act, cleanup, render } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('only pages in the window and the page holding focus carry a mirror', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const displayList: DisplayList = {
    pages: Array.from({ length: 40 }, (_, pageIndex) => ({
      pageIndex,
      width: 100,
      height: 100,
      primitives: [],
    })),
  };
  render(
    <CanvasPagesView displayList={displayList} hostRef={hostRef} glyphOutlineProvider={() => ''} />
  );
  await act(async () => {});
  const host = hostRef.current!;
  const withMirror = () =>
    Array.from(host.querySelectorAll<HTMLElement>('.canvas-page'))
      .filter((page) => page.querySelector('.canvas-page-mirror'))
      .map((page) => Number(page.dataset.pageIndex));

  const mirrored = withMirror();
  expect(mirrored[0]).toBe(0);
  expect(mirrored.length).toBeLessThan(displayList.pages.length);
  expect(host.querySelectorAll('.canvas-page canvas')).toHaveLength(displayList.pages.length);

  const far = host.querySelector<HTMLElement>('.canvas-page[data-page-index="35"] canvas')!;
  await act(async () => {
    far.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
  });
  expect(withMirror()).toContain(35);
  await act(async () => {
    far.dispatchEvent(new FocusEvent('focusout', { bubbles: true, relatedTarget: null }));
  });
  expect(withMirror()).not.toContain(35);
});
