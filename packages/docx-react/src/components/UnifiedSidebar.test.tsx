import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { StrictMode } from 'react';
import type { ReactSidebarItem } from '../plugin-api/types';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const { UnifiedSidebar } = await import('./UnifiedSidebar');
const { act, cleanup, render, waitFor } = await import('@testing-library/react');

type Observed = { callback: ResizeObserverCallback; targets: Set<Element> };
const observers = new Set<Observed>();
class RecordingResizeObserver {
  private readonly observed: Observed;
  constructor(callback: ResizeObserverCallback) {
    this.observed = { callback, targets: new Set() };
    observers.add(this.observed);
  }
  observe(target: Element) { this.observed.targets.add(target); }
  unobserve(target: Element) { this.observed.targets.delete(target); }
  disconnect() { this.observed.targets.clear(); observers.delete(this.observed); }
}
const originalResizeObserver = globalThis.ResizeObserver;
globalThis.ResizeObserver = RecordingResizeObserver as unknown as typeof ResizeObserver;

afterEach(cleanup);
afterAll(async () => {
  globalThis.ResizeObserver = originalResizeObserver;
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a card that grows pushes the next card down, also after a StrictMode effect replay', async () => {
  const heights = new Map([['a', 40], ['b', 40]]);
  const item = (id: string, fixedY: number): ReactSidebarItem => ({
    id, anchorPos: 0, fixedY,
    render: ({ measureRef }) => <div data-testid={id} ref={element => {
      if (element) Object.defineProperty(element, 'offsetHeight', { configurable: true, get: () => heights.get(id) });
      measureRef(element);
    }}>{id}</div>,
  });
  const view = render(<StrictMode><UnifiedSidebar items={[item('a', 100), item('b', 120)]} anchorPositions={new Map()}
    renderedDomContext={null} pageWidth={800} zoom={1} editorContainerRef={{ current: null }} /></StrictMode>);
  const top = (id: string) => Number.parseFloat(view.getByTestId(id).parentElement!.style.top);
  await waitFor(() => expect(top('b')).toBeGreaterThanOrEqual(140));
  expect(top('b')).toBeLessThan(200);
  heights.set('a', 200);
  const card = view.getByTestId('a');
  const watching = [...observers].filter(observer => observer.targets.has(card));
  expect(watching.length).toBeGreaterThan(0);
  act(() => { for (const observer of watching) observer.callback([], {} as ResizeObserver); });
  await waitFor(() => expect(top('b')).toBeGreaterThanOrEqual(300));
});

test('position updates move sidebar wrappers without rendering unchanged card content', () => {
  const renderCard = mock(({ measureRef, isExpanded, onToggleExpand }: Parameters<ReactSidebarItem['render']>[0]) =>
    <div ref={measureRef} onClick={onToggleExpand}>{String(isExpanded)}</div>);
  const items: ReactSidebarItem[] = Array.from({ length: 120 }, (_, index) => ({
    id: `revision-${index}`, anchorPos: index, anchorKey: `revision-${index}`, render: renderCard,
  }));
  const props = { renderedDomContext: null, pageWidth: 800, zoom: 1, editorContainerRef: { current: null } };
  const positions = (offset: number) => new Map(items.map((item, index) => [item.id, index * 100 + offset]));
  const view = render(<UnifiedSidebar {...props} items={items} anchorPositions={positions(100)} />);
  const renders = renderCard.mock.calls.length;
  expect(renders).toBe(120);
  const firstWrapper = () => view.container.querySelector<HTMLElement>('aside > div > div')!;
  expect(firstWrapper().style.top).toBe('100px');
  view.rerender(<UnifiedSidebar {...props} items={items.map(item => ({ ...item }))} anchorPositions={positions(300)} />);
  expect(firstWrapper().style.top).toBe('300px');
  expect(renderCard.mock.calls.length).toBe(renders);
  view.rerender(<UnifiedSidebar {...props} items={items} anchorPositions={positions(300)} activeItemId={items[0]!.id} />);
  expect(renderCard.mock.calls.length).toBe(renders + 1);
  expect(firstWrapper().textContent).toBe('true');
});
