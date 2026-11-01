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

test('missing anchor positions preserve every known card and the visible sidebar', () => {
  const items: ReactSidebarItem[] = ['a', 'b', 'c'].map(id => ({
    id, anchorPos: 0, anchorKey: id, render: ({ measureRef }) => <div data-testid={id} ref={measureRef} />,
  }));
  const props = { items, renderedDomContext: null, pageWidth: 800, zoom: 1, editorContainerRef: { current: null } };
  const positions = new Map([['a', 100], ['b', 200], ['c', 300]]);
  const view = render(<UnifiedSidebar {...props} anchorPositions={positions} />);
  const wrappers = items.map(item => view.getByTestId(item.id).parentElement!);
  const before = wrappers.map(card => ({ top: card.style.top, opacity: card.style.opacity }));
  view.rerender(<UnifiedSidebar {...props} anchorPositions={new Map()} />);
  expect(view.getByRole('complementary').style.opacity).toBe('1');
  expect(wrappers.map(card => ({ top: card.style.top, opacity: card.style.opacity }))).toEqual(before);
  expect(wrappers.every(card => card.style.visibility !== 'hidden')).toBe(true);
  view.rerender(<UnifiedSidebar {...props} anchorPositions={positions} />);
  expect(wrappers.map(card => ({ top: card.style.top, opacity: card.style.opacity }))).toEqual(before);
});

test('edit-driven anchor moves skip transitions while expansion and collapse still animate collisions', async () => {
  const heights = new Map([['a', 40], ['b', 40]]);
  const items: ReactSidebarItem[] = ['a', 'b'].map(id => ({
    id, anchorPos: 0, anchorKey: id,
    render: ({ measureRef, isExpanded }) => <div data-testid={id} ref={measureRef}>{String(isExpanded)}</div>,
  }));
  const props = { items, renderedDomContext: null, pageWidth: 800, zoom: 1, editorContainerRef: { current: null } };
  const positions = (offset: number) => new Map([['a', 100 + offset], ['b', 120 + offset]]);
  const anchors = positions(0);
  const view = render(<UnifiedSidebar {...props} anchorPositions={anchors} />);
  for (const item of items) Object.defineProperty(view.getByTestId(item.id), 'offsetHeight', {
    configurable: true, get: () => heights.get(item.id),
  });
  const wrapper = (id: string) => view.getByTestId(id).parentElement!;
  await waitFor(() => expect(wrapper('b').style.transition).toContain('top 0.15s ease'));
  const moved = positions(200);
  view.rerender(<UnifiedSidebar {...props} anchorPositions={moved} />);
  expect(wrapper('a').style.top).toBe('300px');
  for (const item of items) {
    expect(wrapper(item.id).style.transition).toBe('none');
    expect(wrapper(item.id).style.opacity).toBe('1');
  }
  const collapsedTop = Number.parseFloat(wrapper('b').style.top);
  heights.set('a', 200);
  view.rerender(<UnifiedSidebar {...props} anchorPositions={moved} activeItemId="a" />);
  await waitFor(() => expect(Number.parseFloat(wrapper('b').style.top)).toBeGreaterThanOrEqual(500));
  expect(wrapper('b').style.transition).toContain('top 0.15s ease');
  heights.set('a', 40);
  view.rerender(<UnifiedSidebar {...props} anchorPositions={moved} activeItemId={null} />);
  await waitFor(() => expect(Number.parseFloat(wrapper('b').style.top)).toBe(collapsedTop));
  expect(wrapper('b').style.transition).toContain('top 0.15s ease');
});

test('first appearances and newly added cards retain their opacity transition', () => {
  const item = (id: string, fixedY: number): ReactSidebarItem => ({
    id, anchorPos: 0, fixedY, render: ({ measureRef }) => <div data-testid={id} ref={measureRef} />,
  });
  const first = item('a', 100);
  const props = { anchorPositions: new Map<string, number>(), renderedDomContext: null, pageWidth: 800, zoom: 1, editorContainerRef: { current: null } };
  const view = render(<UnifiedSidebar {...props} items={[first]} />);
  expect(view.getByRole('complementary').style.transition).toBe('opacity 0.15s ease');
  expect(view.getByTestId('a').parentElement!.style.transition).toBe('opacity 0.2s ease');
  view.rerender(<UnifiedSidebar {...props} items={[first, item('b', 300)]} />);
  expect(view.getByTestId('b').parentElement!.style.transition).toBe('opacity 0.2s ease');
});
