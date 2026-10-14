import { expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createT, en } from '@betteroffice/vsdx-i18n';
import { useState } from 'react';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { cleanup, fireEvent, render } = await import('@testing-library/react');
const { StatusBar, zoomForSliderPosition } = await import('./StatusBar');
const t = createT(en);
const pages = [
  { id: 'one', name: 'Cover' },
  { id: 'two', name: null },
  { id: 'three', name: 'Details' },
];

function renderStatusBar(overrides: Partial<Parameters<typeof StatusBar>[0]> = {}) {
  const selected: number[] = [];
  const zooms: number[] = [];
  function Host() {
    const [zoom, setZoom] = useState(overrides.zoom ?? 1);
    return (
      <StatusBar
        {...overrides}
        pages={overrides.pages ?? pages}
        activeIndex={overrides.activeIndex ?? 1}
        onSelectPage={(index) => selected.push(index)}
        zoom={zoom}
        onZoomChange={(nextZoom) => {
          zooms.push(nextZoom);
          setZoom(nextZoom);
        }}
        onFitToWindow={() => {}}
        t={t}
      />
    );
  }
  const view = render(
    <Host />,
  );
  return { ...view, selected, zooms };
}

test('renders active page tabs and selects the clicked page', () => {
  const view = renderStatusBar();
  const tabs = view.getAllByRole('tab');
  expect(tabs[1].getAttribute('aria-selected')).toBe('true');
  fireEvent.click(tabs[2]);
  expect(view.selected).toEqual([2]);
  cleanup();
});

test('moves between page tabs with arrow keys', () => {
  const view = renderStatusBar();
  const tabs = view.getAllByRole('tab');
  tabs[1].focus();
  fireEvent.keyDown(tabs[1], { key: 'ArrowRight' });
  fireEvent.keyDown(tabs[1], { key: 'ArrowLeft' });
  expect(view.selected).toEqual([2, 0]);
  cleanup();
});

test('disables previous on the first page and next on the last page', () => {
  const first = renderStatusBar({ activeIndex: 0 });
  expect(first.getByRole('button', { name: 'Previous page' }).hasAttribute('disabled')).toBe(true);
  expect(first.getByRole('button', { name: 'Next page' }).hasAttribute('disabled')).toBe(false);
  cleanup();
  const last = renderStatusBar({ activeIndex: 2 });
  expect(last.getByRole('button', { name: 'Previous page' }).hasAttribute('disabled')).toBe(false);
  expect(last.getByRole('button', { name: 'Next page' }).hasAttribute('disabled')).toBe(true);
  cleanup();
});

test('uses the existing numbered fallback title for unnamed pages', () => {
  const view = renderStatusBar();
  expect(view.getByRole('tab', { name: 'Page 2' })).toBeDefined();
  cleanup();
});

test('walks conventional zoom stops instead of adding a fixed delta', () => {
  const view = renderStatusBar({ zoom: 0.5 });
  const zoomIn = view.getByRole('button', { name: 'Zoom in' });
  const zoomOut = view.getByRole('button', { name: 'Zoom out' });
  fireEvent.click(zoomIn);
  fireEvent.click(zoomIn);
  fireEvent.click(zoomIn);
  fireEvent.click(zoomOut);
  fireEvent.click(zoomOut);
  expect(view.zooms).toEqual([0.75, 1, 1.5, 1, 0.75]);
  cleanup();
});

test('clamps invalid zoom values and both zoom endpoints', () => {
  const low = renderStatusBar({ zoom: 0.1 });
  fireEvent.click(low.getByRole('button', { name: 'Zoom out' }));
  expect(low.zooms).toEqual([0.1]);
  cleanup();
  const high = renderStatusBar({ zoom: -1 });
  fireEvent.click(high.getByRole('button', { name: 'Zoom in' }));
  expect(high.zooms).toEqual([1.5]);
  cleanup();
  const max = renderStatusBar({ zoom: 4 });
  fireEvent.click(max.getByRole('button', { name: 'Zoom in' }));
  expect(max.zooms).toEqual([4]);
  expect(max.zooms.every((zoom) => Number.isFinite(zoom) && zoom > 0)).toBe(true);
  cleanup();
});

test('resets zoom to 100% from the percentage readout', () => {
  const view = renderStatusBar({ zoom: 1.5 });
  fireEvent.click(view.getByRole('button', { name: 'Reset zoom to 100%' }));
  expect(view.zooms).toEqual([1]);
  cleanup();
});

test('maps the slider midpoint to 100% with a logarithmic curve', () => {
  const view = renderStatusBar({ zoom: 0.5 });
  const slider = view.getByRole('slider', { name: 'Zoom level' });
  fireEvent.change(slider, { target: { value: '500' } });
  expect(view.zooms).toEqual([1]);
  expect(zoomForSliderPosition(0.5)).toBe(1);
  cleanup();
});

test('does not render an add page control without an add-page operation', () => {
  const view = renderStatusBar();
  expect(view.queryByRole('button', { name: 'Add page' })).toBeNull();
  cleanup();
});

test('reorders the page the drag started on after the page list shifts underneath', () => {
  const reorders: Array<[string, number]> = [];
  function Host({ items }: { items: typeof pages }) {
    return <StatusBar pages={items} activeIndex={0} onSelectPage={() => {}} onReorderPage={(pageId, toIndex) => reorders.push([pageId, toIndex])} zoom={1} onZoomChange={() => {}} onFitToWindow={() => {}} t={t} />;
  }
  const view = render(<Host items={pages} />);
  fireEvent.dragStart(view.getAllByRole('tab')[2]);
  view.rerender(<Host items={[pages[1], pages[2]]} />);
  fireEvent.drop(view.getAllByRole('tab')[0]);
  expect(reorders).toEqual([['three', 0]]);
  cleanup();
});

test('ignores a drop whose dragged page a peer removed', () => {
  const reorders: Array<[string, number]> = [];
  function Host({ items }: { items: typeof pages }) {
    return <StatusBar pages={items} activeIndex={0} onSelectPage={() => {}} onReorderPage={(pageId, toIndex) => reorders.push([pageId, toIndex])} zoom={1} onZoomChange={() => {}} onFitToWindow={() => {}} t={t} />;
  }
  const view = render(<Host items={pages} />);
  fireEvent.dragStart(view.getAllByRole('tab')[2]);
  view.rerender(<Host items={[pages[0], pages[1]]} />);
  fireEvent.drop(view.getAllByRole('tab')[0]);
  expect(reorders).toEqual([]);
  cleanup();
});
