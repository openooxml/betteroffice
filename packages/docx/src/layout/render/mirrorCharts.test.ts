import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { ChartData, DisplayPage, DisplayPrimitive } from './displayList';
import { buildMirrorPage, buildMirrorPageText, updateMirrorPage } from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const label = 'Sales, column chart, 2 series, 2 categories';
const data: ChartData = {
  categories: ['Q1', 'Q2'],
  series: [{ name: 'North', values: ['12', null] }, { values: ['8', '14'] }],
};

/** A chart block: a background rect carrying the data, then an axis label. */
function page(chartData: unknown, extra: object = {}): DisplayPage {
  const chart = { label, frame: { x: 0, y: 0, w: 200, h: 100 } };
  return {
    pageIndex: 0,
    width: 300,
    height: 200,
    primitives: [
      { kind: 'rect', x: 0, y: 0, w: 200, h: 100, fill: '#fff', blockId: 'c', chart: { ...chart, data: chartData }, ...extra },
      { kind: 'text', x: 10, y: 90, baselineY: 95, width: 20, text: 'Q1', font: '10px sans-serif', color: '#000', blockId: 'c', chart, ...extra },
    ] as unknown as DisplayPrimitive[],
  };
}

const figure = (root: HTMLElement) => root.querySelector('[role="figure"]') as HTMLElement | null;
const rows = (root: HTMLElement) =>
  Array.from(root.querySelectorAll('.layout-chart-row'), (row) =>
    Array.from(row.children, (cell) => [cell.getAttribute('role'), cell.textContent]));

test('a chart with data reads as a figure holding its values as a table', () => {
  const mirror = buildMirrorPage(page(data));
  const chart = figure(mirror)!;
  expect(chart.getAttribute('aria-label')).toBe(label);
  const table = chart.querySelector('[role="table"]')!;
  expect([table.getAttribute('aria-rowcount'), table.getAttribute('aria-colcount')]).toEqual(['3', '3']);
  expect(rows(mirror)).toEqual([
    [['cell', ''], ['columnheader', 'North'], ['cell', '']],
    [['rowheader', 'Q1'], ['cell', '12'], ['cell', '8']],
    [['rowheader', 'Q2'], ['cell', ''], ['cell', '14']],
  ]);
  const painted = Array.from(chart.children).filter((child) => child !== table);
  expect(painted.length).toBeGreaterThan(0);
  expect(painted.every((child) => child.getAttribute('aria-hidden') === 'true')).toBe(true);
});

test('the text-only mirror keeps the chart table and drops the painted labels', () => {
  const mirror = buildMirrorPageText(page(data));
  expect(rows(mirror)).toEqual(rows(buildMirrorPage(page(data))));
  expect(figure(mirror)!.textContent).toBe('NorthQ1128Q214');
});

test('a chart without data, or with data that does not fit, keeps its summary image', () => {
  for (const broken of [undefined, { categories: [], series: [{ values: [] }] }, { categories: ['Q1'], series: [{ values: ['1', '2'] }] },
    { categories: ['Q1'], series: [{ values: [3] }] }, { categories: [5], series: [{ values: ['1'] }] }]) {
    const mirror = buildMirrorPage(page(broken));
    expect(figure(mirror)).toBeNull();
    expect(mirror.querySelector('[role="img"]')?.getAttribute('aria-label')).toBe(label);
    expect(mirror.querySelector('[role="table"]')).toBeNull();
    expect(mirror.querySelector('[role="img"]')!.hasAttribute('aria-description')).toBe(false);
  }
});

test('a chart that caches no points keeps its name and says it has no data in the given language', () => {
  const empty = { categories: [], series: [] };
  const mirror = buildMirrorPage(page(empty), { labels: { chartNoData: 'Keine Daten' } });
  const chart = mirror.querySelector('[role="img"]')!;
  expect(chart.getAttribute('aria-label')).toBe(label);
  expect(chart.getAttribute('aria-description')).toBe('Keine Daten');
  expect(mirror.querySelector('[role="table"]')).toBeNull();
  const described = buildMirrorPage(page(empty, { ariaDescription: 'Survey results' }), { labels: { chartNoData: 'No data' } });
  expect(described.querySelector('[role="img"]')!.getAttribute('aria-description')).toBe('Survey results. No data');
  expect(buildMirrorPage(page(empty)).querySelector('[role="img"]')!.hasAttribute('aria-description')).toBe(false);
});

test('updating a mirror in another language redescribes a chart that caches no points', () => {
  const empty = page({ categories: [], series: [] });
  const mirror = buildMirrorPage(empty, { labels: { chartNoData: 'No data' } });
  expect(updateMirrorPage(mirror, empty, { labels: { chartNoData: 'Keine Daten' } })).toBe(true);
  expect(mirror.querySelector('[role="img"]')!.getAttribute('aria-description')).toBe('Keine Daten');
});

test('the painted parts of a chart carry no name of their own', () => {
  for (const chartData of [data, undefined]) {
    const mirror = buildMirrorPage(page(chartData, { ariaLabel: 'Chart title', ariaDescription: 'Alt text' }));
    const block = mirror.querySelector('.layout-run-chart')!;
    expect(block.getAttribute('aria-label')).toBe('Chart title');
    const painted = Array.from(block.children).filter((child) => !child.matches('.layout-chart-table'));
    expect(painted.length).toBeGreaterThan(0);
    for (const child of painted) {
      expect(child.getAttribute('aria-hidden')).toBe('true');
      expect(child.hasAttribute('aria-label')).toBe(false);
      expect(child.hasAttribute('aria-description')).toBe(false);
    }
  }
});

test('a decorative chart stays hidden with its table', () => {
  const mirror = buildMirrorPage(page(data, { decorative: true }));
  const block = mirror.querySelector('.layout-run-chart')!;
  expect(block.getAttribute('aria-hidden')).toBe('true');
  expect(block.hasAttribute('role')).toBe(false);
});
