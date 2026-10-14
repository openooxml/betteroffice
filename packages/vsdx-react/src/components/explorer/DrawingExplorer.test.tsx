import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createT, en } from '@betteroffice/vsdx-i18n';
import type { CellSnapshot, DiagramSnapshot, ShapeSnapshot } from '@betteroffice/vsdx';
import type { VsdxShapeSelection } from '../../VsdxEditor';
import { DrawingExplorer, MAX_EXPLORER_DEPTH } from './DrawingExplorer';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { cleanup, fireEvent, render } = await import('@testing-library/react');
const t = createT(en);

afterEach(() => cleanup());

type RowRef = { index: number } | { name: string };

function shapeCell(shapeId: number, name: string, formula: string | null, value: string | null): CellSnapshot {
  return { locator: { sheet: { page: 1 }, shapeId, section: null, row: null, cellName: name }, name, formula, value };
}

function sectionCell(shapeId: number, section: string, row: RowRef, rowType: string | undefined, cellName: string, formula: string | null, value: string | null): CellSnapshot {
  const cell: CellSnapshot = { locator: { sheet: { page: 1 }, shapeId, section, row, cellName }, name: cellName, formula, value };
  if (rowType !== undefined) cell.rowType = rowType;
  return cell;
}

function shape(id: string, sourceId: number, name: string | null, cells: CellSnapshot[], children: ShapeSnapshot[] = []): ShapeSnapshot {
  return { id, sourceId, name, cells, children };
}

function snapshot(): DiagramSnapshot {
  const group = shape('g1', 5, 'Alpha', [
    shapeCell(5, 'PinX', '2', '2'),
    shapeCell(5, 'Width', '4', '4'),
    sectionCell(5, 'Geometry', { index: 0 }, 'MoveTo', 'X', '0', '0'),
    sectionCell(5, 'Geometry', { index: 0 }, 'MoveTo', 'Y', '0', '0'),
    sectionCell(5, 'Geometry', { index: 1 }, 'LineTo', 'X', 'Width*0.5', '2'),
    sectionCell(5, 'Geometry', { index: 1 }, 'LineTo', 'Y', '1', '1'),
    sectionCell(5, 'User', { name: 'visVersion' }, undefined, 'Value', null, '15'),
  ], [
    shape('c1', 6, 'Beta', [shapeCell(6, 'PinX', '3', '3')]),
    shape('c2', 7, 'Gamma', []),
  ]);
  const connector = shape('k1', 12, 'Link', [
    shapeCell(12, 'OneD', null, '1'),
    shapeCell(12, 'BeginX', null, '0'),
    shapeCell(12, 'EndX', null, '1'),
    sectionCell(12, 'Geometry', { index: 0 }, 'MoveTo', 'X', '0', '0'),
  ]);
  return {
    pages: [
      { id: 'page:1', sourcePartPath: 'visio/pages/page1.xml', name: 'Page-1', shapes: [group, connector] },
      { id: 'page:2', sourcePartPath: 'visio/pages/page2.xml', name: null, shapes: [] },
    ],
  };
}

interface Harness {
  selections: VsdxShapeSelection[];
  pages: number[];
}

function harness() {
  const state: Harness = { selections: [], pages: [] };
  const view = render(
    <DrawingExplorer
      snapshot={snapshot()}
      activePageIndex={0}
      selection={null}
      onSelectPage={(index) => state.pages.push(index)}
      onSelectShape={(selection) => { if (selection) state.selections.push(selection); }}
      collapsed={false}
      onToggleCollapsed={() => {}}
      t={t}
    />,
  );
  return { ...view, state };
}

function button(view: ReturnType<typeof harness>, name: string): HTMLButtonElement {
  return view.getByRole('button', { name }) as HTMLButtonElement;
}

test('nests pages, shapes and group children with real ids and names', () => {
  const view = harness();
  expect(button(view, 'Document')).toBeDefined();
  expect(button(view, 'Pages')).toBeDefined();
  expect(button(view, 'Page-1')).toBeDefined();
  expect(button(view, 'Shapes')).toBeDefined();
  expect(button(view, 'Shape 5 "Alpha" (Group)')).toBeDefined();
  expect(button(view, 'Shape 12 "Link" (1-D)')).toBeDefined();
  expect(view.queryByRole('button', { name: 'Shape 6 "Beta"' })).toBeNull();
  fireEvent.click(button(view, 'Shape 5 "Alpha" (Group)'));
  expect(button(view, 'Shape 6 "Beta"')).toBeDefined();
  expect(button(view, 'Shape 7 "Gamma"')).toBeDefined();
});

test('renders cells lazily with formula and value side by side', () => {
  const view = harness();
  expect(view.queryByText('Geometry')).toBeNull();
  expect(view.queryByText('Width*0.5')).toBeNull();
  fireEvent.click(button(view, 'Shape 5 "Alpha" (Group)'));
  expect(button(view, 'Geometry')).toBeDefined();
  expect(button(view, 'User')).toBeDefined();
  expect(button(view, 'Shape')).toBeDefined();
  expect(view.queryByText('Width*0.5')).toBeNull();
  fireEvent.click(button(view, 'Geometry'));
  expect(view.getByText('Row 1 (MoveTo)')).toBeDefined();
  expect(view.getByText('Row 2 (LineTo)')).toBeDefined();
  expect(view.getByText('Formula')).toBeDefined();
  expect(view.getByText('Value')).toBeDefined();
  expect(view.getByText('Width*0.5')).toBeDefined();
  expect(view.getAllByText('1').length).toBeGreaterThan(1);
  fireEvent.click(button(view, 'User'));
  expect(view.getByText('visVersion')).toBeDefined();
  expect(view.getByText('15')).toBeDefined();
  fireEvent.click(button(view, 'Shape'));
  expect(view.getByText('PinX')).toBeDefined();
});

test('keeps sections with the same name on different shapes independent', () => {
  const view = harness();
  fireEvent.click(button(view, 'Shape 5 "Alpha" (Group)'));
  fireEvent.click(button(view, 'Shape 12 "Link" (1-D)'));
  const sections = view.getAllByRole('button', { name: 'Geometry' });
  expect(sections).toHaveLength(2);
  fireEvent.click(sections[0] as HTMLButtonElement);
  expect(sections[0]?.getAttribute('aria-expanded')).toBe('true');
  expect(sections[1]?.getAttribute('aria-expanded')).toBe('false');
});

test('selecting a node reuses the editor selection model', () => {
  const view = harness();
  fireEvent.click(button(view, 'Shape 5 "Alpha" (Group)'));
  fireEvent.click(button(view, 'Shape 6 "Beta"'));
  expect(view.state.selections).toEqual([
    { pageId: 'page:1', shapeId: 'g1', hit: { kind: 'shape', shapeId: 'g1' } },
    { pageId: 'page:1', shapeId: 'c1', hit: { kind: 'shape', shapeId: 'c1' } },
  ]);
  expect(view.state.pages).toEqual([]);
});

test('a canvas selection reveals and highlights its node', () => {
  const view = render(
    <DrawingExplorer
      snapshot={snapshot()}
      activePageIndex={0}
      selection={{ pageId: 'page:1', shapeId: 'c1', hit: { kind: 'shape', shapeId: 'c1' } }}
      onSelectPage={() => {}}
      onSelectShape={() => {}}
      collapsed={false}
      onToggleCollapsed={() => {}}
      t={t}
    />,
  );
  const node = view.getByRole('button', { name: 'Shape 6 "Beta"' });
  expect(node.closest('[role="treeitem"]')?.getAttribute('aria-selected')).toBe('true');
  expect(view.getByRole('button', { name: 'Shape 5 "Alpha" (Group)' }).closest('[role="treeitem"]')?.getAttribute('aria-selected')).toBe('false');
});

test('navigates pages and falls back for unnamed or empty pages', () => {
  const view = harness();
  fireEvent.click(button(view, 'Page 2'));
  expect(view.state.pages).toEqual([1]);
  const shapesNodes = view.getAllByRole('button', { name: 'Shapes' });
  expect(shapesNodes).toHaveLength(2);
  fireEvent.click(shapesNodes[1] as HTMLButtonElement);
  expect(view.getByText(t('explorer.emptyShapes'))).toBeDefined();
});

test('reflects an active page change in the tree', () => {
  const props = {
    snapshot: snapshot(),
    selection: null,
    onSelectPage: () => {},
    onSelectShape: () => {},
    collapsed: false,
    onToggleCollapsed: () => {},
    t,
  };
  const view = render(<DrawingExplorer {...props} activePageIndex={0} />);
  view.rerender(<DrawingExplorer {...props} activePageIndex={1} />);
  const page = view.getByRole('button', { name: 'Page 2' });
  expect(page.closest('[role="treeitem"]')?.getAttribute('aria-selected')).toBe('true');
  expect(page.getAttribute('aria-expanded')).toBe('true');
  expect(view.getAllByRole('button', { name: 'Shapes' })).toHaveLength(1);
});

test('does not reopen a selected shape after the user collapses it', () => {
  const props = {
    snapshot: snapshot(),
    activePageIndex: 0,
    onSelectPage: () => {},
    onSelectShape: () => {},
    collapsed: false,
    onToggleCollapsed: () => {},
    t,
  };
  const selection = { pageId: 'page:1', shapeId: 'g1', hit: { kind: 'shape' as const, shapeId: 'g1' } };
  const view = render(<DrawingExplorer {...props} selection={selection} />);
  const node = view.getByRole('button', { name: 'Shape 5 "Alpha" (Group)' });
  fireEvent.click(node);
  fireEvent.click(node);
  view.rerender(<DrawingExplorer {...props} selection={{ ...selection }} />);
  expect(node.getAttribute('aria-expanded')).toBe('false');
});

test('bounds the depth of a hostile document', () => {
  let child: ShapeSnapshot = shape('deep-leaf', 1000, null, []);
  for (let level = MAX_EXPLORER_DEPTH + 15; level >= 0; level -= 1) {
    child = shape(`deep-${level}`, level, null, [shapeCell(level, 'PinX', '1', '1')], [child]);
  }
  const deep: DiagramSnapshot = { pages: [{ id: 'page:1', sourcePartPath: 'page', name: 'Page-1', shapes: [child] }] };
  const view = render(
    <DrawingExplorer
      snapshot={deep}
      activePageIndex={0}
      selection={null}
      onSelectPage={() => {}}
      onSelectShape={() => {}}
      collapsed={false}
      onToggleCollapsed={() => {}}
      t={t}
    />,
  );
  for (let level = 0; level < MAX_EXPLORER_DEPTH; level += 1) {
    const node = view.container.querySelector(`[data-explorer-key="shape:page:1:deep-${level}"]`);
    expect(node?.textContent).toBe(`Shape ${level} (Group)`);
    fireEvent.click(node as HTMLButtonElement);
  }
  expect(view.getByText(t('explorer.depthLimit'))).toBeDefined();
  expect(view.container.querySelectorAll('[data-explorer-key^="shape:"]')).toHaveLength(MAX_EXPLORER_DEPTH);
});

test('keeps a collapsed branch collapsed when a new snapshot arrives for the same selection', () => {
  const props = {
    activePageIndex: 0,
    selection: { pageId: 'page:1', shapeId: 'c1', hit: { kind: 'shape' as const, shapeId: 'c1' } },
    onSelectPage: () => {},
    onSelectShape: () => {},
    collapsed: false,
    onToggleCollapsed: () => {},
    t,
  };
  const view = render(<DrawingExplorer {...props} snapshot={snapshot()} />);
  const shapes = view.getByRole('button', { name: 'Shapes' });
  fireEvent.click(shapes);
  expect(shapes.getAttribute('aria-expanded')).toBe('false');
  view.rerender(<DrawingExplorer {...props} snapshot={snapshot()} />);
  expect(shapes.getAttribute('aria-expanded')).toBe('false');
});
