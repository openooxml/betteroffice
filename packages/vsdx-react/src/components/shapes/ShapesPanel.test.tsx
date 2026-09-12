import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createT, en } from '@betteroffice/vsdx-i18n';
import { standardShapes } from './shapeLibrary';
import type { StandardShape } from './shapeLibrary';
import { ShapesPanel } from './ShapesPanel';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { cleanup, fireEvent, render } = await import('@testing-library/react');
const t = createT(en);

afterEach(() => cleanup());

function panel(overrides: { collapsed?: boolean; shapes?: readonly StandardShape[] } = {}) {
  const inserted: string[] = [];
  const toggles: boolean[] = [];
  const view = render(<ShapesPanel shapes={standardShapes} collapsed={false} onToggleCollapsed={() => toggles.push(true)} onInsert={(shape) => inserted.push(shape.id)} t={t} {...overrides} />);
  return { ...view, inserted, toggles };
}

function tiles(view: ReturnType<typeof panel>) {
  return view.getAllByRole('gridcell').map((cell) => cell.querySelector('button') as HTMLButtonElement);
}

test('renders every exported standard shape', () => {
  const view = panel();
  expect(view.getAllByRole('gridcell')).toHaveLength(standardShapes.length);
  for (const shape of standardShapes) expect(view.getByRole('button', { name: t(shape.nameKey) })).toBeDefined();
});

test('wraps grid cells in rows so the grid keeps a valid ARIA structure', () => {
  const view = panel();
  const rows = view.getAllByRole('row');
  expect(rows).toHaveLength(Math.ceil(standardShapes.length / 3));
  for (const row of rows) expect(row.parentElement?.getAttribute('role')).toBe('grid');
});

test('exposes a single tab stop and moves it with the roving focus', () => {
  const view = panel();
  const buttons = tiles(view);
  expect(buttons.filter((button) => button.tabIndex === 0)).toHaveLength(1);
  expect(buttons[0].tabIndex).toBe(0);
  buttons[0].focus();
  fireEvent.keyDown(buttons[0], { key: 'ArrowRight' });
  expect(tiles(view).filter((button) => button.tabIndex === 0)).toHaveLength(1);
  expect(tiles(view)[1].tabIndex).toBe(0);
});

test('filters case-insensitively and shows an empty state', () => {
  const view = panel();
  const search = view.getByRole('searchbox', { name: t('shapesPanel.searchLabel') });
  fireEvent.change(search, { target: { value: 'HEXAGON' } });
  expect(view.getAllByRole('gridcell')).toHaveLength(1);
  expect(view.getByRole('button', { name: 'Hexagon' })).toBeDefined();
  fireEvent.change(search, { target: { value: 'not a shape' } });
  expect(view.queryAllByRole('gridcell')).toHaveLength(0);
  expect(view.getByText(t('shapesPanel.empty'))).toBeDefined();
});

test('inserts the selected shape once per activation', () => {
  const view = panel();
  const rectangle = view.getByRole('button', { name: 'Rectangle' });
  fireEvent.click(rectangle);
  fireEvent.click(rectangle);
  expect(view.inserted).toEqual(['rectangle', 'rectangle']);
});

test('moves focus through the grid with arrow keys', () => {
  const view = panel();
  const [rectangle, square, , ellipse, rightTriangle] = tiles(view);
  rectangle.focus();
  fireEvent.keyDown(rectangle, { key: 'ArrowRight' });
  expect(document.activeElement).toBe(square);
  fireEvent.keyDown(square, { key: 'ArrowDown' });
  expect(document.activeElement).toBe(rightTriangle);
  fireEvent.keyDown(rightTriangle, { key: 'ArrowLeft' });
  expect(document.activeElement).toBe(ellipse);
});

test('collapsing hides the gallery and exposes the collapsed state', () => {
  const view = panel({ collapsed: true });
  expect(view.queryByRole('grid')).toBeNull();
  const expand = view.getByRole('button', { name: t('shapesPanel.expand') });
  expect(expand.getAttribute('aria-expanded')).toBe('false');
  fireEvent.click(expand);
  expect(view.toggles).toEqual([true]);
});
