import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { TableGridInline } from './TableGridInline';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const { act, cleanup, fireEvent, render } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('provides rows and a single tab stop and clamps arrows at every grid edge', () => {
  const view = render(<TableGridInline gridRows={2} gridColumns={3} onInsert={() => {}} />);
  const grid = view.getByRole('grid');
  expect(view.getAllByRole('row').length).toBe(2);
  const cells = view.getAllByRole('gridcell');
  expect(cells.every((cell) => cell.parentElement?.getAttribute('role') === 'row')).toBe(true);
  expect(grid.querySelectorAll('[tabindex="0"]').length).toBe(1);
  act(() => cells[0].focus());
  for (const key of ['ArrowLeft', 'ArrowUp']) {
    fireEvent.keyDown(document.activeElement!, { key });
    expect(document.activeElement === cells[0]).toBe(true);
  }
  for (const key of ['ArrowRight', 'ArrowRight', 'ArrowRight', 'ArrowDown', 'ArrowDown']) {
    fireEvent.keyDown(document.activeElement!, { key });
  }
  expect(document.activeElement === cells[5]).toBe(true);
  expect(grid.querySelectorAll('[tabindex="0"]').length).toBe(1);
  expect(view.getByText('3 × 2')).not.toBeNull();
});

test('Home and End navigate within the row and Control navigates the whole grid', () => {
  const view = render(<TableGridInline gridRows={3} gridColumns={4} onInsert={() => {}} />);
  const cells = view.getAllByRole('gridcell');
  act(() => cells[5].focus());
  fireEvent.keyDown(cells[5], { key: 'End' });
  expect(document.activeElement === cells[7]).toBe(true);
  fireEvent.keyDown(cells[7], { key: 'Home' });
  expect(document.activeElement === cells[4]).toBe(true);
  fireEvent.keyDown(cells[4], { key: 'End', ctrlKey: true });
  expect(document.activeElement === cells[11]).toBe(true);
  fireEvent.keyDown(cells[11], { key: 'Home', ctrlKey: true });
  expect(document.activeElement === cells[0]).toBe(true);
});

test.each(['Enter', ' '])('inserts the focused dimensions with %s', (key) => {
  const inserted: number[][] = [];
  const view = render(
    <TableGridInline
      gridRows={2}
      gridColumns={3}
      onInsert={(rows, cols) => inserted.push([rows, cols])}
    />
  );
  const cell = view.getByRole('gridcell', { name: '3 columns, 2 rows' });
  act(() => cell.focus());
  fireEvent.keyDown(cell, { key });
  expect(inserted).toEqual([[2, 3]]);
});

test('mouse hover previews and clicking inserts the hovered dimensions', () => {
  const inserted: number[][] = [];
  const view = render(<TableGridInline onInsert={(rows, cols) => inserted.push([rows, cols])} />);
  const cell = view.getByRole('gridcell', { name: '4 columns, 3 rows' });
  fireEvent.mouseEnter(cell);
  expect(view.getByText('4 × 3')).not.toBeNull();
  expect(view.getAllByRole('gridcell', { selected: true }).length).toBe(12);
  fireEvent.click(cell);
  expect(inserted).toEqual([[3, 4]]);
});
