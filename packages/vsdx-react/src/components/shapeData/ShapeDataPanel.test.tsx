import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createT, en } from '@betteroffice/vsdx-i18n';
import type { CellSnapshot, ShapeSnapshot } from '@betteroffice/vsdx';
import type { ShapeDataRow } from '@betteroffice/vsdx';
import { ShapeDataPanel } from './ShapeDataPanel';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { cleanup, fireEvent, render } = await import('@testing-library/react');
const t = createT(en);

afterEach(() => cleanup());

function cell(section: string, row: string, name: string, formula: string | null, value: string | null): CellSnapshot {
  return {
    locator: { sheet: { page: 1 }, shapeId: 1, section, row: { name: row }, cellName: name },
    name, formula, value,
  };
}

function shape(cells: CellSnapshot[]): ShapeSnapshot {
  return { id: 'page:1:shape:1', sourceId: 1, name: null, cells, children: [] };
}

function panel(shapeValue: ShapeSnapshot | null) {
  const commits: Array<{ row: ShapeDataRow; formula: string }> = [];
  const errors: unknown[] = [];
  const view = render(<ShapeDataPanel shape={shapeValue} onCommit={(row, formula) => commits.push({ row, formula })} onError={(error) => errors.push(error)} t={t} />);
  return { ...view, commits, errors };
}

test('lists visible properties as labels and values', () => {
  const view = panel(shape([
    cell('Property', 'Device', 'Label', null, 'Device name'),
    cell('Property', 'Device', 'Value', '"Amp"', 'Amp'),
    cell('Property', 'Device', 'Type', null, '0'),
    cell('Property', 'Hidden', 'Value', null, 'x'),
    cell('Property', 'Hidden', 'Invisible', null, '1'),
  ]));
  expect(view.getByText('Device name')).toBeDefined();
  expect(view.getByDisplayValue('Amp')).toBeDefined();
  expect(view.queryByDisplayValue('x')).toBe(null);
});

test('explains empty states', () => {
  expect(panel(null).getByText(t('shapeData.noSelection'))).toBeDefined();
  expect(panel(shape([])).getByText(t('shapeData.empty'))).toBeDefined();
});

test('commits text edits as quoted formulas on blur', () => {
  const view = panel(shape([
    cell('Property', 'Device', 'Label', null, 'Device name'),
    cell('Property', 'Device', 'Value', '"Amp"', 'Amp'),
  ]));
  const input = view.getByDisplayValue('Amp') as HTMLInputElement;
  fireEvent.change(input, { target: { value: 'Mixer' } });
  fireEvent.blur(input);
  expect(view.commits).toEqual([{ row: expect.objectContaining({ rowName: 'Device' }), formula: '"Mixer"' }]);
});

test('commits boolean toggles as 1 or 0', () => {
  const view = panel(shape([
    cell('Property', 'Flag', 'Label', null, 'Flag'),
    cell('Property', 'Flag', 'Value', '1', '1'),
    cell('Property', 'Flag', 'Type', null, '3'),
  ]));
  const box = view.getByRole('checkbox') as HTMLInputElement;
  expect(box.checked).toBe(true);
  fireEvent.click(box);
  expect(view.commits).toEqual([{ row: expect.objectContaining({ rowName: 'Flag' }), formula: '0' }]);
});

test('offers list options from the format cell', () => {
  const view = panel(shape([
    cell('Property', 'Use', 'Label', null, 'Use'),
    cell('Property', 'Use', 'Value', 'INDEX(0,Prop.Use.Format)', 'Conference'),
    cell('Property', 'Use', 'Type', null, '1'),
    cell('Property', 'Use', 'Format', null, 'Conference;Office'),
  ]));
  const select = view.getByDisplayValue('Conference') as HTMLSelectElement;
  fireEvent.change(select, { target: { value: 'Office' } });
  expect(view.commits).toEqual([{ row: expect.objectContaining({ rowName: 'Use' }), formula: '"Office"' }]);
});

test('reports commit failures instead of throwing', () => {
  const errors: unknown[] = [];
  const view = render(<ShapeDataPanel
    shape={shape([cell('Property', 'Device', 'Value', '"Amp"', 'Amp')])}
    onCommit={() => { throw new Error('GUARD protects the requested cell'); }}
    onError={(error) => errors.push(error)}
    t={t}
  />);
  const input = view.getByDisplayValue('Amp');
  fireEvent.change(input, { target: { value: 'Mixer' } });
  fireEvent.blur(input);
  expect(errors.length).toBe(1);
  expect(String(errors[0])).toContain('GUARD');
});
