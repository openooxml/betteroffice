import { expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createT, en } from '@betteroffice/vsdx-i18n';
import type { ShapeSnapshot } from '@betteroffice/vsdx';
import type { BindOutcome, ImportedTable } from './bindTable';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { cleanup, fireEvent, render } = await import('@testing-library/react');
const { DataBindingPanel } = await import('./DataBindingPanel');

const t = createT(en);

function cell(name: string, label: string, value: string) {
  return { locator: { sheet: { page: 1 } as const, shapeId: 1, section: 'Property', row: { name }, cellName: label }, name: label, formula: value, value };
}

const shape: ShapeSnapshot = {
  id: 'shape', sourceId: 1, name: null, children: [],
  cells: [cell('Device', 'Label', 'Device'), cell('Device', 'Value', '"Amp"')],
};

const table: ImportedTable = { headers: ['Device'], rows: [['Mixer'], ['Desk']] };

test('it asks for a table before anything else', () => {
  const view = render(<DataBindingPanel table={null} shape={shape} onImport={() => {}} onBind={() => {}} t={t} />);
  try {
    expect(view.getByText(en.dataBinding.noTable)).toBeDefined();
  } finally { cleanup(); }
});

test('it asks for a selection once a table is loaded', () => {
  const view = render(<DataBindingPanel table={table} shape={null} onImport={() => {}} onBind={() => {}} t={t} />);
  try {
    expect(view.getByText(en.dataBinding.noSelection)).toBeDefined();
  } finally { cleanup(); }
});

test('it says so when no column matches, rather than offering a bind that does nothing', () => {
  const view = render(<DataBindingPanel table={{ headers: ['Nothing'], rows: [['x']] }} shape={shape} onImport={() => {}} onBind={() => {}} t={t} />);
  try {
    expect(view.getByText(en.dataBinding.unmatched)).toBeDefined();
    expect(view.queryByText(en.dataBinding.bind)).toBeNull();
  } finally { cleanup(); }
});

test('binding a row reports how many rows it wrote', () => {
  const bound: number[] = [];
  const outcome: BindOutcome = { receipts: [{ pageId: 'p', shapeId: 's', rowName: 'Device', rowIndex: null, sectionIndex: null, before: null, after: '"Mixer"', refusal: null }], refusals: [], applied: true };
  const view = render(<DataBindingPanel table={table} shape={shape} onImport={() => {}} onBind={(index) => { bound.push(index); return outcome; }} t={t} />);
  try {
    fireEvent.click(view.getAllByText(en.dataBinding.bind)[1]);
    expect(bound).toEqual([1]);
    expect(view.getByRole('status').textContent).toBe('Bound 1 rows.');
  } finally { cleanup(); }
});

test('a refusal is announced as an alert and names the rows', () => {
  const refusal = { pageId: 'p', shapeId: 's', rowName: 'Due', rowIndex: null, sectionIndex: null, before: null, after: null, refusal: 'a date shape-data row keeps its typed value' };
  const outcome: BindOutcome = { receipts: [refusal], refusals: [refusal], applied: false };
  const view = render(<DataBindingPanel table={table} shape={shape} onImport={() => {}} onBind={() => outcome} t={t} />);
  try {
    fireEvent.click(view.getAllByText(en.dataBinding.bind)[0]);
    const alert = view.getByRole('alert');
    expect(alert.textContent).toContain('Due');
    expect(alert.textContent).toContain('Nothing was written');
  } finally { cleanup(); }
});

test('an empty bind says nothing was written rather than reporting success', () => {
  const outcome: BindOutcome = { receipts: [], refusals: [], applied: false };
  const view = render(<DataBindingPanel table={table} shape={shape} onImport={() => {}} onBind={() => outcome} t={t} />);
  try {
    fireEvent.click(view.getAllByText(en.dataBinding.bind)[0]);
    expect(view.getByRole('status').textContent).toBe(en.dataBinding.nothingToBind);
  } finally { cleanup(); }
});

test('the bind result does not follow the user to another shape', () => {
  const outcome: BindOutcome = { receipts: [{ pageId: 'p', shapeId: 's', rowName: 'Device', rowIndex: null, sectionIndex: null, before: null, after: '"Mixer"', refusal: null }], refusals: [], applied: true };
  const view = render(<DataBindingPanel table={table} shape={shape} onImport={() => {}} onBind={() => outcome} t={t} />);
  try {
    fireEvent.click(view.getAllByText(en.dataBinding.bind)[0]);
    expect(view.queryByRole('status')).not.toBeNull();
    view.rerender(<DataBindingPanel table={table} shape={{ ...shape, id: 'other' }} onImport={() => {}} onBind={() => outcome} t={t} />);
    expect(view.queryByRole('status')).toBeNull();
  } finally { cleanup(); }
});

test('the bind result survives the refresh the bind itself causes', () => {
  const outcome: BindOutcome = { receipts: [{ pageId: 'p', shapeId: 's', rowName: 'Device', rowIndex: null, sectionIndex: null, before: null, after: '"Mixer"', refusal: null }], refusals: [], applied: true };
  const view = render(<DataBindingPanel table={table} shape={shape} onImport={() => {}} onBind={() => outcome} t={t} />);
  try {
    fireEvent.click(view.getAllByText(en.dataBinding.bind)[0]);
    expect(view.getByRole('status').textContent).toBe('Bound 1 rows.');
    // A refresh hands back a fresh object for the same shape; the message must stay.
    view.rerender(<DataBindingPanel table={table} shape={{ ...shape }} onImport={() => {}} onBind={() => outcome} t={t} />);
    expect(view.getByRole('status').textContent).toBe('Bound 1 rows.');
  } finally { cleanup(); }
});
