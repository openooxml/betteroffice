import { expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createT, en } from '@betteroffice/vsdx-i18n';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import * as vsdx from '@betteroffice/vsdx';
import type { ShapeSnapshot } from '@betteroffice/vsdx';
import { bindRowToShape, matchColumnsToRows, type BindOutcome, type ImportedTable } from './bindTable';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { cleanup, fireEvent, render } = await import('@testing-library/react');
const { DataBindingPanel } = await import('./DataBindingPanel');

const t = createT(en);
const root = resolve(import.meta.dir, '../../../../..');
const wasm = await readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm'));
const fixture = await readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/data-binding.vsdx'));
await vsdx.initWasm(wasm);

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

test('binding a row reports how many fields it wrote', () => {
  const bound: number[] = [];
  const outcome: BindOutcome = { receipts: [{ pageId: 'p', shapeId: 's', rowName: 'Device', rowIndex: null, sectionIndex: null, before: null, after: '"Mixer"', refusal: null }], refusals: [], applied: true };
  const view = render(<DataBindingPanel table={table} shape={shape} onImport={() => {}} onBind={(index) => { bound.push(index); return outcome; }} t={t} />);
  try {
    fireEvent.click(view.getAllByText(en.dataBinding.bind)[1]);
    expect(bound).toEqual([1]);
    expect(view.getByRole('status').textContent).toBe('Wrote 1 fields.');
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

test.each([new Error('stale selection'), 'bad row index'])('a throwing bind reports a refusal instead of escaping: %s', (error) => {
  const view = render(<DataBindingPanel table={table} shape={shape} onImport={() => {}} onBind={() => { throw error; }} t={t} />);
  try {
    expect(() => fireEvent.click(view.getAllByText(en.dataBinding.bind)[0])).not.toThrow();
    const reason = error instanceof Error ? error.message : error;
    expect(view.getByRole('alert').textContent).toBe(t('dataBinding.refused', { rows: reason }));
    expect(view.queryByRole('status')).toBeNull();
  } finally { cleanup(); }
});

test('the panel binds through the real engine and reports written fields', () => {
  const handle = vsdx.openDiagram(fixture);
  const page = handle.snapshot().pages[0];
  const selected = page.shapes[0];
  const imported: ImportedTable = { headers: ['Device', 'Count'], rows: [['Mixer', '7']] };
  const bindings = matchColumnsToRows(imported, selected);
  const view = render(<DataBindingPanel table={imported} shape={selected} onImport={() => {}} onBind={(index) => bindRowToShape(handle, page.id, selected.id, imported, index, bindings)} t={t} />);
  try {
    fireEvent.click(view.getByText(en.dataBinding.bind));
    expect(view.getByRole('status').textContent).toBe('Wrote 2 fields.');
    const rows = vsdx.shapeDataRows(handle.snapshot().pages[0].shapes[0]);
    expect(rows.find((row) => row.rowName === 'Device')?.formula).toBe('"Mixer"');
    expect(rows.find((row) => row.rowName === 'Count')?.formula).toBe('7');
    expect(handle.undo().applied).toBe(true);
    expect(handle.canUndo()).toBe(false);
  } finally { cleanup(); handle.dispose(); }
});

test('an engine refusal reaches the panel and leaves the shape and undo history unchanged', () => {
  const handle = vsdx.openDiagram(fixture);
  const page = handle.snapshot().pages[0];
  const selected = page.shapes[0];
  const imported: ImportedTable = { headers: ['Device', 'Due date'], rows: [['Mixer', '4 April']] };
  const before = handle.snapshot();
  const view = render(<DataBindingPanel
    table={imported} shape={selected} onImport={() => {}}
    onBind={(index) => bindRowToShape(handle, page.id, selected.id, imported, index, [{ column: 0, rowName: 'Device' }, { column: 1, rowName: 'Due' }])}
    t={t}
  />);
  try {
    fireEvent.click(view.getByText(en.dataBinding.bind));
    expect(view.getByRole('alert').textContent).toBe(t('dataBinding.refused', { rows: 'Due' }));
    expect(handle.snapshot()).toEqual(before);
    expect(handle.canUndo()).toBe(false);
  } finally { cleanup(); handle.dispose(); }
});

test.each(['page', 'shape', 'row'])('a stale %s is caught on the panel-to-engine path', (target) => {
  const handle = vsdx.openDiagram(fixture);
  const page = handle.snapshot().pages[0];
  const selected = page.shapes[0];
  const bindings = matchColumnsToRows(table, selected);
  const before = handle.snapshot();
  const view = render(<DataBindingPanel
    table={table} shape={selected} onImport={() => {}}
    onBind={(index) => bindRowToShape(handle, target === 'page' ? 'missing-page' : page.id, target === 'shape' ? 'missing-shape' : selected.id, table, target === 'row' ? table.rows.length : index, bindings)}
    t={t}
  />);
  try {
    expect(() => fireEvent.click(view.getAllByText(en.dataBinding.bind)[0])).not.toThrow();
    expect(view.getByRole('alert').textContent).toContain('Nothing was written');
    expect(view.queryByRole('status')).toBeNull();
    expect(handle.snapshot()).toEqual(before);
    expect(handle.canUndo()).toBe(false);
  } finally { cleanup(); handle.dispose(); }
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
    expect(view.getByRole('status').textContent).toBe('Wrote 1 fields.');
    // A refresh hands back a fresh object for the same shape; the message must stay.
    view.rerender(<DataBindingPanel table={table} shape={{ ...shape }} onImport={() => {}} onBind={() => outcome} t={t} />);
    expect(view.getByRole('status').textContent).toBe('Wrote 1 fields.');
  } finally { cleanup(); }
});
