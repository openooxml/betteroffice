import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { fireEvent } from '@testing-library/react';
import type { XlsxEditRequest, XlsxValidationResult } from '@betteroffice/xlsx';
import { LocaleProvider } from '../i18n';
import { ReplaceDialog } from './ReplaceDialog';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const target = {
  sheetId: 'sheet:0',
  range: { kind: 'rowCol' as const, start: { row: 0, col: 0 }, end: { row: 9, col: 2 } },
  label: 'A1:C10',
};

function preview(overrides: Partial<XlsxValidationResult> = {}): XlsxValidationResult {
  return {
    ok: true,
    baseVersion: 'v1',
    wouldApply: true,
    previews: [{
      stepIndex: 0, target: { sheetId: target.sheetId, range: target.range }, wouldChange: true,
      changedCellCount: 2,
      replace: { matched: 3, replaced: 2, skippedFormulas: 1, skippedLocked: 0 },
    }],
    ...overrides,
  } as XlsxValidationResult;
}

function type(input: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
  act(() => {
    setter.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  });
}

function view(api: Parameters<typeof ReplaceDialog>[0]['api'], onApplied?: () => void) {
  return render(<LocaleProvider>
    <ReplaceDialog open target={target} api={api} onClose={() => {}} onApplied={onApplied} />
  </LocaleProvider>);
}

test('previews stats and applies the selected look-in mode', async () => {
  const requests: XlsxEditRequest[] = [];
  let applied = 0;
  const { getByRole, findByText } = view({
    version: () => 'v1',
    find: () => ({ ok: true, version: 'v1', matches: [], truncated: false }),
    validate: (request) => { requests.push(request); return preview(); },
    apply: (request) => {
      requests.push(request);
      return {
        ok: true, applied: true, baseVersion: 'v1', version: 'v2',
        source: 'host', receipts: [{ stepIndex: 0, changed: true, target: request.steps[0].target,
          changedCells: [], replace: { matched: 3, replaced: 2, skippedFormulas: 1, skippedLocked: 0 } }],
        changedSheets: [], calculation: { changed: [], cycleCells: [], limitedCells: [], truncated: false },
      };
    },
  }, () => { applied += 1; });

  expect((getByRole('combobox', { name: 'Search in' }) as HTMLSelectElement).value).toBe('selection');

  const find = getByRole('textbox', { name: 'Find' }) as HTMLInputElement;
  type(find, 'alpha');
  expect(find.value).toBe('alpha');
  await waitFor(() => expect((getByRole('button', { name: 'Preview' }) as HTMLButtonElement).disabled).toBe(false));
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, 'beta');
  expect(find.value).toBe('alpha');
  const previewButton = getByRole('button', { name: 'Preview' }) as HTMLButtonElement;
  expect(previewButton.disabled).toBe(false);
  fireEvent.click(previewButton);
  await findByText(/Matched 3, replaced 2/);

  fireEvent.click(getByRole('radio', { name: 'Formulas' }));
  fireEvent.click(getByRole('button', { name: 'Replace all' }));
  await findByText(/Matched 3, replaced 2/);

  expect(applied).toBe(1);
  expect(requests[requests.length - 1]?.steps[0]).toMatchObject({
    op: 'replaceText', query: 'alpha', replacement: 'beta', lookIn: 'formulas',
  });
});

test('shows engine refusals without applying', async () => {
  let applied = 0;
  const { getByRole, findByText } = view({
    version: () => 'v1',
    find: () => ({ ok: true, version: 'v1', matches: [], truncated: false }),
    validate: () => ({ ok: false, version: 'v1', failure: {
      code: 'stale-version', message: 'the workbook moved',
    } }),
    apply: () => { applied += 1; throw new Error('must not apply'); },
  });

  const find = getByRole('textbox', { name: 'Find' }) as HTMLInputElement;
  type(find, 'alpha');
  expect(find.value).toBe('alpha');
  fireEvent.click(getByRole('button', { name: 'Replace all' }));
  await findByText(/Not replaced: the workbook moved/);
  expect(applied).toBe(0);
});

test('finds one match at a time and replaces only the current cell', async () => {
  const cells = [
    { sheetId: target.sheetId, row: 0, col: 0, a1: 'A1' },
    { sheetId: target.sheetId, row: 1, col: 1, a1: 'B2' },
  ];
  const revealed: string[] = [];
  const requests: XlsxEditRequest[] = [];
  let applied = 0;
  const { getByRole, findByText } = view({
    version: () => 'v1',
    find: () => ({
      ok: true, version: 'v1', truncated: false,
      matches: cells.map((cell) => ({ cell, text: 'alpha' })),
    }),
    reveal: (cell) => { revealed.push(cell.a1); },
    validate: (request) => {
      requests.push(request);
      return {
        ok: true, baseVersion: 'v1', wouldApply: true,
        previews: [{
          stepIndex: 0, target: request.steps[0].target, wouldChange: true,
          changedCellCount: 1,
          replace: { matched: 1, replaced: 1, skippedFormulas: 0, skippedLocked: 0 },
        }],
      };
    },
    apply: (request) => {
      applied += 1;
      return {
        ok: true, applied: true, baseVersion: 'v1', version: 'v2', source: 'host',
        receipts: [{ stepIndex: 0, changed: true, target: request.steps[0].target,
          changedCells: [], replace: { matched: 1, replaced: 1, skippedFormulas: 0, skippedLocked: 0 } }],
        changedSheets: [], calculation: { changed: [], cycleCells: [], limitedCells: [], truncated: false },
      };
    },
  });

  const findInput = getByRole('textbox', { name: 'Find' }) as HTMLInputElement;
  type(findInput, 'alpha');
  const replaceInput = getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement;
  type(replaceInput, 'beta');

  fireEvent.keyDown(findInput, { key: 'Enter' });
  await findByText('1 of 2');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await findByText('2 of 2');
  fireEvent.click(getByRole('button', { name: 'Find previous' }));
  await findByText('1 of 2');

  fireEvent.keyDown(replaceInput, { key: 'Enter' });
  await findByText(/2 of 2/);

  expect(revealed).toEqual(['A1', 'B2', 'A1', 'B2']);
  expect(applied).toBe(1);
  expect(requests[0].steps[0]).toMatchObject({
    op: 'replaceText', query: 'alpha', replacement: 'beta',
    target: { sheetId: target.sheetId, range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } } },
  });
});

test('drags the dialog by its header', () => {
  const { getByText } = view({
    version: () => 'v1',
    find: () => ({ ok: true, version: 'v1', matches: [], truncated: false }),
    validate: () => preview(),
    apply: () => { throw new Error('must not apply'); },
  });
  const header = getByText('Find and replace').parentElement!;
  const dialog = getByText('Find and replace').closest('[role="dialog"]') as HTMLElement;
  expect(dialog.style.transform).toBe('translate(0px, 0px)');

  fireEvent.mouseDown(header, { button: 0, clientX: 100, clientY: 100 });
  fireEvent.mouseMove(window, { clientX: 150, clientY: 130 });
  fireEvent.mouseUp(window);

  expect(dialog.style.transform).toBe('translate(50px, 30px)');
});

test('replaces every sheet when the scope is all sheets', async () => {
  const requests: XlsxEditRequest[] = [];
  const { getByRole, findByText } = render(
    <LocaleProvider>
      <ReplaceDialog
        open
        target={{
          ...target,
          sheets: [
            { sheetId: 'sheet:0', name: 'Data' },
            { sheetId: 'sheet:1', name: 'Summary' },
          ],
        }}
        api={{
          version: () => 'v1',
          find: () => ({ ok: true, version: 'v1', matches: [], truncated: false }),
          validate: (request) => {
            requests.push(request);
            return {
              ok: true, baseVersion: 'v1', wouldApply: true,
              previews: request.steps.map((step, index) => ({
                stepIndex: index, target: step.target, wouldChange: true,
                changedCellCount: 1,
                replace: { matched: 1, replaced: 1, skippedFormulas: 0, skippedLocked: 0 },
              })),
            };
          },
          apply: (request) => ({
            ok: true, applied: true, baseVersion: 'v1', version: 'v2', source: 'host',
            receipts: request.steps.map((step, index) => ({
              stepIndex: index, changed: true, target: step.target, changedCells: [],
              replace: { matched: 1, replaced: 1, skippedFormulas: 0, skippedLocked: 0 },
            })),
            changedSheets: [], calculation: { changed: [], cycleCells: [], limitedCells: [], truncated: false },
          }),
        }}
        onClose={() => {}}
      />
    </LocaleProvider>,
  );

  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, 'beta');
  fireEvent.change(getByRole('combobox', { name: 'Search in' }), { target: { value: 'workbook' } });
  fireEvent.click(getByRole('button', { name: 'Replace all' }));
  await findByText(/Matched 2, replaced 2/);

  expect(requests[0].steps.map((step) => step.target.sheetId)).toEqual(['sheet:0', 'sheet:1']);
  expect(requests[0].steps[0].target.range).toMatchObject({
    kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 1048575, col: 16383 },
  });
});

test('defaults to all sheets when only one cell is selected', () => {
  const { getByRole } = render(
    <LocaleProvider>
      <ReplaceDialog
        open
        target={{
          sheetId: 'sheet:0',
          range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          label: 'A1',
          sheets: [
            { sheetId: 'sheet:0', name: 'Data' },
            { sheetId: 'sheet:1', name: 'Summary' },
          ],
        }}
        api={{
          version: () => 'v1',
          find: () => ({ ok: true, version: 'v1', matches: [], truncated: false }),
          validate: () => preview(),
          apply: () => { throw new Error('must not apply'); },
        }}
        onClose={() => {}}
      />
    </LocaleProvider>,
  );
  expect((getByRole('combobox', { name: 'Search in' }) as HTMLSelectElement).value).toBe('workbook');
});
