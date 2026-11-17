import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { fireEvent } from '@testing-library/react';
import type {
  XlsxEditRequest, XlsxEditResult, XlsxFindRequest, XlsxFindResult, XlsxValidationResult,
} from '@betteroffice/xlsx';
import { LocaleProvider } from '../i18n';
import { ReplaceDialog, type ReplaceDialogApi, type ReplaceDialogTarget } from './ReplaceDialog';

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

test('names the values a replace leaves unchanged', async () => {
  const skippedCells = ['A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8'];
  const { getByRole, findByText } = view({
    version: () => 'v1',
    find: () => ({ ok: true, version: 'v1', matches: [], truncated: false }),
    validate: () => preview({
      previews: [{
        stepIndex: 0, target: { sheetId: target.sheetId, range: target.range }, wouldChange: true,
        changedCellCount: 1,
        replace: {
          matched: 8, replaced: 1, skippedFormulas: 0, skippedLocked: 0,
          skippedValues: skippedCells.length, skippedCells,
        },
      }],
    }),
    apply: () => { throw new Error('must not apply'); },
  });

  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, '1');
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, '2');
  fireEvent.click(getByRole('button', { name: 'Preview' }));
  await findByText(/7 matching dates, TRUE\/FALSE or error values were left unchanged: A2, A3, A4, A5, A6, …/);
});

test('names the cells a replace refuses while the rest applies', async () => {
  const { getByRole, findByText } = view({
    version: () => 'v1',
    find: () => ({ ok: true, version: 'v1', matches: [], truncated: false }),
    validate: () => preview({
      previews: [{
        stepIndex: 0, target: { sheetId: target.sheetId, range: target.range }, wouldChange: true,
        changedCellCount: 1,
        replace: {
          matched: 2, replaced: 1, skippedFormulas: 0, skippedLocked: 0,
          refusedCells: [{ cell: 'B1', reason: 'Data!B1: the new formula =A1(10 does not parse' }],
        },
      }],
    }),
    apply: () => { throw new Error('must not apply'); },
  });

  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, '+');
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, '(');
  fireEvent.click(getByRole('button', { name: 'Preview' }));
  await findByText(/Not replaced: Data!B1: the new formula =A1\(10 does not parse/);
});

test('counts every refused cell, not only the listed ones', async () => {
  const { getByRole, findByText } = view({
    version: () => 'v1',
    find: () => ({ ok: true, version: 'v1', matches: [], truncated: false }),
    validate: () => preview({
      previews: [{
        stepIndex: 0, target: { sheetId: target.sheetId, range: target.range }, wouldChange: true,
        changedCellCount: 1,
        replace: {
          matched: 4, replaced: 1, skippedFormulas: 0, skippedLocked: 0, refused: 3,
          refusedCells: [
            { cell: 'B1', reason: 'Data!B1 holds rich text that a replace would flatten' },
            { cell: 'B2', reason: 'Data!B2 holds rich text that a replace would flatten' },
          ],
        },
      }],
    }),
    apply: () => { throw new Error('must not apply'); },
  });

  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'a');
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, 'b');
  fireEvent.click(getByRole('button', { name: 'Preview' }));
  await findByText(/Data!B2 holds rich text that a replace would flatten; … \+1/);
});

test('replace edits the entered text whatever Look in is set to', async () => {
  const finds: XlsxFindRequest[] = [];
  const requests: XlsxEditRequest[] = [];
  const cell = { sheetId: target.sheetId, row: 0, col: 0, a1: 'A1' };
  const { getByRole, findByText } = view({
    version: () => 'v1',
    find: (request) => {
      finds.push(request);
      return { ok: true, version: 'v1', truncated: false, matches: [{ cell, text: '123' }] };
    },
    validate: (request) => { requests.push(request); return preview(); },
    apply: (request) => ({
      ok: true, applied: true, baseVersion: 'v1', version: 'v2', source: 'host',
      receipts: [{ stepIndex: 0, changed: true, target: request.steps[0].target,
        changedCells: [], replace: { matched: 1, replaced: 1, skippedFormulas: 0, skippedLocked: 0 } }],
      changedSheets: [], calculation: { changed: [], cycleCells: [], limitedCells: [], truncated: false },
    }),
  });

  expect((getByRole('radio', { name: 'Values' }) as HTMLInputElement).checked).toBe(true);
  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, '1');
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, '2');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await findByText('1 of 1');
  expect(finds[0]).toMatchObject({ text: '1', lookIn: 'values' });

  fireEvent.click(getByRole('button', { name: 'Preview' }));
  await findByText(/Matched 3, replaced 2/);
  fireEvent.click(getByRole('button', { name: 'Replace' }));
  await waitFor(() => expect(requests.length).toBe(2));

  for (const request of requests) {
    expect(request.steps[0]).toMatchObject({ op: 'replaceText', lookIn: 'formulas' });
  }
  expect(finds.slice(1).every((request) => request.lookIn === 'formulas')).toBe(true);
  expect(finds.length).toBeGreaterThan(1);
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
  await findByText('2 of 2');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await findByText('1 of 2');
  fireEvent.click(getByRole('button', { name: 'Find previous' }));
  await findByText('2 of 2');

  fireEvent.keyDown(replaceInput, { key: 'Enter' });
  await findByText(/1 of 2/);

  expect(revealed).toEqual(['B2', 'A1', 'B2', 'A1']);
  expect(applied).toBe(1);
  expect(requests[0].steps[0]).toMatchObject({
    op: 'replaceText', query: 'alpha', replacement: 'beta',
    target: { sheetId: target.sheetId, range: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } } },
  });
});

test('replace on a cell that does not match only selects the next match', async () => {
  const valueMatch = { sheetId: target.sheetId, row: 0, col: 0, a1: 'A1' };
  const entryMatch = { sheetId: target.sheetId, row: 1, col: 1, a1: 'B2' };
  const revealed: string[] = [];
  const changed: string[] = [];
  const validated: XlsxEditRequest[] = [];
  const { getByRole, findByText } = view({
    version: () => 'v1',
    find: (request) => ({
      ok: true, version: 'v1', truncated: false,
      matches: [{ cell: request.lookIn === 'formulas' ? entryMatch : valueMatch, text: 'alpha' }],
    }),
    reveal: (cell) => { revealed.push(cell.a1); },
    validate: (request) => {
      validated.push(request);
      return {
        ok: true, baseVersion: 'v1', wouldApply: true,
        previews: [{
          stepIndex: 0, target: request.steps[0].target, wouldChange: true, changedCellCount: 1,
          replace: { matched: 1, replaced: 1, skippedFormulas: 0, skippedLocked: 0 },
        }],
      };
    },
    apply: (request) => {
      const range = request.steps[0].target.range;
      if (range.kind === 'rowCol') changed.push(`${range.start.row}:${range.start.col}`);
      return {
        ok: true, applied: true, baseVersion: 'v1', version: 'v2', source: 'host',
        receipts: [{ stepIndex: 0, changed: true, target: request.steps[0].target,
          changedCells: [], replace: { matched: 1, replaced: 1, skippedFormulas: 0, skippedLocked: 0 } }],
        changedSheets: [], calculation: { changed: [], cycleCells: [], limitedCells: [], truncated: false },
      };
    },
  });

  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, 'beta');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await findByText('1 of 1');
  expect(revealed).toEqual(['A1']);

  fireEvent.click(getByRole('button', { name: 'Replace' }));
  await waitFor(() => expect(revealed.length).toBeGreaterThan(1));
  expect(changed).toEqual([]);
  expect(validated).toEqual([]);
  expect(revealed[revealed.length - 1]).toBe('B2');

  fireEvent.click(getByRole('button', { name: 'Replace' }));
  await waitFor(() => expect(changed).toEqual(['1:1']));
});

test('find next starts after the active cell and find previous before it', async () => {
  const cells = [
    { sheetId: target.sheetId, row: 0, col: 0, a1: 'A1' },
    { sheetId: target.sheetId, row: 4, col: 2, a1: 'C5' },
    { sheetId: target.sheetId, row: 8, col: 0, a1: 'A9' },
  ];
  const revealed: string[] = [];
  const api: ReplaceDialogApi = {
    version: () => 'v1',
    find: () => ({
      ok: true, version: 'v1', truncated: false,
      matches: cells.map((cell) => ({ cell, text: 'alpha' })),
    }),
    reveal: (cell) => { revealed.push(cell.a1); },
    validate: () => preview(),
    apply: () => { throw new Error('must not apply'); },
  };
  const activeTarget = { ...target, activeCell: { row: 4, col: 1 } };

  const forward = render(<LocaleProvider>
    <ReplaceDialog open target={activeTarget} api={api} onClose={() => {}} />
  </LocaleProvider>);
  type(forward.getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  fireEvent.click(forward.getByRole('button', { name: 'Find next' }));
  await forward.findByText('2 of 3');
  forward.unmount();

  const backward = render(<LocaleProvider>
    <ReplaceDialog open target={activeTarget} api={api} onClose={() => {}} />
  </LocaleProvider>);
  type(backward.getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  fireEvent.click(backward.getByRole('button', { name: 'Find previous' }));
  await backward.findByText('1 of 3');

  expect(revealed).toEqual(['C5', 'A1']);
});

test('says the match count is truncated at the find limit', async () => {
  const { getByRole, findByText, queryByText } = view({
    version: () => 'v1',
    find: () => ({
      ok: true, version: 'v1', truncated: true,
      matches: Array.from({ length: 10_000 }, (_, row) => ({
        cell: { sheetId: target.sheetId, row, col: 0, a1: `A${row + 1}` }, text: 'alpha',
      })),
    }),
    validate: () => preview(),
    apply: () => { throw new Error('must not apply'); },
  });

  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await findByText('2 of more than 10,000');
  expect(queryByText(/2 of 10,?000$/)).toBeNull();
});

test('finds afresh after replace all', async () => {
  let finds = 0;
  const { getByRole, findByText } = view({
    version: () => 'v1',
    find: () => {
      finds += 1;
      return {
        ok: true, version: 'v1', truncated: false,
        matches: [{ cell: { sheetId: target.sheetId, row: 1, col: 1, a1: 'B2' }, text: 'alpha' }],
      };
    },
    validate: () => preview(),
    apply: (request) => ({
      ok: true, applied: true, baseVersion: 'v1', version: 'v2', source: 'host',
      receipts: [{ stepIndex: 0, changed: true, target: request.steps[0].target,
        changedCells: [], replace: { matched: 3, replaced: 2, skippedFormulas: 1, skippedLocked: 0 } }],
      changedSheets: [], calculation: { changed: [], cycleCells: [], limitedCells: [], truncated: false },
    }),
  });

  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, 'beta');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await findByText('1 of 1');
  fireEvent.click(getByRole('button', { name: 'Replace all' }));
  await findByText(/Matched 3, replaced 2/);
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await findByText('1 of 1');

  expect(finds).toBe(2);
});

test('finds afresh when the workbook version moves between presses', async () => {
  let version = 'v1';
  let finds = 0;
  const { getByRole, findByText } = view({
    version: () => version,
    find: () => {
      finds += 1;
      return {
        ok: true, version, truncated: false,
        matches: [
          { cell: { sheetId: target.sheetId, row: 1, col: 1, a1: 'B2' }, text: 'alpha' },
          { cell: { sheetId: target.sheetId, row: 2, col: 2, a1: 'C3' }, text: 'alpha' },
        ],
      };
    },
    validate: () => preview(),
    apply: () => { throw new Error('must not apply'); },
  });

  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await findByText('1 of 2');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await findByText('2 of 2');
  expect(finds).toBe(1);

  version = 'v2';
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await findByText('1 of 2');
  expect(finds).toBe(2);
});

test('drops a find result when the options change while it is in flight', async () => {
  const finds: XlsxFindRequest[] = [];
  const revealed: string[] = [];
  const pending: ((result: XlsxFindResult) => void)[] = [];
  const { getByRole, queryByText, findByText } = view({
    version: () => 'v1',
    find: (request) => {
      finds.push(request);
      return new Promise<XlsxFindResult>((resolve) => { pending.push(resolve); });
    },
    reveal: (cell) => { revealed.push(cell.a1); },
    validate: () => preview(),
    apply: () => { throw new Error('must not apply'); },
  });
  const settle = async (index: number, result: XlsxFindResult) => {
    await act(async () => {
      pending[index](result);
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
  };
  const match = { cell: { sheetId: target.sheetId, row: 1, col: 1, a1: 'B2' }, text: 'Alpha' };

  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await waitFor(() => expect(pending.length).toBe(1));
  fireEvent.click(getByRole('checkbox', { name: 'Match case' }));
  await settle(0, { ok: true, version: 'v1', truncated: false, matches: [match] });

  expect(revealed).toEqual([]);
  expect(queryByText('1 of 1')).toBeNull();

  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await waitFor(() => expect(pending.length).toBe(2));
  await settle(1, { ok: true, version: 'v1', truncated: false, matches: [] });
  await findByText(/No matches/);
  expect(finds[1]).toMatchObject({ text: 'alpha', matchCase: true });
  expect(revealed).toEqual([]);
});

test('sends matchCase when Match case is ticked', async () => {
  const finds: XlsxFindRequest[] = [];
  const requests: XlsxEditRequest[] = [];
  const { getByRole, findByText } = view({
    version: () => 'v1',
    find: (request) => {
      finds.push(request);
      return { ok: true, version: 'v1', matches: [], truncated: false };
    },
    validate: (request) => { requests.push(request); return preview(); },
    apply: () => { throw new Error('must not apply'); },
  });

  const findInput = getByRole('textbox', { name: 'Find' }) as HTMLInputElement;
  type(findInput, 'alpha');
  fireEvent.click(getByRole('checkbox', { name: 'Match case' }));
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await findByText(/No matches/);
  fireEvent.click(getByRole('button', { name: 'Preview' }));
  await findByText(/Matched 3, replaced 2/);

  expect(finds[0]).toMatchObject({ text: 'alpha', matchCase: true });
  expect(requests[0].steps[0]).toMatchObject({ op: 'replaceText', query: 'alpha', matchCase: true });
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

test('workbook scope sends no hidden sheets or chartsheets', async () => {
  const finds: XlsxFindRequest[] = [];
  const requests: XlsxEditRequest[] = [];
  const { getByRole, getByText, findByText } = render(
    <LocaleProvider>
      <ReplaceDialog
        open
        target={{
          ...target,
          sheets: [
            { sheetId: 'sheet:0', name: 'Data' },
            { sheetId: 'sheet:1', name: 'Secret', hidden: true },
            { sheetId: 'sheet:2', name: 'Chart', kind: 'chartsheet' },
            { sheetId: 'sheet:3', name: 'Summary', kind: 'worksheet' },
          ],
        }}
        api={{
          version: () => 'v1',
          find: (request) => {
            finds.push(request);
            return { ok: true, version: 'v1', matches: [], truncated: false };
          },
          validate: (request) => {
            requests.push(request);
            return {
              ok: true, baseVersion: 'v1', wouldApply: true,
              previews: request.steps.map((step, index) => ({
                stepIndex: index, target: step.target, wouldChange: false,
                changedCellCount: 0,
                replace: { matched: 0, replaced: 0, skippedFormulas: 0, skippedLocked: 0 },
              })),
            };
          },
          apply: () => { throw new Error('must not apply'); },
        }}
        onClose={() => {}}
      />
    </LocaleProvider>,
  );

  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, 'beta');
  fireEvent.change(getByRole('combobox', { name: 'Search in' }), { target: { value: 'workbook' } });
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await findByText(/No matches/);
  fireEvent.click(getByRole('button', { name: 'Preview' }));
  await waitFor(() => expect(requests.length).toBe(1));

  expect(finds[0].sheetIds).toBeUndefined();
  expect(finds[0].range).toBeUndefined();
  expect(requests[0].steps.map((step) => step.target.sheetId))
    .toEqual(['sheet:0', 'sheet:3']);
  expect(requests[0].steps.every((step) => step.op === 'replaceText' && step.within === 'workbook')).toBe(true);
  expect(getByText('All sheets (2)')).toBeTruthy();
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

function foundAt(row: number, col: number, a1: string) {
  return { cell: { sheetId: target.sheetId, row, col, a1 }, text: 'alpha' };
}

function applySuccess(request: XlsxEditRequest): XlsxEditResult {
  return {
    ok: true, applied: true, baseVersion: request.expectVersion, version: 'v9', source: 'host',
    receipts: request.steps.map((step, index) => ({
      stepIndex: index, changed: true, target: step.target, changedCells: [],
      replace: { matched: 1, replaced: 1, skippedFormulas: 0, skippedLocked: 0 },
    })),
    changedSheets: [], calculation: { changed: [], cycleCells: [], limitedCells: [], truncated: false },
  };
}

function staleFind(version: () => string) {
  const finds: XlsxFindRequest[] = [];
  const revealed: string[] = [];
  const validated: XlsxEditRequest[] = [];
  const applied: XlsxEditRequest[] = [];
  const pending: ((result: XlsxFindResult) => void)[] = [];
  const api: ReplaceDialogApi = {
    version,
    find: (request) => {
      finds.push(request);
      if (finds.length === 1) {
        return new Promise<XlsxFindResult>((resolve) => { pending.push(resolve); });
      }
      return {
        ok: true, version: version(), truncated: false,
        matches: [foundAt(1, 1, 'B2'), foundAt(2, 2, 'C3')],
      };
    },
    reveal: (cell) => { revealed.push(cell.a1); },
    validate: (request) => { validated.push(request); return preview(); },
    apply: (request) => { applied.push(request); return applySuccess(request); },
  };
  const settleStale = async () => {
    await waitFor(() => expect(pending.length).toBe(1));
    await act(async () => {
      pending[0]({ ok: true, version: 'v1', truncated: false, matches: [foundAt(1, 1, 'B2')] });
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
  };
  const replaceSelectsWithoutEditing = async (
    getByRole: (role: string, options: { name: string }) => HTMLElement,
  ) => {
    fireEvent.click(getByRole('button', { name: 'Replace' }));
    await waitFor(() => expect(revealed).toEqual(['B2']));
    expect(validated).toEqual([]);
    expect(applied).toEqual([]);
  };
  return { api, finds, revealed, applied, settleStale, replaceSelectsWithoutEditing };
}

function dialog(api: ReplaceDialogApi, current: ReplaceDialogTarget, open = true) {
  return <LocaleProvider>
    <ReplaceDialog open={open} target={current} api={api} onClose={() => {}} />
  </LocaleProvider>;
}

test('a find answered after the query was edited selects nothing and Replace never edits its match', async () => {
  const stale = staleFind(() => 'v1');
  const { getByRole, queryByText } = render(dialog(stale.api, target));
  const findInput = getByRole('textbox', { name: 'Find' }) as HTMLInputElement;
  type(findInput, 'alpha');
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, 'beta');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await waitFor(() => expect(stale.finds.length).toBe(1));
  type(findInput, 'gamma');
  await stale.settleStale();

  expect(stale.revealed).toEqual([]);
  expect(queryByText('1 of 1')).toBeNull();
  await stale.replaceSelectsWithoutEditing(getByRole);
  expect(stale.finds[1]).toMatchObject({ text: 'gamma' });

  fireEvent.click(getByRole('button', { name: 'Replace all' }));
  await waitFor(() => expect(stale.applied.length).toBe(1));
  expect(stale.applied[0].steps[0]).toMatchObject({ query: 'gamma', target: { range: target.range } });
});

test('a find answered after the target changed selects nothing and Replace never edits its match', async () => {
  const stale = staleFind(() => 'v1');
  const moved: ReplaceDialogTarget = {
    sheetId: target.sheetId,
    range: { kind: 'rowCol', start: { row: 0, col: 1 }, end: { row: 9, col: 3 } },
    label: 'B1:D10',
  };
  const { getByRole, queryByText, rerender } = render(dialog(stale.api, target));
  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, 'beta');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await waitFor(() => expect(stale.finds.length).toBe(1));
  rerender(dialog(stale.api, moved));
  await stale.settleStale();

  expect(stale.revealed).toEqual([]);
  expect(queryByText('1 of 1')).toBeNull();
  await stale.replaceSelectsWithoutEditing(getByRole);
  expect(stale.finds[1]).toMatchObject({ text: 'alpha', range: { sheetId: moved.sheetId, range: moved.range } });

  fireEvent.click(getByRole('button', { name: 'Replace all' }));
  await waitFor(() => expect(stale.applied.length).toBe(1));
  expect(stale.applied[0].steps[0]).toMatchObject({ query: 'alpha', target: { range: moved.range } });
});

test('a find answered after the workbook version moved selects nothing and Replace never edits its match', async () => {
  let version = 'v1';
  const stale = staleFind(() => version);
  const { getByRole, queryByText } = render(dialog(stale.api, target));
  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, 'beta');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await waitFor(() => expect(stale.finds.length).toBe(1));
  version = 'v2';
  await stale.settleStale();

  expect(stale.revealed).toEqual([]);
  expect(queryByText('1 of 1')).toBeNull();
  await stale.replaceSelectsWithoutEditing(getByRole);

  fireEvent.click(getByRole('button', { name: 'Replace all' }));
  await waitFor(() => expect(stale.applied.length).toBe(1));
  expect(stale.applied[0]).toMatchObject({ expectVersion: 'v2', steps: [{ query: 'alpha' }] });
});

test('closing the dialog with a find in flight shows and selects no match on reopen', async () => {
  const stale = staleFind(() => 'v1');
  const { getByRole, queryByText, rerender } = render(dialog(stale.api, target));
  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  fireEvent.click(getByRole('button', { name: 'Find next' }));
  await waitFor(() => expect(stale.finds.length).toBe(1));
  rerender(dialog(stale.api, target, false));
  await stale.settleStale();
  rerender(dialog(stale.api, target));

  expect(stale.revealed).toEqual([]);
  expect(queryByText('1 of 1')).toBeNull();
  expect(queryByText('Checking…')).toBeNull();
});

type Gated = 'version' | 'validate';
type Change = 'query' | 'target' | 'sheet' | 'scope' | 'close';

const sheetTarget: ReplaceDialogTarget = {
  ...target,
  defaultScope: 'sheet',
  sheets: [
    { sheetId: 'sheet:0', name: 'Data' },
    { sheetId: 'sheet:1', name: 'Summary' },
  ],
};

function gated() {
  const revealed: string[] = [];
  const applied: XlsxEditRequest[] = [];
  const skips: Record<Gated, number> = { version: -1, validate: -1 };
  const pending: Record<Gated, (() => void)[]> = { version: [], validate: [] };
  const answer = <T,>(call: Gated, value: T): T | Promise<T> => {
    if (skips[call] < 0) return value;
    if (skips[call] > 0) {
      skips[call] -= 1;
      return value;
    }
    skips[call] = -1;
    return new Promise<T>((resolve) => { pending[call].push(() => resolve(value)); });
  };
  const api: ReplaceDialogApi = {
    version: () => answer('version', 'v1'),
    find: () => ({
      ok: true, version: 'v1', truncated: false,
      matches: [foundAt(1, 1, 'B2'), foundAt(2, 2, 'C3')],
    }),
    reveal: (cell) => { revealed.push(cell.a1); },
    validate: () => answer('validate', preview()),
    apply: (request) => { applied.push(request); return applySuccess(request); },
  };
  const hold = (call: Gated, skip = 0) => { skips[call] = skip; };
  const paused = (call: Gated) => waitFor(() => expect(pending[call].length).toBe(1));
  const resume = (call: Gated) => act(async () => {
    pending[call][0]();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
  return { api, revealed, applied, hold, paused, resume };
}

function gatedDialog(start: ReplaceDialogTarget = target) {
  const gate = gated();
  const ui = render(dialog(gate.api, start));
  const findInput = ui.getByRole('textbox', { name: 'Find' }) as HTMLInputElement;
  type(findInput, 'alpha');
  type(ui.getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, 'beta');
  const click = async (name: string) => {
    const button = ui.getByRole('button', { name }) as HTMLButtonElement;
    await waitFor(() => expect(button.disabled).toBe(false));
    fireEvent.click(button);
  };
  const change = (kind: Change) => {
    if (kind === 'query') {
      type(findInput, 'gamma');
    } else if (kind === 'target') {
      ui.rerender(dialog(gate.api, {
        ...start, label: 'B1:D10',
        range: { kind: 'rowCol', start: { row: 0, col: 1 }, end: { row: 9, col: 3 } },
      }));
    } else if (kind === 'sheet') {
      ui.rerender(dialog(gate.api, { ...start, sheetId: 'sheet:1' }));
    } else if (kind === 'scope') {
      const select = ui.getByRole('combobox', { name: 'Search in' }) as HTMLSelectElement;
      fireEvent.change(select, { target: { value: select.value === 'workbook' ? 'sheet' : 'workbook' } });
    } else {
      ui.rerender(dialog(gate.api, start, false));
    }
  };
  return { ...gate, click, change, queryByText: ui.queryByText };
}

for (const kind of ['query', 'close'] as const) {
  test(`a cached find answered after a ${kind} change reveals no match`, async () => {
    const run = gatedDialog();
    await run.click('Find next');
    await waitFor(() => expect(run.revealed).toEqual(['B2']));
    run.hold('version');
    await run.click('Find next');
    await run.paused('version');
    run.change(kind);
    await run.resume('version');

    expect(run.revealed).toEqual(['B2']);
  });
}

for (const [kind, start] of [['scope', sheetTarget], ['sheet', sheetTarget], ['target', target]] as const) {
  test(`a fresh find answered after a ${kind} change reveals no match`, async () => {
    const run = gatedDialog(start);
    run.hold('version');
    await run.click('Find next');
    await run.paused('version');
    run.change(kind);
    await run.resume('version');

    expect(run.revealed).toEqual([]);
    expect(run.queryByText('1 of 2')).toBeNull();
  });
}

for (const [call, phase] of [['version', 'reads the version'], ['validate', 'validates']] as const) {
  for (const kind of ['query', 'target', 'close'] as const) {
    test(`Replace edits nothing when a ${kind} change lands while it ${phase}`, async () => {
      const run = gatedDialog();
      await run.click('Find next');
      await waitFor(() => expect(run.revealed).toEqual(['B2']));
      run.hold(call, call === 'version' ? 1 : 0);
      await run.click('Replace');
      await run.paused(call);
      run.change(kind);
      await run.resume(call);

      expect(run.applied).toEqual([]);
    });
  }

  for (const kind of ['query', 'scope', 'close'] as const) {
    test(`Replace all edits nothing when a ${kind} change lands while it ${phase}`, async () => {
      const run = gatedDialog();
      run.hold(call);
      await run.click('Replace all');
      await run.paused(call);
      run.change(kind);
      await run.resume(call);

      expect(run.applied).toEqual([]);
    });
  }
}

test('workbook replace all over one visible and 128 hidden sheets stays within the step limit', async () => {
  const validated: XlsxEditRequest[] = [];
  const applied: XlsxEditRequest[] = [];
  const sheets = [
    { sheetId: 'sheet:0', name: 'Data' },
    ...Array.from({ length: 128 }, (_, index) => ({
      sheetId: `sheet:${index + 1}`, name: `Hidden ${index + 1}`, hidden: true,
    })),
  ];
  const overLimit = (request: XlsxEditRequest) => request.steps.length > 128
    ? { ok: false as const, version: 'v1', failure: {
      code: 'limit-exceeded' as const, message: 'a batch has at most 128 steps',
    } }
    : null;
  const { getByRole, findByText } = render(
    <LocaleProvider>
      <ReplaceDialog
        open
        target={{ ...target, defaultScope: 'workbook', sheets }}
        api={{
          version: () => 'v1',
          find: () => ({ ok: true, version: 'v1', matches: [], truncated: false }),
          validate: (request) => {
            validated.push(request);
            return overLimit(request) ?? {
              ok: true, baseVersion: 'v1', wouldApply: true,
              previews: request.steps.map((step, index) => {
                const visible = step.target.sheetId === 'sheet:0';
                return {
                  stepIndex: index, target: step.target, wouldChange: visible,
                  changedCellCount: visible ? 1 : 0,
                  replace: {
                    matched: visible ? 1 : 0, replaced: visible ? 1 : 0,
                    skippedFormulas: 0, skippedLocked: 0,
                  },
                };
              }),
            };
          },
          apply: (request) => {
            applied.push(request);
            return overLimit(request) ?? applySuccess(request);
          },
        }}
        onClose={() => {}}
      />
    </LocaleProvider>,
  );

  expect((getByRole('combobox', { name: 'Search in' }) as HTMLSelectElement).value).toBe('workbook');
  type(getByRole('textbox', { name: 'Find' }) as HTMLInputElement, 'alpha');
  type(getByRole('textbox', { name: 'Replace with' }) as HTMLInputElement, 'beta');
  fireEvent.click(getByRole('button', { name: 'Preview' }));
  await waitFor(() => expect(validated.length).toBe(1));
  expect(validated[0].steps.length).toBeLessThanOrEqual(128);
  await findByText(/Matched 1, replaced 1/);

  fireEvent.click(getByRole('button', { name: 'Replace all' }));
  await waitFor(() => expect(applied.length).toBe(1));
  expect(applied[0].steps.map((step) => step.target.sheetId)).toEqual(['sheet:0']);
  await findByText(/Matched 1, replaced 1/);
});
