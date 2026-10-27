import { expect, test, type Locator, type Page } from 'playwright/test';
import type {} from './xlsx-worker-editor-probe';
import { pinnedEvaluate } from '../pinned-evaluate';

async function openCollaborativeEditors(page: Page) {
  await page.goto(
    '/xlsx-worker-editor.html?collaboration=1',
  );
  await page.waitForFunction(() => '__xlsxWorkerEditor' in window);
  await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.ready);
  await expect.poll(() => page.evaluate(() => [
    window.__xlsxWorkerEditor.paintedTexts(0),
    window.__xlsxWorkerEditor.paintedTexts(1),
  ])).toEqual([
    expect.arrayContaining(['initial']),
    expect.arrayContaining(['initial']),
  ]);
}

function peerEditor(page: Page, index: number): Locator {
  return page.locator(`[data-peerless-editor="${index}"]`);
}

test.describe('XlsxEditor two-peer worker collaboration', () => {
  test('mounts two collaborative editors on the worker route with no main-thread engine', async ({
    page,
  }) => {
    await openCollaborativeEditors(page);

    await expect(peerEditor(page, 0).getByRole('application')).toBeVisible();
    await expect(peerEditor(page, 1).getByRole('application')).toBeVisible();
    await expect(
      peerEditor(page, 0).locator('[data-paint-source="worker"]'),
    ).toBeVisible();
    await expect(
      peerEditor(page, 1).locator('[data-paint-source="worker"]'),
    ).toBeVisible();

    const state = await page.evaluate(() => ({
      mainThreadEngines: window.__xlsxWorkerEditor.mainThreadEngines(),
      workerCount: window.__xlsxWorkerEditor.peerless?.workerCount() ?? 0,
      generationA: window.__xlsxWorkerEditor.generation(0),
      generationB: window.__xlsxWorkerEditor.generation(1),
      hydratedA: window.__xlsxWorkerEditor.hydrated(0),
      hydratedB: window.__xlsxWorkerEditor.hydrated(1),
      failureA: window.__xlsxWorkerEditor.peerless?.failure(0) ?? null,
      failureB: window.__xlsxWorkerEditor.peerless?.failure(1) ?? null,
      formulaA: window.__xlsxWorkerEditor.formulaBar(0),
      formulaB: window.__xlsxWorkerEditor.formulaBar(1),
      paintedA: window.__xlsxWorkerEditor.paintedTexts(0),
      paintedB: window.__xlsxWorkerEditor.paintedTexts(1),
      errors: window.__xlsxWorkerEditor.errors,
    }));

    expect(state.mainThreadEngines).toBe(0);
    expect(state.workerCount).toBe(2);
    expect(state.generationA).toBe(1);
    expect(state.generationB).toBe(1);
    expect(state.hydratedA).toBe(true);
    expect(state.hydratedB).toBe(true);
    expect(state.failureA).toBeNull();
    expect(state.failureB).toBeNull();
    expect(state.formulaA).toBe('initial');
    expect(state.formulaB).toBe('initial');
    expect(state.paintedA).toContain('initial');
    expect(state.paintedB).toContain('initial');
    expect(state.errors).toEqual([]);
  });

  test('edits to different cells on each peer converge in the probe and painted grid', async ({
    page,
  }) => {
    await openCollaborativeEditors(page);

    const formulaA = peerEditor(page, 0).getByTestId('xlsx-formula-input');
    await formulaA.fill('peer-a-a1');
    await formulaA.press('Enter');

    const editorB = peerEditor(page, 1);
    const scrollB = editorB.getByTestId('xlsx-scroll');
    await scrollB.focus();
    await page.keyboard.press('ArrowRight');
    const formulaB = editorB.getByTestId('xlsx-formula-input');
    await expect(formulaB).toHaveValue('1');
    await formulaB.fill('peer-b-b1');
    await formulaB.press('Enter');

    await expect
      .poll(() =>
        page.evaluate(() =>
          Promise.all([
            window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B1', 0),
            window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B1', 1),
            window.__xlsxWorkerEditor.paintedTexts(0),
            window.__xlsxWorkerEditor.paintedTexts(1),
          ]),
        ),
      )
      .toEqual([
        [['peer-a-a1', 'peer-b-b1']],
        [['peer-a-a1', 'peer-b-b1']],
        expect.arrayContaining(['peer-a-a1', 'peer-b-b1']),
        expect.arrayContaining(['peer-a-a1', 'peer-b-b1']),
      ]);

    expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual(
      [],
    );
  });

  test('concurrent edits to the same cell converge to one value on both peers', async ({
    page,
  }) => {
    await openCollaborativeEditors(page);
    await page.evaluate(() => window.__xlsxWorkerEditor.pauseDelivery());

    const formulaA = peerEditor(page, 0).getByTestId('xlsx-formula-input');
    await formulaA.fill('concurrent-peer-a');
    await formulaA.press('Enter');

    const formulaB = peerEditor(page, 1).getByTestId('xlsx-formula-input');
    await formulaB.fill('concurrent-peer-b');
    await formulaB.press('Enter');

    await expect
      .poll(() =>
        page.evaluate(() =>
          Promise.all([
            window.__xlsxWorkerEditor.cell(0),
            window.__xlsxWorkerEditor.cell(1),
            window.__xlsxWorkerEditor.pendingUpdates(0),
            window.__xlsxWorkerEditor.pendingUpdates(1),
          ]),
        ),
      )
      .toEqual(['concurrent-peer-a', 'concurrent-peer-b', 1, 1]);

    await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.resumeDelivery());

    await expect
      .poll(async () => {
        const [cellA, cellB, paintedA, paintedB] = await page.evaluate(() =>
          Promise.all([
            window.__xlsxWorkerEditor.cell(0),
            window.__xlsxWorkerEditor.cell(1),
            window.__xlsxWorkerEditor.paintedTexts(0),
            window.__xlsxWorkerEditor.paintedTexts(1),
          ]),
        );
        const valid =
          cellA === 'concurrent-peer-a' || cellA === 'concurrent-peer-b';
        return (
          valid &&
          cellA === cellB &&
          paintedA.includes(cellA!) &&
          paintedB.includes(cellB!)
        );
      })
      .toBe(true);

    const [cellA, cellB] = await page.evaluate(() =>
      Promise.all([
        window.__xlsxWorkerEditor.cell(0),
        window.__xlsxWorkerEditor.cell(1),
      ]),
    );
    expect(cellA).toBe(cellB);
    expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual(
      [],
    );
  });

  test('undo on one peer reverts only its own edit while the remote edit stays on both peers', async ({
    page,
  }) => {
    await openCollaborativeEditors(page);

    const formulaA = peerEditor(page, 0).getByTestId('xlsx-formula-input');
    await formulaA.fill('peer-a-edit');
    await formulaA.press('Enter');

    const editorB = peerEditor(page, 1);
    const scrollB = editorB.getByTestId('xlsx-scroll');
    await scrollB.focus();
    await page.keyboard.press('ArrowRight');
    const formulaB = editorB.getByTestId('xlsx-formula-input');
    await expect(formulaB).toHaveValue('1');
    await formulaB.fill('peer-b-edit');
    await formulaB.press('Enter');

    await expect
      .poll(() =>
        page.evaluate(() =>
          Promise.all([
            window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B1', 0),
            window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B1', 1),
          ]),
        ),
      )
      .toEqual([
        [['peer-a-edit', 'peer-b-edit']],
        [['peer-a-edit', 'peer-b-edit']],
      ]);

    expect(await page.evaluate(() => window.__xlsxWorkerEditor.undo(0))).toBe(
      true,
    );

    await expect
      .poll(() =>
        page.evaluate(() =>
          Promise.all([
            window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B1', 0),
            window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B1', 1),
            window.__xlsxWorkerEditor.paintedTexts(0),
            window.__xlsxWorkerEditor.paintedTexts(1),
          ]),
        ),
      )
      .toEqual([
        [['initial', 'peer-b-edit']],
        [['initial', 'peer-b-edit']],
        expect.arrayContaining(['initial', 'peer-b-edit']),
        expect.arrayContaining(['initial', 'peer-b-edit']),
      ]);

    expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual(
      [],
    );
  });

  test('remote update arriving during cell editing preserves in-progress local input', async ({
    page,
  }) => {
    await openCollaborativeEditors(page);
    await page.evaluate(() => window.__xlsxWorkerEditor.pauseDelivery());

    const formulaA = peerEditor(page, 0).getByTestId('xlsx-formula-input');
    await formulaA.fill('remote-from-a');
    await formulaA.press('Enter');

    await expect
      .poll(() =>
        page.evaluate(() =>
          Promise.all([
            window.__xlsxWorkerEditor.cell(0),
            window.__xlsxWorkerEditor.pendingUpdates(0),
          ]),
        ),
      )
      .toEqual(['remote-from-a', 1]);

    const editorB = peerEditor(page, 1);
    const scrollB = editorB.getByTestId('xlsx-scroll');
    await scrollB.focus();
    await page.keyboard.press('ArrowRight');
    await expect(editorB.getByTestId('xlsx-formula-input')).toHaveValue('1');

    await page.keyboard.press('F2');
    const cellEditorB = editorB.getByTestId('xlsx-cell-editor');
    await expect(cellEditorB).toBeVisible();
    await expect(cellEditorB).toBeFocused();
    await page.keyboard.press('Backspace');
    await page.keyboard.type('local-');
    await expect(cellEditorB).toHaveValue('local-');

    await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.resumeDelivery());

    await expect
      .poll(() =>
        page.evaluate(() =>
          Promise.all([
            window.__xlsxWorkerEditor.cell(1),
            window.__xlsxWorkerEditor.paintedTexts(1),
          ]),
        ),
      )
      .toEqual(['remote-from-a', expect.arrayContaining(['remote-from-a'])]);

    await expect(cellEditorB).toBeVisible();
    await expect(cellEditorB).toBeFocused();
    await expect(cellEditorB).toHaveValue('local-');

    await page.keyboard.type('typed-b');
    await expect(cellEditorB).toHaveValue('local-typed-b');
    await page.keyboard.press('Enter');

    await expect
      .poll(() =>
        page.evaluate(() =>
          Promise.all([
            window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B1', 0),
            window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B1', 1),
            window.__xlsxWorkerEditor.paintedTexts(0),
            window.__xlsxWorkerEditor.paintedTexts(1),
          ]),
        ),
      )
      .toEqual([
        [['remote-from-a', 'local-typed-b']],
        [['remote-from-a', 'local-typed-b']],
        expect.arrayContaining(['remote-from-a', 'local-typed-b']),
        expect.arrayContaining(['remote-from-a', 'local-typed-b']),
      ]);

    expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual(
      [],
    );
  });
});
