import { test, expect, type Page } from 'playwright/test';
import type { WorkerEditorProbe } from './xlsx-worker-editor-probe';

declare global {
  interface Window { __xlsxWorkerEditor: WorkerEditorProbe }
}

test.setTimeout(120_000);

async function open(page: Page, holdHydration = false) {
  await page.goto(holdHydration ? '/xlsx-worker-editor.html?holdHydration=1' : '/xlsx-worker-editor.html');
  await page.waitForFunction(() => '__xlsxWorkerEditor' in window);
  await page.evaluate(() => window.__xlsxWorkerEditor.ready);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('initial');
}

async function edit(page: Page, value: string, capturePreview = false) {
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.focus();
  await scroll.press('F2');
  const input = page.getByTestId('xlsx-cell-editor');
  await input.fill('');
  await input.pressSequentially(value);
  await expect(input).toHaveValue(value);
  await expect(input).toBeFocused();
  expect(await input.evaluate((element) => (element as HTMLInputElement).selectionStart)).toBe(value.length);
  if (capturePreview) await page.evaluate(() => window.__xlsxWorkerEditor.holdPreview());
  await input.press('Enter');
  if (capturePreview) {
    const preview = page.getByTestId('xlsx-commit-preview');
    await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.previewHeld())).toBe(true);
    expect(await page.evaluate(() => window.__xlsxWorkerEditor.hydrated())).toBe(false);
    await expect(preview).toHaveText(value);
    await expect(preview).toBeVisible();
    expect(await page.evaluate(() => window.__xlsxWorkerEditor.commitOrder)).toEqual([]);
    expect((await preview.screenshot()).byteLength).toBeGreaterThan(0);
    expect(await page.evaluate(() => window.__xlsxWorkerEditor.previewHeld())).toBe(true);
    expect(await page.evaluate(() => window.__xlsxWorkerEditor.hydrated())).toBe(false);
    expect(await page.evaluate(() => window.__xlsxWorkerEditor.commitOrder)).toEqual([]);
    await page.evaluate((text) => window.__xlsxWorkerEditor.releasePreview(text), value);
  }
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.cell())).toBe(value);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain(value);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.hydrated())).toBe(true);
  await expect(page.getByTestId('xlsx-commit-preview')).toHaveCount(0);
}

test('xlsx worker editor echoes the caret, paints a commit preview and saves editable bytes for reopening', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await open(page, true);
  await edit(page, 'worker edit', true);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.commitOrder)).toEqual([
    { kind: 'painted-preview', text: 'worker edit' },
    { kind: 'mutator-entry', text: 'worker edit' },
  ]);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.previews)).toContain('worker edit');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.previewFrames)).toContain('worker edit');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.saveAndReopen())).toBeGreaterThan(0);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.cell())).toBe('worker edit');
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('worker edit');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
  expect(errors).toEqual([]);
});

test('xlsx worker editor undo restores the original cell and canvas value', async ({ page }) => {
  await open(page);
  await edit(page, 'undo value');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.undo())).toBe(true);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.cell())).toBe('initial');
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('initial');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).not.toContain('undo value');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
});


for (const dpr of [1, 2]) {
  for (const zoom of [1, 1.25]) {
    for (const variant of ['plain', 'styled', 'formatted', 'merged', 'overflow']) {
      test(`held commit preview is byte-identical to the adopted worker frame in the edited cell region (${variant}, zoom ${zoom}, DPR ${dpr})`, async ({ browser, baseURL }) => {
        const context = await browser.newContext({ baseURL, deviceScaleFactor: dpr });
        const page = await context.newPage();
        try {
          await page.goto(`/xlsx-worker-editor.html?holdHydration=1&variant=${variant}`);
          await page.waitForFunction(() => '__xlsxWorkerEditor' in window);
          await page.evaluate(() => window.__xlsxWorkerEditor.ready);
          await page.evaluate((scale) => window.__xlsxWorkerEditor.zoom(scale), zoom);
          await expect.poll(() => page.locator('[data-paint-source="worker"]').evaluate((canvas) => canvas.clientWidth)).toBeGreaterThan(0);
          await expect(page.locator('[data-paint-source="worker"]')).toHaveAttribute('data-worker-zoom', String(zoom));
          const clip = await page.evaluate(() => window.__xlsxWorkerEditor.cellClip());
          const value = variant === 'formatted' ? '123.5' : variant === 'overflow' ? 'worker edit '.repeat(12) : 'worker edit';
          await page.getByTestId('xlsx-scroll').focus();
          await page.getByTestId('xlsx-scroll').press('F2');
          await page.getByTestId('xlsx-cell-editor').fill(value);
          await page.evaluate(() => window.__xlsxWorkerEditor.holdPreview());
          await page.getByTestId('xlsx-cell-editor').press('Enter');
          await expect(page.getByTestId('xlsx-commit-preview')).toHaveText(value);
          await expect(page.getByTestId('xlsx-commit-preview')).toBeVisible();
          await expect(page.getByTestId('xlsx-commit-preview')).toHaveAttribute('data-preview-ready', 'true');
          expect(await page.evaluate(() => window.__xlsxWorkerEditor.hydrated())).toBe(false);
          const before = await page.screenshot({ clip, animations: 'disabled' });
          await page.evaluate((text) => window.__xlsxWorkerEditor.releasePreview(text), value);
          await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.adoptedSequence())).toBeGreaterThanOrEqual(1);
          await expect(page.getByTestId('xlsx-commit-preview')).toHaveCount(0);
          const after = await page.screenshot({ clip, animations: 'disabled' });
          expect(before.equals(after)).toBe(true);
          expect(await page.evaluate(() => window.__xlsxWorkerEditor.peerEntries.filter((entry) => entry.method === 'editCell').length)).toBe(1);
          await expect(page.locator('[data-paint-source="worker"]')).toHaveAttribute('data-worker-sequence', '1');
        } finally { await context.close(); }
      });
    }
  }
}

async function replayMatchesPeer(page: Page) {
  await page.evaluate(() => window.__xlsxWorkerEditor.flush());
  const { peers, replays } = await page.evaluate(() => ({ peers: window.__xlsxWorkerEditor.peerEntries, replays: window.__xlsxWorkerEditor.replayEntries }));
  expect(peers).toEqual(replays.map(({ sequence: _sequence, ...entry }) => entry));
  for (const generation of new Set(replays.map((entry) => entry.generation))) {
    const sequence = replays.filter((entry) => entry.generation === generation).map((entry) => entry.sequence);
    expect(sequence).toEqual(sequence.map((_, index) => index + 1));
  }
}

for (const phase of ['before', 'during', 'hydrated']) {
  test(`keeps peer application order equal to worker replay order across every input route (${phase})`, async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await page.goto(`/xlsx-worker-editor.html${phase === 'hydrated' ? '' : `?holdHydration=${phase}`}`);
    await page.waitForFunction(() => '__xlsxWorkerEditor' in window);
    await page.evaluate(() => window.__xlsxWorkerEditor.ready);
    const scroll = page.getByTestId('xlsx-scroll');
    if (phase !== 'hydrated') {
      await page.evaluate(() => window.__xlsxWorkerEditor.queueBulkFill());
      const chart = await page.evaluate(() => window.__xlsxWorkerEditor.chartClip());
      await page.mouse.move(chart.x + 8, chart.y + 8);
      await page.mouse.down();
      await page.mouse.move(chart.x + 18, chart.y + 18);
      await page.mouse.up();
      await scroll.press('ArrowRight');
      const cell = await page.evaluate(() => window.__xlsxWorkerEditor.cellClip());
      await page.mouse.click(cell.x + 8, cell.y + 8);
      await page.evaluate(() => navigator.clipboard.writeText('queued clipboard'));
      await scroll.focus();
      await scroll.press('ControlOrMeta+v');
    }
    await scroll.focus();
    await scroll.press('F2');
    await page.getByTestId('xlsx-cell-editor').fill('cell route');
    await page.getByTestId('xlsx-cell-editor').press('Tab');
    await page.getByTestId('xlsx-formula-input').fill('=24');
    await page.getByTestId('xlsx-formula-input').press('Enter');
    await page.evaluate(() => { window.__xlsxWorkerEditor.queueHostEdit('host route'); });
    if (phase !== 'hydrated') {
      await scroll.focus();
      await scroll.press('ControlOrMeta+b');
      await expect(page.getByTestId('xlsx-commit-preview')).toBeVisible();
      expect(await page.evaluate(() => window.__xlsxWorkerEditor.peerEntries)).toEqual([]);
      await page.evaluate(() => window.__xlsxWorkerEditor.releaseHydration());
    }
    await page.evaluate(() => window.__xlsxWorkerEditor.flush());
    await page.getByRole('tab', { name: 'Other' }).click();
    await page.getByRole('tab', { name: 'Sheet', exact: true }).click();
    await page.evaluate(() => window.__xlsxWorkerEditor.flush());
    await scroll.focus();
    await scroll.press('ControlOrMeta+b');
    await page.evaluate(() => window.__xlsxWorkerEditor.flush());
    await page.evaluate(() => window.__xlsxWorkerEditor.formatCells());
    await page.evaluate(async () => {
      await window.__xlsxWorkerEditor.bulkFill();
      await window.__xlsxWorkerEditor.pluginRoutes();
      await window.__xlsxWorkerEditor.history();
      await window.__xlsxWorkerEditor.proposals();
      await navigator.clipboard.writeText('clipboard route');
    });
    await scroll.focus();
    await scroll.press('ControlOrMeta+v');
    await page.evaluate(() => window.__xlsxWorkerEditor.flush());
    await scroll.press('ControlOrMeta+c');
    await page.evaluate(() => window.__xlsxWorkerEditor.flush());
    await scroll.press('ControlOrMeta+x');
    await page.evaluate(() => window.__xlsxWorkerEditor.flush());
    await scroll.press('Delete');
    await page.evaluate(() => window.__xlsxWorkerEditor.flush());
    const chart = await page.evaluate(() => window.__xlsxWorkerEditor.chartClip());
    await page.mouse.move(chart.x + 8, chart.y + 8);
    await page.mouse.down();
    await page.mouse.move(chart.x + 18, chart.y + 18);
    await page.mouse.up();
    await scroll.press('ArrowRight');
    await replayMatchesPeer(page);
    const methods = await page.evaluate(() => window.__xlsxWorkerEditor.replayEntries.map((entry) => entry.method));
    for (const method of ['editCell', 'editCells', 'setActiveSheet', 'patchRangeStyle', 'applyFormat', 'setNumberFormat', 'applyEdits', 'undo', 'redo', 'moveChart', 'propose', 'acceptProposal']) {
      expect(methods).toContain(method);
    }
    const beforeFailure = await page.evaluate(() => window.__xlsxWorkerEditor.replayEntries.length);
    await page.evaluate(() => { window.__xlsxWorkerEditor.queueHostEdit('accepted before failure'); window.__xlsxWorkerEditor.fail(); });
    expect(await page.evaluate(() => window.__xlsxWorkerEditor.recover())).toBeGreaterThan(0);
    const recovered = await page.evaluate(() => ({ peers: window.__xlsxWorkerEditor.peerEntries, replays: window.__xlsxWorkerEditor.replayEntries }));
    expect(recovered.replays).toHaveLength(beforeFailure);
    expect(recovered.peers.filter((entry) => entry.method === 'editCell' && entry.args[3] === 'accepted before failure')).toHaveLength(1);
  });
}

for (const lifecycle of ['replacement', 'disposal', 'StrictMode']) {
  test(`keeps peer application order equal to worker replay order across every input route (${lifecycle})`, async ({ page }) => {
    await page.goto(`/xlsx-worker-editor.html?holdHydration=before${lifecycle === 'StrictMode' ? '&strict=1' : ''}`);
    await page.waitForFunction(() => '__xlsxWorkerEditor' in window);
    await page.evaluate(() => window.__xlsxWorkerEditor.ready);
    const retiring = await page.evaluate(() => window.__xlsxWorkerEditor.generation());
    await page.getByTestId('xlsx-scroll').focus();
    await page.getByTestId('xlsx-scroll').press('F2');
    await page.getByTestId('xlsx-cell-editor').fill('retiring generation');
    await page.getByTestId('xlsx-cell-editor').press('Enter');
    await expect(page.getByTestId('xlsx-commit-preview')).toBeVisible();
    if (lifecycle === 'disposal') await page.evaluate(() => window.__xlsxWorkerEditor.dispose());
    else if (lifecycle === 'replacement') await page.evaluate(() => { void window.__xlsxWorkerEditor.replace(); });
    await page.evaluate(() => window.__xlsxWorkerEditor.releaseHydration());
    await expect.poll(() => page.evaluate((generation) => window.__xlsxWorkerEditor.replayEntries.filter((entry) => entry.generation === generation).length, retiring)).toBe(1);
    const entries = await page.evaluate(() => ({ peers: window.__xlsxWorkerEditor.peerEntries, replays: window.__xlsxWorkerEditor.replayEntries }));
    expect(entries.peers).toEqual(entries.replays.map(({ sequence: _sequence, ...entry }) => entry));
    expect(entries.peers.filter((entry) => entry.args[3] === 'retiring generation')).toEqual([
      { generation: retiring, method: 'editCell', args: [0, 0, 0, 'retiring generation'] },
    ]);
    expect(entries.replays[0].sequence).toBe(1);
    if (lifecycle === 'replacement') {
      await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.generation())).toBeGreaterThan(retiring);
      await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.hydrated())).toBe(true);
      expect(await page.evaluate(() => window.__xlsxWorkerEditor.cell())).toBe('initial');
      await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('initial');
      expect(await page.evaluate((retired) => window.__xlsxWorkerEditor.peerEntries.filter((entry) => entry.generation !== retired), retiring)).toEqual([]);
    }
  });
}

for (const route of ['cell', 'host', 'paste'] as const) {
  for (const [input, normalized] of [['001', '1'], ['true', 'TRUE'], ['false', 'FALSE'], ["'quoted", 'quoted'],
    ['=RANDBETWEEN(1,1000000)', '=RANDBETWEEN(1,1000000)']]) {
    test(`adopts a byte-identical normalized or volatile preview (${route}, ${input})`, async ({ page, context }) => {
      await context.grantPermissions(['clipboard-read', 'clipboard-write']);
      await open(page, true);
      const clip = await page.evaluate(() => window.__xlsxWorkerEditor.cellClip());
      if (route === 'host') await page.evaluate((value) => window.__xlsxWorkerEditor.queueCellHostEdit(value), input);
      else if (route === 'paste') {
        await page.evaluate((value) => navigator.clipboard.writeText(value), input);
        await page.getByTestId('xlsx-scroll').focus();
        await page.getByTestId('xlsx-scroll').press('ControlOrMeta+v');
      } else {
        await page.getByTestId('xlsx-scroll').focus();
        await page.getByTestId('xlsx-scroll').press('F2');
        await page.getByTestId('xlsx-cell-editor').fill(input);
        await page.getByTestId('xlsx-cell-editor').press('Enter');
      }
      await expect(page.getByTestId('xlsx-commit-preview')).toHaveAttribute('data-preview-ready', 'true');
      expect(await page.evaluate(() => window.__xlsxWorkerEditor.hydrated())).toBe(false);
      const before = await page.screenshot({ clip, animations: 'disabled' });
      await page.evaluate(() => window.__xlsxWorkerEditor.releaseHydration());
      await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.adoptedSequence())).toBeGreaterThanOrEqual(1);
      await expect(page.getByTestId('xlsx-commit-preview')).toHaveCount(0);
      const after = await page.screenshot({ clip, animations: 'disabled' });
      expect(before.equals(after)).toBe(true);
      await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.cell())).toBe(normalized);
      expect(await page.evaluate(() => window.__xlsxWorkerEditor.peerEntries)).toHaveLength(1);
      expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
    });
  }
}

test('shows no speculative canvas for a host batch containing a style step', async ({ page }) => {
  await open(page, true);
  await page.evaluate(() => window.__xlsxWorkerEditor.queueStyledBatch());
  await page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.hydrated())).toBe(false);
  await expect(page.getByTestId('xlsx-commit-preview')).toHaveCount(0);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.peerEntries)).toEqual([]);
  await page.evaluate(() => window.__xlsxWorkerEditor.releaseHydration());
  await page.evaluate(() => window.__xlsxWorkerEditor.flush());
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.adoptedSequence())).toBeGreaterThanOrEqual(1);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.cell())).toBe('styled batch');
  await expect(page.getByTestId('xlsx-commit-preview')).toHaveCount(0);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.peerEntries[0])).toMatchObject({ method: 'applyEdits', args: [{
    steps: [{ op: 'setCellInputs', inputs: [['styled batch']] }, { op: 'patchStyle', patch: { fontSize: 24 } }],
  }] });
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
});
