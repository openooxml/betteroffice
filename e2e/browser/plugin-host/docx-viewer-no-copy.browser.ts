import { expect, test, type Page } from 'playwright/test';

interface PersistedParagraph {
  kind: 'persisted';
  story: string;
  paraId: string;
}

interface ParagraphIdentity {
  session?: { story: string };
  persisted?: PersistedParagraph;
}

interface ProposalSnapshot {
  version: string;
  previewVersion: number;
}

type ProposalResult =
  | { ok: true; snapshot: ProposalSnapshot }
  | { ok: false; failure: { message: string } };

interface ViewerEditor {
  whenLayoutComplete(): Promise<number>;
  flushPendingInput(): Promise<void>;
  exportStructuredWithPages(options: { revisionView: 'markup'; stories: string[] }): Promise<
    { ok: true; content: { layout: { pages: unknown[] } } } | { ok: false }
  >;
  readSelectionInfo(): Promise<{ selectedText: string } | null>;
  listContentControls(): Promise<{ ok: boolean }>;
  commands: { execute(command: string, payload: unknown): unknown };
  getDocument(): unknown;
  getEditorRef(): unknown;
  setParagraphStyle(options: { paraId: string; styleId: string }): boolean;
  getParagraphIdentities(): Promise<{ paragraphs: ParagraphIdentity[] }>;
  resolveParagraphAnchors(anchors: PersistedParagraph[]): Promise<{ version: string }>;
  proposeChanges(request: {
    expectVersion: string;
    proposals: Array<{
      id: string;
      paragraph: PersistedParagraph;
      suggest: { author: string; date: string };
    } & (
      { op: 'replaceText'; search: string; replaceWith: string }
      | { op: 'insertText'; at: 'end'; text: string }
    )>;
  }): Promise<ProposalResult>;
  setProposalStates(request: {
    expectVersion: string;
    expectPreviewVersion: number;
    changes: Array<{ id: string; state: 'accepted' | 'rejected' }>;
  }): Promise<ProposalResult>;
  getProposals(): Promise<{ proposals: Array<{ id: string; state: string }> }>;
}

interface DocumentLoads {
  sessionsCaptured: number;
  total: number;
  sessions: Record<string, number>[];
  events: Array<{ session: number; method: string; at: number }>;
}

interface ViewerSidebarsProbe {
  editor: ViewerEditor | null;
  copies: string[];
  errors: string[];
  reportedErrors: Error[];
  sidebarOpen: boolean;
  mainDocumentLoads(): DocumentLoads;
  workersOpened(): number;
  crashResidentWorker(): void;
  saveForTest(): Promise<{
    before: DocumentLoads;
    after: DocumentLoads;
    saveStarted: number;
    byteLength: number;
    signature: number[];
    validDocument: boolean;
  }>;
}

interface ViewerWindow {
  __viewerSidebarsProbe: ViewerSidebarsProbe;
  __wasmInstantiations: number;
}

async function open(page: Page, kind = 'readOnly') {
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.addInitScript(() => {
    const target = window as unknown as ViewerWindow;
    target.__wasmInstantiations = 0;
    for (const name of ['instantiate', 'instantiateStreaming', 'compile', 'compileStreaming'] as const) {
      const original = WebAssembly[name] as (...args: unknown[]) => unknown;
      (WebAssembly as unknown as Record<string, unknown>)[name] = (...args: unknown[]) => {
        target.__wasmInstantiations += 1;
        return original.apply(WebAssembly, args);
      };
    }
  });
  await page.goto(`/docx-viewer-sidebars.html?noCopy=1&kind=${kind}`);
  await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
  await expect.poll(() => page.evaluate(() =>
    (window as unknown as ViewerWindow).__viewerSidebarsProbe.mainDocumentLoads().sessionsCaptured
  )).toBeGreaterThan(0);
  await page.evaluate(() => (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!.whenLayoutComplete());
  await expectNoCopy(page);
  return page.evaluate(() => (window as unknown as ViewerWindow).__wasmInstantiations);
}

async function expectNoCopy(page: Page, wasm?: number) {
  const status = await page.evaluate(() => {
    const target = window as unknown as ViewerWindow;
    return {
      loads: target.__viewerSidebarsProbe.mainDocumentLoads(),
      errors: target.__viewerSidebarsProbe.errors,
      wasm: target.__wasmInstantiations,
    };
  });
  expect(status.loads.sessionsCaptured).toBeGreaterThan(0);
  expect(status.loads.sessions).toHaveLength(status.loads.sessionsCaptured);
  expect(status.loads.total).toBe(0);
  expect(status.loads.events).toEqual([]);
  for (const counts of status.loads.sessions) {
    expect(counts).toEqual({ openDocx: 0, openDocxPreview: 0, loadState: 0, applyUpdate: 0 });
  }
  expect(status.errors).toEqual([]);
  if (wasm !== undefined) expect(status.wasm).toBe(wasm);
}

async function focusPages(page: Page) {
  const canvas = page.locator('canvas[data-page-index="0"]');
  await canvas.scrollIntoViewIfNeeded();
  const box = await canvas.boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.click(box!.x + box!.width * 0.2, box!.y + box!.height * (96 / 1056 + 0.008));
  await expect(page.getByTestId('yrs-input')).toBeFocused();
}

async function exportPages(page: Page) {
  return page.evaluate(async () => {
    const result = await (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!.exportStructuredWithPages({
      revisionView: 'markup', stories: ['body'],
    });
    return { ok: result.ok, pages: result.ok ? result.content.layout.pages.length : 0 };
  });
}

for (const kind of ['readOnly', 'viewing']) {
  test(`a ${kind} viewer keeps every main session empty across reads, proposals and idle gates`, async ({ page }) => {
    const wasm = await open(page, kind);
    await focusPages(page);
    await page.evaluate(async () => {
      (window as unknown as ViewerWindow).__viewerSidebarsProbe.copies.length = 0;
      await navigator.clipboard.writeText('copy sentinel');
    });
    await page.keyboard.press('ControlOrMeta+A');
    await expect(page.locator('[data-testid^="selection-rect-"]').first()).toBeVisible();
    await page.keyboard.press('ControlOrMeta+C');
    await expect.poll(() => page.evaluate(async () => {
      const copies = (window as unknown as ViewerWindow).__viewerSidebarsProbe.copies;
      const clipboard = await navigator.clipboard.readText();
      return clipboard.includes('Commented text') || copies.some((text) => text.includes('Commented text'));
    })).toBe(true);
    await expectNoCopy(page, wasm);

    await page.keyboard.press('ControlOrMeta+f');
    const dialog = page.locator('.docx-find-replace-dialog');
    await expect(dialog).toBeVisible();
    const input = dialog.locator('.docx-find-replace-dialog-input').first();
    await input.fill('Commented text');
    await expect(dialog.locator('.docx-find-replace-dialog-status')).toHaveText(/\b1\b/);
    await expect.poll(() => page.evaluate(() =>
      (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!.readSelectionInfo()
    )).toMatchObject({ selectedText: 'Commented text' });
    await input.press('Escape');
    await expect(dialog).not.toBeVisible();
    await expectNoCopy(page, wasm);

    const sidebarOpen = await page.evaluate(() =>
      (window as unknown as ViewerWindow).__viewerSidebarsProbe.sidebarOpen
    );
    if (!sidebarOpen) {
      expect(await page.evaluate(() =>
        (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!.commands.execute('commentsSidebar', null)
      )).toMatchObject({ ok: true });
    }
    await expect(page.locator('.docx-comment-card')).toContainText('Check this page one comment.');
    await expect(page.getByTestId('viewer-plugin-card')).toBeVisible();
    await page.locator('.docx-outline-toggle').click();
    await expect(page.locator('.docx-outline-nav').getByText('First heading', { exact: true })).toBeVisible();
    await expect(page.locator('.docx-outline-nav').getByText('Page three heading', { exact: true })).toBeVisible();
    await expectNoCopy(page, wasm);

    expect(await exportPages(page)).toEqual({ ok: true, pages: 3 });
    expect(await page.evaluate(() =>
      (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!.listContentControls()
    )).toMatchObject({ ok: true, content: { controls: [] } });
    const refused = await page.evaluate(() => {
      const editor = (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!;
      let errorName: string | null = null;
      try {
        editor.getDocument();
      } catch (error) {
        errorName = (error as Error).name;
      }
      return {
        editorRef: editor.getEditorRef(),
        paragraphStyle: editor.setParagraphStyle({ paraId: '00000002', styleId: 'Heading1' }),
        errorName,
      };
    });
    expect(refused).toEqual({ editorRef: null, paragraphStyle: false, errorName: 'DocxAsyncOnlyError' });
    await expectNoCopy(page, wasm);

    const decisions = await page.evaluate(async () => {
      const editor = (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!;
      const identities = await editor.getParagraphIdentities();
      const first = identities.paragraphs.find((entry) => entry.session?.story === 'body')!;
      const resolved = await editor.resolveParagraphAnchors([first.persisted!]);
      const proposed = await editor.proposeChanges({
        expectVersion: resolved.version,
        proposals: [
          {
            id: 'accept-one',
            paragraph: first.persisted!,
            suggest: { author: 'Reviewer', date: '2026-10-01T00:00:00Z' },
            op: 'replaceText', search: 'First', replaceWith: 'Reviewed',
          },
          {
            id: 'reject-one',
            paragraph: first.persisted!,
            suggest: { author: 'Reviewer', date: '2026-10-01T00:00:00Z' },
            op: 'insertText', at: 'end', text: ' Extra heading.',
          },
        ],
      });
      if (!proposed.ok) throw new Error(proposed.failure.message);
      const accepted = await editor.setProposalStates({
        expectVersion: proposed.snapshot.version,
        expectPreviewVersion: proposed.snapshot.previewVersion,
        changes: [{ id: 'accept-one', state: 'accepted' }],
      });
      if (!accepted.ok) throw new Error(accepted.failure.message);
      const rejected = await editor.setProposalStates({
        expectVersion: accepted.snapshot.version,
        expectPreviewVersion: accepted.snapshot.previewVersion,
        changes: [{ id: 'reject-one', state: 'rejected' }],
      });
      if (!rejected.ok) throw new Error(rejected.failure.message);
      const snapshot = await editor.getProposals();
      return snapshot.proposals.map(({ id, state }) => ({ id, state }));
    });
    expect(decisions).toEqual([{ id: 'accept-one', state: 'accepted' }, { id: 'reject-one', state: 'rejected' }]);
    await expectNoCopy(page, wasm);
    await page.waitForTimeout(7000);
    await expectNoCopy(page, wasm);

    await page.evaluate(() => {
      Object.defineProperty(document, 'visibilityState', { configurable: true, get: () => 'hidden' });
      document.dispatchEvent(new Event('visibilitychange'));
    });
    expect(await page.evaluate(() => document.visibilityState)).toBe('hidden');
    await page.waitForTimeout(7000);
    await expectNoCopy(page, wasm);
    await page.evaluate(() => {
      delete (document as unknown as { visibilityState?: string }).visibilityState;
      document.dispatchEvent(new Event('visibilitychange'));
    });
    await focusPages(page);
    for (const key of ['ArrowRight', 'ArrowDown', 'Home', 'End', 'a', 'Backspace', 'Enter']) {
      await page.keyboard.press(key);
    }
    await page.evaluate(() => (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!.flushPendingInput());
    expect(await exportPages(page)).toEqual({ ok: true, pages: 3 });
    await expectNoCopy(page, wasm);
  });
}

test('viewer save is the only action that loads the main document and returns a valid DOCX', async ({ page }) => {
  await open(page);
  await page.waitForTimeout(7000);
  await expectNoCopy(page);
  const saved = await page.evaluate(() => (window as unknown as ViewerWindow).__viewerSidebarsProbe.saveForTest());
  expect(saved.before.sessionsCaptured).toBeGreaterThan(0);
  expect(saved.before.total).toBe(0);
  expect(saved.before.events).toEqual([]);
  expect(saved.byteLength).toBeGreaterThan(0);
  expect(saved.signature).toEqual([80, 75, 3, 4]);
  expect(saved.validDocument).toBe(true);
  expect(saved.after.total).toBeGreaterThan(0);
  expect(saved.after.events.length).toBe(saved.after.total);
  for (const event of saved.after.events) expect(event.at).toBeGreaterThanOrEqual(saved.saveStarted);
  expect(await exportPages(page)).toEqual({ ok: true, pages: 3 });
  expect(await page.evaluate(() =>
    (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!.listContentControls()
  )).toMatchObject({ ok: true });
  expect(await page.evaluate(() => {
    const editor = (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!;
    let errorName: string | null = null;
    try { editor.getDocument(); } catch (error) { errorName = (error as Error).name; }
    return { editorRef: editor.getEditorRef(), errorName };
  })).toEqual({ editorRef: null, errorName: 'DocxAsyncOnlyError' });
  await focusPages(page);
  await page.keyboard.press('ArrowRight');
  const after = await page.evaluate(() => (window as unknown as ViewerWindow).__viewerSidebarsProbe.mainDocumentLoads());
  expect(after).toEqual(saved.after);
});

async function crashAndRecover(page: Page) {
  await page.waitForTimeout(500);
  const opened = await page.evaluate(() => {
    const probe = (window as unknown as ViewerWindow).__viewerSidebarsProbe;
    const before = probe.workersOpened();
    probe.crashResidentWorker();
    return before;
  });
  await expect.poll(() => page.evaluate(() =>
    (window as unknown as ViewerWindow).__viewerSidebarsProbe.workersOpened()
  )).toBeGreaterThan(opened);
  await expect.poll(() => exportPages(page)).toEqual({ ok: true, pages: 3 });
  await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible();
  await expect(page.getByTestId('canvas-renderer-error')).toHaveCount(0);
}

test('one idle worker crash recovers with no main document copy', async ({ page }) => {
  const wasm = await open(page);
  await crashAndRecover(page);
  await focusPages(page);
  await page.keyboard.press('ArrowRight');
  await expectNoCopy(page, wasm);
});

test('a second worker crash reports DocxWorkerError and shows the same terminal failure', async ({ page }) => {
  await open(page);
  await crashAndRecover(page);
  await page.waitForTimeout(500);
  await page.evaluate(() => (window as unknown as ViewerWindow).__viewerSidebarsProbe.crashResidentWorker());
  const alert = page.getByTestId('canvas-renderer-error');
  await expect(alert).toHaveAttribute('role', 'alert');
  await expect(alert).toBeVisible();
  await expect.poll(() => page.evaluate(() =>
    (window as unknown as ViewerWindow).__viewerSidebarsProbe.reportedErrors.map((error) => error.name)
  )).toEqual(['DocxWorkerError']);
  const message = await page.evaluate(() =>
    (window as unknown as ViewerWindow).__viewerSidebarsProbe.reportedErrors[0]!.message
  );
  await expect(alert).toContainText(message);
  const loads = await page.evaluate(() => (window as unknown as ViewerWindow).__viewerSidebarsProbe.mainDocumentLoads());
  expect(loads.sessionsCaptured).toBeGreaterThan(0);
  expect(loads.total).toBe(0);
  expect(loads.events).toEqual([]);
});
