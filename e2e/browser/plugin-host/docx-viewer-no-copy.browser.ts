import { expect, test, type Page } from 'playwright/test';
import type { DocxLayoutMap, DocxPagedStructuredContent } from '../../../packages/docx/src/yrs/pagedExport';
import type { DocxExportInline } from '../../../packages/docx/src/yrs/structuredExport';

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
    { ok: true; content: DocxPagedStructuredContent<DocxLayoutMap> } | { ok: false; failure: { code: string; message: string } }
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
    changes: Array<{ id: string; state: 'proposed' | 'accepted' | 'rejected' }>;
  }): Promise<ProposalResult>;
  getProposals(): Promise<ProposalSnapshot & { proposals: Array<{ id: string; state: string }> }>;
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

async function open(page: Page, kind = 'readOnly', source?: 'bytes' | 'prop' | 'loadDocument') {
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
  const fixture = source ? `&fixture=public&source=${source}` : '';
  await page.goto(`/docx-viewer-sidebars.html?noCopy=1&kind=${kind}${fixture}`);
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
    return {
      ok: result.ok, pages: result.ok ? result.content.layout.pages.length : 0,
      failure: result.ok ? null : { code: result.failure.code, message: result.failure.message },
    };
  });
}

async function pageContents(page: Page) {
  return page.evaluate(async () => {
    const result = await (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!.exportStructuredWithPages({
      revisionView: 'markup', stories: ['body'],
    });
    if (!result.ok) throw new Error(result.failure.message);
    const { structured, layout } = result.content;
    const nodes = new Map<string, Extract<DocxExportInline, { kind: 'text' }>>();
    const collect = (value: unknown): void => {
      if (!value || typeof value !== 'object') return;
      const node = value as DocxExportInline;
      if (node.kind === 'text') nodes.set(node.id, node);
      for (const child of Object.values(value)) collect(child);
    };
    collect(structured);
    return {
      pages: layout.pages,
      text: layout.pages.map((page) => layout.fragments
        .filter((fragment) => fragment.pageIndex === page.pageIndex)
        .flatMap((fragment) => {
          const node = nodes.get(fragment.nodeId);
          if (!node || node.anchor.kind !== 'range' || fragment.slice.kind !== 'text') return [];
          return [node.text.slice(
            fragment.slice.range.start.offset - node.anchor.start.offset,
            fragment.slice.range.end.offset - node.anchor.start.offset
          )];
        }).join('')),
    };
  });
}

for (const kind of ['readOnly', 'viewing']) {
  for (const entry of ['prop', 'loadDocument'] as const) {
    test(`a ${kind} viewer opened by ${entry} renders the public fixture like its serialized bytes without a main copy`, async ({ page }) => {
      await open(page, kind, 'bytes');
      const expected = await pageContents(page);
      expect(expected.pages.length).toBeGreaterThan(0);
      expect(expected.text.join('')).toContain('Parsed document host edit.');
      await open(page, kind, entry);
      const actual = await pageContents(page);
      expect(actual.pages).toEqual(expected.pages);
      expect(actual.text).toEqual(expected.text);
      await expectNoCopy(page);
    });
  }

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

    expect(await exportPages(page)).toEqual({ ok: true, pages: 3, failure: null });
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
    expect(await exportPages(page)).toEqual({
      ok: false, pages: 0,
      failure: {
        code: 'unsupported-revision-layout',
        message: 'The retained layout previews revision decisions instead of their markup.',
      },
    });
    await expectNoCopy(page, wasm);
    const markup = await page.evaluate(async () => {
      const editor = (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!;
      const snapshot = await editor.getProposals();
      const result = await editor.setProposalStates({
        expectVersion: snapshot.version,
        expectPreviewVersion: snapshot.previewVersion,
        changes: snapshot.proposals.map(({ id }) => ({ id, state: 'proposed' })),
      });
      if (!result.ok) throw new Error(result.failure.message);
      return (await editor.getProposals()).proposals.map(({ id, state }) => ({ id, state }));
    });
    expect(markup).toEqual([{ id: 'accept-one', state: 'proposed' }, { id: 'reject-one', state: 'proposed' }]);
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
    expect(await exportPages(page)).toEqual({ ok: true, pages: 3, failure: null });
    await expectNoCopy(page, wasm);
  });
}

test('viewer save loads no main document and returns a valid DOCX', async ({ page }) => {
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
  expect(saved.after.total).toBe(0);
  expect(saved.after.events).toEqual([]);
  expect(await exportPages(page)).toEqual({ ok: true, pages: 3, failure: null });
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
  await expect.poll(() => exportPages(page)).toEqual({ ok: true, pages: 3, failure: null });
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
