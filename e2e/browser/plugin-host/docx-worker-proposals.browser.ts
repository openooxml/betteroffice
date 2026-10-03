import { test, expect, type Page } from 'playwright/test';

interface DocxSessionParagraphAnchor {
  kind: 'session';
  sessionId: string;
  story: string;
  paraId: string;
}

interface DocxPersistedParagraphAnchor {
  kind: 'persisted';
  story:
    | { partUri: string; kind: 'body' | 'header' | 'footer' }
    | { partUri: string; kind: 'footnote' | 'endnote' | 'comment'; itemId: string };
  paraId: string;
}

interface DocxParagraphIdentitySnapshot {
  sessionId: string;
  paragraphs: {
    session: DocxSessionParagraphAnchor | null;
    persisted: DocxPersistedParagraphAnchor | null;
  }[];
}

type DocxProposalState = 'proposed' | 'accepted' | 'rejected';

type DocxProposalInput = {
  id: string;
  paragraph: DocxPersistedParagraphAnchor;
  suggest: { author: string; date: string };
} & (
  | { op: 'replaceText'; search: string; replaceWith: string }
  | { op: 'insertText'; at: 'start' | 'end' | { offset: number }; text: string }
);

interface DocxProposalSnapshot {
  version: string;
  previewVersion: number;
  proposals: readonly {
    id: string;
    state: DocxProposalState;
    revisionIds: readonly string[];
    changed: boolean;
  }[];
}

type DocxProposalResult =
  | { ok: true; snapshot: DocxProposalSnapshot }
  | { ok: false; failure: { message: string } };

interface YrsLoc {
  story: string;
  paraId: string;
  offset: number;
}

interface ViewState {
  scrollTop: number;
  selection: { anchor: YrsLoc; head: YrsLoc } | null;
}

interface WorkerProposalProbe {
  editor: {
    getTotalPages(): number;
    getParagraphIdentities(): Promise<DocxParagraphIdentitySnapshot>;
    resolveParagraphAnchors(anchors: readonly DocxPersistedParagraphAnchor[]): Promise<{
      version: string;
      results: unknown[];
    }>;
    proposeChanges(request: {
      expectVersion: string;
      proposals: readonly DocxProposalInput[];
    }): Promise<DocxProposalResult>;
    setProposalStates(request: {
      expectVersion: string;
      expectPreviewVersion: number;
      changes: readonly { id: string; state: DocxProposalState }[];
    }): Promise<DocxProposalResult>;
    withdrawProposals(request: {
      expectVersion: string;
      ids: readonly string[];
    }): Promise<DocxProposalResult>;
    getProposals(): Promise<DocxProposalSnapshot>;
  } | null;
  session: { selection(): ViewState['selection'] } | null;
  status(): {
    pending: boolean | null;
    captures: number;
    hydratedBeforeSidebar: boolean;
    sidebarOpen: boolean;
    sidebarOpenChanges: boolean[];
    beforeSidebarOpen: ViewState | null;
    layoutComplete: number | null;
    renderedDomContextCalls: number;
    contentChanges: { bodyContainsProposedText: boolean }[];
    documentChanges: string[];
    load: { version: string; sessionVersion: string; snapshotVersion: string } | null;
    events: { load: number; 'proposal-change': number; 'layout-change': number };
    errors: string[];
  };
  view(): ViewState;
  navigate(target: { story: string; paraId: string }, expectVersion: string): Promise<{ ok: boolean } | null>;
  toggleSidebar(): Promise<{ ok: boolean }>;
}

interface ProbeWindow {
  __workerProposalProbe: WorkerProposalProbe;
  __workerRequests: string[];
  __workerInstrumentation: {
    firstEncodeState: number | null;
    deliberateScroll: boolean;
    scrollEvents: number;
    unexpectedScrolls: number;
  };
}

const IDS = Array.from({ length: 10 }, (_, index) => `host-${index + 1}`);
const SUGGEST = { author: 'Host reviewer', date: '2026-09-30T12:00:00Z' };

async function instrument(page: Page) {
  await page.addInitScript(() => {
    const requests: string[] = [];
    const windowProbe = window as unknown as ProbeWindow;
    windowProbe.__workerRequests = requests;
    const instrumentation = windowProbe.__workerInstrumentation = {
      firstEncodeState: null as number | null,
      deliberateScroll: false,
      scrollEvents: 0,
      unexpectedScrolls: 0,
    };
    const postMessage = Worker.prototype.postMessage;
    Worker.prototype.postMessage = function (message: unknown, ...args: unknown[]) {
      if (message && typeof message === 'object' && 'type' in message) {
        requests.push(String(message.type));
        if (message.type === 'encodeState') {
          instrumentation.firstEncodeState ??= performance.now();
        }
      }
      return Reflect.apply(postMessage, this, [message, ...args]);
    };
    const scrollPositions = new WeakMap<HTMLElement, number>();
    document.addEventListener('scroll', (event) => {
      const scroller = event.target;
      if (!(scroller instanceof HTMLElement) || !scroller.matches('.docx-editor__scroll-container')) return;
      const previous = scrollPositions.get(scroller) ?? 0;
      if (!instrumentation.deliberateScroll && Math.abs(scroller.scrollTop - previous) > 1) {
        instrumentation.unexpectedScrolls += 1;
      }
      scrollPositions.set(scroller, scroller.scrollTop);
      instrumentation.scrollEvents += 1;
    }, true);
  });
}

async function status(page: Page) {
  return page.evaluate(() => {
    const windowProbe = window as unknown as ProbeWindow;
    return {
      ...windowProbe.__workerProposalProbe.status(),
      ...windowProbe.__workerInstrumentation,
      encodeState: windowProbe.__workerRequests.filter((type) => type === 'encodeState').length,
      revisionCount: windowProbe.__workerRequests.filter((type) => type === 'revisionCount').length,
      openedInWorker: windowProbe.__workerRequests.includes('open'),
    };
  });
}

async function assertReplica(page: Page, readOnly: boolean, sidebarOpen = false) {
  const current = await status(page);
  expect(current.openedInWorker).toBe(true);
  expect(current.captures).toBe(1);
  expect(current.errors).toEqual([]);
  expect(current.pending).toBe(readOnly);
  expect(current.encodeState).toBe(readOnly ? 0 : 1);
  expect(current.sidebarOpen).toBe(sidebarOpen);
  if (readOnly) {
    expect(current.hydratedBeforeSidebar).toBe(false);
    expect(current.layoutComplete).not.toBeNull();
  }
}

async function view(page: Page) {
  const scroller = page.locator('.docx-editor__scroll-container');
  return scroller.evaluate((element) => ({
    scrollTop: element.scrollTop,
    selection: (window as unknown as ProbeWindow).__workerProposalProbe.session!.selection(),
  }));
}

async function assertHydration(
  page: Page,
  before: ReturnType<WorkerProposalProbe['view']>,
  options: { checkLayoutOrder?: boolean } = {}
) {
  await expect.poll(async () => {
    const current = await status(page);
    return { encodeState: current.encodeState, pending: current.pending };
  }).toEqual({ encodeState: 1, pending: false });
  const current = await status(page);
  expect(current.layoutComplete).not.toBeNull();
  expect(current.firstEncodeState).not.toBeNull();
  if (options.checkLayoutOrder !== false) {
    expect(current.firstEncodeState!).toBeGreaterThan(current.layoutComplete!);
  }
  const after = await view(page);
  expect(Math.abs(after.scrollTop - before.scrollTop)).toBeLessThanOrEqual(1);
  expect(after.selection).toEqual(before.selection);
}

async function openEditor(page: Page, readOnly: boolean, options = '') {
  await page.goto(`/docx-worker-proposals.html?readOnly=${readOnly}&${options}`);
  await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
  await expect.poll(async () => (await status(page)).load).not.toBeNull();
  if (!readOnly) {
    await expect.poll(async () => (await status(page)).pending).toBe(false);
  }
  const current = await status(page);
  expect(current.load!.version).toBe(current.load!.sessionVersion);
  expect(current.load!.version).toBe(current.load!.snapshotVersion);
  expect(current.events.load).toBeGreaterThan(0);
  await expect.poll(() => page.evaluate(() =>
    (window as unknown as ProbeWindow).__workerProposalProbe.editor!.getTotalPages()
  )).toBe(6);
  await expect.poll(async () => (await status(page)).layoutComplete).not.toBeNull();
  await expect.poll(async () => (await status(page)).renderedDomContextCalls).toBeGreaterThan(0);
}

async function open(page: Page, readOnly: boolean, options = '') {
  await openEditor(page, readOnly, options);
  await assertReplica(page, readOnly);
}

async function getProposals(page: Page) {
  return page.evaluate(() =>
    (window as unknown as ProbeWindow).__workerProposalProbe.editor!.getProposals()
  );
}

async function prepare(page: Page, readOnly: boolean) {
  const identities = await page.evaluate(() =>
    (window as unknown as ProbeWindow).__workerProposalProbe.editor!.getParagraphIdentities()
  );
  const afterIdentities = await getProposals(page);
  await assertReplica(page, readOnly);
  const body = identities.paragraphs.filter((paragraph) => paragraph.session?.story === 'body');
  expect(body).toHaveLength(72);
  const anchors = body.slice(0, 10).map((paragraph) => paragraph.persisted!);
  expect(anchors.every(Boolean)).toBe(true);
  const resolved = await page.evaluate(
    (input) => (window as unknown as ProbeWindow).__workerProposalProbe.editor!.resolveParagraphAnchors(input),
    anchors
  );
  expect(resolved.results).toHaveLength(10);
  expect(resolved.results).toEqual(body.slice(0, 10).map((paragraph) => ({
    status: 'found',
    anchor: paragraph.session,
  })));
  const afterAnchors = await getProposals(page);
  expect(afterAnchors).toEqual(afterIdentities);
  await assertReplica(page, readOnly);
  const proposals: DocxProposalInput[] = anchors.map((paragraph, index) =>
    index % 2 === 0
      ? {
          id: IDS[index],
          paragraph,
          suggest: SUGGEST,
          op: 'replaceText',
          search: 'Original',
          replaceWith: `Reviewed ${index + 1}`,
        }
      : {
          id: IDS[index],
          paragraph,
          suggest: SUGGEST,
          op: 'insertText',
          at: 'end',
          text: ` Added ${index + 1}.`,
        }
  );
  return { identities, body, resolved, proposals, afterIdentities, afterAnchors };
}

function snapshotOf(result: DocxProposalResult): DocxProposalSnapshot {
  expect(result.ok).toBe(true);
  if (!result.ok) throw new Error(result.failure.message);
  return result.snapshot;
}

async function overlay(
  page: Page,
  snapshot: DocxProposalSnapshot,
  afterEventSerial = -1,
  requireSuccess = true
) {
  const layer = page.getByTestId('proposal-overlay');
  await expect(layer).toHaveAttribute('data-version', snapshot.version);
  await expect(layer).toHaveAttribute('data-preview-version', String(snapshot.previewVersion));
  await expect(layer).toHaveAttribute('data-layout-version', snapshot.version);
  await expect(layer).toHaveAttribute('data-layout-preview-version', String(snapshot.previewVersion));
  await expect.poll(async () => Number(await layer.getAttribute('data-event-serial'))).toBeGreaterThan(afterEventSerial);
  const expected = snapshot.proposals.map(({ id, state }) => ({ id, state, ok: true, failure: '' }));
  const read = () => layer.locator('[data-proposal-id]').evaluateAll((elements) => elements.map((element) => ({
    id: element.getAttribute('data-proposal-id'),
    state: element.getAttribute('data-state'),
    ok: element.getAttribute('data-ok') === 'true',
    failure: element.getAttribute('data-failure'),
  })));
  await expect.poll(async () => (await read()).map(({ id, state }) => ({ id, state }))).toEqual(
    expected.map(({ id, state }) => ({ id, state }))
  );
  if (requireSuccess) {
    await expect.poll(read).toEqual(expected);
  } else {
    let previous = '';
    await expect.poll(async () => {
      const current = JSON.stringify(await read());
      const settled = current === previous;
      previous = current;
      return settled;
    }).toBe(true);
  }
  return read();
}

async function eventSerial(page: Page) {
  return Number(await page.getByTestId('proposal-overlay').getAttribute('data-event-serial'));
}

function normalize(identities: DocxParagraphIdentitySnapshot) {
  const versions = new Map<string, number>();
  const revisions = new Map<string, number>();
  const sessions = new Map([[identities.sessionId, 0]]);
  const paragraphs = new Map(identities.paragraphs.flatMap((paragraph, index) =>
    paragraph.session ? [[paragraph.session.paraId, index] as const] : []
  ));
  const token = (map: Map<string, number>, value: string) => {
    if (!map.has(value)) map.set(value, map.size);
    return map.get(value);
  };
  const visit = (value: unknown, key = ''): unknown => {
    if (typeof value === 'string') {
      if (key === 'version') return token(versions, value);
      if (key === 'sessionId') return token(sessions, value);
      if (key === 'paraId' && paragraphs.has(value)) return paragraphs.get(value);
      if (key === 'revisionIds') return token(revisions, value);
      return value;
    }
    if (Array.isArray(value)) return value.map((entry) => visit(entry, key));
    if (value && typeof value === 'object') {
      return Object.fromEntries(Object.entries(value).map(([name, entry]) => [name, visit(entry, name)]));
    }
    return value;
  };
  return visit;
}

async function flow(page: Page, readOnly: boolean, options = '') {
  await open(page, readOnly, options);
  if (options.includes('revisions=1')) {
    await expect.poll(async () => (await status(page)).sidebarOpenChanges).toEqual([true]);
    expect((await status(page)).revisionCount).toBe(1);
    await assertReplica(page, readOnly);
  }
  const initial = await getProposals(page);
  expect(initial.proposals).toEqual([]);
  await assertReplica(page, readOnly);
  const prepared = await prepare(page, readOnly);
  expect(prepared.resolved.version).toBe(initial.version);
  const canonical = normalize(prepared.identities);
  expect(prepared.afterIdentities).toEqual(initial);
  const transcript: unknown[] = [
    canonical(initial),
    canonical(prepared.afterIdentities),
    canonical(prepared.resolved),
    canonical(prepared.afterAnchors),
  ];
  const record = async (name: string, result: DocxProposalResult, serial: number, sidebarOpen = false) => {
    const snapshot = snapshotOf(result);
    await assertReplica(page, readOnly, sidebarOpen);
    const read = await getProposals(page);
    expect(read).toEqual(snapshot);
    const geometry = await overlay(page, read, serial);
    await assertReplica(page, readOnly, sidebarOpen);
    transcript.push({ name, result: canonical(result), read: canonical(read), geometry });
    return read;
  };
  const serial = await eventSerial(page);
  const result = await page.evaluate(
    (request) => (window as unknown as ProbeWindow).__workerProposalProbe.editor!.proposeChanges(request),
    { expectVersion: prepared.resolved.version, proposals: prepared.proposals }
  );
  let snapshot = await record('propose', result, serial);
  expect(snapshot.version).not.toBe(initial.version);
  expect(snapshot.proposals.map(({ id }) => id)).toEqual(IDS);
  expect(snapshot.proposals.every((proposal) => proposal.changed && proposal.state === 'proposed')).toBe(true);
  expect(snapshot.proposals.map(({ revisionIds }) => revisionIds.length)).toEqual([1, 1, 1, 1, 1, 1, 1, 1, 1, 1]);
  const setStates = async (name: string, ids: string[], state: DocxProposalState) => {
    const before = snapshot;
    const serial = await eventSerial(page);
    const changes = ids.map((id) => ({ id, state }));
    const result = await page.evaluate(
      (request) => (window as unknown as ProbeWindow).__workerProposalProbe.editor!.setProposalStates(request),
      { expectVersion: before.version, expectPreviewVersion: before.previewVersion, changes }
    );
    snapshot = await record(name, result, serial);
    expect(snapshot.version).toBe(before.version);
    expect(snapshot.previewVersion).toBe(before.previewVersion + 1);
    expect(snapshot.proposals).toEqual(before.proposals.map((proposal) =>
      ids.includes(proposal.id) ? { ...proposal, state } : proposal
    ));
  };
  for (const id of IDS.slice(0, 4)) {
    for (const state of ['accepted', 'rejected', 'proposed'] as const) {
      await setStates(`${id}-${state}`, [id], state);
    }
  }
  await setStates('accept-all', IDS, 'accepted');
  const scroller = page.locator('.docx-editor__scroll-container');
  const scrollTop = await scroller.evaluate((element) => element.scrollTop);
  for (const [name, paragraph] of [
    ['end', prepared.body.at(-2)!],
    ['start', prepared.body[0]],
  ] as const) {
    const target = paragraph.session!;
    const scrollEvents = await page.evaluate(() => {
      const instrumentation = (window as unknown as ProbeWindow).__workerInstrumentation;
      instrumentation.deliberateScroll = true;
      return instrumentation.scrollEvents;
    });
    const result = await page.evaluate(
      ({ target, version }) => (window as unknown as ProbeWindow).__workerProposalProbe.navigate(target, version),
      { target: { story: target.story, paraId: target.paraId }, version: snapshot.version }
    );
    expect(result).toEqual({ ok: true });
    if (name === 'end') {
      await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeGreaterThan(scrollTop + 1000);
      await expect(page.locator('canvas[data-page-index="5"]')).toBeInViewport();
    } else {
      await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeLessThan(scrollTop + 100);
      await expect(page.locator('canvas[data-page-index="0"]')).toBeInViewport();
    }
    await expect.poll(async () => (await status(page)).scrollEvents).toBeGreaterThan(scrollEvents);
    await page.evaluate(() => new Promise<void>((resolve) => {
      const scroller = document.querySelector<HTMLElement>('.docx-editor__scroll-container')!;
      let previous = scroller.scrollTop;
      let stableFrames = 0;
      const sample = () => {
        const current = scroller.scrollTop;
        stableFrames = current === previous ? stableFrames + 1 : 0;
        previous = current;
        if (stableFrames < 10) {
          requestAnimationFrame(sample);
          return;
        }
        (window as unknown as ProbeWindow).__workerInstrumentation.deliberateScroll = false;
        resolve();
      };
      requestAnimationFrame(sample);
    }));
    const read = await getProposals(page);
    expect(read).toEqual(snapshot);
    await assertReplica(page, readOnly);
    const geometry = await overlay(page, read, -1, name === 'start');
    await assertReplica(page, readOnly);
    transcript.push({ name: `navigate-${name}`, result, read: canonical(read), geometry });
  }
  const withdrawalSerial = await eventSerial(page);
  const withdrawn = await page.evaluate(
    (request) => (window as unknown as ProbeWindow).__workerProposalProbe.editor!.withdrawProposals(request),
    { expectVersion: snapshot.version, ids: IDS }
  );
  const sidebarOpenAfterWithdrawal = !readOnly;
  await expect.poll(async () => (await status(page)).sidebarOpen).toBe(sidebarOpenAfterWithdrawal);
  const final = await record('withdraw', withdrawn, withdrawalSerial, sidebarOpenAfterWithdrawal);
  expect(final.proposals).toEqual([]);
  expect(final.version).not.toBe(snapshot.version);
  expect((await status(page)).events['proposal-change']).toBeGreaterThanOrEqual(15);
  expect((await status(page)).events['layout-change']).toBeGreaterThan(0);
  await assertReplica(page, readOnly, sidebarOpenAfterWithdrawal);
  expect((await status(page)).unexpectedScrolls).toBe(0);
  return transcript;
}

test('host proposals, overlay geometry and plugin navigation match the hydrated editor without hydrating the worker stub', async ({ page }) => {
  test.setTimeout(360_000);
  await instrument(page);
  const worker = await flow(page, true);
  const main = await flow(page, false);
  expect(worker).toEqual(main);
  expect((await status(page)).unexpectedScrolls).toBe(0);
});

test('opening the built-in sidebar hydrates the worker replica once and preserves host proposal records', async ({ page }) => {
  await instrument(page);
  await open(page, true);
  const prepared = await prepare(page, true);
  const result = await page.evaluate(
    (request) => (window as unknown as ProbeWindow).__workerProposalProbe.editor!.proposeChanges(request),
    { expectVersion: prepared.resolved.version, proposals: prepared.proposals }
  );
  const before = snapshotOf(result);
  expect(before.proposals).toHaveLength(10);
  await overlay(page, before);
  expect(await getProposals(page)).toEqual(before);
  await assertReplica(page, true);
  const openSidebar = () => page.evaluate(() =>
    (window as unknown as ProbeWindow).__workerProposalProbe.toggleSidebar()
  );
  const beforeHydration = await view(page);
  expect(await openSidebar()).toMatchObject({ ok: true });
  await expect.poll(async () => (await status(page)).sidebarOpen).toBe(true);
  await expect.poll(async () => (await status(page)).pending).toBe(false);
  expect((await status(page)).encodeState).toBe(1);
  await assertHydration(page, beforeHydration);
  const after = await getProposals(page);
  expect(after.previewVersion).toBe(before.previewVersion);
  expect(after.proposals).toEqual(before.proposals);
  const sidebar = page.locator('.docx-unified-sidebar');
  await expect(sidebar).toBeAttached();
  await expect(sidebar).toHaveCSS('opacity', '1');
  const cards = sidebar.locator('.docx-tracked-change-card');
  await expect(cards).toHaveCount(after.proposals.length);
  for (let index = 0; index < after.proposals.length; index += 1) {
    await expect(cards.nth(index)).toBeVisible();
    await expect(cards.nth(index)).toContainText(SUGGEST.author);
  }
  await overlay(page, after);
  expect(await openSidebar()).toMatchObject({ ok: true });
  await expect.poll(async () => (await status(page)).sidebarOpen).toBe(false);
  expect(await openSidebar()).toMatchObject({ ok: true });
  await expect.poll(async () => (await status(page)).sidebarOpen).toBe(true);
  expect(await getProposals(page)).toEqual(after);
  const current = await status(page);
  expect(current.encodeState).toBe(1);
  expect(current.hydratedBeforeSidebar).toBe(false);
  expect(current.errors).toEqual([]);
  expect(current.unexpectedScrolls).toBe(0);
});

test('existing revisions keep the controlled sidebar and host replica closed', async ({ page }) => {
  test.setTimeout(360_000);
  await instrument(page);
  await flow(page, true, 'revisions=1&sidebar=closed');
  const current = await status(page);
  expect(current.sidebarOpenChanges).toEqual([true]);
  expect(current.revisionCount).toBe(1);
  expect(current.sidebarOpen).toBe(false);
  expect(current.pending).toBe(true);
  expect(current.encodeState).toBe(0);
  expect(current.layoutComplete).not.toBeNull();
  expect(current.hydratedBeforeSidebar).toBe(false);
  expect(current.renderedDomContextCalls).toBeGreaterThan(0);
  expect(current.errors).toEqual([]);
  await expect(page.locator('.docx-unified-sidebar')).toHaveCount(0);
  expect((await status(page)).unexpectedScrolls).toBe(0);
});

test('existing revisions open the sidebar and hydrate once', async ({ page }) => {
  await instrument(page);
  await openEditor(page, true, 'revisions=1');
  await expect.poll(async () => (await status(page)).sidebarOpen).toBe(true);
  await expect.poll(async () => (await status(page)).pending).toBe(false);
  const beforeHydration = (await status(page)).beforeSidebarOpen;
  expect(beforeHydration).not.toBeNull();
  await assertHydration(page, beforeHydration!, { checkLayoutOrder: false });
  const requests = await page.evaluate(() => (window as unknown as ProbeWindow).__workerRequests);
  expect(requests.indexOf('revisionCount')).toBeGreaterThanOrEqual(0);
  expect(requests.indexOf('encodeState')).toBeGreaterThan(requests.indexOf('revisionCount'));
  const sidebar = page.locator('.docx-unified-sidebar');
  await expect(sidebar).toBeAttached();
  await expect(sidebar).toHaveCSS('opacity', '1');
  await expect(sidebar.locator('.docx-tracked-change-card')).toBeVisible();
  await expect(page.locator('.docx-tracked-change-card')).toHaveCount(1);
  await expect(page.locator('.docx-tracked-change-card')).toContainText('Document reviewer');
  expect((await getProposals(page)).proposals).toEqual([]);
  const current = await status(page);
  expect(current.openedInWorker).toBe(true);
  expect(current.captures).toBe(1);
  expect(current.sidebarOpenChanges).toEqual([true]);
  expect(current.revisionCount).toBe(1);
  expect(current.encodeState).toBe(1);
  expect(current.hydratedBeforeSidebar).toBe(false);
  expect(current.errors).toEqual([]);
  expect(current.unexpectedScrolls).toBe(0);
});

test('a worker viewer reports a proposal through onDocumentChange without opening the main-thread copy', async ({ page }) => {
  await instrument(page);
  await open(page, true, 'onChange=1');
  const prepared = await prepare(page, true);
  const result = await page.evaluate(
    (request) =>
      (window as unknown as ProbeWindow).__workerProposalProbe.editor!.proposeChanges(request),
    { expectVersion: prepared.resolved.version, proposals: prepared.proposals.slice(0, 1) }
  );
  const proposed = snapshotOf(result);
  expect(proposed.proposals.map(({ id }) => id)).toEqual([IDS[0]]);
  expect(proposed.proposals[0]!.changed).toBe(true);
  await expect.poll(async () => (await status(page)).documentChanges.length).toBeGreaterThan(0);
  const after = await getProposals(page);
  expect(after.previewVersion).toBe(proposed.previewVersion);
  expect(after.proposals).toEqual(proposed.proposals);
  await overlay(page, after);
  expect(await getProposals(page)).toEqual(after);
  expect(await getProposals(page)).toEqual(after);
  const current = await status(page);
  expect(current.openedInWorker).toBe(true);
  expect(current.captures).toBe(1);
  expect(current.pending).toBe(true);
  expect(current.encodeState).toBe(0);
  expect(current.contentChanges).toEqual([]);
  expect(current.sidebarOpen).toBe(false);
  expect(current.errors).toEqual([]);
  expect(current.unexpectedScrolls).toBe(0);
});
