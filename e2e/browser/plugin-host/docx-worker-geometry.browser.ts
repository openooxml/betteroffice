import { test, expect, type Page } from 'playwright/test';

interface SourceAnchor {
  kind: 'source';
  packageSha256: string;
  partUri: string;
  paragraphOrdinal: number;
}

interface Paragraph {
  story: string;
  paraId: string;
  text: string;
  anchor: SourceAnchor | null;
}

interface TextRange {
  story: string;
  start: { paraId: string; offset: number };
  end: { paraId: string; offset: number };
  view: 'accepted';
}

type GeometryTarget =
  | { kind: 'paragraph'; paragraph: SourceAnchor }
  | { kind: 'search'; paragraph: SourceAnchor; text: string; occurrence?: 'first' | 'all' }
  | { kind: 'range'; version: string; range: TextRange }
  | { kind: 'revision'; revisionId: string }
  | { kind: 'proposal'; id: string };

interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

type GeometryResult =
  | {
      ok: true;
      layoutId: string;
      version: string;
      previewVersion: number;
      rects: (Rect & { pageIndex: number })[];
      anchor: Rect & { pageIndex: number };
      pageRect: Rect;
    }
  | { ok: false; failure: { code: string; message: string } };

type ProposalInput = {
  id: string;
  paragraph: SourceAnchor;
  suggest: { author: string; date: string };
} & (
  | { op: 'replaceText'; search: string; replaceWith: string; occurrence: 'first' | 'all' }
  | { op: 'insertText'; at: 'end'; text: string }
);

interface ProposalSnapshot {
  version: string;
  previewVersion: number;
  proposals: { id: string; state: 'proposed' | 'accepted' | 'rejected'; revisionIds: string[] }[];
}

type ProposalResult =
  | { ok: true; snapshot: ProposalSnapshot }
  | { ok: false; failure: { code: string; message: string } };

interface GeometryProbe {
  editor: {
    getTotalPages(): number;
    scrollToPage(pageNumber: number): void;
    proposeChanges(request: { expectVersion: string; proposals: ProposalInput[] }): Promise<ProposalResult>;
    getProposals(): Promise<ProposalSnapshot>;
    setProposalStates(request: {
      expectVersion: string;
      expectPreviewVersion: number;
      changes: { id: string; state: 'accepted' }[];
    }): Promise<ProposalResult>;
  } | null;
  context: {
    snapshot: { version: string; readOnly: boolean };
    geometry: { layout: { version: string; previewVersion: number } } | null;
  } | null;
  layoutComplete: number | null;
  pending(): boolean | null;
  events: { load: number; 'proposal-change': number; 'layout-change': number };
  errors: string[];
  readParagraphs(): Promise<{ version: string; paragraphs: Paragraph[] }>;
  anchorGeometry(targets: GeometryTarget[]): (GeometryResult | null)[];
}

interface ProbeWindow {
  __workerGeometryProbe: GeometryProbe;
}

interface TargetEntry {
  name: string;
  target: GeometryTarget;
  failure?: 'missing-target' | 'stale-version';
}

interface ArmTranscript {
  worker: boolean;
  paragraphs?: { ordinal: number; length: number }[];
  proposals?: ProposalInput[];
  targets?: string[];
  scrolls?: { proposed: number[]; accepted?: number[] };
  pending?: { proposed: boolean | null; afterProposedReads?: boolean | null; afterAcceptedReads?: boolean | null };
  firstAttempt?: { proposed?: (string | null)[]; accepted?: (string | null)[] };
  geometry?: { proposed?: ReturnType<typeof normalize>; accepted?: ReturnType<typeof normalize> };
  events?: GeometryProbe['events'];
  errors?: string[];
}

const IDS = ['geometry-replace', 'geometry-insert', 'geometry-all'];
const SUGGEST = { author: 'Reviewer', date: '2026-09-30T12:00:00Z' };

async function status(page: Page) {
  return page.evaluate(() => {
    const probe = (window as unknown as ProbeWindow).__workerGeometryProbe;
    return {
      pending: probe.pending(),
      layout: probe.context?.geometry?.layout ?? null,
      readOnly: probe.context?.snapshot.readOnly,
      events: { ...probe.events },
      errors: [...probe.errors],
    };
  });
}

function word(text: string) {
  const words = text.match(/\b[a-z]{3,}\b/g) ?? [];
  expect(words.length).toBeGreaterThan(0);
  return [...words].sort((a, b) => text.split(b).length - text.split(a).length)[0]!;
}

function pickParagraphs(paragraphs: Paragraph[]) {
  const candidates = paragraphs.filter((paragraph) =>
    paragraph.anchor && paragraph.text.length >= 12 && /\b[a-z]{3,}\b/.test(paragraph.text)
  );
  expect(candidates.length).toBeGreaterThanOrEqual(6);
  const picked = Array.from({ length: 6 }, (_, index) =>
    candidates[Math.round(index * (candidates.length - 1) / 5)]!
  );
  if (picked[2].text.split(word(picked[2].text)).length <= 2) {
    const repeated = candidates.filter((candidate) =>
      !picked.includes(candidate) && candidate.text.split(word(candidate.text)).length > 2
    ).sort((a, b) =>
      Math.abs(a.anchor!.paragraphOrdinal - picked[2].anchor!.paragraphOrdinal) -
      Math.abs(b.anchor!.paragraphOrdinal - picked[2].anchor!.paragraphOrdinal)
    );
    expect(repeated.length).toBeGreaterThan(0);
    picked[2] = repeated[0];
  }
  return picked;
}

function proposalsFor(paragraphs: Paragraph[]): ProposalInput[] {
  expect(paragraphs[2].text.split(word(paragraphs[2].text)).length).toBeGreaterThan(2);
  return [
    {
      id: IDS[0], paragraph: paragraphs[0].anchor!, suggest: SUGGEST,
      op: 'replaceText', search: word(paragraphs[0].text), replaceWith: 'reviewed', occurrence: 'first',
    },
    {
      id: IDS[1], paragraph: paragraphs[1].anchor!, suggest: SUGGEST,
      op: 'insertText', at: 'end', text: ' Added text.',
    },
    {
      id: IDS[2], paragraph: paragraphs[2].anchor!, suggest: SUGGEST,
      op: 'replaceText', search: word(paragraphs[2].text), replaceWith: 'revised', occurrence: 'all',
    },
  ];
}

function targetsFor(paragraphs: Paragraph[], snapshot: ProposalSnapshot, staleVersion: string): TargetEntry[] {
  const paragraph = paragraphs.slice(3).find((candidate) => candidate.text.split(word(candidate.text)).length > 2)
    ?? paragraphs[3];
  const text = word(paragraph.text);
  const range: TextRange = {
    story: paragraph.story,
    start: { paraId: paragraph.paraId, offset: 0 },
    end: { paraId: paragraph.paraId, offset: Math.min(12, paragraph.text.length) },
    view: 'accepted',
  };
  return [
    ...paragraphs.map((candidate, index): TargetEntry => ({
      name: `paragraph-${index + 1}`,
      target: { kind: 'paragraph', paragraph: candidate.anchor! },
    })),
    { name: 'search-first', target: { kind: 'search', paragraph: paragraph.anchor!, text, occurrence: 'first' } },
    { name: 'search-all', target: { kind: 'search', paragraph: paragraph.anchor!, text, occurrence: 'all' } },
    {
      name: 'search-absent', failure: 'missing-target',
      target: { kind: 'search', paragraph: paragraph.anchor!, text: 'absent-geometry-probe-text' },
    },
    { name: 'range-current', target: { kind: 'range', version: snapshot.version, range } },
    { name: 'range-stale', failure: 'stale-version', target: { kind: 'range', version: staleVersion, range } },
    ...IDS.map((id): TargetEntry => {
      const proposal = snapshot.proposals.find((record) => record.id === id)!;
      expect(proposal.revisionIds.length).toBeGreaterThan(0);
      return { name: `revision-${id}`, target: { kind: 'revision', revisionId: proposal.revisionIds[0] } };
    }),
    ...IDS.map((id): TargetEntry => ({ name: `proposal-${id}`, target: { kind: 'proposal', id } })),
  ];
}

function normalize(results: (GeometryResult | null)[]) {
  return results.map((result) => {
    if (!result?.ok) return result;
    const { layoutId: _layoutId, version: _version, ...geometry } = result;
    return geometry;
  });
}

function snapshotOf(result: ProposalResult) {
  expect(result.ok, JSON.stringify(result)).toBe(true);
  if (!result.ok) throw new Error(result.failure.message);
  return result.snapshot;
}

async function getProposals(page: Page) {
  return page.evaluate(() => (window as unknown as ProbeWindow).__workerGeometryProbe.editor!.getProposals());
}

async function waitForPreview(page: Page, snapshot: ProposalSnapshot) {
  await expect.poll(async () => (await status(page)).layout).toMatchObject({
    version: snapshot.version,
    previewVersion: snapshot.previewVersion,
  });
}

async function renderPages(page: Page, plan?: number[]) {
  const pages = await page.evaluate(() => (window as unknown as ProbeWindow).__workerGeometryProbe.editor!.getTotalPages());
  const scrolls: number[] = [];
  for (let pageNumber = 1; pageNumber <= pages; pageNumber += 1) {
    const canvas = page.locator(`canvas[data-page-index="${pageNumber - 1}"]`);
    if (plan ? plan.includes(pageNumber) : !(await canvas.isVisible())) {
      await page.evaluate((target) => (window as unknown as ProbeWindow).__workerGeometryProbe.editor!.scrollToPage(target), pageNumber);
      scrolls.push(pageNumber);
    }
    await expect(canvas).toBeVisible();
  }
  if (scrolls.length > 0) {
    await page.evaluate(() => (window as unknown as ProbeWindow).__workerGeometryProbe.editor!.scrollToPage(1));
    await expect(page.locator('canvas[data-page-index="0"]')).toBeInViewport();
  }
  return scrolls;
}

async function geometry(
  page: Page,
  entries: TargetEntry[],
  snapshot: ProposalSnapshot,
  transcript: ArmTranscript,
  phase: 'proposed' | 'accepted'
) {
  const targets = entries.map(({ target }) => target);
  const read = () => page.evaluate((input) =>
    (window as unknown as ProbeWindow).__workerGeometryProbe.anchorGeometry(input), targets
  );
  const before = await status(page);
  const first = await read();
  transcript.firstAttempt ??= {};
  transcript.firstAttempt[phase] = first.map((result) => result?.ok ? 'ok' : result?.failure.code ?? null);
  transcript.geometry ??= {};
  let results = first;
  await expect.poll(async () => {
    results = await read();
    transcript.geometry![phase] = normalize(results);
    return results.filter((result) => !result || (!result.ok && result.failure.code === 'layout-unavailable')).length;
  }, { timeout: 30_000 }).toBe(0);
  if (transcript.worker && first.some((result) => result && !result.ok && result.failure.code === 'layout-unavailable')) {
    await expect.poll(async () => (await status(page)).events['layout-change']).toBeGreaterThan(before.events['layout-change']);
  }
  for (const [index, entry] of entries.entries()) {
    const result = results[index]!;
    if (entry.failure) {
      expect(result, entry.name).toMatchObject({ ok: false, failure: { code: entry.failure } });
    } else {
      expect(result, entry.name).toMatchObject({ ok: true, version: snapshot.version, previewVersion: snapshot.previewVersion });
      if (result?.ok) {
        expect(result.pageRect.width, entry.name).toBeGreaterThan(0);
        expect(result.pageRect.height, entry.name).toBeGreaterThan(0);
      }
    }
  }
  const paragraphPages = results.slice(0, 6).flatMap((result) => {
    expect(result?.ok).toBe(true);
    if (!result?.ok) return [];
    expect(result.rects.length).toBeGreaterThan(0);
    return result.rects.map(({ pageIndex }) => pageIndex);
  });
  expect(new Set(paragraphPages).size).toBeGreaterThanOrEqual(2);
  const after = await status(page);
  expect.soft(after.pending, `${phase}: replica pending after geometry`).toBe(transcript.worker);
  expect(after.errors).toEqual([]);
}

test('worker-held proposals preserve plugin anchor geometry without hydrating the read-only replica', async ({ page }) => {
  const arms: ArmTranscript[] = [];
  for (const worker of [false, true]) {
    const transcript: ArmTranscript = { worker };
    arms.push(transcript);
    try {
      await page.goto(`/docx-worker-geometry.html?worker=${worker ? 1 : 0}`);
      await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
      await expect.poll(() => page.evaluate(() => {
        const probe = (window as unknown as ProbeWindow).__workerGeometryProbe;
        return !!probe.editor && !!probe.context?.geometry && probe.layoutComplete !== null;
      })).toBe(true);
      expect((await status(page)).readOnly).toBe(true);
      const read = await page.evaluate(() => (window as unknown as ProbeWindow).__workerGeometryProbe.readParagraphs());
      const paragraphs = pickParagraphs(read.paragraphs);
      const proposals = proposalsFor(paragraphs);
      transcript.paragraphs = paragraphs.map((paragraph) => ({ ordinal: paragraph.anchor!.paragraphOrdinal, length: paragraph.text.length }));
      transcript.proposals = proposals;
      const proposed = snapshotOf(await page.evaluate((request) =>
        (window as unknown as ProbeWindow).__workerGeometryProbe.editor!.proposeChanges(request),
        { expectVersion: read.version, proposals }
      ));
      await expect.poll(async () => (await getProposals(page)).proposals.map(({ id }) => id).sort()).toEqual([...IDS].sort());
      const snapshot = await getProposals(page);
      expect(snapshot).toEqual(proposed);
      expect(snapshot.version).not.toBe(read.version);
      expect(snapshot.proposals.every(({ state }) => state === 'proposed')).toBe(true);
      const afterProposing = await status(page);
      transcript.pending = { proposed: afterProposing.pending };
      expect.soft(afterProposing.pending, 'replica pending after proposing').toBe(worker);
      const targets = targetsFor(paragraphs, snapshot, read.version);
      transcript.targets = targets.map(({ name }) => name);
      await waitForPreview(page, snapshot);
      transcript.scrolls = { proposed: await renderPages(page, worker ? arms[0].scrolls!.proposed : undefined) };
      await geometry(page, targets, snapshot, transcript, 'proposed');
      transcript.pending.afterProposedReads = (await status(page)).pending;
      const accepted = snapshotOf(await page.evaluate((request) =>
        (window as unknown as ProbeWindow).__workerGeometryProbe.editor!.setProposalStates(request),
        {
          expectVersion: snapshot.version,
          expectPreviewVersion: snapshot.previewVersion,
          changes: [{ id: IDS[0], state: 'accepted' as const }],
        }
      ));
      expect(accepted.previewVersion).toBeGreaterThan(snapshot.previewVersion);
      expect(accepted.version).toBe(snapshot.version);
      await expect.poll(() => getProposals(page)).toEqual(accepted);
      await waitForPreview(page, accepted);
      transcript.scrolls.accepted = await renderPages(page, worker ? arms[0].scrolls!.accepted : undefined);
      await geometry(page, targets, accepted, transcript, 'accepted');
      transcript.pending.afterAcceptedReads = (await status(page)).pending;
    } finally {
      const final = await status(page).catch(() => null);
      transcript.events = final?.events;
      transcript.errors = final?.errors;
      await test.info().attach(`geometry-${worker ? 'worker' : 'default'}.json`, {
        body: JSON.stringify(transcript),
        contentType: 'application/json',
      });
    }
  }
  expect(arms[1].paragraphs).toEqual(arms[0].paragraphs);
  expect(arms[1].proposals).toEqual(arms[0].proposals);
  expect(arms[1].targets).toEqual(arms[0].targets);
  expect(arms[1].scrolls).toEqual(arms[0].scrolls);
  expect(arms[1].geometry!.proposed).toEqual(arms[0].geometry!.proposed);
  expect(arms[1].geometry!.accepted).toEqual(arms[0].geometry!.accepted);
});
