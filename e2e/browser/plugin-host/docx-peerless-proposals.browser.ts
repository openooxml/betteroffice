import { expect, test, type Page } from 'playwright/test';
import { pinnedEvaluate } from '../pinned-evaluate';

interface PeerlessProposalProbe {
  editor: { getTotalPages(): number; readSelectionInfo(): Promise<unknown> } | null;
  propose(): Promise<{ ok: boolean }>;
  decide(ids: string[], state: 'proposed' | 'accepted' | 'rejected'): Promise<{ ok: boolean }>;
  undoDecision(): Promise<{ ok: boolean }>;
  next(index: number): Promise<{ ok: boolean } | null>;
  savedText(): Promise<string>;
  events: { generation: string; version: string; previewVersion: number }[];
  errors: string[];
}

type ProbeWindow = Window & { __peerlessProposalProbe: PeerlessProposalProbe };
const ids = ['change-0', 'change-1', 'change-2', 'change-3'];

async function overlays(page: Page) {
  const marks = page.locator('[data-proposal-id]');
  await expect(marks).toHaveCount(4);
  await expect.poll(() => marks.evaluateAll(elements => elements.every(element => element.getAttribute('data-ok') === 'true'))).toBe(true);
  await expect.poll(() => page.getByTestId('host-proposals').evaluate(element => element.getAttribute('data-preview') === element.getAttribute('data-layout-preview'))).toBe(true);
  return marks;
}

test('editable worker and hybrid hosts produce the same proposal decisions and saved text', async ({ page }) => {
  test.setTimeout(300_000);
  const arms: { states: string[][]; saved: string[] }[] = [];
  for (const arm of ['worker', 'hybrid']) {
    await page.goto(`/docx-peerless-proposals.html?arm=${arm}`);
    await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
    await expect.poll(() => page.evaluate(() => (window as unknown as ProbeWindow).__peerlessProposalProbe.editor?.getTotalPages())).toBe(3);
    const proposed = await pinnedEvaluate(page, () => (window as unknown as ProbeWindow).__peerlessProposalProbe.propose());
    expect(proposed.ok).toBe(true);
    const marks = await overlays(page);
    await expect.poll(() => marks.evaluateAll(elements => elements.every(element => element.getAttribute('data-text-aligned') === 'true'))).toBe(true);
    const pages = await marks.evaluateAll(elements => [...new Set(elements.map(element => element.getAttribute('data-page')))]);
    expect(pages.length).toBeGreaterThanOrEqual(2);
    await expect(page.locator('.docx-tracked-change-card')).toHaveCount(4);
    const states: string[][] = [];
    const saved: string[] = [];
    const record = async () => {
      await overlays(page);
      states.push(await marks.evaluateAll(elements => elements.map(element => `${element.getAttribute('data-proposal-id')}:${element.getAttribute('data-state')}`)));
      const text = await pinnedEvaluate(page, () => (window as unknown as ProbeWindow).__peerlessProposalProbe.savedText());
      expect(typeof text).toBe('string');
      expect(text.length).toBeGreaterThan(0);
      saved.push(text);
    };
    await record();
    for (const state of ['accepted', 'proposed'] as const) {
      expect((await pinnedEvaluate(page, state => (window as unknown as ProbeWindow).__peerlessProposalProbe.decide(['change-1'], state), state)).ok).toBe(true);
      await record();
    }
    for (const state of ['accepted', 'rejected'] as const) {
      const result = await pinnedEvaluate(page, state => (window as unknown as ProbeWindow).__peerlessProposalProbe.decide(['change-0'], state), state);
      expect(result.ok).toBe(true);
      await expect(marks.first()).toHaveAttribute('data-state', state);
      await record();
      expect((await pinnedEvaluate(page, () => (window as unknown as ProbeWindow).__peerlessProposalProbe.undoDecision())).ok).toBe(true);
      await expect(marks.first()).toHaveAttribute('data-state', 'proposed');
      await record();
    }
    for (const state of ['accepted', 'rejected'] as const) {
      expect((await pinnedEvaluate(page, ({ ids, state }) => (window as unknown as ProbeWindow).__peerlessProposalProbe.decide(ids, state), { ids, state })).ok).toBe(true);
      await expect.poll(() => marks.evaluateAll(elements => elements.map(element => element.getAttribute('data-state')))).toEqual(ids.map(() => state));
      await record();
    }
    const beforeSelection = await pinnedEvaluate(page, () => (window as unknown as ProbeWindow).__peerlessProposalProbe.editor!.readSelectionInfo());
    for (const index of [2, 3, 0]) {
      expect(await pinnedEvaluate(page, index => (window as unknown as ProbeWindow).__peerlessProposalProbe.next(index), index)).toEqual({ ok: true });
      await expect(marks.nth(index)).toBeInViewport();
      expect(await pinnedEvaluate(page, () => (window as unknown as ProbeWindow).__peerlessProposalProbe.editor!.readSelectionInfo())).toEqual(beforeSelection);
    }
    expect(await page.evaluate(() => (window as unknown as ProbeWindow).__peerlessProposalProbe.events.every(event => event.generation && event.version && Number.isInteger(event.previewVersion)))).toBe(true);
    expect(await page.evaluate(() => (window as unknown as ProbeWindow).__peerlessProposalProbe.errors)).toEqual([]);
    arms.push({ states, saved });
  }
  expect(arms[0]).toEqual(arms[1]);
});
