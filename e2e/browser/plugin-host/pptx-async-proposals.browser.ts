import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';
import type { ProposalEdit } from '../../../packages/pptx/src/proposals';
import { pinnedEvaluate } from '../pinned-evaluate';
import type {} from './pptx-peerless-editor-probe';

test.use({ launchOptions: { executablePath: process.env.CHROME || undefined } });
test.setTimeout(60_000);
test.use({ actionTimeout: 10_000 });

const pageErrors = new WeakMap<Page, string[]>();
const panel = (page: Page) => page.getByTestId('pptx-proposals-panel');
const row = (page: Page, note: string) => panel(page).locator(':scope > div, :scope > section').filter({ hasText: note });
const input = (page: Page) => page.getByRole('textbox', { name: 'Slide text' });
const painted = (page: Page) => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts());

async function content(page: Page) {
  return pinnedEvaluate(page, () => window.__pptxPeerlessEditor.content());
}

async function stage(page: Page, note: string, edits: ProposalEdit[]) {
  return pinnedEvaluate(page, async ({ note, edits }) => {
    const api = window.__pptxPeerlessEditor.editor();
    const proposal = await (await api.handleAsync()).propose('browser', note, edits);
    await api.flushPendingInput();
    return proposal;
  }, { note, edits });
}

async function stageText(page: Page, note: string, text: string, slide = 0) {
  const target = (await content(page)).stories[slide]!;
  return stage(page, note, [{ type: 'replaceText', storyId: target.storyId, start: 0, end: target.text.length, text }]);
}

async function openPanel(page: Page) {
  expect(await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.editor().commands.execute('proposalsPanel', { open: true })))
    .toMatchObject({ ok: true });
  await expect(panel(page)).toBeVisible();
}

async function replace(page: Page, value: string) {
  const point = await page.evaluate(() => window.__pptxPeerlessEditor.point(0));
  await page.mouse.click(point.x, point.y);
  await input(page).press('Home');
  await input(page).press('Shift+End');
  await input(page).pressSequentially(value);
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.flush());
  await expect.poll(async () => (await content(page)).stories[0]!.text).toBe(value);
  await expect.poll(() => painted(page)).toContain(value);
}

async function previewPainted(page: Page) {
  const diff = page.getByTestId('pptx-proposal-diff');
  await expect(diff).toBeVisible();
  await expect.poll(() => diff.evaluate((canvas: HTMLCanvasElement) => {
    const context = canvas.getContext('2d');
    if (!context || !canvas.width || !canvas.height) return false;
    const { data } = context.getImageData(0, 0, canvas.width, canvas.height);
    for (let alpha = 3; alpha < data.length; alpha += 4) if (data[alpha]) return true;
    return false;
  })).toBe(true);
}

async function pending(page: Page) {
  return pinnedEvaluate(page, async () => (await window.__pptxPeerlessEditor.editor().handleAsync()).listProposals());
}

async function save(page: Page) {
  await page.evaluate(() => window.__pptxPeerlessEditor.takeSaved());
  await page.getByTestId('pptx-save').click();
  let bytes: number[] | null = null;
  await expect.poll(async () => {
    bytes = await page.evaluate(() => window.__pptxPeerlessEditor.takeSaved());
    return bytes !== null;
  }).toBe(true);
  const zip = await JSZip.loadAsync(Uint8Array.from(bytes!));
  const parts = Object.values(zip.files).filter((part) => !part.dir && (
    /^ppt\/slides\/slide\d+\.xml$/.test(part.name) ||
    /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(part.name) || part.name === 'ppt/presentation.xml'
  ));
  const xml = await Promise.all(parts.map(async (part) => ({ name: part.name, xml: await part.async('string') })));
  return page.evaluate((parts) => {
    const documents = parts.map((part) => {
      const document = new DOMParser().parseFromString(part.xml, 'application/xml');
      if (document.querySelector('parsererror')) throw new Error(`Invalid XML in ${part.name}`);
      return { ...part, document };
    });
    const first = documents.find((part) => part.name === 'ppt/slides/slide1.xml')!.document;
    const shape = first.getElementsByTagNameNS('*', 'sp')[0]!;
    const off = shape.getElementsByTagNameNS('*', 'off')[0]!;
    const ext = shape.getElementsByTagNameNS('*', 'ext')[0]!;
    const text = (document: Document) => Array.from(document.getElementsByTagNameNS('*', 't')).map((node) => node.textContent ?? '').join('');
    return {
      firstText: text(first),
      slideTexts: documents.filter((part) => part.name.startsWith('ppt/slides/')).map((part) => text(part.document)),
      rect: { x: Number(off.getAttribute('x')), y: Number(off.getAttribute('y')),
        width: Number(ext.getAttribute('cx')), height: Number(ext.getAttribute('cy')) },
      notes: documents.filter((part) => part.name.startsWith('ppt/notesSlides/')).map((part) => text(part.document)).join('\n'),
      slideCount: documents.find((part) => part.name === 'ppt/presentation.xml')!.document.getElementsByTagNameNS('*', 'sldId').length,
      slideParts: documents.filter((part) => part.name.startsWith('ppt/slides/')).length,
    };
  }, xml);
}

async function reopen(page: Page, text: string) {
  expect(await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.saveAndReopen())).toBeGreaterThan(0);
  await expect.poll(() => painted(page)).toContain(text);
  expect((await content(page)).stories[0]!.text).toBe(text);
  return save(page);
}

test.beforeEach(async ({ page }) => {
  const errors: string[] = [];
  pageErrors.set(page, errors);
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/pptx-worker-editor.html?asyncApi=1');
  await page.waitForFunction(() => '__pptxPeerlessEditor' in window);
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.ready);
  await expect.poll(() => painted(page)).toContain('initial');
  expect(errors).toEqual([]);
});

test.afterEach(async ({ page }) => {
  expect.soft(pageErrors.get(page), 'Page errors').toEqual([]);
  expect.soft(await page.evaluate(() => window.__pptxMainEngineProbe?.count())).toBe(0);
  expect.soft(await page.evaluate(() => window.__pptxPeerlessEditor?.errors ?? [])).toEqual([]);
});

test('accepting one proposal leaves another available to reject independently', async ({ page }) => {
  await stageText(page, 'Keep first change', 'accepted first slide');
  const second = await stageText(page, 'Discard second change', 'rejected second slide', 1);
  await openPanel(page);
  await row(page, 'Keep first change').getByRole('button', { name: 'Accept proposal', exact: true }).click();
  await expect(row(page, 'Keep first change')).toHaveCount(0);
  await expect(row(page, 'Discard second change')).toBeVisible();
  expect((await pending(page)).map((proposal) => proposal.id)).toEqual([second.id]);
  await expect.poll(() => painted(page)).toContain('accepted first slide');
  await row(page, 'Discard second change').getByRole('button', { name: 'Reject', exact: true }).click();
  await expect(row(page, 'Discard second change')).toHaveCount(0);
  expect(await pending(page)).toEqual([]);
  expect((await save(page)).slideTexts).toEqual(expect.arrayContaining(['accepted first slide', 'other']));
  const reopened = await reopen(page, 'accepted first slide');
  expect(reopened.slideTexts).toEqual(expect.arrayContaining(['accepted first slide', 'other']));
  expect(reopened.slideTexts).not.toContain('rejected second slide');
});

test('accepting a text proposal supports toolbar undo, redo and save after reopening', async ({ page }) => {
  await stageText(page, 'Text to accept', 'accepted text');
  await openPanel(page);
  await row(page, 'Text to accept').getByRole('button', { name: 'Accept proposal', exact: true }).click();
  await expect(row(page, 'Text to accept')).toHaveCount(0);
  await expect.poll(() => painted(page)).toContain('accepted text');
  expect(await pending(page)).toEqual([]);
  expect((await save(page)).firstText).toBe('accepted text');
  await page.getByTestId('pptx-undo').click();
  await expect.poll(() => painted(page)).toContain('initial');
  expect((await save(page)).firstText).toBe('initial');
  await page.getByTestId('pptx-redo').click();
  await expect.poll(() => painted(page)).toContain('accepted text');
  expect((await reopen(page, 'accepted text')).firstText).toBe('accepted text');
});

test('rejecting a text proposal keeps the original canvas and reopened slide XML', async ({ page }) => {
  await stageText(page, 'Text to reject', 'discard this text');
  await openPanel(page);
  await row(page, 'Text to reject').getByRole('button', { name: 'Reject', exact: true }).click();
  await expect(row(page, 'Text to reject')).toHaveCount(0);
  await expect.poll(() => painted(page)).toContain('initial');
  expect(await pending(page)).toEqual([]);
  expect((await save(page)).firstText).toBe('initial');
  expect((await reopen(page, 'initial')).firstText).toBe('initial');
});

test('accepting a geometry proposal moves the shape and persists its transform through undo and redo', async ({ page }) => {
  const target = (await content(page)).stories[0]!;
  const original = await save(page);
  await openPanel(page);
  const point = await page.evaluate(() => window.__pptxPeerlessEditor.shapePoint());
  const rect = { x: 1828800, y: 1371600, width: 5486400, height: 1371600 };
  await stage(page, 'Move and resize', [{ type: 'setShapeRect', slideId: target.slideId, shapeId: target.shapeId, rect }]);
  expect((await save(page)).rect).toEqual(original.rect);
  await row(page, 'Move and resize').getByRole('button', { name: 'Accept proposal', exact: true }).click();
  await expect(row(page, 'Move and resize')).toHaveCount(0);
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.shapePoint().x)).toBeGreaterThan(point.x);
  expect((await save(page)).rect).toEqual(rect);
  await page.getByTestId('pptx-undo').click();
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.shapePoint().x)).toBeCloseTo(point.x, 1);
  expect((await save(page)).rect).toEqual(original.rect);
  await page.getByTestId('pptx-redo').click();
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.shapePoint().x)).toBeGreaterThan(point.x);
  expect((await save(page)).rect).toEqual(rect);
  expect((await reopen(page, 'initial')).rect).toEqual(rect);
});

test('accepting a notes proposal updates speaker notes and saved notes XML with keyboard undo and redo', async ({ page }) => {
  const target = (await content(page)).stories[0]!;
  await stage(page, 'Speaker notes', [{ type: 'setSlideNotes', slideId: target.slideId, text: 'Proposed speaker notes' }]);
  const notes = page.getByTestId('pptx-notes-textarea');
  await expect(notes).toHaveValue('');
  expect((await save(page)).notes).not.toContain('Proposed speaker notes');
  await openPanel(page);
  await row(page, 'Speaker notes').getByRole('button', { name: 'Accept proposal', exact: true }).click();
  await expect(notes).toHaveValue('Proposed speaker notes');
  expect((await save(page)).notes).toContain('Proposed speaker notes');
  await page.getByRole('application').focus();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(notes).toHaveValue('');
  expect((await save(page)).notes).not.toContain('Proposed speaker notes');
  await page.getByRole('application').focus();
  await page.keyboard.press('ControlOrMeta+Shift+z');
  await expect(notes).toHaveValue('Proposed speaker notes');
  expect((await reopen(page, 'initial')).notes).toContain('Proposed speaker notes');
  await expect(notes).toHaveValue('Proposed speaker notes');
});

test('rejecting a notes proposal preserves the original notes after reopening', async ({ page }) => {
  const target = (await content(page)).stories[0]!;
  await stage(page, 'Discard notes', [{ type: 'setSlideNotes', slideId: target.slideId, text: 'Rejected notes' }]);
  await openPanel(page);
  await row(page, 'Discard notes').getByRole('button', { name: 'Reject', exact: true }).click();
  await expect(row(page, 'Discard notes')).toHaveCount(0);
  await expect(page.getByTestId('pptx-notes-textarea')).toHaveValue('');
  expect((await save(page)).notes).not.toContain('Rejected notes');
  expect((await reopen(page, 'initial')).notes).not.toContain('Rejected notes');
});

test('previewing a proposal navigates to its slide without committing its text', async ({ page }) => {
  await stageText(page, 'Second slide review', 'proposed second slide', 1);
  await openPanel(page);
  await row(page, 'Second slide review').getByRole('button', { name: 'Preview changes', exact: true }).click();
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.activeSlide())).toBe(2);
  await expect(page.getByRole('complementary', { name: 'Slides', exact: true }).getByRole('button').nth(1)).toHaveAttribute('aria-current', 'page');
  await previewPainted(page);
  expect((await content(page)).stories[1]!.text).toBe('other');
  const saved = await save(page);
  expect(saved.slideTexts).toEqual(expect.arrayContaining(['initial', 'other']));
  expect(saved.slideTexts).not.toContain('proposed second slide');
  await row(page, 'Second slide review').getByRole('button', { name: 'Reject', exact: true }).click();
  await expect(page.getByTestId('pptx-proposal-diff')).toHaveCount(0);
  await expect.poll(() => painted(page)).toContain('other');
  const reopened = await reopen(page, 'initial');
  expect(reopened.slideCount).toBe(2);
  expect(reopened.slideTexts).toEqual(expect.arrayContaining(['initial', 'other']));
});

test('the proposals button counts staged proposals and toggles the panel', async ({ page }) => {
  const button = page.getByTestId('pptx-proposals-button');
  await expect(button.getByTestId('pptx-proposals-count')).toHaveText('0');
  await stageText(page, 'Button suggestion', 'from the button');
  await expect(button.getByTestId('pptx-proposals-count')).toHaveText('1');
  await button.click();
  await expect(panel(page)).toBeVisible();
  await expect(button).toHaveAttribute('aria-expanded', 'true');
  await expect(row(page, 'Button suggestion')).toBeVisible();
  await button.click();
  await expect(panel(page)).toHaveCount(0);
  await expect(button).toHaveAttribute('aria-expanded', 'false');
});

test('staged proposals show their count, author and text changes without changing saved slides', async ({ page }) => {
  await replace(page, 'typed before review');
  const first = await stageText(page, 'First suggestion', 'accepted text');
  const second = await stageText(page, 'Second suggestion', 'other suggestion');
  expect(first.id).not.toBe(second.id);
  expect(first.changes[0]).toMatchObject({ oldText: 'typed before review', newText: 'accepted text' });
  await openPanel(page);
  await expect(panel(page).locator('strong').first()).toHaveText('Agent proposals 2');
  await expect(panel(page).getByTestId('pptx-proposal')).toHaveCount(2);
  await expect(row(page, 'First suggestion')).toBeVisible();
  await expect(row(page, 'Second suggestion')).toBeVisible();
  await expect(row(page, 'First suggestion')).toContainText('browser');
  await expect(row(page, 'First suggestion').locator('del')).toHaveText('typed before review');
  await expect(row(page, 'First suggestion').locator('ins')).toHaveText('accepted text');
  expect((await content(page)).stories[0]!.text).toBe('typed before review');
  expect((await save(page)).firstText).toBe('typed before review');
});

test('notes preview shows the proposed notes while the saved notes remain unchanged', async ({ page }) => {
  const target = (await content(page)).stories[0]!;
  await stage(page, 'Preview notes', [{ type: 'setSlideNotes', slideId: target.slideId, text: 'Review these notes' }]);
  await openPanel(page);
  await row(page, 'Preview notes').getByRole('button', { name: 'Preview changes', exact: true }).click();
  await previewPainted(page);
  const diff = page.getByTestId('pptx-proposal-notes-diff');
  await expect(diff).toBeVisible();
  await expect(diff.locator('del')).toHaveText('Empty');
  await expect(diff.locator('ins')).toHaveText('Review these notes');
  await expect(page.getByTestId('pptx-notes-textarea')).toHaveValue('');
  expect((await save(page)).notes).not.toContain('Review these notes');
  await row(page, 'Preview notes').getByRole('button', { name: 'Reject', exact: true }).click();
  await expect(row(page, 'Preview notes')).toHaveCount(0);
  expect((await reopen(page, 'initial')).notes).not.toContain('Review these notes');
});

test('a conflicting local edit marks a proposal stale and rejection preserves the local text', async ({ page }) => {
  const proposal = await stageText(page, 'Stale review', 'suggestion');
  await replace(page, 'changed');
  await expect.poll(async () => (await pending(page))[0]!.staleTargets.length).toBeGreaterThan(0);
  await openPanel(page);
  await expect(panel(page).getByTestId('pptx-proposal-stale')).toContainText('changed after the proposal was created');
  await expect(row(page, 'Stale review').getByRole('button', { name: 'Accept proposal', exact: true })).toBeDisabled();
  expect(await pinnedEvaluate(page, (proposalId) => window.__pptxPeerlessEditor.editor().commands.getState('proposalAccept', { proposalId }), proposal.id))
    .toMatchObject({ enabled: false, disabledReason: { code: 'proposal-stale' } });
  expect((await save(page)).firstText).toBe('changed');
  await row(page, 'Stale review').getByRole('button', { name: 'Reject', exact: true }).click();
  await expect(row(page, 'Stale review')).toHaveCount(0);
  expect((await reopen(page, 'changed')).firstText).toBe('changed');
});

test('a stale proposal can be force accepted after previewing the updated slide', async ({ page }) => {
  await stageText(page, 'Force review', 'suggestion');
  await replace(page, 'changed');
  await expect.poll(async () => (await pending(page))[0]!.staleTargets.length).toBeGreaterThan(0);
  await openPanel(page);
  await row(page, 'Force review').getByRole('button', { name: 'Preview changes', exact: true }).click();
  await page.getByTestId('pptx-proposal-force').click();
  await expect.poll(() => painted(page)).toContain('suggestion');
  expect(await pending(page)).toEqual([]);
  expect((await save(page)).firstText).toBe('suggestion');
  await page.getByTestId('pptx-undo').click();
  await expect.poll(() => painted(page)).toContain('changed');
  expect((await save(page)).firstText).toBe('changed');
  await page.getByTestId('pptx-redo').click();
  await expect.poll(() => painted(page)).toContain('suggestion');
  expect((await reopen(page, 'suggestion')).firstText).toBe('suggestion');
});

test('slide insertion proposals are refused and the editor keeps working', async ({ page }) => {
  const slides = page.getByRole('complementary', { name: 'Slides', exact: true }).getByRole('button');
  await expect(slides).toHaveCount(2);
  const refusal = await pinnedEvaluate(page, async (edits) => {
    try {
      await (await window.__pptxPeerlessEditor.editor().handleAsync()).propose('browser', 'Insert slide', edits);
      return null;
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    }
  }, [{ type: 'insertSlide', index: 1 } as unknown as ProposalEdit]);
  expect(refusal).toBe('invalid-arguments');
  expect(await pending(page)).toEqual([]);
  await expect(slides).toHaveCount(2);
  await stageText(page, 'After refusal', 'still editable');
  await openPanel(page);
  await expect(row(page, 'Insert slide')).toHaveCount(0);
  await row(page, 'After refusal').getByRole('button', { name: 'Accept proposal', exact: true }).click();
  await expect.poll(() => painted(page)).toContain('still editable');
  expect(await save(page)).toMatchObject({ firstText: 'still editable', slideCount: 2, slideParts: 2 });
  expect(await reopen(page, 'still editable')).toMatchObject({ firstText: 'still editable', slideCount: 2, slideParts: 2 });
});

test('a refused slide insertion proposal is not reported to onError', async ({ page }) => {
  const refused = await pinnedEvaluate(page, async (edits) => {
    try {
      await (await window.__pptxPeerlessEditor.editor().handleAsync()).propose('browser', 'Insert slide', edits);
      return false;
    } catch {
      return true;
    }
  }, [{ type: 'insertSlide', index: 1 } as unknown as ProposalEdit]);
  expect(refused).toBe(true);
  await stageText(page, 'After refusal', 'still editable');
  expect(await page.evaluate(() => window.__pptxPeerlessEditor.errors)).toEqual([]);
});
