import { test, expect } from 'playwright/test';
import { pinnedEvaluate } from '../pinned-evaluate';
import type {} from './pptx-peerless-editor-probe';

test.use({ launchOptions: { executablePath: process.env.CHROME || undefined } });

test('async editor props notify versions, set zoom and opt into saving downloads', async ({ page }) => {
  for (const download of [false, true]) {
    await page.goto(`/pptx-worker-editor.html?asyncApi=1&propsZoom=2${download ? '&propsDownload=1' : ''}`);
    await page.waitForFunction(() => '__pptxPeerlessEditor' in window);
    await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.ready);
    const canvas = page.getByTestId('pptx-slide-canvas');
    await expect(canvas).toBeVisible();
    expect(await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.editor().commands.getState('zoom'))).toMatchObject({ value: 2 });
    expect(await page.locator('html').getAttribute('data-document-versions')).toBeNull();
    const before = await pinnedEvaluate(page, async () => (await window.__pptxPeerlessEditor.content()).version);
    await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.select(0, 0));
    await page.getByRole('textbox', { name: 'Slide text' }).pressSequentially('props');
    await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.flush());
    const committed = await pinnedEvaluate(page, async () => (await window.__pptxPeerlessEditor.content()).version);
    await expect.poll(async () => JSON.parse((await page.locator('html').getAttribute('data-document-versions')) ?? '[]').at(-1)).toBe(committed);
    const versions = JSON.parse((await page.locator('html').getAttribute('data-document-versions'))!) as string[];
    expect(versions.length).toBeGreaterThan(0);
    expect(new Set(versions).size).toBe(versions.length);
    expect(versions).not.toContain(before);
    expect(versions.at(-1)).toBe(committed);
    const downloads: string[] = [];
    page.on('download', (event) => downloads.push(event.suggestedFilename()));
    const downloaded = download ? page.waitForEvent('download') : undefined;
    await page.getByRole('textbox', { name: 'Slide text' }).press('ControlOrMeta+s');
    if (downloaded) await downloaded;
    await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.saves.length)).toBe(1);
    expect(downloads).toHaveLength(download ? 1 : 0);
  }
});
