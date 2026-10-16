import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const sample = resolve(import.meta.dirname, '../../apps/demo/public/betteroffice-demo.docx');

async function saved(page: Page, info: TestInfo, name: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const part = async (name: string) => await zip.file(name)?.async('string') ?? '';
  return { comments: await part('word/comments.xml'), extended: await part('word/commentsExtended.xml'), document: await part('word/document.xml') };
}

async function selectLine(page: Page, down: number) {
  const input = page.locator('.paged-editor__peerless-input');
  await input.press('ControlOrMeta+Home');
  for (let line = 0; line < down; line++) await input.press('ArrowDown');
  await input.press('Home');
  await input.press('Shift+End');
  await expect(page.getByRole('button', { name: 'Add comment' })).toBeVisible();
}

async function compose(page: Page, text: string) {
  const composer = page.getByPlaceholder('Add a comment...');
  await expect(composer).toBeVisible();
  await composer.fill(text);
  await composer.press('Enter');
  await expect(page.locator('.docx-comment-card', { hasText: text })).toBeVisible();
}

test('worker comments are added from the context menu and the floating button, replied to, resolved, reopened, deleted and saved', async ({ page }, info) => {
  test.setTimeout(240_000);
  await page.setViewportSize({ width: 2560, height: 1200 });
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles(sample);
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await page.locator('.paged-editor__page canvas').first().click({ position: { x: 150, y: 150 } });
  await expect(page.locator('.paged-editor__peerless-input')).toBeEditable();
  const toggle = page.getByRole('button', { name: 'Toggle comments sidebar' });
  await expect(toggle).toBeEnabled();

  await selectLine(page, 0);
  await page.locator('.paged-editor__peerless-input').dispatchEvent('contextmenu', { button: 2, clientX: 400, clientY: 300 });
  await page.getByRole('menuitem', { name: 'Comment', exact: true }).click();
  await compose(page, 'First note');

  await selectLine(page, 2);
  await page.getByRole('button', { name: 'Add comment' }).dispatchEvent('mousedown');
  await compose(page, 'Second note');

  const first = page.locator('.docx-comment-card', { hasText: 'First note' });
  await first.click();
  await first.getByPlaceholder('Reply or add others with @').click();
  await first.getByPlaceholder('Reply or add others with @').fill('A reply');
  await first.getByPlaceholder('Reply or add others with @').press('Enter');
  await expect(first).toContainText('A reply');

  await first.getByTitle('Resolve', { exact: true }).click();
  await expect(page.locator('.docx-comment-card', { hasText: 'First note' })).toHaveCount(0);
  const marker = page.locator('[data-comment-resolved="true"]');
  await expect(marker).toHaveCount(1);
  const resolved = await saved(page, info, 'resolved');
  expect(resolved.extended).toMatch(/w15:done="1"/);
  expect(resolved.extended).toMatch(/w15:paraIdParent="[0-9A-F]{8}"/);

  await marker.click();
  await page.locator('.docx-comment-card', { hasText: 'First note' }).getByTitle('Reopen', { exact: true }).click();
  await expect(page.locator('.docx-comment-card', { hasText: 'First note' })).toBeVisible();

  await toggle.click();
  await expect(page.locator('.docx-comment-card')).toHaveCount(0);
  await toggle.click();
  await expect(page.locator('.docx-comment-card', { hasText: 'Second note' })).toBeVisible();

  const second = page.locator('.docx-comment-card', { hasText: 'Second note' });
  await second.click();
  await second.getByTitle('More options', { exact: true }).click();
  await page.getByRole('menuitem', { name: 'Delete', exact: true }).click();
  await expect(page.locator('.docx-comment-card', { hasText: 'Second note' })).toHaveCount(0);

  const final = await saved(page, info, 'final');
  expect(final.comments).toContain('First note');
  expect(final.comments).toContain('A reply');
  expect(final.comments).not.toContain('Second note');
  expect(final.extended).not.toMatch(/w15:done="1"/);
  expect([...final.comments.matchAll(/<w:comment [^>]*w:id="\d+"/g)]).toHaveLength(2);
  const anchored = final.comments.match(/<w:comment [^>]*w:id="(\d+)"(?:(?!<\/w:comment>)[\s\S])*First note/)?.[1];
  expect(anchored).toBeDefined();
  expect(final.document).toContain(`<w:commentRangeStart w:id="${anchored}"`);
});
