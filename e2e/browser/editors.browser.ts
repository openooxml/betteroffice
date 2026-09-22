import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

type Format = 'docx' | 'xlsx' | 'pptx';
const root = resolve(import.meta.dirname, '../..');
const marker = 'BrowserE2Eprobe';

async function open(page: Page, format: Format, file?: string) {
  await page.goto('/?format=' + format);
  await page
    .locator('input[type=file]')
    .last()
    .setInputFiles(
      file ??
        resolve(
          root,
          'apps/demo/public',
          format === 'xlsx' ? 'showcase.xlsx' : 'betteroffice-demo.' + format
        )
    );
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({
    timeout: 60_000,
  });
  await expect(page.locator('.notice')).toHaveCount(0);
  await expect(page.locator('[data-testid$="-error"]')).toHaveCount(0);
}

async function save(page: Page, format: Format, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const download = await downloading;
  const file = info.outputPath(label + '.' + format);
  await download.saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const parts = Object.values(zip.files).filter(
    (part) => !part.dir && part.name.endsWith('.xml')
  );
  const xml = await Promise.all(parts.map((part) => part.async('string')));
  const text = xml.map((value) => value.replace(/<[^>]*>/g, '')).join('\n');
  return { file, text };
}

for (const format of ['docx', 'xlsx', 'pptx'] as const) {
  test(
    format + ': edit, undo, redo, save, reopen, and reject a broken file',
    async ({ page }, info) => {
      const errors: string[] = [];
      page.on('pageerror', (error) => errors.push(error.message));
      page.on('response', (response) => {
        if (response.status() >= 400)
          errors.push(response.status() + ' ' + response.url());
      });
      page.on('requestfailed', (request) =>
        errors.push(request.url() + ': ' + request.failure()?.errorText)
      );
      await page.route('**/*', async (route) => {
        const url = new URL(route.request().url());
        if (
          ['http:', 'https:'].includes(url.protocol) &&
          url.hostname !== '127.0.0.1'
        ) {
          errors.push('unexpected external request: ' + url.href);
          await route.abort();
        } else await route.continue();
      });
      await open(page, format);
      const original = await save(page, format, info, 'original');
      expect(original.text).not.toContain(marker);

      if (format === 'docx') {
        await page
          .locator('.canvas-page canvas')
          .first()
          .click({ position: { x: 150, y: 150 } });
        const input = page.getByTestId('yrs-input');
        await expect(input).toBeEditable();
        await input.press('ControlOrMeta+Home');
        await input.pressSequentially(marker);
        await expect(page.locator('.canvas-page-mirror').first()).toContainText(
          marker
        );
      } else if (format === 'xlsx') {
        await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
        const editor = page.getByTestId('xlsx-scroll');
        await editor.press('F2');
        const input = page.getByTestId('xlsx-cell-editor');
        await input.press('ControlOrMeta+A');
        await input.pressSequentially(marker);
        await expect(input).toHaveValue(marker);
        await input.press('Enter');
        await expect(
          page.getByRole('gridcell').filter({ hasText: marker })
        ).toHaveCount(1);
      } else {
        await page.getByTestId('pptx-tool-text-box').click();
        const bounds = await page
          .getByTestId('pptx-slide-canvas')
          .boundingBox();
        expect(bounds).not.toBeNull();
        await page.mouse.move(
          bounds!.x + bounds!.width * 0.1,
          bounds!.y + bounds!.height * 0.65
        );
        await page.mouse.down();
        await page.mouse.move(
          bounds!.x + bounds!.width * 0.6,
          bounds!.y + bounds!.height * 0.85,
          { steps: 5 }
        );
        await page.mouse.up();
        await page.keyboard.type(marker);
      }

      const edited = await save(page, format, info, 'edited');
      expect(edited.text).toContain(marker);
      if (format === 'docx')
        await page.getByTestId('yrs-input').press('ControlOrMeta+z');
      else await page.getByTestId(format + '-undo').click();
      expect((await save(page, format, info, 'undone')).text).not.toContain(
        marker
      );
      if (format === 'docx')
        await page.getByTestId('yrs-input').press('ControlOrMeta+Shift+z');
      else await page.getByTestId(format + '-redo').click();
      const redone = await save(page, format, info, 'redone');
      expect(redone.text).toContain(marker);

      await open(page, format, redone.file);
      if (format === 'docx')
        await expect(page.locator('.canvas-page-mirror').first()).toContainText(
          marker
        );
      if (format === 'xlsx')
        await expect(
          page.getByRole('gridcell').filter({ hasText: marker })
        ).toHaveCount(1);
      expect((await save(page, format, info, 'reopened')).text).toContain(
        marker
      );

      await page
        .locator('input[type=file]')
        .last()
        .setInputFiles({
          name: 'broken.' + format,
          mimeType: 'application/octet-stream',
          buffer: Buffer.from('not an OOXML archive'),
        });
      await expect(page.locator('.notice')).toBeVisible();
      await expect(page.locator('.editor-stage canvas').first()).toBeVisible();
      expect(
        (await save(page, format, info, 'after-failed-open')).text
      ).toContain(marker);
      await info.attach('editor-after-reopen', {
        body: await page.screenshot(),
        contentType: 'image/png',
      });
      expect(errors).toEqual([]);
    }
  );
}
