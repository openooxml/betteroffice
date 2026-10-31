import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

async function open(page: Page, arm: string) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Seed</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  await page.goto('/?format=docx' + arm);
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'clear-formatting.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await zip.generateAsync({ type: 'nodebuffer' }),
  });
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('formatting-bar')).toBeVisible();
  const input = page.locator('[data-testid=yrs-input], .paged-editor__input');
  await input.focus();
  return input;
}

async function savedRuns(page: Page, info: TestInfo, name: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate((source) => {
    const document = new DOMParser().parseFromString(source, 'application/xml');
    const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    return Array.from(document.getElementsByTagNameNS(w, 'r')).map((run) => ({
      text: Array.from(run.getElementsByTagNameNS(w, 't')).map((node) => node.textContent ?? '').join(''),
      bold: run.getElementsByTagNameNS(w, 'b').length > 0,
    }));
  }, xml);
}

for (const arm of ['&peerless=0', '']) {
  const label = arm ? 'hybrid' : 'default';

  test(`docx clear formatting shortcut removes saved bold formatting (${label})`, async ({ page }, info) => {
    const input = await open(page, arm);
    const text = 'Clear formatting shortcut';
    await input.press('ControlOrMeta+A');
    await input.pressSequentially(text);
    await input.press('ControlOrMeta+A');
    await input.press('ControlOrMeta+B');
    const bold = page.getByTestId('formatting-bar').getByRole('button', { name: 'Bold', exact: true });
    await expect(bold).toHaveAttribute('aria-pressed', 'true');
    const before = await savedRuns(page, info, `${label}-bold`);
    expect(before.map((run) => run.text).join('')).toBe(text);
    expect(before.some((run) => run.bold)).toBe(true);

    await input.focus();
    await input.press('ControlOrMeta+A');
    const mac = await page.evaluate(() => /Mac|iPod|iPhone|iPad/.test(navigator.platform));
    await input.press(mac ? 'Meta+Backslash' : 'Control+Space');
    await expect(bold).toHaveAttribute('aria-pressed', 'false');
    const after = await savedRuns(page, info, `${label}-cleared`);
    expect(after.map((run) => run.text).join('')).toBe(text);
    expect(after.every((run) => !run.bold)).toBe(true);
  });
}
