import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

async function open(page: Page, arm: string) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Target paragraph</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  await page.setViewportSize({ width: 1600, height: 1000 });
  await page.goto('/?format=docx' + arm);
  await page.locator('input[type=file]').last().setInputFiles({ name: 'paste-options.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  const canvas = page.locator('.canvas-page canvas, .paged-editor__page canvas').first();
  await expect(canvas).toBeVisible({ timeout: 60000 });
  await canvas.click({ position: { x: 150, y: 150 } });
  const input = page.locator('[data-testid=yrs-input], .paged-editor__input');
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  return input;
}

async function savedXml(page: Page, info: TestInfo, name: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(name);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  return zip.file('word/document.xml')!.async('string');
}

function xmlText(xml: string): string {
  return xml.replace(/<[^>]+>/g, '');
}

for (const arm of ['&peerless=0', '']) {
  test(`paste options merge, text only, undo and dismissal (${arm ? 'legacy' : 'peerless'})`, async ({ page, context }, info) => {
    const input = await open(page, arm);
    await context.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: new URL(page.url()).origin });
    await page.evaluate(async () => {
      await navigator.clipboard.write([new ClipboardItem({
        'text/html': new Blob(
          ['<span style="font-family:Georgia;font-size:20pt;color:#FF0000;font-weight:bold">Rich paste</span>'],
          { type: 'text/html' },
        ),
        'text/plain': new Blob(['Rich paste'], { type: 'text/plain' }),
      })]);
    });
    await input.focus();
    await page.keyboard.press('ControlOrMeta+V');
    const control = page.getByRole('button', { name: 'Paste Options (Ctrl)' });
    await expect(control).toBeVisible({ timeout: 30000 });

    await page.keyboard.down('Control');
    await page.keyboard.up('Control');
    await expect(page.getByRole('menu')).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByRole('menu')).toBeHidden();
    await expect(control).toBeVisible();

    await control.click();
    await page.getByRole('menuitemradio', { name: 'Keep Text Only' }).click();
    await expect(control).toBeVisible();
    await control.click();
    await page.getByRole('menuitemradio', { name: 'Merge Formatting' }).click();
    await expect(control).toBeVisible();
    const merged = await savedXml(page, info, 'merge.docx');
    expect(merged).toContain('Rich paste');
    expect(merged).toContain('<w:b/>');
    expect(merged).not.toContain('Georgia');
    expect(merged).not.toMatch(/w:sz w:val="40"/);
    expect(merged).not.toContain('FF0000');
    await expect(control).toBeHidden();

    await input.focus();
    await input.press('ControlOrMeta+z');
    const undone = await savedXml(page, info, 'undone.docx');
    expect(xmlText(undone)).toContain('Target paragraph');
    expect(xmlText(undone)).not.toContain('Rich paste');

    await input.focus();
    await page.keyboard.press('ControlOrMeta+V');
    await expect(control).toBeVisible({ timeout: 30000 });
    await page.keyboard.type('x');
    await expect(control).toBeHidden();
  });

  test(`paste options over a selection replace it and one undo restores it (${arm ? 'legacy' : 'peerless'})`, async ({ page, context }, info) => {
    const input = await open(page, arm);
    await context.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: new URL(page.url()).origin });
    await page.evaluate(async () => {
      await navigator.clipboard.write([new ClipboardItem({
        'text/html': new Blob(
          ['<span style="font-weight:bold">BAR</span>'],
          { type: 'text/html' },
        ),
        'text/plain': new Blob(['BAR'], { type: 'text/plain' }),
      })]);
    });
    await input.focus();
    await input.press('ControlOrMeta+Home');
    for (let index = 0; index < 6; index++) await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('ControlOrMeta+V');
    const control = page.getByRole('button', { name: 'Paste Options (Ctrl)' });
    await expect(control).toBeVisible({ timeout: 30000 });

    await control.click();
    await page.getByRole('menuitemradio', { name: 'Keep Text Only' }).click();
    const textOnly = await savedXml(page, info, 'text-only-over-selection.docx');
    expect(xmlText(textOnly)).toContain('BAR paragraph');
    expect(xmlText(textOnly)).not.toContain('Target');
    expect(textOnly).not.toContain('<w:b/>');

    await input.focus();
    await input.press('ControlOrMeta+z');
    const undone = await savedXml(page, info, 'undone-over-selection.docx');
    expect(xmlText(undone)).toContain('Target paragraph');
    expect(xmlText(undone)).not.toContain('BAR');
  });
}
