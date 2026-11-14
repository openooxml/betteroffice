import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const item = (num: number, level: number, text: string) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${num}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

async function docxFile() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="num" Type="${r}/numbering" Target="numbering.xml"/></Relationships>`);
  zip.file('word/numbering.xml', `<w:numbering xmlns:w="${w}">`
    + '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="2"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>'
    + '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="Appendix %1"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>'
    + '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\u2022"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>'
    + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>');
  zip.file('word/document.xml', `<w:document xmlns:w="${w}"><w:body>`
    + item(1, 0, 'Alpha') + item(1, 1, 'Beta') + item(2, 0, 'Gamma') + item(3, 0, 'Delta')
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, worker: boolean) {
  await page.addInitScript(() => {
    document.addEventListener('copy', (event) => {
      const text = event.clipboardData?.getData('text/plain');
      if (text) (window as unknown as { copied: string }).copied = text;
    });
  });
  await page.goto(`/?format=docx&peerless=${worker ? 1 : 0}`);
  await page.locator('input[type=file]').last().setInputFiles({ name: 'lists.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await docxFile() });
  await expect(page.locator('[data-page-index="0"] canvas').first()).toBeVisible({ timeout: 60000 });
  await expect(page.locator('.canvas-page-mirror').getByText('Delta')).toBeAttached();
}

for (const worker of [false, true]) {
  test(`docx mirror lists read their numbers, labels and the word for a bullet (${worker ? 'worker' : 'in-thread'} editor)`, async ({ page }) => {
    await open(page, worker);
    const exposed = page.locator('.canvas-page-mirror .layout-list-marker:not([aria-hidden="true"])');
    await expect(exposed).toHaveText(['3.', '3.2', 'Appendix 1', 'bullet'].map((text) => new RegExp(`^${text.replace('.', '\\.')}\\s*$`)));
    await expect(page.locator('.canvas-page-mirror .layout-list-marker[aria-hidden="true"]')).toHaveCount(0);
    const snapshot = await page.locator('.canvas-page-mirror').first().ariaSnapshot();
    for (const text of ['3.', '3.2', 'Appendix 1', 'bullet', 'Alpha', 'Beta', 'Gamma', 'Delta']) expect(snapshot).toContain(text);
    expect(snapshot).not.toContain('\u2022');
  });

  test(`docx list numbers stay out of what a selection copies (${worker ? 'worker' : 'in-thread'} editor)`, async ({ page }) => {
    await open(page, worker);
    await page.locator('[data-page-index="0"] canvas').first().click({ position: { x: 150, y: 150 } });
    const input = page.locator(worker ? '.paged-editor__input' : '[data-testid="yrs-input"]');
    await expect(input).toBeEditable();
    await input.focus();
    await page.keyboard.press('ControlOrMeta+a');
    await page.keyboard.press('ControlOrMeta+c');
    await expect.poll(() => page.evaluate(() => (window as unknown as { copied?: string }).copied ?? '')).toContain('Delta');
    const copied = await page.evaluate(() => (window as unknown as { copied: string }).copied);
    expect(copied).not.toContain('Appendix');
    expect(copied).not.toContain('3.2');
    expect(copied).not.toContain('bullet');
  });

  test(`docx find does not match a list number (${worker ? 'worker' : 'in-thread'} editor)`, async ({ page }) => {
    await open(page, worker);
    const canvas = page.locator('[data-page-index="0"] canvas').first();
    await canvas.click({ position: { x: 150, y: 150 } });
    const input = page.locator(worker ? '.paged-editor__input' : '[data-testid="yrs-input"]');
    await expect(input).toBeEditable();
    await input.press('ControlOrMeta+f');
    const dialog = page.getByTestId('find-replace-dialog');
    await expect(dialog).toBeVisible();
    const find = dialog.getByRole('textbox', { name: 'Find text', exact: true });
    const status = dialog.locator('.docx-find-replace-dialog-status');
    await find.fill('Appendix');
    await find.press('Enter');
    await expect(status).toHaveText('No results found');
    await find.fill('bullet');
    await find.press('Enter');
    await expect(status).toHaveText('No results found');
    await find.fill('Gamma');
    await find.press('Enter');
    await expect(status).toHaveText('1 of 1 matches');
  });
}
