import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const TBL_PR = '<w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblCaption w:val="caption"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>';
const TC_PR = '<w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:tcMar><w:left w:w="99" w:type="dxa"/></w:tcMar></w:tcPr>';
const BODY = `<w:p><w:r><w:rPr><w:noProof/></w:rPr><w:t>Before</w:t></w:r></w:p><w:tbl>${TBL_PR}<w:tr><w:tc>${TC_PR}<w:p><w:r><w:t>Inside the cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p>`;

async function open(page: Page, arm: string) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${BODY}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  await page.goto('/?format=docx' + arm);
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'select-all-table.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await zip.generateAsync({ type: 'nodebuffer' }),
  });
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('formatting-bar')).toBeVisible();
  const input = page.locator('[data-testid=yrs-input], .paged-editor__input');
  await input.focus();
  return input;
}

async function saved(page: Page, info: TestInfo, name: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  return (await JSZip.loadAsync(await readFile(file))).file('word/document.xml')!.async('string');
}
const dialog = (page: Page) => page.getByTestId('find-replace-dialog');
const findInput = (page: Page) => dialog(page).getByRole('textbox', { name: 'Find text', exact: true });
const replaceInput = (page: Page) => dialog(page).getByRole('textbox', { name: 'Replace text', exact: true });
const documentInput = (page: Page) => page.locator('[data-docx-worker-editor] .paged-editor__input');
const status = (page: Page) => dialog(page).locator('.docx-find-replace-dialog-status');
async function openFind(page: Page, replace = false) {
  await documentInput(page).press(replace ? 'ControlOrMeta+h' : 'ControlOrMeta+f');
  await expect(dialog(page)).toBeVisible();
  await expect(dialog(page).getByRole('heading', { name: replace ? 'Find and Replace' : 'Find', exact: true })).toBeVisible();
  await expect(findInput(page)).toBeFocused();
  if (replace) await expect(replaceInput(page)).toBeVisible();
}
const runWith = (xml: string, text: string) => (xml.match(/<w:r\b[\s\S]*?<\/w:r>/g) ?? []).find(run => run.includes(`>${text}<`)) ?? '';

for (const arm of ['&peerless=0', '']) {
  const label = arm ? 'hybrid' : 'default';
  test(`select all then bold formats the table text and keeps the table (${label})`, async ({ page }, info) => {
    const input = await open(page, arm);
    await input.press('ControlOrMeta+A');
    await input.press('ControlOrMeta+B');
    const xml = await saved(page, info, `${label}-bold`);
    for (const text of ['Before', 'Inside the cell', 'After']) expect(runWith(xml, text)).toContain('<w:b/>');
    expect(xml).toContain(TBL_PR);
    expect(xml).toContain(TC_PR);
  });
}

test('replace all reaches table text and keeps the table and untouched runs', async ({ page }, info) => {
  await open(page, '');
  await openFind(page, true);
  await findInput(page).fill('e');
  await expect(status(page)).not.toHaveText('No results found');
  await replaceInput(page).fill('3');
  await dialog(page).getByRole('button', { name: 'Replace All', exact: true }).click();
  await expect(status(page)).toHaveText('No results found');
  const xml = await saved(page, info, 'replace-all');
  for (const text of ['B3for3', 'Insid3 th3 c3ll', 'Aft3r']) expect(xml).toContain(`>${text}<`);
  expect(xml).toContain(TBL_PR);
  expect(xml).toContain(TC_PR);
  expect(runWith(xml, 'B3for3')).toContain('<w:noProof/>');
});
