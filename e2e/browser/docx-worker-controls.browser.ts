import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

async function open(page: Page) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  const checkbox = '<w:sdt><w:sdtPr><w:id w:val="7"/><w:alias w:val="Flag"/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>';
  const dropdown = '<w:sdt><w:sdtPr><w:id w:val="7"/><w:alias w:val="Choose option"/><w:dropDownList><w:listItem w:displayText="First" w:value="one"/><w:listItem w:displayText="Second" w:value="two"/></w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>First</w:t></w:r></w:sdtContent></w:sdt>';
  const date = '<w:sdt><w:sdtPr><w:id w:val="7"/><w:alias w:val="Choose date"/><w:date w:fullDate="2026-10-12T00:00:00Z"><w:dateFormat w:val="yyyy-MM-dd"/></w:date></w:sdtPr><w:sdtContent><w:r><w:t>2026-10-12</w:t></w:r></w:sdtContent></w:sdt>';
  const repeat = '<w:sdt><w:sdtPr><w:id w:val="8"/><w15:repeatingSection/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="9"/><w:alias w:val="Line"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Item text</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt>';
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('xmlns:w14=', 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w14=').replace('<w:r><w:t>Seed</w:t></w:r>', checkbox + '<w:r><w:t>gap</w:t></w:r>' + checkbox + dropdown + date).replace('<w:sectPr>', repeat + '<w:sectPr>'));
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'worker-properties.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
  const input = page.locator('.paged-editor__input');
  await input.focus();
  await input.press('ControlOrMeta+Home');
  return input;
}

async function savedXml(page: Page, info: TestInfo, part = 'word/document.xml') {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('properties.docx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  return zip.file(part)!.async('string');
}

test('worker widgets distinguish duplicate IDs and save checkbox dropdown and date values', async ({ page }, info) => {
  await open(page);
  const widgets = page.locator('[data-worker-control]');
  const flags = widgets.getByRole('checkbox', { name: 'Flag', exact: true });
  await expect(flags).toHaveCount(2);
  await flags.nth(1).click();
  await expect(flags.nth(0)).not.toBeChecked();
  await expect(flags.nth(1)).toBeChecked();
  await widgets.getByRole('combobox', { name: 'Choose option', exact: true }).selectOption('two');
  await widgets.getByLabel('Choose date', { exact: true }).fill('2026-11-03');
  const xml = await savedXml(page, info);
  const checks = [...xml.matchAll(/w14:checked w14:val="(\d)"/g)].map(match => match[1]);
  expect(checks).toEqual(['0', '1']);
  expect(xml).toContain('Second');
  expect(xml).toContain('2026-11-03');
});

test('worker repeating controls refresh physical handles after cloning and removal', async ({ page }, info) => {
  await open(page);
  const add = page.getByRole('button', { name: 'Add item', exact: true });
  const remove = page.getByRole('button', { name: 'Remove item', exact: true });
  await expect(add).toHaveCount(1);
  await add.click();
  await expect(add).toHaveCount(2);
  await remove.nth(1).click();
  await expect(add).toHaveCount(1);
  const xml = await savedXml(page, info);
  expect([...xml.matchAll(/<w15:repeatingSectionItem/g)]).toHaveLength(1);
});
