import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

async function open(page: Page, content: 'images' | 'table' | 'text' = 'images') {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  const image = (description: string) => '<w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="7" name="Picture" descr="' + description + '"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:body>', '<w:body xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">').replace('<w:r><w:t>Seed</w:t></w:r>', image('First') + image('Second') + '<w:r><w:t>Seed</w:t></w:r>'));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'));
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.png"/></Relationships>');
  zip.file('word/media/image.png', Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==', 'base64'));
  if (content !== 'images') {
    let document = (await zip.file('word/document.xml')!.async('string')).replace(image('First') + image('Second'), '');
    if (content === 'table') document = document.replace('<w:sectPr>', '<w:tbl><w:tblPr><w:tblW w:w="5760" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>First cell</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Second cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr>');
    zip.file('word/document.xml', document);
  }
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'worker-properties.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
  const input = page.locator('.paged-editor__peerless-input');
  await input.focus();
  await input.press('ControlOrMeta+Home');
  return input;
}

async function objectMenu(page: Page, x: number, y: number, action: string) {
  const surface = page.locator('.paged-editor__page').first();
  const box = (await surface.boundingBox())!;
  await surface.click({ button: 'right', position: { x: x * box.width / 816, y: y * box.height / 1056 }, force: true });
  await page.getByRole('menuitem', { name: action, exact: true }).click();
  return page.getByRole('dialog', { name: action, exact: true });
}

test('worker image context resize changes the clicked duplicate image and saves it', async ({ page }, info) => {
  await open(page);
  const dialog = await objectMenu(page, 246, 121, 'Resize image');
  await dialog.getByLabel('Width (px)').fill('80');
  await dialog.getByLabel('Height (px)').fill('40');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(dialog).toBeHidden();
  const xml = await savedXml(page, info);
  const sizes = await page.evaluate(xml => [...new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS('*', 'drawing')].map(node => {
    const extent = node.getElementsByTagNameNS('*', 'extent')[0]!;
    return [extent.getAttribute('cx'), extent.getAttribute('cy')];
  }), xml);
  expect(sizes).toEqual([['952500', '476250'], ['762000', '381000']]);
});

test('worker image context move uses the clicked physical image', async ({ page }, info) => {
  await open(page);
  const dialog = await objectMenu(page, 246, 121, 'Move image');
  await dialog.getByLabel('Left (px)').fill('210');
  await dialog.getByLabel('Top (px)').fill('180');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(dialog).toBeHidden();
  const xml = await savedXml(page, info);
  const positions = await page.evaluate(xml => [...new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS('*', 'drawing')].map(node => ({
    kind: node.firstElementChild?.localName,
    offsets: [...node.getElementsByTagNameNS('*', 'posOffset')].map(value => value.textContent),
  })), xml);
  expect(positions).toEqual([{ kind: 'inline', offsets: [] }, { kind: 'anchor', offsets: ['2000250', '1714500'] }]);
});

test('worker cell context width targets the clicked column and saves it', async ({ page }, info) => {
  await open(page, 'table');
  const dialog = await objectMenu(page, 384, 125, 'Column width');
  await dialog.getByLabel('Width (px)').fill('160');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(dialog).toBeHidden();
  const xml = await savedXml(page, info);
  const widths = await page.evaluate(xml => [...new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS('*', 'gridCol')].map(node => node.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w')), xml);
  expect(widths).toEqual(['2880', '2400']);
});

async function color(page: Page, title: string, choice: string) {
  const picker = page.locator(`.docx-color-picker-arrow[aria-label="${title}"]`);
  if (await picker.isVisible()) {
    await picker.click();
    await page.getByRole('dialog', { name: `${title} picker`, exact: true }).getByRole('button', { name: choice, exact: true }).click();
  } else {
    await page.getByTestId('toolbar-more').click();
    const entry = page.getByRole('menu', { name: 'More actions', exact: true }).locator(`[data-label="${title}"]`);
    await entry.focus();
    await page.keyboard.press('ArrowRight');
    let menu = page.getByRole('menu', { name: title, exact: true });
    if (choice.startsWith('Accent 1,')) {
      for (const label of ['Theme Colors', 'Accent 1']) {
        await menu.locator(`[data-label="${label}"]`).focus();
        await page.keyboard.press('ArrowRight');
        menu = page.getByRole('menu', { name: label, exact: true });
      }
    }
    await menu.locator(`[data-label="${choice}"]`).click();
  }
}

test('worker text color saves theme tint and automatic removes it', async ({ page }, info) => {
  const input = await open(page, 'text');
  await input.press('Shift+End');
  await color(page, 'Font Color', 'Accent 1, Lighter 80%');
  const colored = await savedXml(page, info);
  expect(colored).toContain('w:themeColor="accent1"');
  expect(colored).toContain('w:themeTint="CC"');
  await input.focus();
  await input.press('ControlOrMeta+Home');
  await input.press('Shift+End');
  await color(page, 'Font Color', 'Automatic');
  const cleared = await savedXml(page, info);
  expect(cleared).not.toContain('<w:color');
  expect(cleared).toContain('Seed');
});


async function savedXml(page: Page, info: TestInfo, part = 'word/document.xml') {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('properties.docx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  return zip.file(part)!.async('string');
}
