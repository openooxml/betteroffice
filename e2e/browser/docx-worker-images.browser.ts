import { expect, test, type Locator, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

async function open(page: Page) {
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
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'worker-properties.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
  const input = page.locator('.paged-editor__peerless-input');
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

async function imageToolbarTarget(inline: Locator, more: Locator): Promise<'inline' | 'more'> {
  await expect(inline).toBeAttached();
  await expect(inline).not.toHaveAttribute('aria-disabled', 'true');
  let target: 'inline' | 'more' | null = null;
  await expect.poll(async () => {
    if ((await inline.getAttribute('aria-disabled')) === 'true') return (target = null);
    if (await inline.isVisible()) return (target = 'inline');
    if (await more.isVisible()) return (target = 'more');
    return (target = null);
  }).not.toBeNull();
  return target!;
}

async function imageDialog(page: Page) {
  const input = await open(page);
  await input.press('Shift+ArrowRight');
  const inline = page.getByRole('button', { name: 'Image properties', exact: true });
  const more = page.getByRole('button', { name: 'More actions', exact: true });
  const target = await imageToolbarTarget(inline, more);
  if (target === 'inline') {
    await expect(inline).toBeVisible();
    await expect(inline).toBeEnabled();
    await inline.click();
  } else {
    await expect(more).toBeVisible();
    await expect(more).toBeEnabled();
    await more.click();
    const item = page.getByRole('menuitem', { name: 'Image properties', exact: true });
    await expect(item).toBeVisible();
    await expect(item).not.toHaveAttribute('aria-disabled', 'true');
    await item.click();
  }
  const dialog = page.getByRole('dialog', { name: 'Image Properties', exact: true });
  await expect(dialog).toBeVisible();
  await expect(dialog.locator('textarea')).toHaveValue('First');
  return dialog;
}

test('worker image properties edit the first of two images with a shared authored ID', async ({ page }, info) => {
  const dialog = await imageDialog(page);
  await dialog.locator('textarea').fill('Edited first');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  const xml = await savedXml(page, info);
  const descriptions = await page.evaluate(xml => [...new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS('*', 'docPr')].map(node => node.getAttribute('descr')), xml);
  expect(descriptions).toEqual(['Edited first', 'Second']);
});

test('worker image position applies to the same captured image', async ({ page }, info) => {
  const dialog = await imageDialog(page);
  await dialog.getByRole('button', { name: 'Image Position', exact: true }).click();
  const position = page.getByRole('dialog', { name: 'Image Position', exact: true });
  await expect(position).toBeVisible();
  await position.getByRole('button', { name: 'Apply', exact: true }).click();
  const xml = await savedXml(page, info);
  const kinds = await page.evaluate(xml => [...new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS('*', 'drawing')].map(node => node.firstElementChild?.localName), xml);
  expect(kinds).toEqual(['anchor', 'inline']);
});

async function selectImageWrap(page: Page, optionTestId: string, label: string) {
  const input = page.locator('.paged-editor__peerless-input');
  const inline = page.getByTestId('toolbar-image-wrap');
  const more = page.getByRole('button', { name: 'More actions', exact: true });
  const target = await imageToolbarTarget(inline, more);
  if (target === 'inline') {
    await expect(inline).toBeVisible();
    await expect(inline).toBeEnabled();
    await inline.click();
    const option = page.getByTestId(`toolbar-image-wrap-${optionTestId}`);
    await expect(option).toBeVisible();
    await input.blur();
    await option.click();
  } else {
    await expect(more).toBeVisible();
    await expect(more).toBeEnabled();
    await more.click();
    const submenu = page.getByRole('menuitem', { name: 'Text wrapping', exact: true });
    await expect(submenu).toBeVisible();
    await expect(submenu).not.toHaveAttribute('aria-disabled', 'true');
    await submenu.hover();
    const option = page.getByRole('menuitemradio', { name: label, exact: true });
    await expect(option).toBeVisible();
    await input.blur();
    await option.click();
  }
  await expect(input).toBeFocused();
  await expect(inline).toHaveAttribute('aria-label', `Wrap: ${label}`);
  await expect(inline).not.toHaveAttribute('aria-disabled', 'true');
}

async function selectImageTransform(page: Page, action: string, label: string) {
  const input = page.locator('.paged-editor__peerless-input');
  const inline = page.getByTestId('toolbar-image-transform');
  const more = page.getByRole('button', { name: 'More actions', exact: true });
  const target = await imageToolbarTarget(inline, more);
  if (target === 'inline') {
    await expect(inline).toBeVisible();
    await expect(inline).toBeEnabled();
    await inline.click();
    const option = page.getByTestId(`toolbar-image-transform-${action}`);
    await expect(option).toBeVisible();
    await input.blur();
    await option.click();
  } else {
    await expect(more).toBeVisible();
    await expect(more).toBeEnabled();
    await more.click();
    const submenu = page.getByRole('menuitem', { name: 'Transform', exact: true });
    await expect(submenu).toBeVisible();
    await expect(submenu).not.toHaveAttribute('aria-disabled', 'true');
    await submenu.hover();
    const option = page.getByRole('menuitem', { name: label, exact: true });
    await expect(option).toBeVisible();
    await input.blur();
    await option.click();
  }
  await expect(input).toBeFocused();
  await expect(inline).toBeAttached();
  await expect(inline).not.toHaveAttribute('aria-disabled', 'true');
}

test('worker image wrap and transform round-trip through saved OOXML', async ({ page }, info) => {
  const input = await open(page);
  await input.press('Shift+ArrowRight');
  await selectImageWrap(page, 'wrapRight', 'Square Left');
  await selectImageWrap(page, 'behind', 'Behind Text');
  await selectImageTransform(page, 'rotateCW', 'Rotate clockwise');
  await selectImageTransform(page, 'flipH', 'Flip horizontal');
  const xml = await savedXml(page, info);
  expect(xml).toContain('<wp:anchor');
  expect(xml).toContain('behindDoc="1"');
  expect(xml).toContain('<wp:wrapNone/>');
  expect(xml).toContain('<a:xfrm');
  expect(xml).toContain('rot="5400000"');
  expect(xml).toContain('flipH="1"');
});
