import { expect, test, type Locator, type Page } from 'playwright/test';
import JSZip from 'jszip';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const caret = '[data-docx-worker-editor] [data-docx-caret]';
const run = (text: string, bold = false) => `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t>${text}</w:t></w:r>`;
const textBody = `<w:p>${run('Plain words here')}</w:p><w:p>${run('Bold words here', true)}</w:p><w:p>${run('pl')}${run('BOLD', true)}</w:p>`;

function picture(description: string) {
  return '<w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="7" name="Picture" descr="' + description + '"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
}

async function open(page: Page, content: 'text' | 'pictures' = 'text'): Promise<Locator> {
  const zip = new JSZip();
  const pictures = content === 'pictures';
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    (pictures ? '<Default Extension="png" ContentType="image/png"/>' : '') + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const namespaces = pictures ? ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' : '';
  const body = pictures ? `<w:p>${picture('First')}${picture('Second')}${run('Seed')}</w:p>` : textBody;
  zip.file('word/document.xml', `<w:document ${W}><w:body${namespaces}>${body}` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  if (pictures) {
    zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.png"/></Relationships>');
    zip.file('word/media/image.png', Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==', 'base64'));
  }
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'drag-reads.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60_000 });
  const input = page.locator('.paged-editor__input');
  await input.focus();
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  await expect(page.locator(caret)).toHaveCount(1);
  return input;
}

async function recordNativeCopies(page: Page) {
  await page.addInitScript(() => {
    const denied = () => Promise.reject(new DOMException('Write permission denied.', 'NotAllowedError'));
    Object.defineProperty(Navigator.prototype, 'clipboard', { configurable: true,
      get: () => ({ write: denied, writeText: denied, read: denied, readText: denied }) });
    document.addEventListener('copy', (event: ClipboardEvent) => {
      const text = event.clipboardData?.getData('text/plain');
      if (text) (window as unknown as { nativeClipboardText: string }).nativeClipboardText = text;
    });
  });
}

async function caretBox(page: Page) {
  await expect(page.locator(caret)).toHaveCount(1);
  const box = await page.locator(caret).first().boundingBox();
  if (!box) throw new Error('The caret has no box');
  return box;
}

async function lineEnd(page: Page) {
  const sheet = await page.locator('.paged-editor__page').first().boundingBox();
  if (!sheet) throw new Error('The page has no box');
  return sheet.x + (816 - 96) * (sheet.width / 816) - 2;
}

async function dragAlongLine(page: Page, from: { x: number; y: number; height: number }) {
  const y = from.y + from.height / 2;
  const end = await lineEnd(page);
  await page.mouse.move(from.x + 1, y);
  await page.mouse.down();
  for (let step = 1; step <= 12; step++) await page.mouse.move(from.x + 1 + (end - from.x - 1) * step / 12, y);
  await page.mouse.up();
}

const nativeCopy = (page: Page) => page.evaluate(() => (window as unknown as { nativeClipboardText?: string }).nativeClipboardText ?? '');

for (const wait of [0, 200]) test(`a copy ${wait ? `${wait} ms after` : 'at'} the end of a drag holds the dragged text`, async ({ page }) => {
  await recordNativeCopies(page);
  const input = await open(page);
  await dragAlongLine(page, await caretBox(page));
  if (wait) await page.waitForTimeout(wait);
  await input.press('ControlOrMeta+c');
  await expect.poll(() => nativeCopy(page)).toBe('Plain words here');
});

test('the toolbar Bold state follows a drag at release and Shift+Arrow at each step', async ({ page }) => {
  const input = await open(page);
  const bold = page.getByTestId('formatting-bar').getByRole('button', { name: 'Bold', exact: true });
  const before = await caretBox(page);
  await input.press('ArrowDown');
  await expect.poll(async () => (await caretBox(page)).y).toBeGreaterThan(before.y);
  const boldLine = await caretBox(page);
  await input.press('ArrowUp');
  await expect(bold).toHaveAttribute('aria-pressed', 'false');
  await dragAlongLine(page, boldLine);
  await expect(bold).toHaveAttribute('aria-pressed', 'true', { timeout: 1_000 });

  await input.press('ControlOrMeta+End');
  await input.press('Home');
  await input.press('ArrowRight');
  await input.press('ArrowRight');
  for (const [key, pressed] of [['Shift+ArrowRight', 'true'], ['Shift+ArrowRight', 'true'], ['Shift+ArrowLeft', 'true'], ['Shift+ArrowLeft', null], ['Shift+ArrowLeft', 'false']] as const) {
    await input.press(key);
    if (pressed) await expect(bold).toHaveAttribute('aria-pressed', pressed, { timeout: 1_000 });
  }
});

test('clicking a picture shows its properties', async ({ page }) => {
  await open(page, 'pictures');
  const surface = page.locator('.paged-editor__page').first();
  const box = (await surface.boundingBox())!;
  await surface.click({ position: { x: 146 * box.width / 816, y: 121 * box.height / 1056 }, force: true });
  const inline = page.getByRole('button', { name: 'Image properties', exact: true });
  const more = page.getByRole('button', { name: 'More actions', exact: true });
  await expect(inline).toBeAttached();
  await expect(inline).not.toHaveAttribute('aria-disabled', 'true');
  if (await inline.isVisible()) await inline.click();
  else {
    await more.click();
    await page.getByRole('menuitem', { name: 'Image properties', exact: true }).click();
  }
  const dialog = page.getByRole('dialog', { name: 'Image Properties', exact: true });
  await expect(dialog).toBeVisible();
  await expect(dialog.locator('textarea')).toHaveValue('First');
});
