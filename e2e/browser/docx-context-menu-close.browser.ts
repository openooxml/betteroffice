import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

async function open(page: Page, arm: string) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Context menu target paragraph</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  await page.goto('/?format=docx' + arm);
  await page.locator('input[type=file]').last().setInputFiles({ name: 'context-menu.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  const surface = page.locator('.canvas-page, .paged-editor__page').first();
  await expect(surface.locator('canvas').first()).toBeVisible({ timeout: 60_000 });
  await surface.click({ position: { x: 140, y: 100 } });
  await expect(page.locator('[data-testid=yrs-input], .paged-editor__input').first()).toBeEditable();
  return surface;
}

async function openMenu(page: Page, surface: ReturnType<Page['locator']>) {
  await surface.click({ button: 'right', position: { x: 140, y: 100 } });
  const menu = page.getByRole('menu', { name: 'Text editing menu' });
  await expect(menu).toBeVisible();
  return menu;
}

for (const [arm, label] of [['', 'worker'], ['&peerless=0', 'non-worker']] as const) {
  test(`docx text context menu closes on a click inside the page (${label})`, async ({ page }) => {
    const surface = await open(page, arm);
    const menu = await openMenu(page, surface);
    await surface.click({ position: { x: 300, y: 400 } });
    await expect(menu).toBeHidden();
  });

  test(`docx text context menu closes on a click outside the editor (${label})`, async ({ page }) => {
    const surface = await open(page, arm);
    const menu = await openMenu(page, surface);
    await page.mouse.click(4, 996);
    await expect(menu).toBeHidden();
  });

  test(`docx text context menu stays open while clicking inside it (${label})`, async ({ page }) => {
    const surface = await open(page, arm);
    const menu = await openMenu(page, surface);
    const box = (await menu.boundingBox())!;
    await page.mouse.click(box.x + 2, box.y + 2);
    await expect(menu).toBeVisible();
  });

  test(`docx text context menu closes on a press outside in the same turn it opens (${label})`, async ({ page }) => {
    const surface = await open(page, arm);
    const box = (await surface.boundingBox())!;
    const pressed = await page.evaluate(({ x, y }) => new Promise<boolean>((resolve) => {
      const observer = new MutationObserver(() => {
        if (!document.querySelector('[role="menu"][aria-label="Text editing menu"]')) return;
        observer.disconnect();
        const outside = document.elementFromPoint(4, 996) ?? document.body;
        const init = { bubbles: true, cancelable: true, composed: true, clientX: 4, clientY: 996, button: 0, buttons: 1 };
        outside.dispatchEvent(new PointerEvent('pointerdown', { ...init, pointerType: 'mouse', isPrimary: true }));
        outside.dispatchEvent(new MouseEvent('mousedown', init));
        resolve(true);
      });
      observer.observe(document.body, { childList: true, subtree: true });
      document.elementFromPoint(x, y)!.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true,
        composed: true, clientX: x, clientY: y, button: 2 }));
      setTimeout(() => { observer.disconnect(); resolve(false); }, 5000);
    }), { x: box.x + 140, y: box.y + 100 });
    expect(pressed).toBe(true);
    await expect(page.getByRole('menu', { name: 'Text editing menu' })).toBeHidden();
  });

  test(`docx text context menu stays open while typing, like non-worker mode (${label})`, async ({ page }) => {
    const surface = await open(page, arm);
    const menu = await openMenu(page, surface);
    await page.keyboard.type('x');
    await page.waitForTimeout(300);
    await expect(menu).toBeVisible();
  });
}
