import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';
import { caretBoxAtRunStart, movedCaretBox } from './rendered-box';
import type { DocxEngineWorkerRequest } from '../../packages/docx/src/yrs/residentEngineWorkerProtocol';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

declare global {
  interface Window {
    __docxShiftDelete: { typingHeld(): boolean; releaseTyping(): void; noticeShown: boolean };
  }
}

async function documentBytes(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${DOCX}.main+xml"/></Types>`);
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document ${W}><w:body><w:p><w:r><w:t>ab</w:t></w:r></w:p>` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function holdTyping(page: Page) {
  await page.addInitScript(() => {
    let hold = true;
    const held: (() => void)[] = [];
    const NativeWorker = window.Worker;
    window.Worker = class extends NativeWorker {
      postMessage(message: unknown, transfer?: Transferable[] | StructuredSerializeOptions) {
        const request = message as DocxEngineWorkerRequest;
        const post = () => {
          if (Array.isArray(transfer)) super.postMessage(message, transfer);
          else super.postMessage(message, transfer);
        };
        if (hold && request?.type === 'edit' && request.command.kind === 'insertText') held.push(post);
        else post();
      }
    };
    window.__docxShiftDelete = {
      typingHeld: () => held.length > 0,
      releaseTyping() {
        hold = false;
        for (const post of held.splice(0)) post();
      },
      noticeShown: false,
    };
  });
}

async function openAtCaret(page: Page) {
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'shift-delete.docx',
    mimeType: DOCX, buffer: await documentBytes() });
  const editor = page.locator('[data-docx-worker-editor]');
  await expect(editor.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60_000 });
  const mirror = editor.locator('[data-page-index="0"] .canvas-page-mirror');
  await expect(mirror).toHaveText('ab');
  const input = editor.locator('.paged-editor__input');
  await expect(input).toBeEditable();
  await input.focus();
  const caret = editor.locator('[data-docx-caret]');
  await caretBoxAtRunStart(caret, mirror.locator('.layout-run-text').first(), () => input.press('ControlOrMeta+Home'));
  await movedCaretBox(caret, () => input.press('ArrowRight'));
  return { input, mirror };
}

test('worker Shift+Delete deletes forward after pending typing without a notice', async ({ page }) => {
  await holdTyping(page);
  const { input, mirror } = await openAtCaret(page);
  let nativeAlert = false;
  page.on('dialog', async dialog => {
    nativeAlert ||= dialog.type() === 'alert';
    await dialog.dismiss();
  });
  await page.evaluate(() => {
    const visibleNotices = () => Array.from(document.querySelectorAll('[data-sonner-toast], [role="alert"]')).filter(node =>
      node.textContent?.trim() && node.getClientRects().length > 0 && getComputedStyle(node).visibility === 'visible');
    const existing = new Map(visibleNotices().map(node => [node, node.textContent]));
    new MutationObserver(() => {
      window.__docxShiftDelete.noticeShown ||= visibleNotices().some(node => existing.get(node) !== node.textContent);
    }).observe(document.documentElement, { subtree: true, childList: true, characterData: true, attributes: true,
      attributeFilter: ['role', 'data-sonner-toast', 'hidden', 'class', 'style'] });
  });

  await page.keyboard.type('x');
  await expect.poll(() => page.evaluate(() => window.__docxShiftDelete.typingHeld())).toBe(true);
  await input.press('Shift+Delete');
  await page.evaluate(() => window.__docxShiftDelete.releaseTyping());

  await expect(mirror).toHaveText('ax');
  expect(await page.evaluate(() => window.__docxShiftDelete.noticeShown)).toBe(false);
  expect(nativeAlert).toBe(false);
});

test('worker Shift+Delete deletes forward at a settled caret', async ({ page }) => {
  const { input, mirror } = await openAtCaret(page);
  await input.press('Shift+Delete');
  await expect(mirror).toHaveText('a');
});
