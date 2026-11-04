import { test, expect, type Browser } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');
const PAGES = 3;
const SETTLE = '[data-docx-caret], .paged-editor__input { visibility: hidden !important; } * { caret-color: transparent !important; }';

async function longDocument(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.docx')));
  const body = Array.from({ length: PAGES }, (_, index) => {
    const number = String(index + 1).padStart(2, '0');
    const heading = `<w:p><w:pPr>${index ? '<w:pageBreakBefore/>' : ''}</w:pPr><w:r><w:t>Reading page ${number}</w:t></w:r></w:p>`;
    return heading + Array.from({ length: 20 }, (_, line) =>
      `<w:p><w:r><w:t xml:space="preserve">Marker${number} line ${line + 1} keeps the quick brown fox jumping over the lazy dog.</w:t></w:r></w:p>`).join('');
  }).join('');
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(browser: Browser, worker: boolean) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
  const page = await context.newPage();
  await page.goto(`/?format=docx&peerless=${worker ? 1 : 0}`);
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'reading.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await longDocument(),
  });
  await expect(page.locator('.canvas-page[data-page-index="1"] canvas')).toBeAttached({ timeout: 60_000 });
  await expect(page.getByRole('status').filter({ hasText: /^\d+ of \d+$/ })).toHaveText(`1 of ${PAGES}`, { timeout: 60_000 });
  await expect(page.locator('.notice, [data-testid$="-error"]')).toHaveCount(0);
  if (worker) await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false');
  await page.addStyleTag({ content: SETTLE });
  return { context, page };
}

type AXNode = { ignored?: boolean; name?: { value?: string }; role?: { value?: string } };

for (const worker of [true, false]) {
  test(`docx ${worker ? 'worker' : 'in-thread'} editor: the page mirror reads whole words`, async ({ browser }, info) => {
    test.setTimeout(120_000);
    const { context, page } = await open(browser, worker);
    try {
      const cdp = await context.newCDPSession(page);
      const named = async () => {
        const { nodes } = await cdp.send('Accessibility.getFullAXTree') as { nodes: AXNode[] };
        return nodes.filter((node) => !node.ignored && node.role?.value === 'paragraph')
          .map((node) => String(node.name?.value ?? '')).filter((name) => name.includes('Marker01') || name.includes('M a r k e r'));
      };
      await expect.poll(async () => (await named()).length, { timeout: 30_000 }).toBeGreaterThan(0);
      const names = await named();
      await info.attach('paragraph-names.json', { body: JSON.stringify(names.slice(0, 5), null, 2), contentType: 'application/json' });
      expect(names).toContain('Marker01 line 1 keeps the quick brown fox jumping over the lazy dog.');
      const snapshot = await page.locator('.canvas-page[data-page-index="0"]').first().ariaSnapshot();
      await info.attach('page-1.yml', { body: snapshot, contentType: 'text/yaml' });
      expect(snapshot).toContain('paragraph: Marker01 line 1 keeps the quick brown fox');
    } finally {
      await context.close();
    }
  });
}
