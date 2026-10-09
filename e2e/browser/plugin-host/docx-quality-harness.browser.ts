import { test, expect } from 'playwright/test';
import JSZip from 'jszip';

const PAGES = 60;
const HOLD_MS = 5_000;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function pagedDocx(pages: number): Promise<number[]> {
  const body = Array.from(
    { length: pages },
    (_, index) =>
      `<w:p>${index ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : ''}` +
      `<w:r><w:t>Synthetic page ${index + 1}</w:t></w:r></w:p>`
  ).join('');
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>'
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W}"><w:body>${body}` +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
      '</w:sectPr></w:body></w:document>'
  );
  return [...(await zip.generateAsync({ type: 'uint8array' }))];
}

test('the DOCX quality harness captures every page when the first pages paint first', async ({
  page,
}) => {
  // Holds the rest of the layout past the harness's page-count stability window.
  await page.addInitScript((hold) => {
    const state = window as unknown as { __held: boolean };
    state.__held = false;
    const post = Worker.prototype.postMessage;
    const queued = new WeakMap<Worker, unknown[][]>();
    Worker.prototype.postMessage = function (this: Worker, ...args: unknown[]) {
      const queue = queued.get(this);
      if (queue) return void queue.push(args);
      if ((args[0] as { type?: string } | null)?.type !== 'completeLayout') {
        return post.apply(this, args as Parameters<typeof post>);
      }
      state.__held = true;
      const held = [args];
      queued.set(this, held);
      setTimeout(() => {
        queued.delete(this);
        for (const message of held) post.apply(this, message as Parameters<typeof post>);
      }, hold);
    } as typeof post;
  }, HOLD_MS);
  await page.goto('/docx-quality.html');
  await page.waitForFunction(() => (window as unknown as { oracleReady?: boolean }).oracleReady, null, {
    timeout: 120_000,
  });
  const result = await page.evaluate(
    (bytes) =>
      (
        window as unknown as {
          oracleInit(bytes: number[], fonts: boolean): Promise<{ pages: number }>;
        }
      ).oracleInit(bytes, false),
    await pagedDocx(PAGES)
  );
  expect(await page.evaluate(() => (window as unknown as { __held: boolean }).__held)).toBe(true);
  expect(result.pages).toBe(PAGES);
  const last = await page.evaluate(
    (index) =>
      (window as unknown as { oraclePage(index: number): Promise<string> }).oraclePage(index),
    PAGES - 1
  );
  expect(last).toMatch(/^data:image\/png;base64,/);
});
