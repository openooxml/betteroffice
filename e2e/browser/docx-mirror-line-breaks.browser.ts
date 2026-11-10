import { test, expect } from 'playwright/test';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const t = (text: string) => `<w:t xml:space="preserve">${text}</w:t>`;
const p = (inner: string) => `<w:p><w:r>${inner}</w:r></w:p>`;
const cell = (inner: string) => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p(inner)}</w:tc>`;
const WRAP = 'Alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november oscar papa quebec romeo sierra tango uniform victor whiskey xray yankee zulu.';
const CJK = '日本語のテキストは単語の間に空白がありません。行の折り返しで空白を入れてはいけません。日本語のテキストは単語の間に空白がありません。';

async function breaksDocument(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`);
  const table = `<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr>${
    cell(`${t('Cellspecial')}<w:br/>${t('Cellappropriation')}`)}${cell(`${t('Celltab')}<w:tab/>${t('Cellword')}`)}</w:tr></w:tbl>`;
  const body = p(`${t('Special')}<w:br/>${t('Appropriation')}`) + p(`${t('Tabbed')}<w:tab/>${t('Word')}`)
    + p(`${t('Double')}<w:br/><w:br/>${t('Gap')}`) + table + p(t(WRAP)) + p(t('state-of-the-art well-known')) + p(t(CJK));
  zip.file('word/document.xml', `<w:document xmlns:w="${w}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

type AXNode = { ignored?: boolean; role?: { value?: string }; name?: { value?: string } };

for (const mirror of ['full', 'text'] as const) {
  test(`docx ${mirror} page mirror reads explicit line breaks and tabs as separators`, async ({ browser }) => {
    test.setTimeout(120_000);
    const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
    try {
      const page = await context.newPage();
      await page.goto(`/?format=docx&peerless=1${mirror === 'text' ? '&pageMirror=text' : ''}`);
      await page.locator('input[type=file]').last().setInputFiles({
        name: 'breaks.docx',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        buffer: await breaksDocument(),
      });
      const pageMirror = page.locator('.canvas-page[data-page-index="0"] .canvas-page-mirror');
      await expect(pageMirror).toContainText('Appropriation', { timeout: 60_000 });
      const texts = () => page.evaluate(() => Array.from(
        document.querySelectorAll<HTMLElement>('.canvas-page[data-page-index="0"] .canvas-page-mirror [role="paragraph"]'),
        (paragraph) => paragraph.textContent ?? ''));
      await expect.poll(texts, { timeout: 30_000 }).toContain('Special\nAppropriation');
      const read = await texts();
      for (const text of ['Tabbed\tWord', 'Double\n\nGap', 'Cellspecial\nCellappropriation', 'Celltab\tCellword', WRAP, 'state-of-the-art well-known', CJK]) {
        expect(read).toContain(text);
      }

      const cdp = await context.newCDPSession(page);
      const { nodes } = await cdp.send('Accessibility.getFullAXTree') as { nodes: AXNode[] };
      const names = nodes.filter((node) => !node.ignored).map((node) => String(node.name?.value ?? ''));
      for (const fused of ['SpecialAppropriation', 'TabbedWord', 'DoubleGap', 'CellspecialCellappropriation', 'CelltabCellword']) {
        expect(names.filter((name) => name.includes(fused))).toEqual([]);
      }
      const cells = nodes.filter((node) => !node.ignored && node.role?.value === 'cell').map((node) => String(node.name?.value ?? ''));
      expect(cells.some((name) => /^Cellspecial\s+Cellappropriation$/.test(name))).toBe(true);

      if (mirror === 'text') {
        expect(await page.evaluate(() => (window as unknown as { find(text: string): boolean }).find('Special Appropriation'))).toBe(true);
        const selected = await page.evaluate(() => {
          const paragraph = Array.from(document.querySelectorAll<HTMLElement>('.canvas-page-mirror [role="paragraph"]'))
            .find((element) => (element.textContent ?? '').startsWith('Special'))!;
          const selection = getSelection()!;
          selection.removeAllRanges();
          const range = document.createRange();
          range.selectNodeContents(paragraph);
          selection.addRange(range);
          const text = selection.toString();
          selection.removeAllRanges();
          return text;
        });
        expect(selected).toBe('Special Appropriation');
      }
    } finally {
      await context.close();
    }
  });
}
