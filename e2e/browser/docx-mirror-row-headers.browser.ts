import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (texts: string[], header = false) =>
  `<w:tr>${header ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${texts.map(cell).join('')}</w:tr>`;

/** Header row plus first column (Word's default look 04A0), then a table
 *  with the same look but no header row. */
function table(look: string, rows: string) {
  return `<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/><w:tblLook ${look}/></w:tblPr>` +
    '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' + rows + '</w:tbl><w:p/>';
}

async function docxFile() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`);
  const look = 'w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"';
  zip.file('word/document.xml', `<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body>` +
    '<w:p><w:r><w:t>Sales</w:t></w:r></w:p>' +
    table(look, row(['Region', 'Q1'], true) + row(['North', '12']) + row(['South', '8'])) +
    table(look, row(['Layout', 'grid']) + row(['Left', 'Right'])) +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, worker: boolean) {
  await page.goto(`/?format=docx&peerless=${worker ? 1 : 0}`);
  await page.locator('input[type=file]').last().setInputFiles({ name: 'row-headers.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await docxFile() });
  await expect(page.locator('[data-page-index="0"] canvas').first()).toBeVisible({ timeout: 60000 });
  await expect(page.locator('.canvas-page-mirror').getByText('South')).toBeAttached();
}

type AXNode = { nodeId: string; parentId?: string; ignored?: boolean; role?: { value?: unknown }; name?: { value?: unknown }; childIds?: string[] };
const cellRoles = new Set(['columnheader', 'rowheader', 'cell']);

/** Every table's header and data cells in tree order, as role and name. */
async function tables(page: Page) {
  const cdp = await page.context().newCDPSession(page);
  const { nodes } = await cdp.send('Accessibility.getFullAXTree') as { nodes: AXNode[] };
  await cdp.detach();
  const byId = new Map(nodes.map((node) => [node.nodeId, node]));
  const out: string[][][] = [];
  const visit = (node: AXNode | undefined, cells: string[][] | null) => {
    if (!node) return;
    const role = String(node.role?.value);
    let into = cells;
    if (!node.ignored && role === 'table') out.push(into = []);
    else if (!node.ignored && into && cellRoles.has(role)) into.push([role, String(node.name?.value ?? '').replace(/\s+/g, ' ').trim()]);
    for (const id of node.childIds ?? []) visit(byId.get(id), into);
  };
  visit(nodes.find((node) => !node.parentId && !node.ignored) ?? nodes[0], null);
  return out;
}

for (const worker of [false, true]) {
  test(`docx mirror tables head their rows from the first column only under header rows (${worker ? 'worker' : 'in-thread'} editor)`, async ({ page }) => {
    await open(page, worker);
    const mirror = page.locator('.canvas-page-mirror');
    await expect(mirror.locator('[role="rowheader"]')).toHaveText(['North', 'South']);
    await expect(mirror.locator('[role="columnheader"]')).toHaveText(['Region', 'Q1']);
    const labelled = mirror.locator('[role="cell"]').filter({ hasText: '12' });
    const ids = (await labelled.getAttribute('aria-labelledby'))!.split(' ');
    const names = await Promise.all(ids.map((id) => mirror.locator(`[id="${id}"]`).textContent()));
    expect(names.map((name) => name?.trim())).toEqual(['Q1', 'North']);
    await expect.poll(() => tables(page)).toEqual([
      [['columnheader', 'Region'], ['columnheader', 'Q1'], ['rowheader', 'North'], ['cell', 'Q1 North'], ['rowheader', 'South'], ['cell', 'Q1 South']],
      [['cell', 'Layout'], ['cell', 'grid'], ['cell', 'Left'], ['cell', 'Right']],
    ]);
  });
}
