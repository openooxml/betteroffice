import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';

const drawing = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const presentation = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const namespaces = `xmlns:p="${presentation}" xmlns:a="${drawing}" xmlns:r="${relationship}"`;
const emu = (value: number) => Math.round(value * 9525);
const tree = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';
const run = (text: string, link = '') => `<a:r><a:rPr sz="1800"><a:latin typeface="Liberation Sans"/>${link}</a:rPr><a:t>${text}</a:t></a:r>`;
const title = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(40)}" y="${emu(20)}"/><a:ext cx="${emu(800)}" cy="${emu(60)}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${run('Sales by region')}</a:p></p:txBody></p:sp>`;
const cell = (attributes: string, text: string, link = '') =>
  `<a:tc${attributes}><a:txBody><a:bodyPr/><a:lstStyle/>${text ? `<a:p>${run(text, link)}</a:p>` : '<a:p/>'}</a:txBody><a:tcPr/></a:tc>`;

function table(properties: string) {
  const rows = [
    [cell(' gridSpan="2"', 'Region'), cell(' hMerge="1"', ''), cell('', 'Total')],
    [cell(' rowSpan="2"', 'North'), cell('', 'Q1'), cell('', '')],
    [cell(' vMerge="1"', ''), cell('', 'Q2'), cell('', 'Report', '<a:hlinkClick r:id="url"/>')],
  ].map((cells) => `<a:tr h="${emu(60)}">${cells.join('')}</a:tr>`).join('');
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Sales"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(40)}" y="${emu(120)}"/><a:ext cx="${emu(720)}" cy="${emu(180)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr${properties}/><a:tblGrid>${`<a:gridCol w="${emu(240)}"/>`.repeat(3)}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

async function fixture(properties = ' firstRow="1" firstCol="1"') {
  const zip = new JSZip();
  const relationships = (items: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
  zip.file('_rels/.rels', relationships(`<Relationship Id="rId1" Type="${relationship}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation ${namespaces}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="${emu(960)}" cy="${emu(540)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', relationships(`<Relationship Id="rId1" Type="${relationship}/slide" Target="slides/slide1.xml"/>`));
  zip.file('ppt/slides/_rels/slide1.xml.rels', relationships(
    `<Relationship Id="url" Type="${relationship}/hyperlink" Target="https://example.test/report" TargetMode="External"/>`));
  zip.file('ppt/slides/slide1.xml', `<p:sld ${namespaces}><p:cSld><p:spTree>${tree}${title}${table(properties)}</p:spTree></p:cSld></p:sld>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, properties?: string) {
  await page.goto('/?format=pptx');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'tables.pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', buffer: await fixture(properties),
  });
  await expect(page.getByTestId('pptx-worker-editor')).toBeVisible();
  await expect(page.getByTestId('pptx-worker-status')).toHaveCount(0);
  await expect(page.getByTestId('pptx-slide-canvas')).toHaveAttribute('data-pptx-input-ready', 'true', { timeout: 60_000 });
}

type AXNode = { nodeId: string; ignored?: boolean; role?: { value?: unknown }; name?: { value?: unknown }; childIds?: string[] };

const mirror = (page: Page) => page.getByTestId('pptx-slide-mirror');
const tableRoles = new Set(['table', 'row', 'columnheader', 'rowheader', 'cell']);

/** The table's accessibility subtree in tree order, as role and name; rows by role only. */
async function tableTree(page: Page) {
  const cdp = await page.context().newCDPSession(page);
  const { nodes } = await cdp.send('Accessibility.getFullAXTree') as { nodes: AXNode[] };
  await cdp.detach();
  const byId = new Map(nodes.map((node) => [node.nodeId, node]));
  const out: string[][] = [];
  const visit = (node: AXNode | undefined) => {
    if (!node) return;
    const role = String(node.role?.value);
    if (!node.ignored && tableRoles.has(role)) out.push(role === 'row' ? ['row'] : [role, String(node.name?.value ?? '')]);
    for (const id of node.childIds ?? []) visit(byId.get(id));
  };
  visit(nodes.find((node) => !node.ignored && node.role?.value === 'table'));
  return out;
}

test('pptx mirror tables expose rows, header cells and spans to assistive technology', async ({ page }) => {
  await open(page);
  const document = mirror(page);
  const grid = document.getByRole('table', { name: 'Table, 3 rows, 3 columns' });
  await expect(grid).toBeAttached();
  await expect(grid).toHaveAttribute('aria-rowcount', '3');
  await expect(grid).toHaveAttribute('aria-colcount', '3');
  await expect(grid.getByRole('row')).toHaveCount(3);
  await expect(grid.getByRole('columnheader')).toHaveCount(2);
  await expect(grid.getByRole('rowheader')).toHaveCount(1);
  await expect(grid.getByRole('cell')).toHaveCount(4);
  await expect(grid.getByRole('columnheader', { name: 'Region' })).toHaveAttribute('aria-colspan', '2');
  await expect(grid.getByRole('rowheader', { name: 'North' })).toHaveAttribute('aria-rowspan', '2');
  await expect(grid.getByRole('cell', { name: 'Report' }).getByRole('link', { name: 'Report' }))
    .toHaveAttribute('href', 'https://example.test/report');
  await expect(document.getByRole('heading', { name: 'Sales by region' })).toBeAttached();
  await expect.poll(() => tableTree(page)).toEqual([
    ['table', 'Table, 3 rows, 3 columns'],
    ['row'], ['columnheader', 'Region'], ['columnheader', 'Total'],
    ['row'], ['rowheader', 'North'], ['cell', 'Q1'], ['cell', ''],
    ['row'], ['cell', 'Q2'], ['cell', 'Report'],
  ]);
});

test('pptx mirror tables without header flags read as data cells only', async ({ page }) => {
  await open(page, '');
  const grid = mirror(page).getByRole('table', { name: 'Table, 3 rows, 3 columns' });
  await expect(grid.getByRole('cell')).toHaveCount(7);
  await expect(grid.getByRole('columnheader')).toHaveCount(0);
  await expect(grid.getByRole('rowheader')).toHaveCount(0);
});
