import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const drawing = '<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
  + '<wp:inline><wp:extent cx="4572000" cy="2743200"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
  + '<c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing>';
const series = (index: number, name: string, values: number[]) =>
  `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:v>${name}</c:v></c:tx>`
  + '<c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>'
  + `<c:val><c:numRef><c:numCache><c:formatCode>0%</c:formatCode><c:ptCount val="2"/>${values.map((value, at) => `<c:pt idx="${at}"><c:v>${value}</c:v></c:pt>`).join('')}</c:numCache></c:numRef></c:val></c:ser>`;
const chart = '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>'
  + '<c:title><c:tx><c:rich><a:p><a:r><a:t>Share</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>'
  + series(0, 'North', [0.25, 0.5]) + series(1, 'South', [0.75, 0.5])
  + '<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>'
  + '<c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>';

async function docxFile() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="${r}/chart" Target="charts/chart1.xml"/></Relationships>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body><w:p><w:r><w:t>Quarterly share</w:t></w:r></w:p><w:p><w:r>${drawing}</w:r></w:p>`
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  zip.file('word/charts/chart1.xml', chart);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, worker: boolean) {
  await page.goto(`/?format=docx&peerless=${worker ? 1 : 0}`);
  await page.locator('input[type=file]').last().setInputFiles({ name: 'charts.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await docxFile() });
  await expect(page.locator('[data-page-index="0"] canvas').first()).toBeVisible({ timeout: 60000 });
  await expect(page.locator('.canvas-page-mirror').getByText('Quarterly share')).toBeAttached();
}

type AXNode = { nodeId: string; parentId?: string; ignored?: boolean; role?: { value?: unknown }; name?: { value?: unknown }; childIds?: string[] };
const tableRoles = new Set(['table', 'columnheader', 'rowheader', 'cell']);

/** The chart figure's accessibility subtree in tree order, as role and name. */
async function figureTree(page: Page) {
  const cdp = await page.context().newCDPSession(page);
  const { nodes } = await cdp.send('Accessibility.getFullAXTree') as { nodes: AXNode[] };
  await cdp.detach();
  const byId = new Map(nodes.map((node) => [node.nodeId, node]));
  const out: string[][] = [];
  const visit = (node: AXNode | undefined) => {
    if (!node) return;
    const role = String(node.role?.value);
    if (!node.ignored && (role === 'figure' || tableRoles.has(role))) out.push([role, String(node.name?.value ?? '').trim()]);
    for (const id of node.childIds ?? []) visit(byId.get(id));
  };
  visit(nodes.find((node) => !node.ignored && node.role?.value === 'figure'));
  return out;
}

for (const worker of [false, true]) {
  test(`docx mirror charts expose their values as a table inside a named figure (${worker ? 'worker' : 'in-thread'} editor)`, async ({ page }) => {
    await open(page, worker);
    const mirror = page.locator('.canvas-page-mirror');
    const figure = mirror.locator('[role="figure"]');
    await expect(figure).toHaveCount(1);
    await expect(figure.locator('[role="columnheader"]')).toHaveText(['North', 'South']);
    await expect(figure.locator('[role="rowheader"]')).toHaveText(['Q1', 'Q2']);
    await expect(figure.locator('[role="table"] [role="cell"]')).toHaveText(['', '25%', '75%', '50%', '50%']);
    await expect.poll(() => figureTree(page)).toEqual([
      ['figure', (await figure.getAttribute('aria-label')) ?? ''],
      ['table', ''],
      ['cell', ''], ['columnheader', 'North'], ['columnheader', 'South'],
      ['rowheader', 'Q1'], ['cell', '25%'], ['cell', '75%'],
      ['rowheader', 'Q2'], ['cell', '50%'], ['cell', '50%'],
    ]);
  });
}
