import { test, expect, type CDPSession, type Page } from 'playwright/test';
import JSZip from 'jszip';

const drawing = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const presentation = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const namespaces = `xmlns:p="${presentation}" xmlns:a="${drawing}" xmlns:r="${relationship}"`;
const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const emu = (value: number) => Math.round(value * 9525);
const frame = (x: number, y: number, w: number, h: number) => `<a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/>`;
const run = (text: string, lang?: string) => `<a:r><a:rPr sz="2400"${lang ? ` lang="${lang}"` : ''}><a:latin typeface="Liberation Sans"/></a:rPr><a:t>${text}</a:t></a:r>`;
const textShape = (id: number, name: string, placeholder: string, y: number, runs: string[][], languages?: string[][]) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${placeholder}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm>${frame(40, y, 800, 100)}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${runs.map((parts, paragraph) => `<a:p>${parts.map((text, index) => run(text, languages?.[paragraph]?.[index])).join('')}</a:p>`).join('')}</p:txBody></p:sp>`;
const picture = (id: number, name: string, descr: string, x: number, decorative = false) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}" descr="${descr}">${decorative ? '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>' : ''}</p:cNvPr><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm>${frame(x, 300, 100, 100)}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
const cell = (text: string) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${run(text)}</a:p></a:txBody><a:tcPr/></a:tc>`;
const table = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm>${frame(40, 300, 500, 100)}</p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="${emu(250)}"/><a:gridCol w="${emu(250)}"/></a:tblGrid><a:tr h="${emu(50)}">${cell('North region')}${cell('Twelve percent')}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
const tree = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';
const slide = (body: string) => `<p:sld ${namespaces}><p:cSld><p:spTree>${tree}${body}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

async function fixture(languages = false) {
  const zip = new JSZip();
  const relationships = (items: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
  const part = (n: number) => `<Override PartName="/ppt/slides/slide${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${part(1)}${part(2)}</Types>`);
  zip.file('_rels/.rels', relationships(`<Relationship Id="rId1" Type="${relationship}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation ${namespaces}><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="${emu(960)}" cy="${emu(540)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', relationships(`<Relationship Id="rId1" Type="${relationship}/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="${relationship}/slide" Target="slides/slide2.xml"/>`));
  zip.file('ppt/media/image1.png', png, { base64: true });
  const media = relationships(`<Relationship Id="rId2" Type="${relationship}/image" Target="../media/image1.png"/>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', media);
  zip.file('ppt/slides/_rels/slide2.xml.rels', media);
  zip.file('ppt/slides/slide1.xml', slide(
    textShape(2, 'Title 1', '<p:ph type="title"/>', 40, [['Quarterly results']], languages ? [['en-US']] : undefined) +
    textShape(3, 'Content 2', '', 160, [['Revenue ', 'grew ', 'twelve percent'], ['Costs ', 'stayed ', 'flat']],
      languages ? [['en-US', 'fr-FR', 'en-US'], ['en-US', 'en-US', 'en-US']] : undefined) +
    picture(4, 'Picture 3', 'A rising bar chart', 580) + picture(5, 'Picture 4', 'Should stay silent', 720, true) + table));
  zip.file('ppt/slides/slide2.xml', slide(textShape(2, 'Title 1', '<p:ph type="title"/>', 40, [['Next steps']])));
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, languages = false) {
  await page.goto('/?format=pptx');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'reading.pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', buffer: await fixture(languages),
  });
  await expect(page.getByTestId('pptx-worker-editor')).toBeVisible();
  await expect(page.getByTestId('pptx-worker-status')).toHaveCount(0);
  await expect(page.getByTestId('pptx-slide-canvas')).toHaveAttribute('data-pptx-input-ready', 'true', { timeout: 60_000 });
}

const mirror = (page: Page) => page.getByTestId('pptx-slide-mirror');
const live = (page: Page) => page.getByTestId('pptx-live-region');
const heard = (page: Page, text: string) => expect(live(page)).toHaveText(new RegExp(`^${text}\\s?$`));

type AXNode = { nodeId: string; ignored?: boolean; role?: { value?: string }; name?: { value?: string }; childIds?: string[] };

async function accessibleNames(cdp: CDPSession) {
  const { nodes } = await cdp.send('Accessibility.getFullAXTree') as { nodes: AXNode[] };
  return nodes.filter((node) => !node.ignored).map((node) => String(node.name?.value ?? ''));
}

async function accessibleSlide(cdp: CDPSession, label: string) {
  const { nodes } = await cdp.send('Accessibility.getFullAXTree') as { nodes: AXNode[] };
  const byId = new Map(nodes.map((node) => [node.nodeId, node]));
  const root = nodes.find((node) => !node.ignored && node.name?.value === label && node.role?.value === 'document');
  const found: { role: string; name: string }[] = [];
  const visit = (node: AXNode | undefined) => {
    if (!node) return;
    if (!node.ignored) found.push({ role: String(node.role?.value ?? ''), name: String(node.name?.value ?? '') });
    node.childIds?.forEach((id) => visit(byId.get(id)));
  };
  visit(root);
  return found;
}

test('pptx screen reader: the accessibility tree reads the slide in whole words', async ({ page, context }, info) => {
  await open(page);
  const cdp = await context.newCDPSession(page);
  await expect.poll(() => accessibleNames(cdp), { timeout: 30_000 }).toContain('Revenue grew twelve percent');
  await info.attach('slide-1.yml', { body: await mirror(page).ariaSnapshot(), contentType: 'text/yaml' });
  const nodes = await accessibleSlide(cdp, 'Slide 1 of 2');
  const text = nodes.filter((node) => node.role === 'StaticText').map((node) => node.name);
  expect(text).toEqual(expect.arrayContaining(['Quarterly results', 'Revenue grew twelve percent', 'Costs stayed flat', 'North region', 'Twelve percent']));
  expect(text.filter((name) => name.trim().length === 1)).toEqual([]);
  expect(nodes).toContainEqual({ role: 'heading', name: 'Quarterly results' });
  expect(nodes).toContainEqual({ role: 'image', name: 'A rising bar chart' });
  expect(nodes.map((node) => node.name)).not.toContain('Should stay silent');
  await expect(page.getByRole('application')).toHaveCount(1);
});

test('pptx screen reader: authored languages reach the DOM and one language stays one StaticText', async ({ page, context }) => {
  await open(page, true);
  await expect(mirror(page).getByRole('heading', { name: 'Quarterly results' })).toHaveAttribute('lang', 'en-US');
  const mixed = mirror(page).locator('p').filter({ hasText: /^Revenue grew twelve percent$/ });
  await expect(mixed).toHaveAttribute('lang', 'en-US');
  await expect(mixed.locator('span')).toHaveCount(1);
  await expect(mixed.locator('span')).toHaveAttribute('lang', 'fr-FR');
  await expect(mixed.locator('span')).toHaveText('grew');
  const single = mirror(page).locator('p').filter({ hasText: /^Costs stayed flat$/ });
  await expect(single).toHaveAttribute('lang', 'en-US');
  expect(await single.evaluate((node) => [...node.childNodes].map((child) => child.nodeType))).toEqual([3]);
  const unknown = mirror(page).locator('p').filter({ hasText: /^North region$/ });
  expect(await unknown.getAttribute('lang')).toBeNull();
  const cdp = await context.newCDPSession(page);
  await expect.poll(async () => (await accessibleSlide(cdp, 'Slide 1 of 2'))
    .filter((node) => node.role === 'StaticText' && node.name === 'Costs stayed flat')).toHaveLength(1);
});

test('pptx screen reader: slide changes and shape selection are announced', async ({ page }) => {
  await open(page);
  const stage = page.getByRole('application');
  await stage.focus();
  await stage.press('PageDown');
  await heard(page, 'Slide 2 of 2: Next steps');
  await expect(mirror(page)).toHaveAttribute('aria-label', 'Slide 2 of 2');
  await expect(mirror(page).getByRole('heading', { name: 'Next steps' })).toBeAttached();
  await stage.press('PageUp');
  await heard(page, 'Slide 1 of 2: Quarterly results');
  await expect(page.locator('button[data-slide-id]').first()).toHaveAttribute('aria-label', 'Slide 1 of 2: Quarterly results');
  const bounds = await page.getByTestId('pptx-slide-canvas').boundingBox();
  if (!bounds) throw new Error('Slide canvas is missing');
  await page.mouse.click(bounds.x + 630 * bounds.width / 960, bounds.y + 350 * bounds.height / 540);
  await heard(page, 'A rising bar chart selected');
  await page.mouse.click(bounds.x + 770 * bounds.width / 960, bounds.y + 350 * bounds.height / 540);
  await heard(page, 'Picture 4 selected');
});

test('pptx screen reader: the mirror is a visible, keyboard reachable stop that leaves the slide unchanged', async ({ page }) => {
  await open(page);
  const stage = page.getByRole('application');
  await stage.focus();
  await page.keyboard.press('Tab');
  await expect(mirror(page)).toBeFocused();
  await expect(mirror(page)).toHaveCSS('outline-style', 'solid');
  await page.keyboard.press('PageDown');
  await heard(page, 'Slide 2 of 2: Next steps');
  await expect(mirror(page)).toBeFocused();
  const saved = page.waitForEvent('download');
  await page.keyboard.press('ControlOrMeta+s');
  await saved;
  await page.keyboard.press('Tab');
  await expect(page.getByRole('textbox', { name: 'Slide text', exact: true })).toBeFocused();
});
