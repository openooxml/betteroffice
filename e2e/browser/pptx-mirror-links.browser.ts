import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';

const drawing = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const presentation = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const namespaces = `xmlns:p="${presentation}" xmlns:a="${drawing}" xmlns:r="${relationship}"`;
const emu = (value: number) => Math.round(value * 9525);
const tree = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';
const run = (text: string, link = '') => `<a:r><a:rPr sz="2400"><a:latin typeface="Liberation Sans"/>${link}</a:rPr><a:t>${text}</a:t></a:r>`;
const shape = (id: number, text: string, link = '', title = false) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr/><p:nvPr>${title ? '<p:ph type="title"/>' : ''}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(40)}" y="${emu(id * 60)}"/><a:ext cx="${emu(800)}" cy="${emu(50)}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${run(text, link)}</a:p></p:txBody></p:sp>`;
const slide = (body: string) => `<p:sld ${namespaces}><p:cSld><p:spTree>${tree}${body}</p:spTree></p:cSld></p:sld>`;

async function fixture() {
  const zip = new JSZip();
  const relationships = (items: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
  const part = (n: number) => `<Override PartName="/ppt/slides/slide${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${[1, 2, 3].map(part).join('')}</Types>`);
  zip.file('_rels/.rels', relationships(`<Relationship Id="rId1" Type="${relationship}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation ${namespaces}><p:sldIdLst>${[1, 2, 3].map((index) => `<p:sldId id="${255 + index}" r:id="rId${index}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${emu(960)}" cy="${emu(540)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', relationships([1, 2, 3].map((index) => `<Relationship Id="rId${index}" Type="${relationship}/slide" Target="slides/slide${index}.xml"/>`).join('')));
  zip.file('ppt/slides/_rels/slide1.xml.rels', relationships(
    `<Relationship Id="url" Type="${relationship}/hyperlink" Target="https://example.test/docs" TargetMode="External"/>` +
    `<Relationship Id="hostile" Type="${relationship}/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>` +
    `<Relationship Id="third" Type="${relationship}/slide" Target="slide3.xml"/>`));
  zip.file('ppt/slides/slide1.xml', slide(
    shape(2, 'Slide links', '', true) +
    shape(3, 'Read docs', '<a:hlinkClick r:id="url"/>') +
    shape(4, 'Unsafe target', '<a:hlinkClick r:id="hostile"/>') +
    shape(5, 'Third slide', '<a:hlinkClick r:id="third" action="ppaction://hlinksldjump"/>') +
    shape(6, 'Next slide', '<a:hlinkClick action="ppaction://hlinkshowjump?jump=nextslide"/>')));
  zip.file('ppt/slides/slide2.xml', slide(shape(2, 'Second slide', '', true)));
  zip.file('ppt/slides/slide3.xml', slide(shape(2, 'Third destination', '', true)));
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page) {
  await page.goto('/?format=pptx');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'links.pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', buffer: await fixture(),
  });
  await expect(page.getByTestId('pptx-worker-editor')).toBeVisible();
  await expect(page.getByTestId('pptx-worker-status')).toHaveCount(0);
  await expect(page.getByTestId('pptx-slide-canvas')).toHaveAttribute('data-pptx-input-ready', 'true', { timeout: 60_000 });
}

type AXNode = { ignored?: boolean; role?: { value?: string }; name?: { value?: string } };

const mirror = (page: Page) => page.getByTestId('pptx-slide-mirror');

test('pptx mirror links expose names and safe targets to assistive technology', async ({ page, context }) => {
  await open(page);
  const document = mirror(page);
  await expect(document.getByRole('link')).toHaveCount(3);
  await expect(document.getByRole('link', { name: 'Read docs' })).toHaveAttribute('href', 'https://example.test/docs');
  await expect(document.getByRole('link', { name: 'Third slide' })).toHaveAttribute('href', '#');
  await expect(document.getByRole('link', { name: 'Next slide' })).toHaveAttribute('href', '#');
  await expect(document.getByText('Unsafe target', { exact: true })).toBeAttached();
  await expect(document.getByRole('link', { name: 'Unsafe target' })).toHaveCount(0);
  const cdp = await context.newCDPSession(page);
  await expect.poll(async () => {
    const { nodes } = await cdp.send('Accessibility.getFullAXTree') as { nodes: AXNode[] };
    return nodes.filter((node) => !node.ignored && node.role?.value === 'link').map((node) => node.name?.value);
  }).toEqual(expect.arrayContaining(['Read docs', 'Third slide', 'Next slide']));
});

test('Enter opens an isolated external tab and refuses a replaced hostile href', async ({ page, context }) => {
  await open(page);
  await page.evaluate(() => {
    const target = window as Window & { mirrorOpens?: unknown[][] };
    target.mirrorOpens = [];
    const original = window.open.bind(window);
    window.open = (href, name, features) => {
      target.mirrorOpens!.push([href, name, features]);
      return original('about:blank', name, features);
    };
  });
  const anchor = mirror(page).getByRole('link', { name: 'Read docs' });
  await anchor.focus();
  const opened = context.waitForEvent('page');
  await anchor.press('Enter');
  const tab = await opened;
  await tab.waitForLoadState();
  expect(await tab.evaluate(() => window.opener === null)).toBe(true);
  expect(await page.evaluate(() => (window as Window & { mirrorOpens?: unknown[][] }).mirrorOpens))
    .toEqual([['https://example.test/docs', '_blank', 'noopener,noreferrer']]);
  await tab.close();
  await anchor.evaluate((node) => node.setAttribute('href', 'java\tscript:alert(1)'));
  await anchor.focus();
  await anchor.press('Enter');
  expect(await page.evaluate(() => (window as Window & { mirrorOpens?: unknown[][] }).mirrorOpens)).toHaveLength(1);
});

test('Enter follows relative and explicit slide navigation links', async ({ page }) => {
  await open(page);
  const next = mirror(page).getByRole('link', { name: 'Next slide' });
  await next.focus();
  await next.press('Enter');
  await expect(mirror(page)).toHaveAttribute('aria-label', 'Slide 2 of 3');
  await expect(mirror(page).getByRole('heading', { name: 'Second slide' })).toBeAttached();
  const stage = page.getByRole('application');
  await stage.focus();
  await stage.press('PageUp');
  await expect(mirror(page)).toHaveAttribute('aria-label', 'Slide 1 of 3');
  const third = mirror(page).getByRole('link', { name: 'Third slide' });
  await third.focus();
  await third.press('Enter');
  await expect(mirror(page)).toHaveAttribute('aria-label', 'Slide 3 of 3');
  await expect(mirror(page).getByRole('heading', { name: 'Third destination' })).toBeAttached();
});
