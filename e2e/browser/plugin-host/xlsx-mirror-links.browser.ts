import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const safe = 'https://example.com/safe';
const HOSTILE = {
  js: 'javascript:window.__pwned=1;void 0',
  data: 'data:text/html,<script>window.__pwned=1</script>',
  vb: 'vbscript:msgbox(1)',
};
const escape = (value: string) => value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
const text = (ref: string, value: string) => `<c r="${ref}" t="inlineStr"><is><t>${escape(value)}</t></is></c>`;
const row = (n: number, cells: string) => `<row r="${n}">${cells}</row>`;

async function workbook() {
  const zip = new JSZip();
  const externals = Object.entries({ safe, ...HOSTILE });
  const part = (n: number) => `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${part(1)}${part(2)}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Links" sheetId="1" r:id="rId1"/><sheet name="Notes" sheetId="2" r:id="rId2"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`);
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${externals.map(([id, target]) =>
    `<Relationship Id="${id}" Type="${REL}/hyperlink" Target="${escape(target)}" TargetMode="External"/>`).join('')}</Relationships>`);
  const labels = ['Safe', ...Object.keys(HOSTILE).map((id) => `Rel ${id}`)];
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><sheetData>` +
    labels.map((label, index) => row(index + 1, text(`A${index + 1}`, label))).join('') +
    row(5, text('A5', 'Inner')) +
    row(6, `<c r="A6" t="str"><f>HYPERLINK("${escape(HOSTILE.js)}","Formula js")</f><v>Formula js</v></c>`) +
    '</sheetData><hyperlinks>' +
    externals.map(([id], index) => `<hyperlink ref="A${index + 1}" r:id="${id}"/>`).join('') +
    '<hyperlink ref="A5" location="Notes!B3" display="Inner"/></hyperlinks></worksheet>');
  zip.file('xl/worksheets/sheet2.xml', `<worksheet xmlns="${MAIN}"><sheetData>${row(3, text('B3', 'here'))}</sheetData></worksheet>`);
  return [...await zip.generateAsync({ type: 'uint8array' })];
}

const grid = (page: Page) => page.getByRole('grid').first();
const cell = (page: Page, label: RegExp) => grid(page).getByRole('gridcell', { name: label });

async function open(page: Page, surface: string) {
  await page.goto(`/xlsx-mirror-links.html?surface=${surface}`);
  await page.waitForFunction(() => '__mirrorLinks' in window);
  await page.evaluate((bytes) => (window as unknown as { __mirrorLinks: { load(bytes: number[]): void } }).__mirrorLinks.load(bytes), await workbook());
  await expect(cell(page, /^A1, Safe/)).toBeAttached({ timeout: 60_000 });
  await expect(cell(page, /^A6, Formula js/)).toBeAttached();
}

const hrefs = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll('[role="grid"] [href]'), (element) => element.getAttribute('href')!));
const opened = (page: Page) => page.evaluate(() => (window as unknown as { __mirrorLinks: { opened: unknown[][] } }).__mirrorLinks.opened);

async function activate(page: Page, label: RegExp) {
  const link = cell(page, label).getByRole('link');
  await expect(link).toHaveCount(1);
  await link.focus();
  await expect(link).toBeFocused();
  await page.keyboard.press('Enter');
}

for (const surface of ['peerless', 'in-thread', 'legacy-worker', 'viewer']) {
  test(`${surface}: hyperlink cells are links in the screen-reader grid and hostile targets are not`, async ({ page }) => {
    await open(page, surface);
    for (const href of await hrefs(page)) expect(href).toMatch(/^(https?:|mailto:|tel:|#)/);
    await expect(cell(page, /^A1, Safe/).getByRole('link')).toHaveAttribute('href', safe);
    await expect(cell(page, /^A5, Inner/).getByRole('link')).toHaveAttribute('href', '#Notes!B3');
    for (const label of [/^A2, Rel js/, /^A3, Rel data/, /^A4, Rel vb/, /^A6, Formula js/]) {
      await expect(cell(page, label).getByRole('link')).toHaveCount(0);
    }
  });

  test(`${surface}: following a link from the grid opens it without an opener`, async ({ page }) => {
    await open(page, surface);
    const start = page.url();
    await activate(page, /^A1, Safe/);
    await expect.poll(() => opened(page)).toEqual([[safe, '_blank', 'noopener,noreferrer']]);
    expect(page.url()).toBe(start);
    expect(await page.evaluate(() => (window as unknown as { __pwned?: number }).__pwned ?? 0)).toBe(0);
  });

  test(`${surface}: following a link to a place in the workbook selects it`, async ({ page }) => {
    await open(page, surface);
    const start = page.url();
    await activate(page, /^A5, Inner/);
    await expect(cell(page, /^B3, here, selected/)).toBeAttached();
    expect(page.url()).toBe(start);
    expect(await opened(page)).toEqual([]);
  });
}
