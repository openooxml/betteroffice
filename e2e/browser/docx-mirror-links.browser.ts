import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const safe = 'https://example.com/safe';
const HOSTILE = {
  js: 'javascript:window.__pwned=1;void 0',
  data: 'data:text/html,<script>window.__pwned=1</script>',
  vb: 'vbscript:msgbox(1)',
  tab: 'java\tscript:window.__pwned=1',
};
const escape = (value: string) => value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;')
  .replaceAll('\t', '&#9;').replaceAll('\u0001', '');
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const complexField = (target: string, text: string) => `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "${escape(target)}" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run(text)}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
const simpleField = (target: string, text: string) => `<w:p><w:fldSimple w:instr="${escape(` HYPERLINK "${target}" `)}">${run(text)}</w:fldSimple></w:p>`;

async function docxFile() {
  const zip = new JSZip();
  const rels = Object.entries({ safe, ...HOSTILE }).map(([id, target]) =>
    `<Relationship Id="${id}" Type="${r}/hyperlink" Target="${escape(target)}" TargetMode="External"/>`).join('');
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body>` +
    '<w:p><w:r><w:t>Bare text</w:t></w:r></w:p>' +
    `<w:p><w:hyperlink r:id="safe">${run('Safe link')}</w:hyperlink></w:p>` +
    Object.keys(HOSTILE).map(id => `<w:p><w:hyperlink r:id="${id}">${run(`Rel ${id}`)}</w:hyperlink></w:p>`).join('') +
    Object.entries(HOSTILE).map(([id, target]) => complexField(target, `Field ${id}`) + simpleField(target, `Simple ${id}`)).join('') +
    `<w:p><w:hyperlink w:anchor="target">${run('Jump to target')}</w:hyperlink></w:p>` +
    `<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:bookmarkStart w:id="1" w:name="target"/>${run('Destination')}<w:bookmarkEnd w:id="1"/></w:p>` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, worker: boolean) {
  await page.addInitScript(() => {
    const state = window as typeof window & { opened?: unknown[][] };
    state.opened = [];
    window.open = (...args) => { state.opened!.push(args); return null; };
  });
  await page.goto(`/?format=docx&peerless=${worker ? 1 : 0}`);
  await page.locator('input[type=file]').last().setInputFiles({ name: 'mirror-links.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await docxFile() });
  await expect(page.locator('[data-page-index="0"] canvas').first()).toBeVisible({ timeout: 60000 });
  await expect(page.locator('.canvas-page-mirror').getByText('Jump to target')).toBeAttached();
  const input = page.locator('textarea.paged-editor__input, textarea[data-testid="yrs-input"]').first();
  await input.focus();
  return input;
}

const anchorFor = (page: Page, text: string) =>
  page.locator('.canvas-page-mirror .layout-paragraph').filter({ hasText: text }).first().locator('a[href]').first();
const opened = (page: Page) => page.evaluate(() => (window as typeof window & { opened?: unknown[][] }).opened);
const pwned = (page: Page) => page.evaluate(() => (window as typeof window & { __pwned?: number }).__pwned ?? 0);

async function assertNoHostile(page: Page) {
  const resolved = await page.evaluate(() =>
    Array.from(document.querySelectorAll('.canvas-page-mirror a[href], .canvas-page-mirror [data-href]'), element =>
      [element.getAttribute('href'), element.getAttribute('data-href')].filter((value): value is string => value !== null)
        .map(value => ({ text: element.textContent ?? '', value, protocol: new URL(value, location.href).protocol }))).flat());
  for (const link of resolved) expect(['http:', 'https:', 'mailto:', 'tel:', 'ftp:'], `${link.text}: ${link.value}`).toContain(link.protocol);
}

async function activate(page: Page, text: string) {
  const anchor = anchorFor(page, text);
  await expect(anchor).toBeAttached();
  await anchor.evaluate(node => node.scrollIntoView({ block: 'center' }));
  await anchor.focus();
  await expect(anchor).toBeFocused();
  await page.keyboard.press('Enter');
}

async function savedXml(page: Page, info: TestInfo) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('mirror-links.docx');
  await (await downloading).saveAs(file);
  return (await JSZip.loadAsync(await readFile(file))).file('word/document.xml')!.async('string');
}

for (const worker of [true, false]) {
  const mode = worker ? 'worker' : 'legacy';

  test(`${mode} mirror exposes no hostile hyperlink target from relationships or HYPERLINK fields`, async ({ page }) => {
    await open(page, worker);
    await assertNoHostile(page);
    for (const id of Object.keys(HOSTILE)) {
      for (const text of [`Rel ${id}`, `Field ${id}`, `Simple ${id}`]) {
        await expect(page.locator('.canvas-page-mirror').getByText(text)).toBeAttached();
        await expect(anchorFor(page, text)).toHaveCount(0);
      }
    }
    await expect(anchorFor(page, 'Safe link')).toHaveAttribute('href', safe);
  });

  test(`${mode} mirror link activation opens a safe target without an opener and keeps the page`, async ({ page }) => {
    await open(page, worker);
    const start = page.url();
    await activate(page, 'Safe link');
    await expect.poll(() => opened(page)).toEqual([[safe, '_blank', 'noopener,noreferrer']]);
    expect(page.url()).toBe(start);
    expect(await pwned(page)).toBe(0);
  });

  for (const caret of [false, true]) test(`${mode} mirror bookmark link moves the caret to its bookmark${caret ? ' from a placed caret' : ''}`, async ({ page }, info) => {
    const input = await open(page, worker);
    if (caret) await input.press('ControlOrMeta+Home');
    const start = page.url();
    const scroller = page.locator('.docx-editor__scroll-container').first();
    const before = await scroller.evaluate(node => node.scrollTop);
    await activate(page, 'Jump to target');
    await expect(input).toBeFocused();
    await expect(page.locator('.canvas-page-mirror .layout-paragraph').filter({ hasText: 'Destination' }).first().locator('.layout-run').first()).toBeInViewport({ timeout: 10000 });
    info.annotations.push({ type: 'scroll', description: `${before} -> ${await scroller.evaluate(node => node.scrollTop)}` });
    expect(page.url()).toBe(start);
    expect(await opened(page)).toEqual([]);
    await input.pressSequentially('Arrived');
    const xml = await savedXml(page, info);
    const paragraph = await page.evaluate(xml => Array.from(new DOMParser().parseFromString(xml, 'application/xml')
      .getElementsByTagNameNS('*', 'p')).map(p => p.textContent ?? '').find(text => text.includes('Destination')), xml);
    expect(paragraph).toContain('Arrived');
  });

  test(`${mode} a hostile URL set from the link dialog never becomes a followable mirror link`, async ({ page }) => {
    const input = await open(page, worker);
    await input.press('ControlOrMeta+Home');
    await input.press('Shift+End');
    await input.press('ControlOrMeta+k');
    const dialog = page.getByRole('dialog', { name: 'Insert Hyperlink', exact: true });
    await expect(dialog).toBeVisible();
    await dialog.getByLabel('URL', { exact: true }).fill(HOSTILE.js);
    const insert = dialog.getByRole('button', { name: 'Insert', exact: true });
    if (await insert.isEnabled()) await insert.click();
    const accepted = await dialog.waitFor({ state: 'hidden', timeout: 3000 }).then(() => true, () => false);
    test.info().annotations.push({ type: 'dialog', description: accepted ? 'accepted the hostile URL' : 'refused the hostile URL' });
    if (!accepted) await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
    await expect(dialog).toHaveCount(0);
    await expect(page.locator('.canvas-page-mirror').getByText('Bare text')).toBeAttached();
    await assertNoHostile(page);
    await expect(anchorFor(page, 'Bare text')).toHaveCount(0);
    expect(await pwned(page)).toBe(0);
  });
}
