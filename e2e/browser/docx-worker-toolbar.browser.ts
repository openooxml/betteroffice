import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');
const text = 'Toolbar sample';

async function openWorker(page: Page) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Toolbar sample</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  await page.setViewportSize({ width: 2560, height: 1200 });
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'toolbar.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
  const input = page.locator('.paged-editor__peerless-input');
  await input.focus();
  await input.press('ControlOrMeta+Home');
  return input;
}

async function savedXml(page: Page, info: TestInfo, name: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  return zip.file('word/document.xml')!.async('string');
}

test('worker highlight saves w:highlight and clearing removes it', async ({ page }, info) => {
  const input = await openWorker(page);
  await input.press('Shift+End');
  const highlight = page.getByTestId('formatting-bar').getByRole('button', { name: 'Text Highlight Color', exact: true });
  await expect(highlight.first()).not.toHaveAttribute('aria-disabled', 'true');
  await highlight.first().click();
  expect(await savedXml(page, info, 'highlight')).toContain('<w:highlight w:val="yellow"');
  await highlight.last().click();
  await page.getByRole('button', { name: 'No Color', exact: true }).click();
  expect(await savedXml(page, info, 'cleared')).not.toContain('<w:highlight');
});

test('worker page and section breaks save their breaks', async ({ page }, info) => {
  const input = await openWorker(page);
  await input.press('ArrowRight');
  await input.press('ArrowRight');
  for (const name of ['Page break', 'Section break (next page)', 'Section break (continuous)']) {
    await page.getByRole('menubar').getByRole('menuitem', { name: 'Insert', exact: true }).click();
    await page.getByRole('menu', { name: 'Insert', exact: true }).getByRole('menuitem', { name: 'Break', exact: true }).hover();
    const command = page.getByRole('menuitem', { name, exact: true });
    await expect(command).not.toHaveAttribute('aria-disabled', 'true');
    await command.click();
  }
  const xml = await savedXml(page, info, 'breaks');
  expect(xml).toContain('<w:br w:type="page"');
  expect(xml.match(/<w:sectPr\b/g)).toHaveLength(3);
  expect(xml).toContain('<w:type w:val="nextPage"');
  expect(xml).toContain('<w:type w:val="continuous"');
});

async function savedProperties(page: Page, info: TestInfo, arm: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${arm}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate(({ xml, text }) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    const paragraphs = Array.from(document.getElementsByTagNameNS('*', 'p'));
    const paragraph = paragraphs.find(node => Array.from(node.getElementsByTagNameNS('*', 't'))
      .map(node => node.textContent ?? '').join('').includes(text));
    const value = (name: string) => paragraph?.getElementsByTagNameNS('*', name)[0]?.getAttributeNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'val') ?? null;
    const table = document.getElementsByTagNameNS('*', 'tbl')[0];
    return {
      strike: !!paragraph?.getElementsByTagNameNS('*', 'strike').length,
      alignment: value('jc'),
      list: !!paragraph?.getElementsByTagNameNS('*', 'numPr').length,
      style: value('pStyle'),
      table: table ? Array.from(table.getElementsByTagNameNS('*', 'tr')).map(row => row.getElementsByTagNameNS('*', 'tc').length) : null,
    };
  }, { xml, text });
}

test('docx toolbar saves equal formatting and table properties in both arms', async ({ page }, info) => {
  await page.setViewportSize({ width: 2560, height: 1200 });
  const results: Awaited<ReturnType<typeof savedProperties>>[] = [];
  for (const arm of ['&peerless=0', '']) {
    const label = arm ? 'hybrid' : 'default';
    await test.step(label, async () => {
      await page.goto('/?format=docx' + arm);
      await page.locator('input[type=file]').last().setInputFiles(resolve(root, 'apps/demo/public/betteroffice-demo.docx'));
      await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
      const rail = page.getByTestId('formatting-bar');
      await expect(rail).toBeVisible();
      await page.locator('.canvas-page canvas, .paged-editor__page canvas').first().click({ position: { x: 150, y: 150 } });
      const input = page.locator('[data-testid=yrs-input], .paged-editor__peerless-input');
      await expect(input).toBeEditable();
      await input.press('ControlOrMeta+Home');
      await input.pressSequentially(text);
      await input.press('Shift+Home');
      const strike = rail.getByRole('button', { name: 'Strikethrough', exact: true });
      await expect(strike).not.toHaveAttribute('aria-disabled', 'true');
      await strike.click();
      await expect(strike).toHaveAttribute('aria-pressed', 'true');
      await rail.getByTestId('toolbar-alignment').click();
      await page.getByTestId('alignment-center').click();
      const bullet = rail.getByRole('button', { name: 'Bullet List', exact: true });
      await bullet.click();
      await expect(bullet).toHaveAttribute('aria-pressed', 'true');
      await rail.getByRole('combobox', { name: 'Select paragraph style', exact: true }).click();
      await page.getByRole('option', { name: 'Heading 1', exact: true }).click();
      await input.press('ArrowRight');
      await page.getByRole('menubar').getByRole('menuitem', { name: 'Insert', exact: true }).click();
      await page.getByRole('menu', { name: 'Insert' }).getByRole('menuitem', { name: 'Table', exact: true }).hover();
      const grid = page.getByRole('grid', { name: 'Table size selector' });
      await expect(grid).toBeVisible();
      await grid.getByRole('gridcell').nth(7).click();
      const properties = await savedProperties(page, info, label);
      expect(properties).toEqual({ strike: true, alignment: 'center', list: true, style: 'Heading1', table: [2, 2] });
      results.push(properties);
    });
  }
  expect(results[1]).toEqual(results[0]);
});

async function openSpaced(page: Page, arm: string, source: Buffer | string) {
  await page.goto('/?format=docx' + arm);
  await page.locator('input[type=file]').last().setInputFiles(typeof source === 'string' ? source : {
    name: 'spaced.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: source,
  });
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('formatting-bar')).toBeVisible();
  await page.locator('.canvas-page canvas, .paged-editor__page canvas').first().click({ position: { x: 150, y: 150 } });
  const input = page.locator('[data-testid=yrs-input], .paged-editor__peerless-input');
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  return input;
}

test('docx toolbar line spacing and indent save over direct paragraph properties in both arms', async ({ page }, info) => {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('word/styles.xml', `<w:styles xmlns:w="${w}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${w}"><w:body>` +
    '<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:before="120" w:after="240" w:line="240" w:lineRule="auto"/><w:ind w:left="360" w:hanging="360"/></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Bravo</w:t></w:r></w:p>' +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  const source = await zip.generateAsync({ type: 'nodebuffer' });
  await page.setViewportSize({ width: 2560, height: 1200 });
  for (const arm of ['&peerless=0', '']) {
    const label = arm ? 'hybrid' : 'default';
    await test.step(label, async () => {
      const input = await openSpaced(page, arm, source);
      const rail = page.getByTestId('formatting-bar');
      const spacing = rail.getByRole('combobox', { name: 'Line spacing', exact: true });
      await expect(spacing).toHaveAttribute('title', 'Line spacing: Single');
      for (const step of ['first', 'second']) {
        if (step === 'second') {
          await input.press('ArrowDown');
          await expect(spacing).toHaveAttribute('title', 'Line spacing: Single');
        }
        await spacing.click();
        await page.getByRole('option', { name: '1.15', exact: true }).click();
        await expect(spacing).toHaveAttribute('title', 'Line spacing: 1.15');
        await rail.getByRole('button', { name: 'Increase Indent', exact: true }).click();
      }
      const downloading = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Save', exact: true }).first().click();
      const file = info.outputPath(`${label}-spacing.docx`);
      await (await downloading).saveAs(file);
      const xml = await (await JSZip.loadAsync(await readFile(file))).file('word/document.xml')!.async('string');
      const saved = await page.evaluate(({ xml, w }) => Array.from(new DOMParser().parseFromString(xml, 'application/xml')
        .getElementsByTagNameNS(w, 'p')).map((paragraph) => {
        const element = (name: string) => paragraph.getElementsByTagNameNS(w, name)[0];
        const attrs = (name: string) => Object.fromEntries(Array.from(element(name)?.attributes ?? [])
          .map((attribute) => [attribute.localName, attribute.value]));
        return { spacing: attrs('spacing'), ind: attrs('ind') };
      }), { xml, w });
      expect(saved).toEqual([
        { spacing: { before: '120', after: '240', line: '276', lineRule: 'auto' }, ind: { left: '1080', hanging: '360' } },
        { spacing: { line: '276', lineRule: 'auto' }, ind: { left: '720' } },
      ]);
      await openSpaced(page, arm, file);
      await expect(spacing).toHaveAttribute('title', 'Line spacing: 1.15');
    });
  }
});

test('worker pages show the text cursor over document text', async ({ page }) => {
  await openWorker(page);
  const surface = page.locator('.paged-editor__page').first();
  const box = (await surface.boundingBox())!;
  const x = box.x + box.width * 0.2;
  const y = box.y + box.height * 0.1;
  await page.mouse.move(x, y);
  expect(await page.evaluate(([px, py]) => {
    const target = document.elementFromPoint(px, py);
    return target ? getComputedStyle(target).cursor : null;
  }, [x, y] as const)).toBe('text');
});

test('typing keeps the worker ruler markers mounted', async ({ page }, info) => {
  const input = await openWorker(page);
  const ruler = page.locator('.docx-horizontal-ruler');
  await expect(ruler.getByLabel('Left indent')).toBeVisible();
  await ruler.evaluate((element) => {
    const churn = { count: 0 };
    (window as unknown as { __rulerChurn: typeof churn }).__rulerChurn = churn;
    new MutationObserver((records) => {
      churn.count += records.filter(record => record.type === 'childList' || record.attributeName === 'tabindex').length;
    }).observe(element, { subtree: true, childList: true, attributes: true });
  });
  await input.press('End');
  await input.pressSequentially('typed', { delay: 30 });
  expect(await savedXml(page, info, 'ruler-typing')).toContain('typed');
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  expect(await page.evaluate(() => (window as unknown as { __rulerChurn: { count: number } }).__rulerChurn.count)).toBe(0);
});
