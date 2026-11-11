import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

const SEED = '<w:p><w:r><w:t>Seed</w:t></w:r></w:p>';
const HEADING_STYLES = '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style></w:styles>';
const TABLE_BETWEEN = '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Before words</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Cell words</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>After words here</w:t></w:r></w:p>';

async function open(page: Page, arm: string, body = SEED, styles?: string) {
  const zip = new JSZip();
  const stylesType = styles ? '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' : '';
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${stylesType}</Types>`);
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  if (styles) {
    zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    zip.file('word/styles.xml', styles);
  }
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  await page.goto('/?format=docx' + arm);
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'clear-formatting.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await zip.generateAsync({ type: 'nodebuffer' }),
  });
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('formatting-bar')).toBeVisible();
  const input = page.locator('[data-testid=yrs-input], .paged-editor__input');
  await input.focus();
  return input;
}

async function savedRuns(page: Page, info: TestInfo, name: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate((source) => {
    const document = new DOMParser().parseFromString(source, 'application/xml');
    const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    return Array.from(document.getElementsByTagNameNS(w, 'r')).map((run) => ({
      text: Array.from(run.getElementsByTagNameNS(w, 't')).map((node) => node.textContent ?? '').join(''),
      bold: run.getElementsByTagNameNS(w, 'b').length > 0,
      properties: Array.from(run.getElementsByTagNameNS(w, 'rPr')).flatMap((properties) => Array.from(properties.children).map((child) => child.localName)),
      paragraphStyle: run.parentElement?.getElementsByTagNameNS(w, 'pStyle')[0]?.getAttributeNS(w, 'val') ?? null,
    }));
  }, xml);
}

for (const arm of ['&peerless=0', '']) {
  const label = arm ? 'hybrid' : 'default';

  test(`docx clear formatting shortcut removes saved bold formatting (${label})`, async ({ page }, info) => {
    const input = await open(page, arm);
    const text = 'Clear formatting shortcut';
    await input.press('ControlOrMeta+A');
    await input.pressSequentially(text);
    await input.press('ControlOrMeta+A');
    await input.press('ControlOrMeta+B');
    const bold = page.getByTestId('formatting-bar').getByRole('button', { name: 'Bold', exact: true });
    await expect(bold).toHaveAttribute('aria-pressed', 'true');
    const before = await savedRuns(page, info, `${label}-bold`);
    expect(before.map((run) => run.text).join('')).toBe(text);
    expect(before.some((run) => run.bold)).toBe(true);

    await input.focus();
    await input.press('ControlOrMeta+A');
    const mac = await page.evaluate(() => /Mac|iPod|iPhone|iPad/.test(navigator.platform));
    await input.press(mac ? 'Meta+Backslash' : 'Control+Space');
    await expect(bold).toHaveAttribute('aria-pressed', 'false');
    const after = await savedRuns(page, info, `${label}-cleared`);
    expect(after.map((run) => run.text).join('')).toBe(text);
    expect(after.every((run) => !run.bold)).toBe(true);
  });

  test(`docx toolbar clear formatting keeps a heading bold and removes its direct formatting (${label})`, async ({ page }, info) => {
    const input = await open(
      page,
      arm,
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:i/><w:color w:val="C00000"/></w:rPr><w:t>Heading words</w:t></w:r></w:p>',
      HEADING_STYLES,
    );
    const rail = page.getByTestId('formatting-bar');
    const bold = rail.getByRole('button', { name: 'Bold', exact: true });
    const italic = rail.getByRole('button', { name: 'Italic', exact: true });
    await input.press('ControlOrMeta+A');
    await expect(italic).toHaveAttribute('aria-pressed', 'true');
    await rail.getByRole('button', { name: 'Clear formatting', exact: true }).click();
    await expect(italic).toHaveAttribute('aria-pressed', 'false');
    await expect(bold).toHaveAttribute('aria-pressed', 'true');
    const runs = await savedRuns(page, info, `${label}-heading`);
    expect(runs.map((run) => run.text).join('')).toBe('Heading words');
    for (const run of runs.filter((entry) => entry.text)) {
      expect(run.paragraphStyle).toBe('Heading1');
      expect(run.properties).toEqual([]);
    }
  });

  for (const how of ['select all', 'drag'] as const) {
    test(`docx clear formatting over a ${how} that spans a table clears the cell text (${label})`, async ({ page }, info) => {
      const input = await open(page, arm, TABLE_BETWEEN);
      if (how === 'select all') {
        await input.press('ControlOrMeta+A');
      } else {
        const caret = page.locator('[data-docx-caret]');
        await input.press('ControlOrMeta+End');
        const end = await caret.boundingBox();
        await input.press('Home');
        const lineStart = await caret.boundingBox();
        await input.press('ControlOrMeta+Home');
        const start = await caret.boundingBox();
        if (!end || !lineStart || !start) throw new Error('Document caret has no bounds');
        await page.mouse.move(start.x + 0.25, start.y + start.height / 2);
        await page.mouse.down();
        await page.mouse.move((lineStart.x + end.x) / 2, end.y + end.height / 2, { steps: 12 });
        await page.mouse.up();
      }
      const rail = page.getByTestId('formatting-bar');
      const bold = rail.getByRole('button', { name: 'Bold', exact: true });
      await expect(bold).toHaveAttribute('aria-pressed', 'true');
      await rail.getByRole('button', { name: 'Clear formatting', exact: true }).click();
      await expect(bold).toHaveAttribute('aria-pressed', 'false');
      const runs = (await savedRuns(page, info, `${label}-${how}-table`)).filter((run) => run.text);
      expect(runs.find((run) => run.text.includes('Cell'))?.properties).toEqual([]);
      expect(runs.find((run) => run.text.includes('Before'))?.properties).toEqual([]);
      const after = runs.filter((run) => 'After words here'.includes(run.text) && run.text);
      expect(after.map((run) => run.text).join('')).toBe('After words here');
      expect(after.some((run) => run.bold)).toBe(how === 'drag');
      expect(after.some((run) => !run.bold)).toBe(true);
    });
  }
}
