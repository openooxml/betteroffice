import { expect, test, type Locator, type Page } from 'playwright/test';
import JSZip from 'jszip';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const LINK = 'https://example.test/copied';
const run = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (...runs: string[]) => `<w:p>${runs.join('')}</w:p>`;

const SOURCE_BODY = [
  paragraph(
    run('BoldX', '<w:b/>'), run('ItalX', '<w:i/>'), run('UndrX', '<w:u w:val="single"/>'),
    run('FontX', '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="36"/>'),
    run('ColrX', '<w:color w:val="FF0000"/>'),
    `<w:hyperlink r:id="rIdLink">${run('LinkX')}</w:hyperlink>`,
  ),
  paragraph(run('AcrossOne', '<w:b/>')),
  paragraph(run('AcrossTwo', '<w:i/>')),
  paragraph(run('ThemeX', '<w:rStyle w:val="Fancy"/><w:color w:val="4472C4" w:themeColor="accent1"/>')),
  paragraph(run('PlainX')),
  paragraph(run('BoldTarget', '<w:b/>')),
  `<w:tbl><w:tblPr><w:tblW w:w="4680" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid><w:tr><w:tc>${paragraph(run('Cell'))}</w:tc></w:tr></w:tbl>`,
  paragraph(run('End')),
].join('');
const OTHER_BODY = paragraph(run('Other'));
const styles = (fancy: boolean) => `<w:styles ${W}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${
  fancy ? '<w:style w:type="character" w:customStyle="1" w:styleId="Fancy"><w:name w:val="Fancy"/><w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>' : ''}</w:styles>`;
const theme = (accent: string) => `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="${accent}"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="T"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="T"><a:fillStyleLst><a:noFill/><a:noFill/><a:noFill/></a:fillStyleLst><a:lnStyleLst><a:ln/><a:ln/><a:ln/></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:noFill/><a:noFill/><a:noFill/></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

async function docx(body: string, options: { fancy: boolean; accent: string }): Promise<number[]> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${LINK}" TargetMode="External"/></Relationships>`);
  zip.file('word/styles.xml', styles(options.fancy));
  zip.file('word/theme/theme1.xml', theme(options.accent));
  zip.file('word/document.xml', `<w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  return [...await zip.generateAsync({ type: 'uint8array' })];
}
const sourceDocx = () => docx(SOURCE_BODY, { fancy: true, accent: '112233' });

async function load(page: Page, bytes: number[]): Promise<Locator> {
  await page.evaluate(data => (window as unknown as { __copyFormat: { load(data: number[]): void } }).__copyFormat.load(data), bytes);
  await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
  const input = page.locator('.paged-editor__input, [data-testid=yrs-input]').first();
  await input.focus();
  return input;
}

async function open(page: Page, mode: string, query = ''): Promise<Locator> {
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto(`/docx-copy-format.html?mode=${mode}${query}`);
  await expect.poll(() => page.evaluate(() => 'load' in ((window as unknown as { __copyFormat?: object }).__copyFormat ?? {}))).toBe(true);
  return load(page, await sourceDocx());
}

interface Saved { body: string; rels: string; }
async function saved(page: Page): Promise<Saved> {
  const bytes = await page.evaluate(() => (window as unknown as { __copyFormat: { saved(): Promise<number[]> } }).__copyFormat.saved());
  const zip = await JSZip.loadAsync(Uint8Array.from(bytes));
  const xml = await zip.file('word/document.xml')!.async('string');
  return { body: xml.slice(xml.indexOf('<w:body>'), xml.lastIndexOf('<w:sectPr')), rels: await zip.file('word/_rels/document.xml.rels')!.async('string') };
}

function runsWith(body: string, text: string): string[] {
  return (body.match(/<w:r[ >](?:(?!<\/w:r>).)*<\/w:r>/gs) ?? []).filter(entry => entry.includes(text));
}
const lastRun = (body: string, text: string): string => runsWith(body, text).at(-1) ?? '';
const count = (body: string, text: string): number => body.split(text).length - 1;

async function key(input: Locator, combo: string, times = 1): Promise<void> {
  for (let index = 0; index < times; index += 1) await input.press(combo);
}
async function selectFromStart(input: Locator, down: number, skip: number, length: number): Promise<void> {
  await key(input, 'ControlOrMeta+Home');
  await key(input, 'ArrowDown', down);
  await key(input, 'Home');
  await key(input, 'ArrowRight', skip);
  await key(input, 'Shift+ArrowRight', length);
}
async function copy(page: Page, input: Locator, combo = 'ControlOrMeta+c'): Promise<void> {
  await page.evaluate(() => navigator.clipboard.writeText('stale'));
  await input.press(combo);
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).not.toBe('stale');
}
async function pasteAtEnd(input: Locator): Promise<void> {
  await key(input, 'ControlOrMeta+End');
  await input.press('ControlOrMeta+v');
}
async function pasted(page: Page, text: string, copies: number): Promise<Saved> {
  await expect.poll(async () => count((await saved(page)).body, text)).toBe(copies);
  return saved(page);
}

for (const mode of ['worker', 'legacy']) {
  test.describe(mode, () => {
    test('bold, italic, underline, font, size, colour and hyperlink survive', async ({ page }) => {
      const input = await open(page, mode);
      await selectFromStart(input, 0, 0, 30);
      await copy(page, input);
      await pasteAtEnd(input);
      const { body, rels } = await pasted(page, 'LinkX', 2);
      expect(lastRun(body, 'BoldX')).toContain('<w:b/>');
      expect(lastRun(body, 'ItalX')).toContain('<w:i/>');
      expect(lastRun(body, 'UndrX')).toMatch(/<w:u [^>]*w:val="single"/);
      expect(lastRun(body, 'FontX')).toContain('Courier New');
      expect(lastRun(body, 'FontX')).toMatch(/<w:sz [^>]*w:val="36"/);
      expect(lastRun(body, 'ColrX')).toMatch(/<w:color [^>]*w:val="FF0000"/);
      const link = /<w:hyperlink [^>]*r:id="([^"]+)"(?:(?!<\/w:hyperlink>).)*LinkX/gs;
      const ids = [...body.matchAll(link)].map(match => match[1]!);
      expect(ids).toHaveLength(2);
      expect(rels).toContain(`Id="${ids[1]}"`);
      expect(rels).toContain(LINK);
      expect(lastRun(body, 'BoldX')).not.toContain('<w:i/>');
    });

    test('a selection over part of two runs keeps each part', async ({ page }) => {
      const input = await open(page, mode);
      await selectFromStart(input, 0, 2, 5);
      await copy(page, input);
      expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('ldXIt');
      await pasteAtEnd(input);
      const { body } = await pasted(page, '>It<', 1);
      expect(lastRun(body, 'ldX')).toContain('<w:b/>');
      expect(lastRun(body, 'ldX')).not.toContain('<w:i/>');
      expect(lastRun(body, '>It<')).toContain('<w:i/>');
      expect(lastRun(body, '>It<')).not.toContain('<w:b/>');
    });

    test('a selection across paragraphs keeps paragraphs and marks', async ({ page }) => {
      const input = await open(page, mode);
      await selectFromStart(input, 1, 0, 19);
      await copy(page, input);
      await pasteAtEnd(input);
      const { body } = await pasted(page, 'AcrossTwo', 2);
      expect(lastRun(body, 'AcrossOne')).toContain('<w:b/>');
      expect(lastRun(body, 'AcrossTwo')).toContain('<w:i/>');
      expect(body.slice(body.lastIndexOf('AcrossOne'), body.lastIndexOf('AcrossTwo'))).toContain('</w:p>');
    });

    test('pastes formatted into a table cell', async ({ page }) => {
      const input = await open(page, mode);
      await selectFromStart(input, 0, 0, 5);
      await copy(page, input);
      await key(input, 'ControlOrMeta+Home');
      await key(input, 'ArrowDown', 6);
      await key(input, 'End');
      await input.press('ControlOrMeta+v');
      const { body } = await pasted(page, 'BoldX', 2);
      const cell = /<w:tc>.*<\/w:tc>/s.exec(body)?.[0] ?? '';
      expect(lastRun(cell, 'BoldX')).toContain('<w:b/>');
      expect(lastRun(cell, 'Cell')).not.toContain('<w:b/>');
    });

    test('plain source text stays plain inside a bold run', async ({ page }) => {
      const input = await open(page, mode);
      await selectFromStart(input, 4, 0, 6);
      await copy(page, input);
      await key(input, 'ControlOrMeta+Home');
      await key(input, 'ArrowDown', 5);
      await key(input, 'Home');
      await key(input, 'ArrowRight', 4);
      await input.press('ControlOrMeta+v');
      const { body } = await pasted(page, 'PlainX', 2);
      expect(lastRun(body, 'PlainX')).not.toContain('<w:b/>');
      expect(lastRun(body, 'Bold')).toContain('<w:b/>');
    });

    test('one undo removes the whole paste', async ({ page }) => {
      const input = await open(page, mode);
      await selectFromStart(input, 1, 0, 19);
      await copy(page, input);
      await pasteAtEnd(input);
      await pasted(page, 'AcrossTwo', 2);
      await input.press('ControlOrMeta+z');
      const { body } = await pasted(page, 'AcrossTwo', 1);
      expect(count(body, 'AcrossOne')).toBe(1);
      expect(lastRun(body, 'End')).not.toContain('<w:b/>');
    });

    test('cut moves the text with its formatting', async ({ page }) => {
      const input = await open(page, mode);
      await selectFromStart(input, 0, 0, 5);
      await copy(page, input, 'ControlOrMeta+x');
      await expect.poll(async () => count((await saved(page)).body, 'BoldX')).toBe(0);
      await pasteAtEnd(input);
      const { body } = await pasted(page, 'BoldX', 1);
      expect(lastRun(body, 'BoldX')).toContain('<w:b/>');
      expect(body.indexOf('BoldX')).toBeGreaterThan(body.indexOf('Cell'));
    });

    test('the context menu copies, cuts and pastes with formatting', async ({ page }) => {
      const input = await open(page, mode);
      const menu = page.getByRole('menu', { name: 'Text editing menu' });
      const pick = async (name: RegExp) => {
        await page.locator('canvas[data-page-index="0"]').click({ button: 'right', position: { x: 130, y: 104 } });
        await menu.getByRole('menuitem', { name }).click();
      };
      await selectFromStart(input, 0, 0, 5);
      await page.evaluate(() => navigator.clipboard.writeText('stale'));
      await pick(/^Copy/);
      await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('BoldX');
      await input.focus();
      await key(input, 'ControlOrMeta+End');
      await input.press('ControlOrMeta+v');
      const { body } = await pasted(page, 'BoldX', 2);
      expect(lastRun(body, 'BoldX')).toContain('<w:b/>');
    });

    test('rich copies still read as plain text, and paste as plain drops the marks', async ({ page }) => {
      const input = await open(page, mode);
      await selectFromStart(input, 1, 0, 19);
      await copy(page, input);
      expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('AcrossOne\nAcrossTwo');
      const textarea = await page.evaluate(async () => {
        const area = document.createElement('textarea');
        document.body.append(area);
        area.focus();
        document.execCommand('paste');
        const value = area.value;
        area.remove();
        return value;
      });
      expect(textarea === '' || textarea === 'AcrossOne\nAcrossTwo').toBe(true);
      await input.focus();
      await key(input, 'ControlOrMeta+End');
      await page.locator('canvas[data-page-index="0"]').click({ button: 'right', position: { x: 300, y: 400 } });
      await page.getByRole('menu', { name: 'Text editing menu' }).getByRole('menuitem', { name: /^Paste as Plain Text/i }).click();
      const { body } = await pasted(page, 'AcrossTwo', 2);
      expect(lastRun(body, 'AcrossOne')).not.toContain('<w:b/>');
      expect(lastRun(body, 'AcrossTwo')).not.toContain('<w:i/>');
      await input.focus();
      await key(input, 'ControlOrMeta+End');
      await input.press('ControlOrMeta+Shift+V');
      const again = (await pasted(page, 'AcrossTwo', 3)).body;
      expect(lastRun(again, 'AcrossOne')).not.toContain('<w:b/>');
      expect(lastRun(again, 'AcrossTwo')).not.toContain('<w:i/>');
    });

    test('external html pastes as plain text, as before', async ({ page }) => {
      const input = await open(page, mode);
      await page.evaluate(() => navigator.clipboard.write([new ClipboardItem({
        'text/plain': new Blob(['ExternalX'], { type: 'text/plain' }),
        'text/html': new Blob(['<p><b>ExternalX</b></p>'], { type: 'text/html' }),
      })]));
      await pasteAtEnd(input);
      const { body } = await pasted(page, 'ExternalX', 1);
      expect(lastRun(body, 'ExternalX')).not.toContain('<w:b/>');
    });

    test('a crafted clipboard payload falls back to its plain text', async ({ page }) => {
      const input = await open(page, mode);
      await page.evaluate(() => {
        const run = { text: 'HostileX', format: { other: { runStyle: { styleId: 'Fancy' }, bold: true, fontFamily: { ascii: '"/><w:evil/>' } } }, href: 'javascript:alert(1)' };
        const payload = JSON.stringify({ v: 1, fragment: { paragraphs: [{ text: 'HostileX', runs: [run] }] } }).replaceAll('&', '&amp;').replaceAll('"', '&quot;');
        return navigator.clipboard.write([new ClipboardItem({
          'text/plain': new Blob(['HostileX'], { type: 'text/plain' }),
          'text/html': new Blob([`<div data-betteroffice-docx="${payload}"><p>HostileX</p></div>`], { type: 'text/html' }),
        })]);
      });
      await pasteAtEnd(input);
      const { body, rels } = await pasted(page, 'HostileX', 1);
      expect(lastRun(body, 'HostileX')).not.toContain('<w:b/>');
      expect(body).not.toContain('w:evil');
      expect(rels).not.toContain('javascript:');
    });

    test('a second document with other styles and theme gets direct formatting', async ({ page }) => {
      const input = await open(page, mode);
      await selectFromStart(input, 3, 0, 6);
      await copy(page, input);
      const other = await load(page, await docx(OTHER_BODY, { fancy: false, accent: 'ABCDEF' }));
      await pasteAtEnd(other);
      const { body } = await pasted(page, 'ThemeX', 1);
      const copied = lastRun(body, 'ThemeX');
      expect(copied).not.toContain('w:rStyle');
      expect(copied).not.toContain('w:themeColor');
      expect(copied).toContain('<w:b/>');
      expect(copied).toMatch(/<w:sz [^>]*w:val="40"/);
      expect(copied).toMatch(/<w:color [^>]*w:val="112233"/);
    });

    test('suggesting mode keeps the formatting on a tracked insertion', async ({ page }) => {
      const input = await open(page, mode, '&suggesting=1');
      await selectFromStart(input, 0, 0, 5);
      await copy(page, input);
      await pasteAtEnd(input);
      const { body } = await pasted(page, 'BoldX', 2);
      expect(lastRun(body, 'BoldX')).toContain('<w:b/>');
      expect(body.slice(body.lastIndexOf('<w:ins '), body.length)).toContain('BoldX');
    });
  });
}
