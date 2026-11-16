import { expect, test, type Locator, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { pinnedEvaluate } from './pinned-evaluate';

const CUT_REFUSED = 'The selection holds objects that cannot be cut. Nothing was changed.';
const SENTINEL = 'CUT-REFUSAL-SENTINEL';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const CANVAS = '.canvas-page canvas, .paged-editor__page canvas';
const INPUT = '[data-testid=yrs-input], .paged-editor__input';
const CARET = '[data-docx-worker-editor] [data-docx-caret], [data-testid="caret"]';
const HIGHLIGHT = '[data-testid^="selection-rect-"], [data-docx-worker-editor] div[aria-hidden="true"][style*="rgba(59, 130, 246"]';

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
};
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument';

type Part = { path: string; xml: string; contentType: string; relId: string; relType: string; target: string };
type Kind = { id: string; object: string; parts: Part[]; cols: string; body?: string };

function notesPart(tag: 'footnote' | 'endnote', text: string): string {
  const separator = (type: string, id: string, child: string) =>
    `<w:${tag} w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${child}/></w:r></w:p></w:${tag}>`;
  return `<w:${tag}s xmlns:w="${NS.w}">${separator('separator', '-1', 'separator')}${separator('continuationSeparator', '0', 'continuationSeparator')}` +
    `<w:${tag} w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${tag}Ref/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:${tag}></w:${tag}s>`;
}

function noteKind(tag: 'footnote' | 'endnote'): Kind {
  const settings = `<w:settings xmlns:w="${NS.w}"><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>` +
    '<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr></w:settings>';
  return {
    id: tag, cols: '',
    object: `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${tag}Reference w:id="1"/></w:r>`,
    parts: [
      { path: `word/${tag}s.xml`, xml: notesPart(tag, `NOTE-TEXT-${tag}`), contentType: `${CT}.wordprocessingml.${tag}s+xml`, relId: 'rIdNotes', relType: `${NS.r}/${tag}s`, target: `${tag}s.xml` },
      { path: 'word/settings.xml', xml: settings, contentType: `${CT}.wordprocessingml.settings+xml`, relId: 'rIdSettings', relType: `${NS.r}/settings`, target: 'settings.xml' },
    ],
  };
}

function wsp(text: string, textBox: boolean, cx: number, cy: number): string {
  return `<a:graphic><a:graphicData uri="${NS.wps}"><wps:wsp><wps:cNvSpPr txBox="${textBox ? 1 : 0}"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${textBox ? 'rect' : 'roundRect'}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="EEF3FF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="2F5496"/></a:solidFill></a:ln></wps:spPr>` +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720"/></wps:wsp></a:graphicData></a:graphic>`;
}

const series = (index: number, name: string, values: number[]) =>
  `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:v>${name}</c:v></c:tx>` +
  '<c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>' +
  `<c:val><c:numRef><c:numCache><c:formatCode>0%</c:formatCode><c:ptCount val="2"/>${values.map((value, at) => `<c:pt idx="${at}"><c:v>${value}</c:v></c:pt>`).join('')}</c:numCache></c:numRef></c:val></c:ser>`;
const chartXml = `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}"><c:chart>` +
  '<c:title><c:tx><c:rich><a:p><a:r><a:t>CHART-TITLE</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
  series(0, 'North', [0.25, 0.5]) + series(1, 'South', [0.75, 0.5]) +
  '<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
  '<c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>';

const KINDS: Kind[] = [
  noteKind('footnote'),
  noteKind('endnote'),
  { id: 'column-break', cols: '<w:cols w:num="2" w:space="720"/>', parts: [], object: '<w:r><w:br w:type="column"/></w:r>' },
  {
    id: 'inline-shape', cols: '', parts: [],
    object: '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2286000" cy="457200"/><wp:docPr id="2" name="Shape 2"/>' +
      `${wsp('SHAPE-INLINE', false, 2286000, 457200)}</wp:inline></w:drawing></w:r>`,
  },
  {
    id: 'floating-text-box', cols: '', parts: [],
    object: '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>4572000</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="2286000" cy="457200"/><wp:wrapNone/><wp:docPr id="3" name="Text Box 3"/>' +
      `${wsp('SHAPE-FLOAT', true, 2286000, 457200)}</wp:anchor></w:drawing></w:r>`,
  },
  {
    id: 'chart', cols: '',
    object: '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="3657600" cy="1828800"/><wp:docPr id="4" name="Chart 4"/>' +
      `<a:graphic><a:graphicData uri="${NS.c}"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`,
    parts: [{ path: 'word/charts/chart1.xml', xml: chartXml, contentType: `${CT}.drawingml.chart+xml`, relId: 'rIdChart', relType: `${NS.r}/chart`, target: 'charts/chart1.xml' }],
  },
  {
    id: 'horizontal-rule', cols: '', parts: [],
    object: '<w:r><w:pict><v:rect id="_x0000_i1025" style="width:0;height:1.5pt" o:hralign="center" o:hrstd="t" o:hr="t" fillcolor="#a0a0a0" stroked="f"/></w:pict></w:r>',
  },
  { id: 'equation', cols: '', parts: [], object: '<m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>' },
  {
    id: 'content-control', cols: '', parts: [],
    object: '<w:sdt><w:sdtPr><w:alias w:val="Field"/><w:tag w:val="field"/><w:id w:val="101"/></w:sdtPr><w:sdtContent><w:r><w:t>SDT-TEXT</w:t></w:r></w:sdtContent></w:sdt>',
  },
];

const PLAIN: Kind = { id: 'plain', cols: '', parts: [], object: '' };

const breakKind = (type: 'page' | 'column'): Kind => ({
  id: `${type}-break-paragraph`, parts: [], object: '', cols: type === 'column' ? '<w:cols w:num="2" w:space="720"/>' : '',
  body: '<w:p w14:paraId="00000001"><w:r><w:t>ALPHA</w:t></w:r></w:p>' +
    `<w:p w14:paraId="00000002"><w:r><w:t xml:space="preserve">BRAVO before </w:t></w:r><w:r><w:br w:type="${type}"/></w:r><w:r><w:t xml:space="preserve"> after BRAVO</w:t></w:r></w:p>` +
    '<w:p w14:paraId="00000003"><w:r><w:t>CHARLIE</w:t></w:r></w:p>',
});
const BREAKS = [breakKind('page'), breakKind('column')];

const TWO_COLUMNS = '<w:cols w:num="2" w:space="720"/>';
const COLUMN_BREAK = '<w:r><w:br w:type="column"/></w:r>';
const columnBreakAfter = (id: string, lead: string): Kind => ({
  id, parts: [], object: '', cols: TWO_COLUMNS,
  body: '<w:p w14:paraId="00000001"><w:r><w:t>ALPHA</w:t></w:r></w:p>' +
    `<w:p w14:paraId="00000002">${lead}<w:r><w:t xml:space="preserve"> before </w:t></w:r>${COLUMN_BREAK}<w:r><w:t xml:space="preserve"> after BRAVO</w:t></w:r></w:p>` +
    '<w:p w14:paraId="00000003"><w:r><w:t>CHARLIE</w:t></w:r></w:p>',
});
const UNRECORDED = [
  columnBreakAfter('column-break-beside-a-hyperlink', '<w:hyperlink w:anchor="target"><w:r><w:t>LINK</w:t></w:r></w:hyperlink>'),
  columnBreakAfter('column-break-beside-a-field', '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>'),
];
const CELL_BREAK: Kind = {
  id: 'column-break-in-a-table-cell', parts: [], object: '', cols: TWO_COLUMNS,
  body: '<w:p w14:paraId="00000001"><w:r><w:t>ALPHA</w:t></w:r></w:p>' +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
    `<w:p w14:paraId="00000004"><w:r><w:t xml:space="preserve">CELL before </w:t></w:r>${COLUMN_BREAK}<w:r><w:t xml:space="preserve"> after CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
    '<w:p w14:paraId="00000003"><w:r><w:t>CHARLIE</w:t></w:r></w:p>',
};

async function press(page: Page, keys: string[]): Promise<void> {
  for (const key of keys) await page.keyboard.press(key);
}

async function selectSecondParagraphLine(page: Page, input: Locator): Promise<void> {
  await input.focus();
  await page.keyboard.press('ControlOrMeta+Home');
  await expect(page.locator(CARET).first()).toBeVisible({ timeout: 10_000 });
  await press(page, ['ArrowDown', 'Home', 'Shift+End']);
  await expect.poll(() => page.locator(HIGHLIGHT).count(), { timeout: 5000 }).toBeGreaterThan(0);
}

async function typeBeforeTheBreak(page: Page, input: Locator): Promise<void> {
  await input.focus();
  await page.keyboard.press('ControlOrMeta+Home');
  await expect(page.locator(CARET).first()).toBeVisible({ timeout: 10_000 });
  await press(page, ['ArrowDown', 'Home']);
  await page.keyboard.type('TYPED');
}

async function selectIntoLastParagraph(page: Page, input: Locator): Promise<void> {
  await input.focus();
  await page.keyboard.press('ControlOrMeta+Home');
  await expect(page.locator(CARET).first()).toBeVisible({ timeout: 10_000 });
  for (let step = 0; step < 2; step += 1) await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ControlOrMeta+Shift+End');
  for (let step = 0; step < 2; step += 1) await page.keyboard.press('Shift+ArrowLeft');
  await expect.poll(() => page.locator(HIGHLIGHT).count(), { timeout: 5000 }).toBeGreaterThan(1);
}

async function selectAll(page: Page, input: Locator): Promise<void> {
  await input.focus();
  await page.keyboard.press('ControlOrMeta+a');
  await expect.poll(() => page.locator(HIGHLIGHT).count(), { timeout: 5000 }).toBeGreaterThan(1);
}

async function build(kind: Kind): Promise<Buffer> {
  const zip = new JSZip();
  const overrides = [`<Override PartName="/word/document.xml" ContentType="${CT}.wordprocessingml.document.main+xml"/>`,
    ...kind.parts.map(part => `<Override PartName="/${part.path}" ContentType="${part.contentType}"/>`)].join('');
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    `<Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="${PKG}"><Relationship Id="rIdDoc" Type="${NS.r}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="${PKG}">${kind.parts.map(part => `<Relationship Id="${part.relId}" Type="${part.relType}" Target="${part.target}"/>`).join('')}</Relationships>`);
  for (const part of kind.parts) zip.file(part.path, part.xml);
  const declarations = Object.entries(NS).map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`).join(' ');
  zip.file('word/document.xml', `<w:document ${declarations}><w:body>` +
    (kind.body ?? `<w:p w14:paraId="00000001"><w:r><w:t xml:space="preserve">ALPHA before </w:t></w:r>${kind.object}<w:r><w:t xml:space="preserve"> after ALPHA</w:t></w:r></w:p>` +
    '<w:p w14:paraId="00000002"><w:r><w:t>BRAVO target</w:t></w:r></w:p>') +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>${kind.cols}</w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

const EDITORS = [{ name: 'worker', query: '&peerless=1' }, { name: 'default', query: '&peerless=0' }];

async function open(page: Page, query: string, kind: Kind): Promise<Locator> {
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto('/?format=docx' + query);
  await page.locator('input[type=file]').last().setInputFiles({ name: `cut-${kind.id}.docx`, mimeType: DOCX, buffer: await build(kind) });
  await expect(page.locator(CANVAS).first()).toBeVisible({ timeout: 90_000 });
  const input = page.locator(INPUT);
  await input.focus();
  await expect(input).toBeEditable({ timeout: 30_000 });
  return input;
}

async function savedParts(page: Page, info: TestInfo, kind: Kind, name: string): Promise<Record<string, string>> {
  const downloading = page.waitForEvent('download', { timeout: 60_000 });
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${kind.id}-${name}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const parts: Record<string, string> = {};
  for (const path of ['word/document.xml', ...kind.parts.map(part => part.path)]) parts[path] = (await zip.file(path)?.async('string')) ?? '';
  return parts;
}

async function selectFirstParagraph(page: Page, input: Locator): Promise<void> {
  await input.focus();
  await page.keyboard.press('ControlOrMeta+Home');
  const caret = page.locator(CARET).first();
  await expect(caret).toBeVisible({ timeout: 10_000 });
  const box = await caret.boundingBox();
  if (!box) throw new Error('The caret has no box');
  await page.mouse.click(box.x + 12, box.y + box.height / 2, { clickCount: 3 });
  await expect.poll(() => page.locator(HIGHLIGHT).count(), { timeout: 5000 }).toBeGreaterThan(0);
}

const clipboardText = (page: Page) => pinnedEvaluate(page, () => navigator.clipboard.readText().catch(() => ''));

for (const editor of EDITORS) {
  for (const kind of KINDS) {
    test(`a cut over a ${kind.id} is refused in the ${editor.name} editor`, async ({ page }, info) => {
      test.setTimeout(180_000);
      const input = await open(page, editor.query, kind);
      await pinnedEvaluate(page, (text: string) => navigator.clipboard.writeText(text), SENTINEL);
      const before = await savedParts(page, info, kind, 'before');
      expect(before['word/document.xml']).toContain('ALPHA');
      await selectFirstParagraph(page, input);
      await page.keyboard.press('ControlOrMeta+x');
      await expect(page.getByText(CUT_REFUSED)).toBeVisible({ timeout: 15_000 });
      expect(await savedParts(page, info, kind, 'after')).toEqual(before);
      expect(await clipboardText(page)).toBe(SENTINEL);
    });
  }

  for (const kind of BREAKS) {
    for (const [selection, select] of [['into the last paragraph', selectIntoLastParagraph], ['of everything', selectAll]] as const) {
      test(`a cut ${selection} across a ${kind.id} is refused in the ${editor.name} editor`, async ({ page }, info) => {
        test.setTimeout(180_000);
        const input = await open(page, editor.query, kind);
        await pinnedEvaluate(page, (text: string) => navigator.clipboard.writeText(text), SENTINEL);
        const before = await savedParts(page, info, kind, 'before');
        expect(before['word/document.xml']).toContain('CHARLIE');
        await select(page, input);
        await page.keyboard.press('ControlOrMeta+x');
        await expect(page.getByText(CUT_REFUSED)).toBeVisible({ timeout: 15_000 });
        expect(await savedParts(page, info, kind, 'after')).toEqual(before);
        expect(await clipboardText(page)).toBe(SENTINEL);
      });
    }
  }

  for (const kind of [KINDS[0]!, KINDS[2]!]) {
    test(`a Shift+Delete over a ${kind.id} is refused in the ${editor.name} editor`, async ({ page }, info) => {
      test.setTimeout(180_000);
      const input = await open(page, editor.query, kind);
      await pinnedEvaluate(page, (text: string) => navigator.clipboard.writeText(text), SENTINEL);
      const before = await savedParts(page, info, kind, 'before');
      await selectFirstParagraph(page, input);
      await page.keyboard.press('Shift+Delete');
      await expect(page.getByText(CUT_REFUSED)).toBeVisible({ timeout: 15_000 });
      expect(await savedParts(page, info, kind, 'after')).toEqual(before);
      expect(await clipboardText(page)).toBe(SENTINEL);
    });
  }

  test(`a Shift+Delete of plain text still cuts it to the clipboard in the ${editor.name} editor`, async ({ page }, info) => {
    test.setTimeout(180_000);
    const input = await open(page, editor.query, PLAIN);
    await pinnedEvaluate(page, (text: string) => navigator.clipboard.writeText(text), SENTINEL);
    await selectFirstParagraph(page, input);
    await page.keyboard.press('Shift+Delete');
    await expect.poll(() => clipboardText(page), { timeout: 15_000 }).toContain('ALPHA before');
    await expect.poll(async () => (await savedParts(page, info, PLAIN, 'after'))['word/document.xml'], { timeout: 30_000 }).not.toContain('ALPHA');
    await expect(page.getByText(CUT_REFUSED)).toHaveCount(0);
  });

  for (const [kind, selection, select] of [
    ...UNRECORDED.map(kind => [kind, 'of the line', selectSecondParagraphLine] as const),
    [CELL_BREAK, 'of everything', selectAll] as const,
  ]) {
    test(`a cut ${selection} beside a ${kind.id} is refused in the ${editor.name} editor`, async ({ page }, info) => {
      test.setTimeout(180_000);
      const input = await open(page, editor.query, kind);
      await pinnedEvaluate(page, (text: string) => navigator.clipboard.writeText(text), SENTINEL);
      const before = await savedParts(page, info, kind, 'before');
      expect(before['word/document.xml']).toContain('w:type="column"');
      await select(page, input);
      await page.keyboard.press('ControlOrMeta+x');
      await expect(page.getByText(CUT_REFUSED)).toBeVisible({ timeout: 15_000 });
      expect(await savedParts(page, info, kind, 'after')).toEqual(before);
      expect(await clipboardText(page)).toBe(SENTINEL);
    });
  }

  test(`a cut across a column break that typing moved is refused in the ${editor.name} editor`, async ({ page }, info) => {
    test.setTimeout(180_000);
    const kind = breakKind('column');
    const input = await open(page, editor.query, kind);
    await typeBeforeTheBreak(page, input);
    await pinnedEvaluate(page, (text: string) => navigator.clipboard.writeText(text), SENTINEL);
    const before = await savedParts(page, info, kind, 'before');
    expect(before['word/document.xml']).toContain('TYPED');
    await input.focus();
    await press(page, [...Array(9).fill('ArrowRight'), ...Array(8).fill('Shift+ArrowRight')]);
    await expect.poll(() => page.locator(HIGHLIGHT).count(), { timeout: 5000 }).toBeGreaterThan(0);
    await page.keyboard.press('ControlOrMeta+x');
    await expect(page.getByText(CUT_REFUSED)).toBeVisible({ timeout: 15_000 });
    expect(await savedParts(page, info, kind, 'after')).toEqual(before);
    expect(await clipboardText(page)).toBe(SENTINEL);
  });

  test(`a cut of ordinary text after typing before a column break still cuts in the ${editor.name} editor`, async ({ page }, info) => {
    test.setTimeout(180_000);
    const kind = breakKind('column');
    const input = await open(page, editor.query, kind);
    await typeBeforeTheBreak(page, input);
    await pinnedEvaluate(page, (text: string) => navigator.clipboard.writeText(text), SENTINEL);
    await press(page, ['ControlOrMeta+End', ...Array('CHARLIE'.length).fill('Shift+ArrowLeft')]);
    await expect.poll(() => page.locator(HIGHLIGHT).count(), { timeout: 5000 }).toBeGreaterThan(0);
    await page.keyboard.press('ControlOrMeta+x');
    await expect.poll(async () => (await savedParts(page, info, kind, 'after'))['word/document.xml'], { timeout: 30_000 }).not.toContain('CHARLIE');
    await expect(page.getByText(CUT_REFUSED)).not.toBeVisible();
    await expect(page.getByText(CUT_REFUSED)).toHaveCount(0);
    const saved = (await savedParts(page, info, kind, 'final'))['word/document.xml'];
    expect(saved).toContain('TYPED');
    expect(saved).toContain('w:type="column"');
    await expect.poll(() => clipboardText(page), { timeout: 15_000 }).toContain('CHARLIE');
  });

  test(`a cut of plain text still moves it to the clipboard in the ${editor.name} editor`, async ({ page }, info) => {
    test.setTimeout(180_000);
    const input = await open(page, editor.query, PLAIN);
    await pinnedEvaluate(page, (text: string) => navigator.clipboard.writeText(text), SENTINEL);
    await selectFirstParagraph(page, input);
    await page.keyboard.press('ControlOrMeta+x');
    await expect.poll(() => clipboardText(page), { timeout: 15_000 }).toContain('ALPHA before');
    await expect.poll(async () => (await savedParts(page, info, PLAIN, 'after'))['word/document.xml'], { timeout: 30_000 }).not.toContain('ALPHA');
    expect((await savedParts(page, info, PLAIN, 'final'))['word/document.xml']).toContain('BRAVO target');
    await expect(page.getByText(CUT_REFUSED)).toHaveCount(0);
  });
}

const DEFAULT_EDITOR = EDITORS.find(editor => editor.name === 'default')!;

test(`a cut of the first line after a column break selected with End and Shift+Home is refused in the ${DEFAULT_EDITOR.name} editor`, async ({ page }, info) => {
  test.setTimeout(180_000);
  const kind = breakKind('column');
  const input = await open(page, DEFAULT_EDITOR.query, kind);
  await pinnedEvaluate(page, (text: string) => navigator.clipboard.writeText(text), SENTINEL);
  const before = await savedParts(page, info, kind, 'before');
  expect(before['word/document.xml']).toContain('w:type="column"');
  await input.focus();
  await press(page, ['ControlOrMeta+End', 'Shift+Home']);
  await expect.poll(() => page.locator(HIGHLIGHT).count(), { timeout: 5000 }).toBeGreaterThan(0);
  await page.keyboard.press('ControlOrMeta+x');
  await expect(page.getByText(CUT_REFUSED)).toBeVisible({ timeout: 15_000 });
  expect(await savedParts(page, info, kind, 'after')).toEqual(before);
  expect(await clipboardText(page)).toBe(SENTINEL);
});

const extendThenShiftDeleteInOneTask = (page: Page) => pinnedEvaluate(page, (selector: string) => {
  const input = document.querySelector(selector);
  if (!input) throw new Error('The editor input is missing');
  for (const key of ['ArrowRight', 'Delete']) {
    input.dispatchEvent(new KeyboardEvent('keydown', { key, shiftKey: true, bubbles: true, cancelable: true }));
  }
}, INPUT);

test(`a Shift+Delete pressed before a queued selection onto a column break settles is refused in the ${DEFAULT_EDITOR.name} editor`, async ({ page }, info) => {
  test.setTimeout(180_000);
  const kind = breakKind('column');
  const input = await open(page, DEFAULT_EDITOR.query, kind);
  await pinnedEvaluate(page, (text: string) => navigator.clipboard.writeText(text), SENTINEL);
  const before = await savedParts(page, info, kind, 'before');
  await input.focus();
  await page.keyboard.press('ControlOrMeta+Home');
  await expect(page.locator(CARET).first()).toBeVisible({ timeout: 10_000 });
  await press(page, ['ArrowDown', 'Home', ...Array('BRAVO before '.length).fill('ArrowRight')]);
  await expect(page.locator(HIGHLIGHT)).toHaveCount(0);
  await extendThenShiftDeleteInOneTask(page);
  await expect(page.getByText(CUT_REFUSED)).toBeVisible({ timeout: 15_000 });
  expect(await savedParts(page, info, kind, 'after')).toEqual(before);
  expect(await clipboardText(page)).toBe(SENTINEL);
});

test(`a Shift+Delete pressed before a queued selection onto plain text settles cuts it in the ${DEFAULT_EDITOR.name} editor`, async ({ page }, info) => {
  test.setTimeout(180_000);
  const kind = breakKind('column');
  const input = await open(page, DEFAULT_EDITOR.query, kind);
  await pinnedEvaluate(page, (text: string) => navigator.clipboard.writeText(text), SENTINEL);
  await input.focus();
  await page.keyboard.press('ControlOrMeta+End');
  await expect(page.locator(CARET).first()).toBeVisible({ timeout: 10_000 });
  await press(page, ['Home', 'ArrowRight']);
  await expect(page.locator(HIGHLIGHT)).toHaveCount(0);
  await extendThenShiftDeleteInOneTask(page);
  await expect.poll(() => clipboardText(page), { timeout: 15_000 }).toBe('C');
  await expect.poll(async () => (await savedParts(page, info, kind, 'after'))['word/document.xml'], { timeout: 30_000 }).not.toContain('CHARLIE');
  expect((await savedParts(page, info, kind, 'final'))['word/document.xml']).toContain('HARLIE');
  await expect(page.getByText(CUT_REFUSED)).toHaveCount(0);
});

test(`a Shift+Delete whose queued selection holds only the hidden column-break unit changes nothing in the ${DEFAULT_EDITOR.name} editor`, async ({ page }, info) => {
  test.setTimeout(180_000);
  const kind = breakKind('column');
  const input = await open(page, DEFAULT_EDITOR.query, kind);
  await pinnedEvaluate(page, (text: string) => navigator.clipboard.writeText(text), SENTINEL);
  const before = await savedParts(page, info, kind, 'before');
  await input.focus();
  await page.keyboard.press('ControlOrMeta+End');
  await expect(page.locator(CARET).first()).toBeVisible({ timeout: 10_000 });
  await page.keyboard.press('Home');
  await expect(page.locator(HIGHLIGHT)).toHaveCount(0);
  await extendThenShiftDeleteInOneTask(page);
  await page.waitForTimeout(1000);
  expect(await clipboardText(page)).toBe(SENTINEL);
  expect(await savedParts(page, info, kind, 'after')).toEqual(before);
  await expect(page.getByText(CUT_REFUSED)).toHaveCount(0);
});
