import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { initWasm as initXlsx, openWorkbook } from '../../packages/xlsx/src/wasm/loader';
import { initWasm as initPptx, openPresentation } from '../../packages/pptx/src/wasm/loader';
import { createYrsSession, saveYrsDocx } from '../../packages/docx/src/yrs';

type Sheet = { name: string; rows: (string | number)[][]; currency?: string };
type Slide = { title: string; bullets: string[]; notes?: string };
type Paragraph = { text: string; style?: string; list?: boolean };
type Block = Paragraph | { rows: string[][] };

if (!process.argv[2]) throw new Error('Usage: bun eval/formats/generate-fixtures.ts <output-dir>');
const directory = resolve(process.argv[2]);
const tasks: { id: string; format: string }[] = JSON.parse(await readFile(resolve(import.meta.dir, 'tasks.json'), 'utf8'));
const s = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const wordprocessing = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const contentTypes = 'http://schemas.openxmlformats.org/package/2006/content-types';
const fixedDate = new Date('2000-01-01T00:00:00Z');
const escape = (text: string) => text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('"', '&quot;');

const budget = (broken = false): Sheet[] => [{ name: 'Budget', rows: [
  ['Item', 'Planned', 'Actual'], ['Rent', 300, 320], ['Supplies', 120, 100], ['Travel', 200, 210],
  ['Total', '=SUM(B2:B4)', broken ? '=SUM(C2:C3)' : '=SUM(C2:C4)'],
], currency: 'B2:C5' }];
const sales = (): Sheet[] => [{ name: 'Sales', rows: [
  ['Item', 'Revenue', 'Cost'], ['Cap', 120, 50], ['Bag', 250, 100], ['Bottle', 80, 30],
], currency: 'B2:C4' }];
const revenue = (name: string, owner = false): Sheet[] => [{ name, rows: [
  ['Product', owner ? 'Q1' : '2025', owner ? 'Q2' : '2026', owner ? 'Owner' : 'Status'],
  ['Cap', 100, 125, owner ? 'Avery' : 'Stable'], ['Bag', 200, 300, owner ? 'Robin' : 'Growing'],
  ['Bottle', 100, 80, owner ? 'Morgan' : 'Review'],
], currency: 'B2:C4' }];
const workbooks: Sheet[][] = [
  budget(), budget(true),
  [{ name: 'Catalog', rows: [['Product', 'Price'], ['Trail Mix', 6], ['Oat Bar', 3], ['Trail Mix family pack', 12]] },
    { name: 'Orders', rows: [['Order', 'Product', 'Quantity'], [101, 'Trail Mix', 2], [102, 'Oat Bar', 4], [103, 'Trail Mix family pack', 1]] }],
  sales(),
  [{ name: 'Inventory', rows: [['Product', 'SKU', 'Stock'], ['Notebook', 'N-01', 12], ['Pen', 'P-02', 4], ['Folder', 'F-03', 20], ['Tape', 'T-04', 8]] }],
  sales(), revenue('Revenue'),
  [{ name: 'Forecast', rows: [['Metric', 'Value'], ['Revenue', 450]] },
    { name: 'Assumptions', rows: [['Metric', 'Value'], ['Planned', "='Forecast'!B2"], ['Actual', "='Actuals'!B2"]] },
    { name: 'Actuals', rows: [['Metric', 'Value'], ['Revenue', 400]] }],
  [{ name: 'Expenses', rows: [['Item', 'Planned', 'Actual'], ['Cloud', 100, 120], ['Support', 200, 250], ['Hosting', 50, 100]], currency: 'B2:C4' },
    { name: 'Summary', rows: [['Metric', 'Amount'], ['Planned', 0], ['Actual', 0], ['Variance', 0]] }],
  revenue('Quarterly', true),
];
const slide = (title: string, bullets = ['Review progress', 'Discuss risks']): Slide => ({ title, bullets });
const decks: Slide[][] = [
  [slide('Quarterly update'), slide('Agenda'), slide('Results')],
  [slide('Overview'), slide('Results', ['Revenue: $1.2 million', 'Margin: 18%', 'Retention: 91%']), slide('Next steps')],
  [slide('Overview'), slide('Next steps', ['Review draft', 'Wait for feedback', 'Schedule later']), slide('Appendix')],
  [slide('Atlas overview', ['Atlas pilot is ready', 'Atlas has ten participants']),
    { ...slide('Atlas results', ['Atlas retention improved', 'Review Atlas risks']), notes: 'Atlas figures are preliminary.' },
    slide('Next steps', ['Launch Atlas in October', 'Keep the Atlas team informed'])],
  [slide('Overview', ['Draft results will be reviewed', 'Draft next steps will be discussed']),
    slide('Draft results', ['Draft results are preliminary', 'Review with the team']), slide('Draft next steps')],
  [slide('Overview'), slide('Agenda'), slide('Appendix'), slide('Next steps')],
  [slide('Overview'), slide('Next steps'), slide('Agenda'), slide('Results')],
  [slide('Overview'), { ...slide('Results'), notes: 'Draft numbers.' }, slide('Next steps')],
  [slide('Overview'), slide('Results'), slide('Next steps')],
  [slide('Overview'), slide('Agenda'), slide('Results'), slide('Next steps')],
];
const para = (text: string, style?: string): Paragraph => ({ text, style });
const item = (text: string): Paragraph => ({ text, list: true });
const report = (): Block[] => [
  para('Pilot status report', 'Title'), para('Prepared on 2 October 2026 by the project office.'),
  para('Summary', 'Heading1'), para('The pilot started in September with ten participants across two sites.'),
  para('Feedback so far is positive, and the team expects to finish on schedule.'),
  item('Site visits are complete'), item('Training material is in review'), item('Reporting starts next month'),
  para('Budget', 'Heading1'), para('The table lists planned and actual costs in euros.'),
  { rows: [['Item', 'Owner', 'Planned', 'Actual'], ['Rent', 'Avery', '300', '320'], ['Supplies', 'Robin', '120', '100'],
    ['Travel', 'Morgan', '200', '210'], ['Total', '', '620', '630']] },
  para('Next steps', 'Heading1'), para('Collect the remaining feedback forms.'), para('Prepare the final report for the steering group.'),
];
const edit = (blocks: Block[], text: string, ...replacement: Block[]) => blocks.flatMap(block => 'text' in block && block.text === text ? replacement : [block]);
const atlas: Record<string, string> = {
  'Pilot status report': 'Atlas pilot status report',
  'The pilot started in September with ten participants across two sites.': 'The Atlas pilot started in September with ten participants across two sites.',
  'Feedback so far is positive, and the team expects to finish on schedule.': 'Feedback on Atlas is positive, and the Atlas team expects to finish on schedule. The regional atlas of sites is attached.',
  'Training material is in review': 'Atlas training material is in review',
  Supplies: 'Atlas licences',
};
const documents: Block[][] = [
  edit(report(), 'Summary', para('Sumary', 'Heading1')),
  report().map(block => 'rows' in block ? { rows: block.rows.map(row => row.map(cell => atlas[cell] ?? cell)) } : { ...block, text: atlas[block.text] ?? block.text }),
  report(), report(), report(),
  edit(report(), 'Prepare the final report for the steering group.', para('Prepare the final report for the steering group.'),
    para('Risks and mitigations'), para('Supplier delays could move the final report to December.')),
  report(), report(),
  edit(report(), 'The table lists planned and actual costs in euros.', para('The table lists planned and actual costs in euros.'),
    para('Draft note: confirm the travel figures with finance before sending.')),
  edit(edit(report(), 'Pilot status report', para('Pilot status report'), para('DRAFT')),
    'Prepared on 2 October 2026 by the project office.', para('Prepared on [DATE] by the project office.')),
];

function rels(entries: { type: string; target: string; id?: string }[]) {
  return `<Relationships xmlns="${relationships}">${entries.map((entry, i) => `<Relationship Id="${entry.id ?? `rId${i + 1}`}" Type="${entry.type.includes('://') ? entry.type : `${r}/${entry.type}`}" Target="${escape(entry.target)}"/>`).join('')}</Relationships>`;
}

function types(overrides: [string, string][]) {
  return `<Types xmlns="${contentTypes}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.map(([path, type]) => `<Override PartName="/${path}" ContentType="${type.includes('/') ? type : `application/vnd.openxmlformats-officedocument.${type}+xml`}"/>`).join('')}</Types>`;
}

async function normalize(bytes: Uint8Array) {
  const source = await JSZip.loadAsync(bytes);
  const zip = new JSZip();
  for (const path of Object.keys(source.files).sort()) {
    if (!source.files[path].dir) zip.file(path, await source.files[path].async('uint8array'), { date: fixedDate, createFolders: false });
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 9 } });
}

async function workbook(sheets: Sheet[]) {
  const seed = new JSZip();
  seed.file('_rels/.rels', rels([{ type: 'officeDocument', target: 'xl/workbook.xml' }]));
  seed.file('[Content_Types].xml', types([
    ['xl/workbook.xml', 'spreadsheetml.sheet.main'], ['xl/styles.xml', 'spreadsheetml.styles'],
    ...sheets.map((_, i): [string, string] => [`xl/worksheets/sheet${i + 1}.xml`, 'spreadsheetml.worksheet']),
  ]));
  seed.file('xl/workbook.xml', `<workbook xmlns="${s}" xmlns:r="${r}"><bookViews><workbookView/></bookViews><sheets>${sheets.map((sheet, i) => `<sheet name="${escape(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`);
  seed.file('xl/_rels/workbook.xml.rels', rels([
    ...sheets.map((_, i) => ({ type: 'worksheet', target: `worksheets/sheet${i + 1}.xml` })),
    { type: 'styles', target: 'styles.xml' },
  ]));
  seed.file('xl/styles.xml', `<styleSheet xmlns="${s}"><fonts count="1"><font><sz val="11"/><name val="DejaVu Sans"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
  sheets.forEach((_, i) => seed.file(`xl/worksheets/sheet${i + 1}.xml`, `<worksheet xmlns="${s}"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/></worksheet>`));
  const handle = openWorkbook(await seed.generateAsync({ type: 'uint8array' }));
  try {
    sheets.forEach((sheet, i) => {
      handle.editCells(i, sheet.rows.flatMap((row, y) => row.map((value, x) => ({ row: y, col: x, input: String(value) }))));
      if (sheet.currency) handle.setNumberFormat(i, sheet.currency, { type: 'custom', pattern: '$#,##0.00' });
    });
    return await normalize(handle.save());
  } finally { handle.dispose(); }
}

function group() {
  return '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
}

function shape(id: number, name: string, count: number) {
  const title = name === 'Title';
  const paragraphs = Array.from({ length: count }, (_, i) => `<a:p><a:pPr${title ? '' : ' marL="342900" indent="-285750"'}>${title ? '<a:buNone/>' : '<a:buChar char="•"/>'}</a:pPr><a:r><a:rPr sz="${title ? 2800 : 1800}"${title ? ' b="1"' : ''}><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:latin typeface="DejaVu Sans"/></a:rPr><a:t>seed-${name}-${i}</a:t></a:r></a:p>`).join('');
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr><p:ph type="${title ? 'title' : 'body'}" idx="${id - 2}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="${title ? 457200 : 1600200}"/><a:ext cx="8229600" cy="${title ? 685800 : 3429000}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}

async function presentation(slides: Slide[]) {
  const seed = new JSZip();
  seed.file('_rels/.rels', rels([{ type: 'officeDocument', target: 'ppt/presentation.xml' }]));
  seed.file('[Content_Types].xml', types([
    ['ppt/presentation.xml', 'presentationml.presentation.main'],
    ['ppt/slideLayouts/slideLayout1.xml', 'presentationml.slideLayout'], ['ppt/slideMasters/slideMaster1.xml', 'presentationml.slideMaster'],
    ...slides.map((_, i): [string, string] => [`ppt/slides/slide${i + 1}.xml`, 'presentationml.slide']),
  ]));
  seed.file('ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${r}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst><p:sldIdLst>${slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  seed.file('ppt/_rels/presentation.xml.rels', rels([
    ...slides.map((_, i) => ({ type: 'slide', target: `slides/slide${i + 1}.xml` })),
    { type: 'slideMaster', target: 'slideMasters/slideMaster1.xml', id: 'rIdMaster' },
  ]));
  const colorMap = '<p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/>';
  seed.file('ppt/slideMasters/slideMaster1.xml', `<p:sldMaster xmlns:p="${p}" xmlns:a="${a}" xmlns:r="${r}"><p:cSld><p:spTree>${group()}</p:spTree></p:cSld>${colorMap}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle><a:lvl1pPr><a:buChar char="•"/></a:lvl1pPr></p:bodyStyle><p:otherStyle/></p:txStyles></p:sldMaster>`);
  seed.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels([{ type: 'slideLayout', target: '../slideLayouts/slideLayout1.xml' }]));
  seed.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout xmlns:p="${p}" xmlns:a="${a}" xmlns:r="${r}" type="tx" preserve="1"><p:cSld name="Title and bullets"><p:spTree>${group()}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`);
  seed.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels([{ type: 'slideMaster', target: '../slideMasters/slideMaster1.xml' }]));
  slides.forEach((slide, i) => {
    seed.file(`ppt/slides/slide${i + 1}.xml`, `<p:sld xmlns:p="${p}" xmlns:a="${a}" xmlns:r="${r}"><p:cSld name="Slide ${i + 1}"><p:spTree>${group()}${shape(2, 'Title', 1)}${shape(3, 'Bullets', slide.bullets.length)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
    seed.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, rels([{ type: 'slideLayout', target: '../slideLayouts/slideLayout1.xml' }]));
  });
  const handle = openPresentation(await seed.generateAsync({ type: 'uint8array' }), { clientId: 1001 });
  try {
    const content = handle.readContent();
    if (!content.ok) throw new Error(JSON.stringify(content));
    const steps = content.stories.flatMap(story => {
      const index = content.slides.findIndex(slide => slide.id === story.slideId);
      const texts = story.text.startsWith('seed-Title-') ? [slides[index].title] : slides[index].bullets;
      return story.paragraphs.map((paragraph, i) => ({ op: 'replaceText' as const,
        target: { kind: 'range' as const, slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId, start: paragraph.start, end: paragraph.end }, text: texts[i] }));
    });
    const edited = handle.applyEdits({ expectVersion: content.version, steps });
    if (!edited.ok) throw new Error(JSON.stringify(edited));
    content.slides.forEach((slide, i) => handle.setSlideNotes(slide.id, slides[i].notes ?? `Presenter cue: ${slides[i].title}`));
    return await normalize(handle.save());
  } finally { handle.dispose(); }
}

function wordRun(text: string, properties = '') {
  return `<w:r>${properties && `<w:rPr>${properties}</w:rPr>`}<w:t xml:space="preserve">${escape(text)}</w:t></w:r>`;
}

function wordTable(rows: string[][]) {
  const widths = [3200, 2200, 1800, 1800];
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9000" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${widths.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${rows.map((cells, row) => {
    const strong = row === 0 || cells[0] === 'Total' ? '<w:b/>' : '';
    return `<w:tr>${row === 0 ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${cells.map((text, column) => `<w:tc><w:tcPr><w:tcW w:w="${widths[column]}" w:type="dxa"/>${row === 0 ? '<w:shd w:val="clear" w:color="auto" w:fill="D9EAF7"/>' : ''}</w:tcPr><w:p>${column >= 2 ? '<w:pPr><w:jc w:val="right"/></w:pPr>' : ''}${text ? wordRun(text, `${strong}<w:sz w:val="20"/>`) : ''}</w:p></w:tc>`).join('')}</w:tr>`;
  }).join('')}</w:tbl>`;
}

async function wordDocument(blocks: Block[]) {
  const seed = new JSZip();
  const core = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
  seed.file('_rels/.rels', rels([{ type: 'officeDocument', target: 'word/document.xml' }, { type: core, target: 'docProps/core.xml' },
    { type: 'extended-properties', target: 'docProps/app.xml' }]));
  seed.file('[Content_Types].xml', types([
    ['word/document.xml', 'wordprocessingml.document.main'], ['word/styles.xml', 'wordprocessingml.styles'],
    ['word/settings.xml', 'wordprocessingml.settings'], ['word/numbering.xml', 'wordprocessingml.numbering'],
    ['docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'], ['docProps/app.xml', 'extended-properties'],
  ]));
  seed.file('word/_rels/document.xml.rels', rels([{ type: 'styles', target: 'styles.xml' }, { type: 'settings', target: 'settings.xml' },
    { type: 'numbering', target: 'numbering.xml' }]));
  let paragraphIndex = 0;
  const paragraphIds = new Set<number>();
  const body = blocks.map(block => 'rows' in block ? wordTable(block.rows)
    : `<w:p>${block.style || block.list ? `<w:pPr>${block.style ? `<w:pStyle w:val="${block.style}"/>` : ''}${block.list ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : ''}</w:pPr>` : ''}${wordRun(block.text)}</w:p>`).join('').replaceAll('<w:p>', () => {
      const ids = ['paraId', 'textId'].map(name => {
        const value = createHash('sha256').update(`format-eval-docx-${name}-${paragraphIndex}`).digest().readUInt32BE(0) & 0x7fffffff;
        if (value === 0 || paragraphIds.has(value)) throw new Error(`Invalid or colliding DOCX ${name} at paragraph ${paragraphIndex}`);
        paragraphIds.add(value);
        return `w14:${name}="${value.toString(16).toUpperCase().padStart(8, '0')}"`;
      });
      paragraphIndex += 1;
      return `<w:p ${ids.join(' ')}>`;
    });
  seed.file('word/document.xml', `<w:document xmlns:w="${wordprocessing}" xmlns:r="${r}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr></w:body></w:document>`);
  const heading = (id: string, name: string, level: number, size: number) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${level}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('');
  seed.file('word/styles.xml', `<w:styles xmlns:w="${wordprocessing}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans" w:eastAsia="DejaVu Sans" w:cs="DejaVu Sans"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
    + '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr></w:style>'
    + heading('Heading1', 'heading 1', 0, 32) + heading('Heading2', 'heading 2', 1, 26)
    + '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>'
    + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
    + `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>${borders}</w:tblBorders></w:tblPr></w:style>`
    + '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:semiHidden/></w:style></w:styles>');
  seed.file('word/numbering.xml', `<w:numbering xmlns:w="${wordprocessing}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`);
  seed.file('word/settings.xml', `<w:settings xmlns:w="${wordprocessing}"><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
  seed.file('docProps/core.xml', '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Pilot status report</dc:title><dc:creator>Format evaluation</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>');
  seed.file('docProps/app.xml', '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Format evaluation</Application></Properties>');
  const session = await createYrsSession({ clientId: 1001 });
  try {
    session.seedFromDocx(await seed.generateAsync({ type: 'uint8array' }));
    return await normalize((await saveYrsDocx(session, { updateModifiedDate: false })).bytes);
  } finally { session.destroy(); }
}

await Promise.all([
  initXlsx(new Uint8Array(await readFile(resolve(import.meta.dir, '../../packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm')))),
  initPptx(new Uint8Array(await readFile(resolve(import.meta.dir, '../../packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm')))),
]);
await mkdir(directory, { recursive: true });
for (const task of tasks) {
  const index = Number(task.id.split('-')[1]) - 1;
  const bytes = task.format === 'xlsx' ? await workbook(workbooks[index])
    : task.format === 'pptx' ? await presentation(decks[index]) : await wordDocument(documents[index]);
  await writeFile(resolve(directory, `${task.id}.${task.format}`), bytes);
  console.log(`${task.id}.${task.format} ${bytes.length} ${createHash('sha256').update(bytes).digest('hex')}`);
}
