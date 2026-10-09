import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { initWasm as initXlsx, openWorkbook } from '../../../xlsx/src/wasm/loader';
import { initWasm as initPptx, openPresentation } from '../../../pptx/src/wasm/loader';

type Sheet = { name: string; rows: (string | number)[][]; currency?: string };
type Slide = { title: string; bullets: string[]; notes?: string };

const directory = resolve(process.argv[2] ?? resolve(import.meta.dir, 'fixtures'));
const tasks: { id: string; format: string }[] = JSON.parse(await readFile(resolve(import.meta.dir, 'tasks.json'), 'utf8'));
const s = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
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

function rels(entries: { type: string; target: string; id?: string }[]) {
  return `<Relationships xmlns="${relationships}">${entries.map((entry, i) => `<Relationship Id="${entry.id ?? `rId${i + 1}`}" Type="${r}/${entry.type}" Target="${escape(entry.target)}"/>`).join('')}</Relationships>`;
}

function types(overrides: [string, string][]) {
  return `<Types xmlns="${contentTypes}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.map(([path, type]) => `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.${type}+xml"/>`).join('')}</Types>`;
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

await Promise.all([
  initXlsx(new Uint8Array(await readFile(resolve(import.meta.dir, '../../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm')))),
  initPptx(new Uint8Array(await readFile(resolve(import.meta.dir, '../../../pptx/src/wasm/generated/pptx_wasm_bg.wasm')))),
]);
await mkdir(directory, { recursive: true });
for (const task of tasks) {
  const index = Number(task.id.split('-')[1]) - 1;
  const bytes = task.format === 'xlsx' ? await workbook(workbooks[index]) : await presentation(decks[index]);
  await writeFile(resolve(directory, `${task.id}.${task.format}`), bytes);
  console.log(`${task.id}: ${bytes.length} bytes`);
}
