import { beforeAll, describe, expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { preloadEditWasm } from '../../packages/docx/src/wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../../packages/docx/src/wasm/opc';
import { createYrsSession, saveYrsDocx } from '../../packages/docx/src/yrs';

const OUT = process.env.DOCX_CRATE_PARITY_DIR ?? '';
if (OUT === '') throw new Error('DOCX_CRATE_PARITY_DIR is not set');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument';
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const ROOT = `xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="${REL}"`;
const STORY_ROOT = `${ROOT} xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;
const TEXT_BOX_ROOT = `${ROOT} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" mc:Ignorable="w14 wps v"`;
const SECTION = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const NEW_TEXT = 'Replaced paragraph text';

const p = (id: string, text: string) => `<w:p w14:paraId="${id}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const CHART_DRAWING = '<w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId5"/></a:graphicData></a:graphic></wp:inline></w:drawing>';
const BAR_CHART = `${DECLARATION}<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;

const ALTERNATE_TEXT_BOX = '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>1000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>2000</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrapNone/><wp:docPr id="11" name="Text Box 11"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:ind w:firstLineChars="200" w:firstLine="420"/></w:pPr><w:r><w:t>Choice callout</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="_x0000_s1026" type="#_x0000_t202"><v:textbox><w:txbxContent><w:p><w:r><w:t>Fallback callout</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent>';

const TOC_START = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
const TOC_END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const tocEntry = (text: string, page: string) => `<w:r><w:t>${text}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${page}</w:t></w:r>`;

const proofed = (id: string, text: string) =>
  `<w:p w14:paraId="${id}" w:rsidR="00752A6D" w:rsidRDefault="00593C67" w:rsidP="00593C67"><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr><w:proofErr w:type="spellStart"/><w:r w:rsidR="00761B10"><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r w:rsidRPr="00593C67"><w:t xml:space="preserve"> tail</w:t></w:r></w:p>`;

type Parts = Record<string, string>;

interface Case {
  name: string;
  parts: Parts;
  edit: { paragraph: number; paraId: string };
}

function contentTypes(overrides: Record<string, string>): string {
  const entries = Object.entries(overrides).map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join('');
  return `${DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${entries}</Types>`;
}

function relationships(entries: Array<[string, string, string]>): string {
  return `${DECLARATION}<Relationships xmlns="${PKG_REL}">${entries.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`).join('')}</Relationships>`;
}

function bodyOnly(body: string, root = ROOT): Parts {
  return {
    '[Content_Types].xml': contentTypes({ 'word/document.xml': `${OFFICE}.wordprocessingml.document.main+xml` }),
    '_rels/.rels': relationships([['rId1', 'officeDocument', 'word/document.xml']]),
    'word/document.xml': `${DECLARATION}<w:document ${root}><w:body>${body}${SECTION}</w:body></w:document>`,
  };
}

function charted(): Parts {
  return {
    '[Content_Types].xml': contentTypes({
      'word/document.xml': `${OFFICE}.wordprocessingml.document.main+xml`,
      'word/styles.xml': `${OFFICE}.wordprocessingml.styles+xml`,
      'word/header1.xml': `${OFFICE}.wordprocessingml.header+xml`,
      'word/footnotes.xml': `${OFFICE}.wordprocessingml.footnotes+xml`,
      'word/charts/chart1.xml': `${OFFICE}.drawingml.chart+xml`,
    }),
    '_rels/.rels': relationships([['rId1', 'officeDocument', 'word/document.xml']]),
    'word/_rels/document.xml.rels': relationships([
      ['rId1', 'styles', 'styles.xml'],
      ['rId2', 'header', 'header1.xml'],
      ['rId3', 'footnotes', 'footnotes.xml'],
      ['rId5', 'chart', 'charts/chart1.xml'],
    ]),
    'word/styles.xml': `${DECLARATION}<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `${DECLARATION}<w:document ${STORY_ROOT}><w:body>${p('11111111', 'Hello DOCX')}<w:p w14:paraId="33333333"><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:p w14:paraId="22222222"><w:r>${CHART_DRAWING}</w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/header1.xml': `${DECLARATION}<w:hdr ${STORY_ROOT}>${p('44444444', 'Header text')}<w:p w14:paraId="55555555"><w:r>${CHART_DRAWING}</w:r></w:p></w:hdr>`,
    'word/footnotes.xml': `${DECLARATION}<w:footnotes ${STORY_ROOT}><w:footnote w:id="-1" w:type="separator"><w:p w14:paraId="77777777"><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p w14:paraId="66666666"><w:r><w:t>Note text</w:t></w:r><w:r>${CHART_DRAWING}</w:r></w:p></w:footnote></w:footnotes>`,
    'word/charts/chart1.xml': BAR_CHART,
  };
}

const CASES: Case[] = [
  {
    name: 'plain',
    parts: bodyOnly([p('10000001', 'First paragraph.'), p('10000002', 'Second paragraph.'), p('10000003', 'Third paragraph.')].join('')),
    edit: { paragraph: 1, paraId: '10000002' },
  },
  {
    name: 'note-chart',
    parts: charted(),
    edit: { paragraph: 0, paraId: '11111111' },
  },
  {
    name: 'escaped-id',
    parts: bodyOnly([p('0B000001', 'first'), p('0B000002', 'second'), p('0B00000&#x33;', 'third')].join('')),
    edit: { paragraph: 2, paraId: '0B000003' },
  },
  {
    name: 'text-box',
    parts: bodyOnly([
      `<w:p w14:paraId="0D000001"><w:r>${ALTERNATE_TEXT_BOX}</w:r></w:p>`,
      '<w:p w14:paraId="0D000002"><w:pPr><w:ind w:firstLineChars="200" w:firstLine="420"/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p>',
      '<w:p w14:paraId="0D000003"><w:pPr><w:ind w:firstLineChars="0" w:firstLine="0"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>',
    ].join(''), TEXT_BOX_ROOT),
    edit: { paragraph: 1, paraId: '0D000002' },
  },
  {
    name: 'toc',
    parts: bodyOnly([
      `<w:p w14:paraId="0E000001">${TOC_START}${tocEntry('Overview', '1')}</w:p>`,
      `<w:p w14:paraId="0E000002">${tocEntry('Scope', '2')}${TOC_END}</w:p>`,
      p('0E000003', 'Overview'),
      p('0E000004', 'Body text after the contents.'),
    ].join('')),
    edit: { paragraph: 3, paraId: '0E000004' },
  },
  {
    name: 'proofing',
    parts: bodyOnly([
      proofed('0F000001', 'Speling before'),
      '<w:p w14:paraId="0F000002" w:rsidR="00752A6D" w:rsidRDefault="00593C67" w:rsidP="00593C67"><w:r w:rsidR="00761B10" w:rsidRPr="00593C67"><w:t>Run with revision ids</w:t></w:r></w:p>',
      proofed('0F000003', 'Speling after'),
    ].join('')),
    edit: { paragraph: 1, paraId: '0F000002' },
  },
  {
    name: 'repeated-paragraph-id',
    parts: bodyOnly([p('1A000001', 'First paragraph.'), p('1A000001', 'Second paragraph, same id.'), p('1A000003', 'Third paragraph.')].join('')),
    edit: { paragraph: 2, paraId: '1A000003' },
  },
];

async function pack(parts: Parts): Promise<Uint8Array> {
  const zip = new JSZip();
  const date = new Date('2024-01-01T00:00:00Z');
  for (const [name, content] of Object.entries(parts)) zip.file(name, content, { date });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

async function productSave(input: Uint8Array, edit?: { paraId: string; text: string }): Promise<Uint8Array> {
  const session = await createYrsSession({ clientId: 1 });
  try {
    session.seedFromDocx(input.slice(), { generation: createHash('sha256').update(input).digest('hex').slice(0, 32) });
    if (edit) {
      const identity = session.paragraphIdentities().paragraphs.find(
        (entry) => entry.session?.story === 'body' && entry.ooxmlParaId?.toUpperCase() === edit.paraId,
      );
      if (!identity?.session) throw new Error(`no body paragraph saves with ${edit.paraId}`);
      const key = identity.session.paraId;
      const span = session.paragraphSpans('body').find((entry) => entry.paraId === key);
      if (!span) throw new Error(`no span for ${edit.paraId}`);
      session.replaceRange({ story: 'body', start: { paraId: key, offset: 0 }, end: { paraId: key, offset: span.length } }, edit.text);
      expect(session.paragraphs('body').find((paragraph) => paragraph.paraId === key)?.text).toBe(edit.text);
    }
    const { bytes } = await saveYrsDocx(session, { updateModifiedDate: false });
    const xml = unzipContainer(bytes)['word/document.xml'];
    if (!(xml instanceof Uint8Array)) throw new Error('the saved package has no word/document.xml');
    return xml;
  } finally {
    session.destroy();
  }
}

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(new URL('../../packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm', import.meta.url))));
  await preloadOpcWasm(new Uint8Array(readFileSync(new URL('../../packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm', import.meta.url))));
  rmSync(OUT, { recursive: true, force: true });
  mkdirSync(OUT, { recursive: true });
});

describe('product save of generated documents for the crate', () => {
  for (const { name, parts, edit } of CASES) {
    test(name, async () => {
      const dir = join(OUT, name);
      mkdirSync(dir, { recursive: true });
      const input = await pack(parts);
      writeFileSync(join(dir, 'input.docx'), input);
      writeFileSync(join(dir, 'edit.json'), JSON.stringify({ ...edit, text: NEW_TEXT }));
      writeFileSync(join(dir, 'unedited.xml'), await productSave(input));
      writeFileSync(join(dir, 'edited.xml'), await productSave(input, { paraId: edit.paraId, text: NEW_TEXT }));
    }, 30_000);
  }
});

const REPORT_SMART_TAG = '<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place"><w:r><w:t>Ancient</w:t></w:r></w:smartTag>';
const REPORT_CELL = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p w14:paraId="1C000002"><w:r><w:t xml:space="preserve">Near </w:t></w:r>${REPORT_SMART_TAG}</w:p></w:tc></w:tr></w:tbl>`;
const PLAIN = CASES.find(({ name }) => name === 'plain')!;

interface ReportCase {
  name: string;
  parts: Parts;
  edit?: { paragraph: number; paraId: string };
  text?: string;
}

const STRUCTURE_DROPPED_DOCUMENT = `${DECLARATION}\n<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="${REL}" xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:body>${[
  `<w:p w14:paraId="0000000A" w14:textId="77777777"><w:r><w:t xml:space="preserve">Near </w:t></w:r>${REPORT_SMART_TAG}</w:p>`,
  '<w:p w14:paraId="0000000B" w14:textId="77777777"><w:customXml w:uri="urn:pr1" w:element="tagged"><w:r><w:t>Tagged</w:t></w:r></w:customXml></w:p>',
  '<w:p w14:paraId="0000000C" w14:textId="77777777"><w:bookmarkStart w:id="1" w:name="pr1mark"/><w:r><w:t>Marked</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>',
  '<w:p w14:paraId="0000000D" w14:textId="77777777"><w:pPr><mc:AlternateContent><mc:Choice Requires="w14"><w:jc w:val="center"/></mc:Choice><mc:Fallback><w:jc w:val="left"/></mc:Fallback></mc:AlternateContent></w:pPr><w:r><w:t>aaaa</w:t></w:r></w:p>',
  '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
].join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;

const REPORT_CASES: ReportCase[] = [
  { name: 'unedited', parts: PLAIN.parts },
  { name: 'edited', parts: PLAIN.parts, edit: PLAIN.edit },
  {
    name: 'smart-tag-cell',
    parts: bodyOnly([p('1C000001', 'First'), REPORT_CELL, p('1C000003', 'Last')].join('')),
    edit: { paragraph: 2, paraId: '1C000003' },
  },
  {
    name: 'structure-dropped',
    parts: {
      '[Content_Types].xml': `${DECLARATION}\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.wordprocessingml.document.main+xml"/></Types>`,
      '_rels/.rels': `${DECLARATION}\n<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
      'word/document.xml': STRUCTURE_DROPPED_DOCUMENT,
    },
    edit: { paragraph: 3, paraId: '0000000D' },
    text: 'aaa',
  },
];

/** The product save's report, and each session paragraph key's Word paragraph ID. */
async function productReport(input: Uint8Array, edit?: { paraId: string; text: string }) {
  const session = await createYrsSession({ clientId: 1 });
  try {
    session.seedFromDocx(input.slice(), { generation: createHash('sha256').update(input).digest('hex').slice(0, 32) });
    const identities = session.paragraphIdentities().paragraphs;
    if (edit) {
      const identity = identities.find(
        (entry) => entry.session?.story === 'body' && entry.ooxmlParaId?.toUpperCase() === edit.paraId,
      );
      if (!identity?.session) throw new Error(`no body paragraph saves with ${edit.paraId}`);
      const key = identity.session.paraId;
      const span = session.paragraphSpans('body').find((entry) => entry.paraId === key);
      if (!span) throw new Error(`no span for ${edit.paraId}`);
      session.replaceRange({ story: 'body', start: { paraId: key, offset: 0 }, end: { paraId: key, offset: span.length } }, edit.text);
      expect(session.paragraphs('body').find((paragraph) => paragraph.paraId === key)?.text).toBe(edit.text);
    }
    const { keptUnchanged, structureDropped } = await saveYrsDocx(session, { updateModifiedDate: false });
    const ids: Record<string, string> = {};
    for (const entry of identities) if (entry.session && entry.ooxmlParaId) ids[entry.session.paraId] = entry.ooxmlParaId.toUpperCase();
    return { report: { keptUnchanged, structureDropped }, ids };
  } finally {
    session.destroy();
  }
}

describe('product save reports of generated documents for the crate', () => {
  for (const { name, parts, edit, text = NEW_TEXT } of REPORT_CASES) {
    test(name, async () => {
      const dir = join(OUT, 'report', name);
      mkdirSync(dir, { recursive: true });
      const input = await pack(parts);
      writeFileSync(join(dir, 'input.docx'), input);
      writeFileSync(join(dir, 'edit.json'), JSON.stringify(edit ? { ...edit, text } : {}));
      const { report, ids } = await productReport(input, edit && { paraId: edit.paraId, text });
      writeFileSync(join(dir, 'report.json'), JSON.stringify(report));
      writeFileSync(join(dir, 'ids.json'), JSON.stringify(ids));
    }, 30_000);
  }
});
