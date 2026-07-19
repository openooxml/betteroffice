import JSZip from 'jszip';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const output = path.join(root, 'apps/demo/public/betteroffice-demo.pptx');
const zipDate = new Date('2026-01-01T00:00:00Z');

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

type RunOptions = {
  size?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
};

function run(text: string, options: RunOptions = {}): string {
  const size = options.size ?? 2400;
  const color = options.color ?? '172036';
  const bold = options.bold ? ' b="1"' : '';
  const italic = options.italic ? ' i="1"' : '';
  return `<a:r><a:rPr lang="en-US" sz="${size}"${bold}${italic}><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>${escapeXml(text)}</a:t></a:r>`;
}

function paragraph(
  contents: string,
  alignment = 'l',
  level = 0,
  bullet?: string,
): string {
  const bulletXml = bullet ? `<a:buChar char="${escapeXml(bullet)}"/>` : '';
  return `<a:p><a:pPr algn="${alignment}" lvl="${level}">${bulletXml}</a:pPr>${contents}<a:endParaRPr lang="en-US" sz="1800"/></a:p>`;
}

type TextBoxOptions = {
  id: number;
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
  paragraphs: string;
  fill?: string;
  outline?: string;
  geometry?: string;
  anchor?: string;
  radius?: boolean;
};

function textBox(options: TextBoxOptions): string {
  const fill = options.fill
    ? `<a:solidFill><a:srgbClr val="${options.fill}"/></a:solidFill>`
    : '<a:noFill/>';
  const outline = options.outline
    ? `<a:ln w="19050"><a:solidFill><a:srgbClr val="${options.outline}"/></a:solidFill></a:ln>`
    : '<a:ln><a:noFill/></a:ln>';
  const geometry = options.geometry ?? (options.radius ? 'roundRect' : 'rect');
  return `<p:sp>
    <p:nvSpPr><p:cNvPr id="${options.id}" name="${escapeXml(options.name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="${options.x}" y="${options.y}"/><a:ext cx="${options.width}" cy="${options.height}"/></a:xfrm><a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom>${fill}${outline}</p:spPr>
    <p:txBody><a:bodyPr wrap="square" lIns="152400" tIns="91440" rIns="152400" bIns="91440" anchor="${options.anchor ?? 't'}"/><a:lstStyle/>${options.paragraphs}</p:txBody>
  </p:sp>`;
}

type ShapeOptions = {
  id: number;
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
  fill?: string;
  outline?: string;
  geometry?: string;
};

function shape(options: ShapeOptions): string {
  const fill = options.fill
    ? `<a:solidFill><a:srgbClr val="${options.fill}"/></a:solidFill>`
    : '<a:noFill/>';
  const outline = options.outline
    ? `<a:ln w="19050"><a:solidFill><a:srgbClr val="${options.outline}"/></a:solidFill></a:ln>`
    : '<a:ln><a:noFill/></a:ln>';
  return `<p:sp>
    <p:nvSpPr><p:cNvPr id="${options.id}" name="${escapeXml(options.name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="${options.x}" y="${options.y}"/><a:ext cx="${options.width}" cy="${options.height}"/></a:xfrm><a:prstGeom prst="${options.geometry ?? 'rect'}"><a:avLst/></a:prstGeom>${fill}${outline}</p:spPr>
  </p:sp>`;
}

function fixturePicture(): string {
  return `<p:pic>
    <p:nvPicPr><p:cNvPr id="90" name="Media fixture" descr="Transparent parser fixture"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
    <p:blipFill><a:blip r:embed="rId2"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
    <p:spPr><a:xfrm><a:off x="12192000" y="6858000"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln><a:noFill/></a:ln></p:spPr>
  </p:pic>`;
}

function editorMockup(): string {
  return `<p:grpSp>
    <p:nvGrpSpPr><p:cNvPr id="10" name="BetterOffice editor preview"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="7112000" y="610000"/><a:ext cx="4320000" cy="5480000"/><a:chOff x="0" y="0"/><a:chExt cx="4320000" cy="5480000"/></a:xfrm></p:grpSpPr>
    ${shape({ id: 11, name: 'Preview field', x: 0, y: 0, width: 4320000, height: 5480000, fill: 'E6EBFF', geometry: 'roundRect' })}
    ${shape({ id: 12, name: 'Preview accent', x: 3100000, y: 0, width: 1220000, height: 5480000, fill: '315EFB', geometry: 'roundRect' })}
    ${shape({ id: 13, name: 'Editor window', x: 360000, y: 520000, width: 3380000, height: 3860000, fill: 'FFFFFF', outline: 'C7D1FF', geometry: 'roundRect' })}
    ${shape({ id: 14, name: 'Window dot coral', x: 620000, y: 760000, width: 105000, height: 105000, fill: 'FF8066', geometry: 'ellipse' })}
    ${shape({ id: 15, name: 'Window dot gold', x: 780000, y: 760000, width: 105000, height: 105000, fill: 'F9C74F', geometry: 'ellipse' })}
    ${shape({ id: 16, name: 'Window dot mint', x: 940000, y: 760000, width: 105000, height: 105000, fill: '48C78E', geometry: 'ellipse' })}
    ${textBox({ id: 17, name: 'Window address', x: 1160000, y: 690000, width: 2080000, height: 260000, paragraphs: paragraph(run('betteroffice.dev / deck.pptx', { size: 820, color: '667085', bold: true })), anchor: 'ctr' })}
    ${textBox({ id: 18, name: 'Preview title', x: 650000, y: 1320000, width: 2550000, height: 980000, paragraphs: [paragraph(run('Ship ideas,', { size: 2100, color: '101828', bold: true })), paragraph(run('not attachments.', { size: 2100, color: '315EFB', bold: true }))].join('') })}
    ${shape({ id: 19, name: 'Preview rule one', x: 650000, y: 2520000, width: 2220000, height: 105000, fill: 'D0D5DD', geometry: 'roundRect' })}
    ${shape({ id: 20, name: 'Preview rule two', x: 650000, y: 2770000, width: 1670000, height: 105000, fill: 'D0D5DD', geometry: 'roundRect' })}
    ${textBox({ id: 21, name: 'Local badge', x: 650000, y: 3230000, width: 1550000, height: 520000, paragraphs: paragraph(run('LOCAL-FIRST', { size: 900, color: '101828', bold: true }), 'ctr'), fill: 'C8F56A', anchor: 'ctr', radius: true })}
    ${textBox({ id: 22, name: 'DOCX tab', x: 180000, y: 4610000, width: 1160000, height: 560000, paragraphs: paragraph(run('DOCX', { size: 1000, color: 'FFFFFF', bold: true }), 'ctr'), fill: '101828', anchor: 'ctr', radius: true })}
    ${textBox({ id: 23, name: 'XLSX tab', x: 1570000, y: 4610000, width: 1160000, height: 560000, paragraphs: paragraph(run('XLSX', { size: 1000, color: '101828', bold: true }), 'ctr'), fill: 'C8F56A', anchor: 'ctr', radius: true })}
    ${textBox({ id: 24, name: 'PPTX tab', x: 2960000, y: 4610000, width: 1160000, height: 560000, paragraphs: paragraph(run('PPTX', { size: 1000, color: '101828', bold: true }), 'ctr'), fill: 'FF8066', anchor: 'ctr', radius: true })}
  </p:grpSp>`;
}

function tableCell(text: string, fill: string, color: string, bold = false): string {
  return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraph(
    run(text, { size: 1500, color, bold }),
    'ctr',
  )}</a:txBody><a:tcPr><a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR><a:lnT><a:noFill/></a:lnT><a:lnB><a:noFill/></a:lnB><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></a:tcPr></a:tc>`;
}

function architectureTable(): string {
  const rows = [
    ['Boundary', 'OPC', 'Guarded parts', 'No external fetch'],
    ['Model', 'PresentationML', 'Typed slide tree', 'Untouched bytes retained'],
    ['Editing', 'CRDT deck', 'Local-origin undo', 'Raw update exchange'],
    ['Output', 'Slide layout', 'Shaped text', 'Canvas display list'],
  ];
  const header = ['Layer', 'Rust owner', 'Contract', 'Result'];
  return `<p:graphicFrame>
    <p:nvGraphicFramePr><p:cNvPr id="90" name="Architecture table fixture"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
    <p:xfrm><a:off x="12192000" y="6858000"/><a:ext cx="1" cy="1"/></p:xfrm>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
      <a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>
      <a:tblGrid><a:gridCol w="1981200"/><a:gridCol w="2438400"/><a:gridCol w="2895600"/><a:gridCol w="3352800"/></a:tblGrid>
      <a:tr h="609600">${header.map((value) => tableCell(value, '172036', 'FFFFFF', true)).join('')}<a:extLst/></a:tr>
      ${rows
        .map(
          (row, rowIndex) =>
            `<a:tr h="762000">${row
              .map((value, columnIndex) =>
                tableCell(
                  value,
                  rowIndex % 2 === 0 ? 'F4F3FF' : 'FFFFFF',
                  columnIndex === 0 ? '6254E7' : '172036',
                  columnIndex === 0,
                ),
              )
              .join('')}<a:extLst/></a:tr>`,
        )
        .join('')}
    </a:tbl></a:graphicData></a:graphic>
  </p:graphicFrame>`;
}

function slideXml(shapes: string, background = 'FFFFFF'): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${background}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    ${shapes}
  </p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}

const slide1 = slideXml(
  [
    shape({
      id: 2,
      name: 'Cobalt rail',
      x: 0,
      y: 0,
      width: 228600,
      height: 6858000,
      fill: '315EFB',
    }),
    textBox({
      id: 3,
      name: 'Deck label',
      x: 762000,
      y: 533400,
      width: 2971800,
      height: 426720,
      paragraphs: paragraph(
        run('BETTEROFFICE  /  OPEN OOXML', { size: 900, color: 'FFFFFF', bold: true }),
        'ctr',
      ),
      fill: '101828',
      anchor: 'ctr',
      radius: true,
    }),
    textBox({
      id: 4,
      name: 'Title',
      x: 762000,
      y: 1219200,
      width: 6096000,
      height: 1981200,
      paragraphs: [
        paragraph(run('Office files,', { size: 4400, color: '101828', bold: true })),
        paragraph(run('without the office.', { size: 4400, color: '315EFB', bold: true })),
      ].join(''),
    }),
    textBox({
      id: 5,
      name: 'Subtitle',
      x: 762000,
      y: 3276600,
      width: 5715000,
      height: 1066800,
      paragraphs: paragraph(
        run('A Rust-native editing engine for DOCX, XLSX, and PPTX. Local-first, collaborative, and entirely yours.', {
          size: 1700,
          color: '475467',
        }),
      ),
    }),
    textBox({
      id: 6,
      name: 'WASM label',
      x: 762000,
      y: 4572000,
      width: 1188720,
      height: 487680,
      paragraphs: paragraph(run('WASM', { size: 900, color: '101828', bold: true }), 'ctr'),
      fill: 'C8F56A',
      anchor: 'ctr',
      radius: true,
    }),
    textBox({
      id: 7,
      name: 'CRDT label',
      x: 2096760,
      y: 4572000,
      width: 1188720,
      height: 487680,
      paragraphs: paragraph(run('CRDT', { size: 900, color: '101828', bold: true }), 'ctr'),
      fill: 'FFD166',
      anchor: 'ctr',
      radius: true,
    }),
    textBox({
      id: 8,
      name: 'No uploads label',
      x: 3431520,
      y: 4572000,
      width: 1752600,
      height: 487680,
      paragraphs: paragraph(
        run('NO UPLOADS', { size: 900, color: 'FFFFFF', bold: true }),
        'ctr',
      ),
      fill: '101828',
      anchor: 'ctr',
      radius: true,
    }),
    textBox({
      id: 9,
      name: 'Cover statement',
      x: 762000,
      y: 5334000,
      width: 5334000,
      height: 609600,
      paragraphs: paragraph(
        run('Parse. Edit. Collaborate. Render. All in one resident model.', {
          size: 1100,
          color: '667085',
          bold: true,
        }),
      ),
    }),
    editorMockup(),
    fixturePicture(),
  ].join(''),
  'F8F7F2',
);

const slide2 = slideXml(
  [
    shape({
      id: 2,
      name: 'Title accent',
      x: 762000,
      y: 533400,
      width: 99060,
      height: 457200,
      fill: '315EFB',
      geometry: 'roundRect',
    }),
    textBox({
      id: 3,
      name: 'Section label',
      x: 990600,
      y: 510540,
      width: 2514600,
      height: 381000,
      paragraphs: paragraph(
        run('ONE NATIVE STACK', { size: 950, color: '315EFB', bold: true }),
      ),
    }),
    textBox({
      id: 4,
      name: 'Title',
      x: 762000,
      y: 990600,
      width: 10668000,
      height: 762000,
      paragraphs: paragraph(
        run('Three formats. One resident engine.', { size: 3200, color: '101828', bold: true }),
      ),
    }),
    textBox({
      id: 5,
      name: 'Subtitle',
      x: 762000,
      y: 1676400,
      width: 9144000,
      height: 533400,
      paragraphs: paragraph(
        run('No conversions. No round-trips. Every edit stays inside the file model.', {
          size: 1400,
          color: '667085',
        }),
      ),
    }),
    shape({
      id: 6,
      name: 'Format connector',
      x: 2133600,
      y: 3261360,
      width: 7894320,
      height: 22860,
      fill: 'D0D5DD',
    }),
    textBox({
      id: 7,
      name: 'DOCX card',
      x: 762000,
      y: 2286000,
      width: 3200400,
      height: 990600,
      paragraphs: [
        paragraph(run('DOCX', { size: 1800, color: '315EFB', bold: true })),
        paragraph(run('flow layout', { size: 1050, color: '475467', bold: true })),
      ].join(''),
      fill: 'EEF2FF',
      outline: 'C7D1FF',
      radius: true,
      anchor: 'ctr',
    }),
    textBox({
      id: 8,
      name: 'XLSX card',
      x: 4495800,
      y: 2286000,
      width: 3200400,
      height: 990600,
      paragraphs: [
        paragraph(run('XLSX', { size: 1800, color: '16825D', bold: true })),
        paragraph(run('calculation grid', { size: 1050, color: '475467', bold: true })),
      ].join(''),
      fill: 'E7F8EF',
      outline: 'B7E7D0',
      radius: true,
      anchor: 'ctr',
    }),
    textBox({
      id: 9,
      name: 'PPTX card',
      x: 8229600,
      y: 2286000,
      width: 3200400,
      height: 990600,
      paragraphs: [
        paragraph(run('PPTX', { size: 1800, color: 'D94F35', bold: true })),
        paragraph(run('slide canvas', { size: 1050, color: '475467', bold: true })),
      ].join(''),
      fill: 'FFF0EA',
      outline: 'FFD0C4',
      radius: true,
      anchor: 'ctr',
    }),
    shape({
      id: 10,
      name: 'Engine panel',
      x: 762000,
      y: 3657600,
      width: 10668000,
      height: 2438400,
      fill: '101828',
      geometry: 'roundRect',
    }),
    textBox({
      id: 11,
      name: 'Engine label',
      x: 1066800,
      y: 3886200,
      width: 3505200,
      height: 381000,
      paragraphs: paragraph(
        run('RUST CORE  /  ONE MODEL', { size: 950, color: 'C8F56A', bold: true }),
      ),
    }),
    textBox({
      id: 12,
      name: 'Parse contract',
      x: 1066800,
      y: 4495800,
      width: 2057400,
      height: 990600,
      paragraphs: [
        paragraph(run('PARSE', { size: 1500, color: 'C8F56A', bold: true })),
        paragraph(run('guarded OPC parts', { size: 1050, color: 'D0D5DD' })),
      ].join(''),
    }),
    textBox({
      id: 13,
      name: 'Model contract',
      x: 3657600,
      y: 4495800,
      width: 2057400,
      height: 990600,
      paragraphs: [
        paragraph(run('MODEL', { size: 1500, color: '8AB4FF', bold: true })),
        paragraph(run('local CRDT state', { size: 1050, color: 'D0D5DD' })),
      ].join(''),
    }),
    textBox({
      id: 14,
      name: 'Layout contract',
      x: 6248400,
      y: 4495800,
      width: 2057400,
      height: 990600,
      paragraphs: [
        paragraph(run('LAYOUT', { size: 1500, color: 'FF8066', bold: true })),
        paragraph(run('shaped text + pixels', { size: 1050, color: 'D0D5DD' })),
      ].join(''),
    }),
    textBox({
      id: 15,
      name: 'Emit contract',
      x: 8839200,
      y: 4495800,
      width: 2057400,
      height: 990600,
      paragraphs: [
        paragraph(run('EMIT', { size: 1500, color: 'FFD166', bold: true })),
        paragraph(run('OOXML bytes retained', { size: 1050, color: 'D0D5DD' })),
      ].join(''),
    }),
    shape({ id: 16, name: 'Panel divider one', x: 3352800, y: 4381500, width: 15240, height: 1219200, fill: '344054' }),
    shape({ id: 17, name: 'Panel divider two', x: 5943600, y: 4381500, width: 15240, height: 1219200, fill: '344054' }),
    shape({ id: 18, name: 'Panel divider three', x: 8534400, y: 4381500, width: 15240, height: 1219200, fill: '344054' }),
    architectureTable(),
  ].join(''),
  'FCFCFA',
);

const slide3 = slideXml(
  [
    shape({
      id: 2,
      name: 'Title accent',
      x: 762000,
      y: 533400,
      width: 99060,
      height: 457200,
      fill: '315EFB',
      geometry: 'roundRect',
    }),
    textBox({
      id: 3,
      name: 'Section label',
      x: 990600,
      y: 510540,
      width: 2514600,
      height: 381000,
      paragraphs: paragraph(
        run('LIVE COLLABORATION', { size: 950, color: '315EFB', bold: true }),
      ),
    }),
    textBox({
      id: 4,
      name: 'Title',
      x: 762000,
      y: 1066800,
      width: 5334000,
      height: 1600200,
      paragraphs: [
        paragraph(run('Multiplayer,', { size: 3600, color: '101828', bold: true })),
        paragraph(run('minus the server tax.', { size: 3600, color: '315EFB', bold: true })),
      ].join(''),
    }),
    textBox({
      id: 5,
      name: 'Subtitle',
      x: 762000,
      y: 2667000,
      width: 5105400,
      height: 762000,
      paragraphs: paragraph(
        run('Every keystroke lands locally first. Peers exchange only updates, then converge on the same file.', {
          size: 1500,
          color: '475467',
        }),
      ),
    }),
    textBox({
      id: 6,
      name: 'Local first step',
      x: 762000,
      y: 3543300,
      width: 4953000,
      height: 685800,
      paragraphs: [
        paragraph(
          run('01', { size: 1200, color: '315EFB', bold: true }) +
            run('   Edit locally', { size: 1500, color: '101828', bold: true }),
        ),
        paragraph(run('Resident model. Instant feedback.', { size: 1000, color: '667085' })),
      ].join(''),
      fill: 'FFFFFF',
      outline: 'D8E0FF',
      radius: true,
      anchor: 'ctr',
    }),
    textBox({
      id: 7,
      name: 'Exchange updates step',
      x: 762000,
      y: 4381500,
      width: 4953000,
      height: 685800,
      paragraphs: [
        paragraph(
          run('02', { size: 1200, color: 'D94F35', bold: true }) +
            run('   Exchange updates', { size: 1500, color: '101828', bold: true }),
        ),
        paragraph(run('Compact binary deltas. Nothing else.', { size: 1000, color: '667085' })),
      ].join(''),
      fill: 'FFFFFF',
      outline: 'FFD0C4',
      radius: true,
      anchor: 'ctr',
    }),
    textBox({
      id: 8,
      name: 'Converge step',
      x: 762000,
      y: 5219700,
      width: 4953000,
      height: 685800,
      paragraphs: [
        paragraph(
          run('03', { size: 1200, color: '16825D', bold: true }) +
            run('   Converge everywhere', { size: 1500, color: '101828', bold: true }),
        ),
        paragraph(run('One state across every peer.', { size: 1000, color: '667085' })),
      ].join(''),
      fill: 'FFFFFF',
      outline: 'B7E7D0',
      radius: true,
      anchor: 'ctr',
    }),
    shape({
      id: 20,
      name: 'Live session panel',
      x: 6400800,
      y: 609600,
      width: 5029200,
      height: 5486400,
      fill: 'FFFFFF',
      outline: 'C7D1FF',
      geometry: 'roundRect',
    }),
    textBox({
      id: 21,
      name: 'Live session title',
      x: 6781800,
      y: 838200,
      width: 2362200,
      height: 457200,
      paragraphs: paragraph(
        run('Q3 launch deck', { size: 1200, color: '101828', bold: true }),
      ),
      anchor: 'ctr',
    }),
    shape({ id: 22, name: 'Peer one', x: 10058400, y: 838200, width: 342900, height: 342900, fill: '315EFB', geometry: 'ellipse' }),
    shape({ id: 23, name: 'Peer two', x: 10439400, y: 838200, width: 342900, height: 342900, fill: 'FF8066', geometry: 'ellipse' }),
    shape({ id: 24, name: 'Peer three', x: 10820400, y: 838200, width: 342900, height: 342900, fill: '48C78E', geometry: 'ellipse' }),
    shape({
      id: 25,
      name: 'Shared canvas',
      x: 6781800,
      y: 1524000,
      width: 4267200,
      height: 3429000,
      fill: 'F8FAFC',
      outline: 'E4E7EC',
      geometry: 'roundRect',
    }),
    textBox({
      id: 26,
      name: 'Shared canvas heading',
      x: 7086600,
      y: 1828800,
      width: 3048000,
      height: 533400,
      paragraphs: paragraph(
        run('Launch narrative', { size: 1800, color: '101828', bold: true }),
      ),
    }),
    shape({ id: 27, name: 'Canvas line one', x: 7086600, y: 2590800, width: 2819400, height: 91440, fill: 'D0D5DD', geometry: 'roundRect' }),
    shape({ id: 28, name: 'Canvas line two', x: 7086600, y: 2819400, width: 2133600, height: 91440, fill: 'D0D5DD', geometry: 'roundRect' }),
    shape({ id: 29, name: 'Canvas card one', x: 7086600, y: 3352800, width: 1524000, height: 990600, fill: 'E6EBFF', geometry: 'roundRect' }),
    shape({ id: 30, name: 'Canvas card two', x: 8839200, y: 3352800, width: 1905000, height: 990600, fill: 'E7F8EF', geometry: 'roundRect' }),
    textBox({ id: 31, name: 'You cursor', x: 6934200, y: 3048000, width: 762000, height: 381000, paragraphs: paragraph(run('YOU', { size: 800, color: 'FFFFFF', bold: true }), 'ctr'), fill: '315EFB', anchor: 'ctr', radius: true }),
    shape({ id: 32, name: 'You cursor line', x: 7315200, y: 3390900, width: 15240, height: 762000, fill: '315EFB' }),
    textBox({ id: 33, name: 'Maya cursor', x: 9677400, y: 2286000, width: 838200, height: 381000, paragraphs: paragraph(run('MAYA', { size: 800, color: '101828', bold: true }), 'ctr'), fill: 'FF8066', anchor: 'ctr', radius: true }),
    shape({ id: 34, name: 'Maya cursor line', x: 10096500, y: 2628900, width: 15240, height: 914400, fill: 'FF8066' }),
    textBox({
      id: 35,
      name: 'Session status',
      x: 6781800,
      y: 5181600,
      width: 4267200,
      height: 533400,
      paragraphs: paragraph(
        run('3 peers  ·  synced  ·  local-first', { size: 950, color: '101828', bold: true }),
        'ctr',
      ),
      fill: 'C8F56A',
      anchor: 'ctr',
      radius: true,
    }),
  ].join(''),
  'EFF3FF',
);

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const rootRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

const presentation = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;

const presentationRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>
</Relationships>`;

const master = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="BetterOffice Master"><p:bg><p:bgPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    ${shape({ id: 2, name: 'Footer rule', x: 762000, y: 6248400, width: 10668000, height: 15240, fill: 'D0D5DD' })}
    ${textBox({ id: 3, name: 'Brand', x: 762000, y: 6324600, width: 3048000, height: 381000, paragraphs: paragraph(run('betteroffice', { size: 900, color: '101828', bold: true })) })}
    ${textBox({ id: 4, name: 'Open OOXML label', x: 8382000, y: 6324600, width: 3048000, height: 381000, paragraphs: paragraph(run('OPEN OOXML  /  LOCAL-FIRST', { size: 800, color: '667085', bold: true }), 'r') })}
  </p:spTree></p:cSld>
  <p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink" bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="3200" b="1"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr marL="342900" indent="-342900"><a:buChar char="•"/><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles>
</p:sldMaster>`;

const masterRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

const layout = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" matchingName="BetterOffice Blank" preserve="1">
  <p:cSld name="BetterOffice Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;

const layoutRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`;

function slideRelationships(withImage = false): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  ${withImage ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/betteroffice-mark.png"/>' : ''}
</Relationships>`;
}

const theme = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="BetterOffice"><a:themeElements>
  <a:clrScheme name="BetterOffice"><a:dk1><a:srgbClr val="101828"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1D2939"/></a:dk2><a:lt2><a:srgbClr val="F8F7F2"/></a:lt2><a:accent1><a:srgbClr val="315EFB"/></a:accent1><a:accent2><a:srgbClr val="16825D"/></a:accent2><a:accent3><a:srgbClr val="D94F35"/></a:accent3><a:accent4><a:srgbClr val="C8F56A"/></a:accent4><a:accent5><a:srgbClr val="8AB4FF"/></a:accent5><a:accent6><a:srgbClr val="667085"/></a:accent6><a:hlink><a:srgbClr val="315EFB"/></a:hlink><a:folHlink><a:srgbClr val="5D6DCB"/></a:folHlink></a:clrScheme>
  <a:fontScheme name="BetterOffice"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="BetterOffice">
    <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
    <a:lnStyleLst><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
    <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
    <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
  </a:fmtScheme>
</a:themeElements></a:theme>`;

const coreProperties = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>BetterOffice — office files without the office</dc:title><dc:creator>BetterOffice</dc:creator><cp:lastModifiedBy>BetterOffice</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>`;

const appProperties = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>BetterOffice</Application><PresentationFormat>Widescreen</PresentationFormat><Slides>3</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><Company>BetterOffice</Company></Properties>`;

const brandMark = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=',
  'base64',
);

const archive = new JSZip();
const options = { date: zipDate, createFolders: false };
const textParts: Record<string, string> = {
  '[Content_Types].xml': contentTypes,
  '_rels/.rels': rootRelationships,
  'docProps/core.xml': coreProperties,
  'docProps/app.xml': appProperties,
  'ppt/presentation.xml': presentation,
  'ppt/_rels/presentation.xml.rels': presentationRelationships,
  'ppt/slideMasters/slideMaster1.xml': master,
  'ppt/slideMasters/_rels/slideMaster1.xml.rels': masterRelationships,
  'ppt/slideLayouts/slideLayout1.xml': layout,
  'ppt/slideLayouts/_rels/slideLayout1.xml.rels': layoutRelationships,
  'ppt/slides/slide1.xml': slide1,
  'ppt/slides/_rels/slide1.xml.rels': slideRelationships(true),
  'ppt/slides/slide2.xml': slide2,
  'ppt/slides/_rels/slide2.xml.rels': slideRelationships(),
  'ppt/slides/slide3.xml': slide3,
  'ppt/slides/_rels/slide3.xml.rels': slideRelationships(),
  'ppt/theme/theme1.xml': theme,
};

for (const [partPath, contents] of Object.entries(textParts)) {
  archive.file(partPath, contents, options);
}
archive.file('ppt/media/betteroffice-mark.png', brandMark, options);

const buffer = await archive.generateAsync({
  type: 'nodebuffer',
  compression: 'DEFLATE',
  compressionOptions: { level: 9 },
});
fs.mkdirSync(path.dirname(output), { recursive: true });
fs.writeFileSync(output, buffer);
console.log(`Created ${output} (${buffer.length} bytes)`);
