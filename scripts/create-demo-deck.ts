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
  return `<a:r><a:rPr lang="en-US" sz="${size}"${bold}${italic}><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Inter"/></a:rPr><a:t>${escapeXml(text)}</a:t></a:r>`;
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

function picture(): string {
  return `<p:pic>
    <p:nvPicPr><p:cNvPr id="5" name="BetterOffice mark" descr="BetterOffice brand mark"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
    <p:blipFill><a:blip r:embed="rId2"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
    <p:spPr><a:xfrm><a:off x="10058400" y="609600"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:ln><a:noFill/></a:ln></p:spPr>
  </p:pic>`;
}

function groupedPipeline(): string {
  const child = (id: number, x: number, width: number, label: string, color: string) =>
    textBox({
      id,
      name: `${label} stage`,
      x,
      y: 0,
      width,
      height: 914400,
      paragraphs: paragraph(run(label, { size: 1500, color: 'FFFFFF', bold: true }), 'ctr'),
      fill: color,
      anchor: 'ctr',
      radius: true,
    });
  return `<p:grpSp>
    <p:nvGrpSpPr><p:cNvPr id="6" name="Native pipeline"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="914400" y="4876800"/><a:ext cx="10363200" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="10363200" cy="914400"/></a:xfrm></p:grpSpPr>
    ${child(7, 0, 2377440, 'Parse', '6254E7')}
    ${child(8, 2667000, 2377440, 'Collaborate', '4A3FB5')}
    ${child(9, 5334000, 2377440, 'Layout', '2B8A78')}
    ${child(10, 8001000, 2362200, 'Paint', '172036')}
  </p:grpSp>`;
}

function tableCell(text: string, fill: string, color: string, bold = false): string {
  return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraph(
    run(text, { size: 1500, color, bold }),
    'ctr',
  )}</a:txBody><a:tcPr><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR><a:lnT><a:noFill/></a:lnT><a:lnB><a:noFill/></a:lnB></a:tcPr></a:tc>`;
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
    <p:nvGraphicFramePr><p:cNvPr id="3" name="Architecture table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
    <p:xfrm><a:off x="762000" y="1676400"/><a:ext cx="10668000" cy="3657600"/></p:xfrm>
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
    textBox({
      id: 2,
      name: 'Title',
      x: 914400,
      y: 1066800,
      width: 9144000,
      height: 1524000,
      paragraphs: paragraph(run('PPTX, the third format', { size: 4000, color: 'FFFFFF', bold: true })),
    }),
    textBox({
      id: 3,
      name: 'Subtitle',
      x: 914400,
      y: 2667000,
      width: 8382000,
      height: 1219200,
      paragraphs: paragraph(
        run('A Rust engine from guarded package parts to every collaborative pixel.', {
          size: 2000,
          color: 'D9D5FF',
        }),
      ),
    }),
    textBox({
      id: 4,
      name: 'Architecture note',
      x: 914400,
      y: 3733800,
      width: 8382000,
      height: 762000,
      paragraphs: paragraph(
        run('Parse · edit · shape · lay out · emit', { size: 1600, color: 'FFFFFF', bold: true }),
      ),
      fill: '4A3FB5',
      radius: true,
      anchor: 'ctr',
    }),
    picture(),
    groupedPipeline(),
  ].join(''),
  '241E52',
);

const slide2 = slideXml(
  [
    textBox({
      id: 2,
      name: 'Title',
      x: 762000,
      y: 457200,
      width: 10668000,
      height: 914400,
      paragraphs: paragraph(run('One native pipeline, four durable contracts', { size: 3000, color: '172036', bold: true })),
    }),
    architectureTable(),
    textBox({
      id: 4,
      name: 'Footer statement',
      x: 762000,
      y: 5638800,
      width: 10668000,
      height: 609600,
      paragraphs: paragraph(
        run('TypeScript loads the module, decodes the boundary, and replays the display list.', {
          size: 1500,
          color: '5C6478',
          italic: true,
        }),
        'ctr',
      ),
    }),
  ].join(''),
);

const slide3 = slideXml(
  [
    textBox({
      id: 2,
      name: 'Title',
      x: 762000,
      y: 457200,
      width: 10668000,
      height: 914400,
      paragraphs: paragraph(run('Collaboration is a model property', { size: 3000, color: '172036', bold: true })),
    }),
    textBox({
      id: 3,
      name: 'Local first',
      x: 914400,
      y: 1676400,
      width: 3200400,
      height: 2590800,
      paragraphs: [
        paragraph(run('01  Local first', { size: 2000, color: '6254E7', bold: true })),
        paragraph(run('Every keystroke edits the resident deck model.', { size: 1700 }), 'l', 0, '•'),
        paragraph(run('Undo is scoped to the local origin.', { size: 1700 }), 'l', 0, '•'),
      ].join(''),
      fill: 'FFFFFF',
      outline: 'D9D5FF',
      radius: true,
    }),
    textBox({
      id: 4,
      name: 'Exchange updates',
      x: 4495800,
      y: 1676400,
      width: 3200400,
      height: 2590800,
      paragraphs: [
        paragraph(run('02  Exchange updates', { size: 2000, color: '2B8A78', bold: true })),
        paragraph(run('Replicas send compact binary updates.', { size: 1700 }), 'l', 0, '•'),
        paragraph(run('State vectors request only missing work.', { size: 1700 }), 'l', 0, '•'),
      ].join(''),
      fill: 'FFFFFF',
      outline: 'BDE8DF',
      radius: true,
    }),
    textBox({
      id: 5,
      name: 'Converge',
      x: 8077200,
      y: 1676400,
      width: 3200400,
      height: 2590800,
      paragraphs: [
        paragraph(run('03  Converge', { size: 2000, color: 'C65A2E', bold: true })),
        paragraph(run('Slide order and text settle together.', { size: 1700 }), 'l', 0, '•'),
        paragraph(run('Rust reflows the same result everywhere.', { size: 1700 }), 'l', 0, '•'),
      ].join(''),
      fill: 'FFFFFF',
      outline: 'F2D0C3',
      radius: true,
    }),
    textBox({
      id: 6,
      name: 'Closing statement',
      x: 1524000,
      y: 4876800,
      width: 9144000,
      height: 1066800,
      paragraphs: paragraph(
        run('The deck is collaborative before the first toolbar button exists.', {
          size: 2200,
          color: 'FFFFFF',
          bold: true,
        }),
        'ctr',
      ),
      fill: '6254E7',
      radius: true,
      anchor: 'ctr',
    }),
  ].join(''),
  'F7F8FC',
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
    ${textBox({ id: 2, name: 'Brand', x: 762000, y: 6400800, width: 10668000, height: 228600, paragraphs: paragraph(run('betteroffice', { size: 900, color: '6254E7', bold: true })) })}
  </p:spTree></p:cSld>
  <p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rId1"/></p:sldLayoutIdLst>
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
  <a:clrScheme name="BetterOffice"><a:dk1><a:srgbClr val="172036"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="241E52"/></a:dk2><a:lt2><a:srgbClr val="F7F8FC"/></a:lt2><a:accent1><a:srgbClr val="6254E7"/></a:accent1><a:accent2><a:srgbClr val="2B8A78"/></a:accent2><a:accent3><a:srgbClr val="C65A2E"/></a:accent3><a:accent4><a:srgbClr val="4A3FB5"/></a:accent4><a:accent5><a:srgbClr val="D9D5FF"/></a:accent5><a:accent6><a:srgbClr val="5C6478"/></a:accent6><a:hlink><a:srgbClr val="6254E7"/></a:hlink><a:folHlink><a:srgbClr val="4A3FB5"/></a:folHlink></a:clrScheme>
  <a:fontScheme name="BetterOffice"><a:majorFont><a:latin typeface="Inter"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Inter"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="BetterOffice"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
</a:themeElements></a:theme>`;

const coreProperties = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>BetterOffice — Rust-native collaborative slides</dc:title><dc:creator>BetterOffice</dc:creator><cp:lastModifiedBy>BetterOffice</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>`;

const appProperties = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>BetterOffice</Application><PresentationFormat>Widescreen</PresentationFormat><Slides>3</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><Company>BetterOffice</Company></Properties>`;

const brandMark = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
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
