import JSZip from 'jszip';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const output = path.join(root, 'crates/vsdx-parse/tests/fixtures/foundation.vsdx');
const zipDate = new Date('2026-01-01T00:00:00Z');
const ns = "xmlns='http://schemas.microsoft.com/office/visio/2012/main'";
const parts: Record<string, string> = {
  '[Content_Types].xml': "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='xml' ContentType='application/xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>",
  '_rels/.rels': "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>",
  'visio/document.xml': `<VisioDocument ${ns}><DocumentSettings/><Colors><ColorEntry IX='0' RGB='#FFFFFF'/></Colors><FaceNames><FaceName ID='0' Name='Calibri'/></FaceNames><StyleSheets><StyleSheet ID='0' NameU='Normal'><Cell N='LineColor' V='0'/><Cell N='FillForegnd' V='1'/></StyleSheet></StyleSheets><DocumentSheet><Cell N='PageWidth' V='8.5'/><UnknownSheet Flag='yes'>sheet text<Child Value='nested'/></UnknownSheet><Cell N='PageHeight' V='11'/></DocumentSheet></VisioDocument>`,
  'visio/_rels/document.xml.rels': "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/pages' Target='pages/pages.xml'/><Relationship Id='rId2' Type='http://schemas.microsoft.com/visio/2010/relationships/masters' Target='masters/masters.xml'/><Relationship Id='rId3' Type='http://schemas.microsoft.com/visio/2010/relationships/theme' Target='theme/theme1.xml'/><Relationship Id='rId4' Type='http://schemas.microsoft.com/visio/2010/relationships/windows' Target='windows.xml'/></Relationships>",
  'visio/pages/pages.xml': `<Pages ${ns}><Page ID='1' NameU='Page-1' Name='Page-1' r:id='rId1' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><PageSheet><Cell N='PageWidth' F='8.5' V='8.5'/><Trigger N='RecalcColor'><RefBy ID='0' T='Page'/></Trigger></PageSheet></Page></Pages>`,
  'visio/pages/_rels/pages.xml.rels': "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/page' Target='page1.xml'/></Relationships>",
  'visio/pages/page1.xml': `<PageContents ${ns}><Shapes><Shape ID='1' NameU='Process' Type='Shape' Mystery='yes'><Cell N='FOnly' F='Width*2'/><Section N='Geometry'><Row T='RelMoveTo' LocalName='Start'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row><Row T='RelLineTo' N='LineTo'><Cell N='X' V='1'/><Cell N='Y' V='1'/></Row><Row IX='2' Del='1'/><UnknownRowChild Flag='yes'>row text<Child Value='nested'/></UnknownRowChild></Section><Cell N='VOnly' V='5'/><UnknownShape Flag='yes'>shape text<Child Value='nested'/></UnknownShape><Data1 Value='opaque'/><Cell N='Both' F='Height*2' V='2'/><ForeignData ForeignType='Bitmap'/><Section N='User' UnknownSection='kept'><Row N='visVersion' LocalName='Version'><Cell N='Value' V='15'/></Row><UnknownSectionChild Flag='yes'>section text<Child Value='nested'/></UnknownSectionChild><Row IX='3' T='UnknownRow' Weird='kept'><Cell N='UnknownCell' UnknownAttr='kept' V='x'/></Row></Section><Cell N='LineWeight' V='0.01' Del='1'/><Section N='Scratch' Del='1'/><Text> A<cp IX='1'/>B<pp IX='2'/><tp IX='3'/><fld IX='0'/> C </Text></Shape></Shapes><Connects><Connect FromSheet='1' FromCell='BeginX' FromPart='9' ToSheet='1' ToCell='PinX' ToPart='3'/><UnknownConnect Flag='yes'><Child Value='nested'/>connect text</UnknownConnect></Connects></PageContents>`,
  'visio/masters/masters.xml': `<Masters ${ns}><Master ID='1' NameU='Master-1' r:id='rId1' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><PageSheet><Cell N='PageHeight' V='11'/></PageSheet></Master></Masters>`,
  'visio/masters/_rels/masters.xml.rels': "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/master' Target='master1.xml'/></Relationships>",
  'visio/masters/master1.xml': `<MasterContents ${ns}><Shapes/></MasterContents>`,
  'visio/theme/theme1.xml': "<a:theme xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' name='Office Theme'/>",
  'visio/windows.xml': `<Windows ${ns}><Window ID='0'/></Windows>`,
};

const xform = (width: number, height: number, pinX: number, pinY: number, locPinX: number, locPinY: number, angle: number, flipX: number, flipY: number) =>
  `<Cell N='Width' V='${width}'/><Cell N='Height' V='${height}'/><Cell N='PinX' V='${pinX}'/><Cell N='PinY' V='${pinY}'/><Cell N='LocPinX' V='${locPinX}'/><Cell N='LocPinY' V='${locPinY}'/><Cell N='Angle' V='${angle}'/><Cell N='FlipX' V='${flipX}'/><Cell N='FlipY' V='${flipY}'/>`;
const rect = `<Section N='Geometry'><Row IX='0' T='MoveTo'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row><Row IX='1' T='LineTo'><Cell N='X' V='1'/><Cell N='Y' V='0'/></Row><Row IX='2' T='LineTo'><Cell N='X' V='1'/><Cell N='Y' V='1'/></Row><Row IX='3' T='LineTo'><Cell N='X' V='0'/><Cell N='Y' V='1'/></Row><Row IX='4' T='Close'/></Section>`;

async function writeTestFixtures(): Promise<void> {
  const zip = new JSZip();
  for (const [name, contents] of Object.entries(parts)) zip.file(name, contents, { date: zipDate, createFolders: false });
  fs.writeFileSync(output, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', platform: 'DOS' }));

  const nestedOutput = path.join(root, 'crates/vsdx-parse/tests/fixtures/nested-groups.vsdx');
  const nestedParts: Record<string, string | Uint8Array> = {
    ...parts,
    'visio/pages/page1.xml': `<PageContents ${ns}><Shapes><Shape ID='1' Type='Group'>${xform(6, 4, 10, 10, 1, 0.5, 0.5235987755982988, 1, 0)}<Shapes><Shape ID='2' Type='Group'>${xform(3, 5, 2, 1, 0.25, 0.75, -0.7853981633974483, 0, 1)}<Shapes><Shape ID='3' Type='Shape'>${xform(1, 1, 0, 0, 0, 0, 0, 0, 0)}${rect}<Text>deep\nvector</Text></Shape><Shape ID='4' Type='Shape'>${xform(1, 2, 2, 1, 0, 0, 0, 0, 0)}<ForeignData ForeignType='Bitmap'><Rel r:id='rIdImage' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'/></ForeignData></Shape><Shape ID='5' Type='Shape'>${xform(1, 1, 1, 3, 0, 0, 0, 0, 0)}<Section N='Geometry'><Row T='EllipticalArcTo'><Cell N='X' V='1'/></Row></Section></Shape></Shapes></Shape></Shapes></Shape><Shape ID='0' Type='Shape'><Text>page sibling</Text></Shape></Shapes><Connects><Connect FromSheet='3' ToSheet='4'/></Connects></PageContents>`,
    'visio/pages/_rels/page1.xml.rels': "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rIdImage' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='../media/image1.png'/></Relationships>",
    'visio/media/image1.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
  };
  const nestedZip = new JSZip();
  for (const [name, contents] of Object.entries(nestedParts)) nestedZip.file(name, contents, { date: zipDate, createFolders: false });
  fs.writeFileSync(nestedOutput, await nestedZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', platform: 'DOS' }));

  const groupedGlue = new JSZip();
  for (const [part, contents] of Object.entries({
    ...parts,
    'visio/pages/page1.xml': `<PageContents ${ns}><Shapes><Shape ID='1' Type='Shape'>${xform(1, 1, 0, 0, 0, 0, 0, 0, 0)}<Cell N='OneD' V='1'/><Cell N='BeginX' V='0'/><Cell N='BeginY' V='0'/><Cell N='EndX' V='1'/><Cell N='EndY' V='0'/></Shape><Shape ID='10' Type='Group'>${xform(2, 2, 10, 10, 0, 0, 1.5707963267948966, 0, 0)}<Shapes><Shape ID='11' Type='Shape'>${xform(1, 1, 0, 0, 0, 0, 0, 0, 0)}<Section N='Connection'><Row IX='0' T='Connection'><Cell N='X' V='0.5'/><Cell N='Y' V='0.5'/></Row></Section></Shape></Shapes></Shape><Shape ID='20' Type='Group'>${xform(6, 2, 20, 10, 0, 0, 0, 0, 0)}<Shapes><Shape ID='21' Type='Shape'>${xform(3, 1, 0, 0, 0, 0, 0, 0, 0)}<Section N='Connection'><Row IX='0' T='Connection'><Cell N='X' V='0.5'/><Cell N='Y' V='0.5'/></Row></Section></Shape></Shapes></Shape><Shape ID='30' Type='Group'>${xform(4, 4, 30, 0, 0, 0, 0, 0, 0)}<Shapes><Shape ID='31' Type='Group'>${xform(2, 2, 1, 1, 0, 0, 0, 0, 0)}<Shapes><Shape ID='32' Type='Shape'>${xform(1, 1, 0, 0, 0, 0, 0, 0, 0)}<Section N='Connection'><Row IX='0' T='Connection'><Cell N='X' V='0.5'/><Cell N='Y' V='0.5'/></Row></Section></Shape></Shapes></Shape></Shapes></Shape></Shapes><Connects><Connect FromSheet='1' FromCell='BeginX' FromPart='9' ToSheet='11' ToCell='Connections.X1' ToPart='100'/><Connect FromSheet='1' FromCell='EndX' FromPart='9' ToSheet='21' ToCell='Connections.X1' ToPart='100'/><Connect FromSheet='1' FromCell='BeginY' FromPart='9' ToSheet='32' ToCell='Connections.X1' ToPart='100'/><Connect FromSheet='1' FromCell='BeginX' ToSheet='11' ToCell='PinX'/><Connect FromSheet='1' FromCell='EndX' ToSheet='21' ToCell='PinX'/><Connect FromSheet='1' FromCell='BeginY' ToSheet='32' ToCell='PinX'/></Connects></PageContents>`,
  })) groupedGlue.file(part, contents, { date: zipDate, createFolders: false });
  fs.writeFileSync(
    path.join(root, 'crates/vsdx-parse/tests/fixtures/grouped-glue.vsdx'),
    await groupedGlue.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', platform: 'DOS' }),
  );

  for (const [name, rows] of [
    ['geometry-anonymous-rows', "<Row T='MoveTo'><Cell N='X' V='1'/><Cell N='Y' V='2'/></Row><Row T='LineTo'><Cell N='X' V='3'/><Cell N='Y' V='4'/></Row>"],
    ['geometry-duplicate-ix-rows', "<Row IX='0' T='MoveTo'><Cell N='X' V='1'/><Cell N='Y' V='2'/></Row><Row IX='0' T='LineTo'><Cell N='X' V='3'/><Cell N='Y' V='4'/></Row>"],
  ] as const) {
    const fixture = new JSZip();
    for (const [part, contents] of Object.entries({
      ...parts,
      'visio/pages/page1.xml': `<PageContents ${ns}><Shapes><Shape ID='1' Type='Shape'><Section N='Geometry'>${rows}</Section></Shape></Shapes></PageContents>`,
    })) fixture.file(part, contents, { date: zipDate, createFolders: false });
    fs.writeFileSync(path.join(root, `crates/vsdx-parse/tests/fixtures/${name}.vsdx`), await fixture.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', platform: 'DOS' }));
  }

  const realizedGeometryRows = new JSZip();
  for (const [part, contents] of Object.entries({
    ...parts,
    'visio/pages/page1.xml': `<PageContents ${ns}><Shapes><Shape ID='1' Type='Shape'><Section N='Geometry'><Row IX='0' T='MoveTo'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row><Row T='PolylineTo' IX='1'><Cell N='X' V='2'/><Cell N='Y' V='1'/><Cell N='A' V='POLYLINE(1,1,1,0,2,1)'/></Row></Section></Shape><Shape ID='2' Type='Shape'><Section N='Geometry'><Row IX='0' T='InfiniteLine'><Cell N='X' V='-1'/><Cell N='Y' V='0'/><Cell N='A' V='1'/><Cell N='B' V='0.75'/></Row></Section></Shape></Shapes></PageContents>`,
  } as Record<string, string>)) realizedGeometryRows.file(part, contents, { date: zipDate, createFolders: false });
  fs.writeFileSync(
    path.join(root, 'crates/vsdx-parse/tests/fixtures/geometry-polyline-and-infinite-line.vsdx'),
    await realizedGeometryRows.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', platform: 'DOS' }),
  );

  const relativeGeometryRows = new JSZip();
  for (const [part, contents] of Object.entries({
    ...parts,
    'visio/pages/page1.xml': `<PageContents ${ns}><Shapes><Shape ID='1' Type='Shape'>${xform(4, 3, 7, 5, 2, 1.5, 0, 0, 0)}<Section N='Geometry'><Row T='MoveTo' IX='0'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row><Row T='RelLineTo' IX='2'><Cell N='X' V='1'/><Cell N='Y' V='0'/></Row><Row T='RelLineTo' IX='3'><Cell N='X' V='1'/><Cell N='Y' V='1'/></Row><Row T='RelLineTo' IX='4'><Cell N='X' V='0'/><Cell N='Y' V='1'/></Row></Section></Shape><Shape ID='2' Type='Shape'>${xform(4, 3, 14, 5, 2, 1.5, 0, 0, 0)}<Section N='Geometry'><Row T='MoveTo' IX='0'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row><Row T='RelLineTo' IX='1'><Cell N='X' V='1'/><Cell N='Y' V='0'/></Row><Row T='RelMoveTo' IX='2'><Cell N='X' V='0'/><Cell N='Y' V='1'/></Row><Row T='RelLineTo' IX='3'><Cell N='X' V='1'/><Cell N='Y' V='1'/></Row></Section></Shape></Shapes></PageContents>`,
  } as Record<string, string>)) relativeGeometryRows.file(part, contents, { date: zipDate, createFolders: false });
  fs.writeFileSync(
    path.join(root, 'crates/vsdx-parse/tests/fixtures/geometry-relative-rows.vsdx'),
    await relativeGeometryRows.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', platform: 'DOS' }),
  );

  const textAccounting = new JSZip();
  for (const [part, contents] of Object.entries({
    ...parts,
    'visio/document.xml': `<VisioDocument ${ns}><FaceNames><FaceName ID='0' Name='Calibri'/></FaceNames><StyleSheets><StyleSheet ID='1' NameU='Text'><Section N='Character'><Row IX='0'><Cell N='Font' V='0'/><Cell N='Size' V='0.25'/><Cell N='Color' V='RGB(1,2,3)'/></Row></Section></StyleSheet></StyleSheets><DocumentSheet><Cell N='PageWidth' V='8.5'/><Cell N='PageHeight' V='11'/></DocumentSheet></VisioDocument>`,
    'visio/pages/page1.xml': `<PageContents ${ns}><Shapes><Shape ID='1' Type='Shape'>${xform(1, 1, 1, 1, 0, 0, 0, 0, 0)}${rect}<Text><cp IX='0'/><pp IX='0'/></Text></Shape><Shape ID='2' Type='Shape'>${xform(1, 1, 2, 1, 0, 0, 0, 0, 0)}${rect}<Section N='Field'><Row IX='0'><Cell N='Value' V='field value'/></Row></Section><Text><fld IX='0'/></Text></Shape><Shape ID='3' Type='Shape' TextStyle='1'>${xform(1, 1, 3, 1, 0, 0, 0, 0, 0)}${rect}<Text><cp IX='0'/>style text</Text></Shape><Shape ID='4' Type='Shape' Master='1' MasterShape='10'>${xform(1, 1, 4, 1, 0, 0, 0, 0, 0)}${rect}</Shape></Shapes></PageContents>`,
    'visio/masters/master1.xml': `<MasterContents ${ns}><Shapes><Shape ID='10' Type='Shape'><Text>master text</Text></Shape></Shapes></MasterContents>`,
  } as Record<string, string>)) textAccounting.file(part, contents, { date: zipDate, createFolders: false });
  fs.writeFileSync(
    path.join(root, 'crates/vsdx-parse/tests/fixtures/text-accounting.vsdx'),
    await textAccounting.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', platform: 'DOS' }),
  );

}

type BenchmarkFixtureOptions = {
  shapeCount: number;
  groupDepth: number;
  glueDensity: number;
  inheritanceDepth: number;
  pageCount: number;
  textBytes: number;
};

const benchmarkFixtures: Record<string, BenchmarkFixtureOptions> = {
  'shape-heavy': { shapeCount: 1_200, groupDepth: 0, glueDensity: 0, inheritanceDepth: 1, pageCount: 1, textBytes: 16 },
  'nested-groups': { shapeCount: 160, groupDepth: 12, glueDensity: 0, inheritanceDepth: 1, pageCount: 1, textBytes: 16 },
  'dense-glue': { shapeCount: 600, groupDepth: 0, glueDensity: 3, inheritanceDepth: 1, pageCount: 1, textBytes: 16 },
  'deep-inheritance': { shapeCount: 500, groupDepth: 0, glueDensity: 0, inheritanceDepth: 12, pageCount: 1, textBytes: 16 },
  'many-pages': { shapeCount: 600, groupDepth: 0, glueDensity: 0, inheritanceDepth: 1, pageCount: 12, textBytes: 16 },
  'text-heavy': { shapeCount: 320, groupDepth: 0, glueDensity: 0, inheritanceDepth: 1, pageCount: 1, textBytes: 2_048 },
};

function benchmarkPackage(options: BenchmarkFixtureOptions): Record<string, string> {
  const pages = Math.max(1, options.pageCount);
  const inheritanceDepth = Math.max(1, options.inheritanceDepth);
  const pageRelationships = Array.from({ length: pages }, (_, index) =>
    `<Relationship Id='rId${index + 1}' Type='http://schemas.microsoft.com/visio/2010/relationships/page' Target='page${index + 1}.xml'/>`,
  ).join('');
  const pageCatalog = Array.from({ length: pages }, (_, index) =>
    `<Page ID='${index + 1}' NameU='Page-${index + 1}' Name='Page-${index + 1}' r:id='rId${index + 1}' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><PageSheet><Cell N='PageWidth' V='8.5'/><Cell N='PageHeight' V='11'/></PageSheet></Page>`,
  ).join('');
  const masterCatalog = Array.from({ length: inheritanceDepth }, (_, index) =>
    `<Master ID='${index + 1}' NameU='Master-${index + 1}' r:id='rId${index + 1}' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'/>`,
  ).join('');
  const masterRelationships = Array.from({ length: inheritanceDepth }, (_, index) =>
    `<Relationship Id='rId${index + 1}' Type='http://schemas.microsoft.com/visio/2010/relationships/master' Target='master${index + 1}.xml'/>`,
  ).join('');
  const styles = Array.from({ length: inheritanceDepth }, (_, index) => {
    const id = index + 1;
    const basedOn = id > 1 ? ` BasedOn='${id - 1}'` : '';
    return `<StyleSheet ID='${id}' NameU='Bench-${id}'${basedOn}><Cell N='LineColor' V='0'/><Cell N='LineWeight' V='0.01'/><Cell N='FillForegnd' V='1'/><Section N='Character'><Row IX='0'><Cell N='Font' V='0'/><Cell N='Size' V='0.15'/><Cell N='Color' V='0'/></Row></Section></StyleSheet>`;
  }).join('');
  const packageParts: Record<string, string> = {
    '[Content_Types].xml': "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='xml' ContentType='application/xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>",
    '_rels/.rels': "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>",
    'visio/document.xml': `<VisioDocument ${ns}><DocumentSettings/><Colors><ColorEntry IX='0' RGB='#000000'/><ColorEntry IX='1' RGB='#FFFFFF'/></Colors><FaceNames><FaceName ID='0' Name='Calibri'/></FaceNames><StyleSheets>${styles}</StyleSheets><DocumentSheet><Cell N='PageWidth' V='8.5'/><Cell N='PageHeight' V='11'/></DocumentSheet></VisioDocument>`,
    'visio/_rels/document.xml.rels': "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/pages' Target='pages/pages.xml'/><Relationship Id='rId2' Type='http://schemas.microsoft.com/visio/2010/relationships/masters' Target='masters/masters.xml'/></Relationships>",
    'visio/pages/pages.xml': `<Pages ${ns}>${pageCatalog}</Pages>`,
    'visio/pages/_rels/pages.xml.rels': `<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>${pageRelationships}</Relationships>`,
    'visio/masters/masters.xml': `<Masters ${ns}>${masterCatalog}</Masters>`,
    'visio/masters/_rels/masters.xml.rels': `<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>${masterRelationships}</Relationships>`,
  };
  for (let master = 1; master <= inheritanceDepth; master += 1) {
    const parent = master > 1 ? ` Master='${master - 1}' MasterShape='${1_000 + master - 1}'` : '';
    packageParts[`visio/masters/master${master}.xml`] = `<MasterContents ${ns}><Shapes><Shape ID='${1_000 + master}' Type='Shape'${parent}><Cell N='Width' F='1+1' V='2'/><Cell N='Height' V='1'/><Cell N='PinX' V='1'/><Cell N='PinY' V='1'/></Shape></Shapes></MasterContents>`;
  }
  let nextShapeId = 1;
  const text = 'benchmark '.repeat(Math.ceil(Math.max(1, options.textBytes) / 10)).slice(0, Math.max(1, options.textBytes));
  const leafIds: number[][] = Array.from({ length: pages }, () => []);
  const leavesByPage = Array.from({ length: pages }, (_, page) => Math.floor(options.shapeCount / pages) + (page < options.shapeCount % pages ? 1 : 0));
  const shapeXml = (id: number, ordinal: number) => `<Shape ID='${id}' Type='Shape' Master='${inheritanceDepth}' MasterShape='${1_000 + inheritanceDepth}' LineStyle='${inheritanceDepth}' FillStyle='${inheritanceDepth}' TextStyle='${inheritanceDepth}'>${xform(1, 1, (ordinal % 10) + 0.5, Math.floor(ordinal / 10) + 0.5, 0.5, 0.5, 0, 0, 0)}<Cell N='BeginX' F='PinX-Width/2' V='0'/><Cell N='BeginY' F='PinY' V='0'/><Cell N='EndX' F='PinX+Width/2' V='1'/><Cell N='EndY' F='PinY' V='0'/><Section N='Geometry'><Row IX='0' T='MoveTo'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row><Row IX='1' T='LineTo'><Cell N='X' V='1'/><Cell N='Y' V='1'/></Row></Section><Text>${text}</Text></Shape>`;
  const nest = (xml: string, depth: number, ordinal: number): string => {
    if (depth === 0) return xml;
    const groupId = nextShapeId++;
    return `<Shape ID='${groupId}' Type='Group'>${xform(2, 2, (ordinal % 10) + 0.5, Math.floor(ordinal / 10) + 0.5, 1, 1, 0, 0, 0)}<Shapes>${nest(xml, depth - 1, ordinal)}</Shapes></Shape>`;
  };
  for (let page = 0; page < pages; page += 1) {
    const shapes: string[] = [];
    for (let ordinal = 0; ordinal < leavesByPage[page]; ordinal += 1) {
      const id = nextShapeId++;
      leafIds[page].push(id);
      shapes.push(nest(shapeXml(id, ordinal), options.groupDepth, ordinal));
    }
    const connects = leafIds[page].flatMap((from, fromIndex) => Array.from({ length: Math.floor(options.glueDensity) }, (_, offset) => {
      const to = leafIds[page][(fromIndex + offset + 1) % leafIds[page].length];
      return `<Connect FromSheet='${from}' FromCell='BeginX' FromPart='9' ToSheet='${to}' ToCell='PinX' ToPart='3'/><Connect FromSheet='${from}' FromCell='EndX' FromPart='9' ToSheet='${to}' ToCell='PinX' ToPart='3'/>`;
    })).join('');
    packageParts[`visio/pages/page${page + 1}.xml`] = `<PageContents ${ns}><Shapes>${shapes.join('')}</Shapes><Connects>${connects}</Connects></PageContents>`;
  }
  return packageParts;
}

async function writeBenchmarkFixtures(directory: string): Promise<void> {
  fs.mkdirSync(directory, { recursive: true });
  for (const [name, options] of Object.entries(benchmarkFixtures)) {
    const zip = new JSZip();
    for (const [part, contents] of Object.entries(benchmarkPackage(options))) zip.file(part, contents, { date: zipDate, createFolders: false });
    fs.writeFileSync(path.join(directory, `${name}.vsdx`), await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', platform: 'DOS' }));
  }
}

if (import.meta.main) {
  const benchmarkDirectory = process.argv.slice(2).find((argument) => argument.startsWith('--benchmark-dir='));
  if (benchmarkDirectory) {
    await writeBenchmarkFixtures(path.resolve(root, benchmarkDirectory.slice('--benchmark-dir='.length)));
  } else {
    await writeTestFixtures();
  }
}
