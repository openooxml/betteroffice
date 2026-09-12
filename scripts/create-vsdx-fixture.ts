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

const zip = new JSZip();
for (const [name, contents] of Object.entries(parts)) zip.file(name, contents, { date: zipDate, createFolders: false });
fs.writeFileSync(output, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', platform: 'DOS' }));

const nestedOutput = path.join(root, 'crates/vsdx-parse/tests/fixtures/nested-groups.vsdx');
const xform = (width: number, height: number, pinX: number, pinY: number, locPinX: number, locPinY: number, angle: number, flipX: number, flipY: number) =>
  `<Cell N='Width' V='${width}'/><Cell N='Height' V='${height}'/><Cell N='PinX' V='${pinX}'/><Cell N='PinY' V='${pinY}'/><Cell N='LocPinX' V='${locPinX}'/><Cell N='LocPinY' V='${locPinY}'/><Cell N='Angle' V='${angle}'/><Cell N='FlipX' V='${flipX}'/><Cell N='FlipY' V='${flipY}'/>`;
const rect = `<Section N='Geometry'><Row IX='0' T='MoveTo'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row><Row IX='1' T='LineTo'><Cell N='X' V='1'/><Cell N='Y' V='0'/></Row><Row IX='2' T='LineTo'><Cell N='X' V='1'/><Cell N='Y' V='1'/></Row><Row IX='3' T='LineTo'><Cell N='X' V='0'/><Cell N='Y' V='1'/></Row><Row IX='4' T='Close'/></Section>`;
const nestedParts: Record<string, string | Uint8Array> = {
  ...parts,
  'visio/pages/page1.xml': `<PageContents ${ns}><Shapes><Shape ID='1' Type='Group'>${xform(6, 4, 10, 10, 1, 0.5, 0.5235987755982988, 1, 0)}<Shapes><Shape ID='2' Type='Group'>${xform(3, 5, 2, 1, 0.25, 0.75, -0.7853981633974483, 0, 1)}<Shapes><Shape ID='3' Type='Shape'>${xform(1, 1, 0, 0, 0, 0, 0, 0, 0)}${rect}<Text>deep\nvector</Text></Shape><Shape ID='4' Type='Shape'>${xform(1, 2, 2, 1, 0, 0, 0, 0, 0)}<ForeignData ForeignType='Bitmap'><Rel r:id='rIdImage' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'/></ForeignData></Shape><Shape ID='5' Type='Shape'>${xform(1, 1, 1, 3, 0, 0, 0, 0, 0)}<Section N='Geometry'><Row T='EllipticalArcTo'><Cell N='X' V='1'/></Row></Section></Shape></Shapes></Shape></Shapes></Shape></Shapes></PageContents>`,
  'visio/pages/_rels/page1.xml.rels': "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rIdImage' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='../media/image1.png'/></Relationships>",
  'visio/media/image1.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
};
const nestedZip = new JSZip();
for (const [name, contents] of Object.entries(nestedParts)) nestedZip.file(name, contents, { date: zipDate, createFolders: false });
fs.writeFileSync(nestedOutput, await nestedZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', platform: 'DOS' }));

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
