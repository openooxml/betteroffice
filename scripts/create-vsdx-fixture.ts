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
