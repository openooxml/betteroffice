import JSZip from "jszip";
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const output = path.join(root, "apps/demo/public/betteroffice-demo.vsdx");
const zipDate = new Date("2026-01-01T00:00:00Z");

function cells(values: Record<string, string>): string {
  return Object.entries(values).map(([name, value]) => `<Cell N='${name}' V='${value}'/>`).join("");
}

function geometry(): string {
  return `<Section N='Geometry' IX='0'><Row T='MoveTo' IX='0'>${cells({ X: "0", Y: "0" })}</Row><Row T='RelLineTo' IX='1'>${cells({ X: "1", Y: "0" })}</Row><Row T='RelLineTo' IX='2'>${cells({ X: "1", Y: "1" })}</Row><Row T='RelLineTo' IX='3'>${cells({ X: "0", Y: "1" })}</Row><Row T='RelLineTo' IX='4'>${cells({ X: "0", Y: "0" })}</Row></Section>`;
}

function shape(id: number, name: string, x: number, y: number, width: number, height: number, fill: number, text: string): string {
  return `<Shape ID='${id}' Name='${name}' NameU='${name}' Type='Shape'>${cells({ PinX: String(x), PinY: String(y), Width: String(width), Height: String(height), LocPinX: String(width / 2), LocPinY: String(height / 2), FillPattern: "1", FillForegnd: String(fill), LinePattern: "1", LineColor: "2", Angle: "0", FlipX: "0", FlipY: "0", VerticalAlign: "1", LineWeight: "0.018", ThemeIndex: "0" })}${geometry()}<Section N='Character'><Row IX='0'>${cells({ Font: "0", Size: String(14 / 72), Color: fill === 1 ? "0" : "2", Style: "1" })}</Row></Section><Section N='Paragraph'><Row IX='0'>${cells({ HorzAlign: "1" })}</Row></Section><Text><cp IX='0'/><pp IX='0'/>${text}</Text></Shape>`;
}

function connector(): string {
  return `<Shape ID='30' Name='Release connector' NameU='Release connector' Type='Shape'>${cells({ PinX: "5", PinY: "4.25", Width: "3.5", Height: "0.1", BeginX: "3.25", BeginY: "4.25", EndX: "6.75", EndY: "4.25", LinePattern: "1", LineColor: "2", LineWeight: "0.028" })}<Section N='Geometry' IX='0'><Row T='MoveTo' IX='0'>${cells({ X: "0", Y: "0" })}</Row><Row T='RelLineTo' IX='1'>${cells({ X: "1", Y: "0" })}</Row></Section></Shape>`;
}

const contentTypes = `<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='xml' ContentType='application/xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>`;
const rootRelationships = `<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>`;
const document = `<VisioDocument xmlns='http://schemas.microsoft.com/office/visio/2012/main'><DocumentSettings/><Colors><ColorEntry IX='0' RGB='#FFFFFF'/><ColorEntry IX='1' RGB='#315EFB'/><ColorEntry IX='2' RGB='#172036'/><ColorEntry IX='3' RGB='#C8F56A'/><ColorEntry IX='4' RGB='#FF8066'/></Colors><FaceNames><FaceName ID='0' Name='Arial'/></FaceNames><StyleSheets><StyleSheet ID='0' NameU='Normal'>${cells({ LineColor: "2", FillForegnd: "0", TextColor: "2" })}</StyleSheet></StyleSheets><DocumentSheet>${cells({ PageWidth: "10", PageHeight: "7.5" })}</DocumentSheet></VisioDocument>`;
const documentRelationships = `<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/pages' Target='pages/pages.xml'/><Relationship Id='rId2' Type='http://schemas.microsoft.com/visio/2010/relationships/theme' Target='theme/theme1.xml'/></Relationships>`;
const pages = `<Pages xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><Page ID='1' Name='Product map' NameU='Product map' r:id='rId1'><PageSheet>${cells({ PageWidth: "10", PageHeight: "7.5" })}</PageSheet></Page><Page ID='2' Name='Release flow' NameU='Release flow' r:id='rId2'><PageSheet>${cells({ PageWidth: "10", PageHeight: "7.5" })}</PageSheet></Page></Pages>`;
const pagesRelationships = `<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/page' Target='page1.xml'/><Relationship Id='rId2' Type='http://schemas.microsoft.com/visio/2010/relationships/page' Target='page2.xml'/></Relationships>`;
const page1 = `<PageContents xmlns='http://schemas.microsoft.com/office/visio/2012/main'><Shapes>${shape(1, "BetterOffice", 2.25, 5.5, 3.5, 1.15, 1, "BETTEROFFICE")}<Shape ID='10' Name='Capability group' NameU='Capability group' Type='Group'>${cells({ PinX: "6.8", PinY: "3.7", Width: "3.6", Height: "1.1", LocPinX: "1.8", LocPinY: "0.55" })}<Shapes>${shape(11, "Read", 1.2, 1.4, 1.6, 1.1, 3, "READ")}${shape(12, "Edit", 3.2, 1.4, 1.6, 1.1, 4, "EDIT")}</Shapes></Shape>${shape(20, "Rust engine", 5, 1.65, 5.6, 1.15, 1, "RUST ENGINE")}</Shapes></PageContents>`;
const page2 = `<PageContents xmlns='http://schemas.microsoft.com/office/visio/2012/main'><Shapes>${shape(1, "Author", 2, 4.25, 2.5, 1.15, 3, "AUTHOR")}${shape(2, "Browser", 8, 4.25, 2.5, 1.15, 1, "BROWSER")}${connector()}</Shapes><Connects><Connect FromSheet='30' FromCell='BeginX' FromPart='9' ToSheet='1' ToCell='PinX' ToPart='3'/><Connect FromSheet='30' FromCell='EndX' FromPart='12' ToSheet='2' ToCell='PinX' ToPart='3'/></Connects></PageContents>`;
const theme = `<a:theme xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' name='BetterOffice'><a:themeElements><a:clrScheme name='BetterOffice'><a:dk1><a:srgbClr val='172036'/></a:dk1><a:lt1><a:srgbClr val='FFFFFF'/></a:lt1><a:accent1><a:srgbClr val='315EFB'/></a:accent1><a:accent2><a:srgbClr val='C8F56A'/></a:accent2><a:accent3><a:srgbClr val='FF8066'/></a:accent3></a:clrScheme></a:themeElements></a:theme>`;

const archive = new JSZip();
const options = { date: zipDate, createFolders: false };
for (const [part, contents] of Object.entries({
  "[Content_Types].xml": contentTypes,
  "_rels/.rels": rootRelationships,
  "visio/document.xml": document,
  "visio/_rels/document.xml.rels": documentRelationships,
  "visio/pages/pages.xml": pages,
  "visio/pages/_rels/pages.xml.rels": pagesRelationships,
  "visio/pages/page1.xml": page1,
  "visio/pages/page2.xml": page2,
  "visio/theme/theme1.xml": theme,
})) archive.file(part, contents, options);

const buffer = await archive.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 9 } });
fs.mkdirSync(path.dirname(output), { recursive: true });
fs.writeFileSync(output, buffer);
console.log(`Created ${output} (${buffer.length} bytes)`);
