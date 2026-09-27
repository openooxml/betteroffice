import JSZip from 'jszip';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const output = path.join(
  root,
  'packages/docx/src/yrs/__fixtures__/content-controls/template.docx'
);
const zipDate = new Date('2026-01-01T00:00:00Z');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const NS = [
  `xmlns:w="${W}"`,
  `xmlns:r="${R}"`,
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"',
  'mc:Ignorable="w14 w15"',
].join(' ');
const STORE_ITEM = '{6F2C8B5D-3E1A-4C7B-9D2E-1A2B3C4D5E6F}';

const run = (text: string, properties = '') =>
  `<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (id: string, content: string, properties = '') =>
  `<w:p w14:paraId="${id}">${properties ? `<w:pPr>${properties}</w:pPr>` : ''}${content}</w:p>`;

const body = [
  paragraph('10000001', run('Service agreement'), '<w:pStyle w:val="Title"/>'),
  paragraph(
    '10000002',
    run('Customer: ') +
      '<w:sdt><w:sdtPr><w:rPr><w:b/></w:rPr><w:alias w:val="Customer name"/><w:tag w:val="customer.name"/><w:id w:val="101"/><w:showingPlcHdr/><w15:appearance w15:val="tags"/><w:text/></w:sdtPr>' +
      `<w:sdtContent>${run('Click to enter a name.', '<w:rStyle w:val="PlaceholderText"/>')}</w:sdtContent></w:sdt>` +
      run('.')
  ),
  paragraph(
    '10000003',
    run('Reference: ') +
      '<w:sdt><w:sdtPr><w:alias w:val="Account reference"/><w:tag w:val="account.reference"/><w:id w:val="102"/><w:richText/></w:sdtPr><w:sdtEndPr><w:rPr><w:i/></w:rPr></w:sdtEndPr>' +
      `<w:sdtContent>${run('REF-', '<w:b/>')}${run('000')}</w:sdtContent></w:sdt>`
  ),
  paragraph(
    '10000004',
    run('Previous reference: ') +
      '<w:sdt><w:sdtPr><w:tag w:val="account.reference"/><w:id w:val="103"/><w:text/></w:sdtPr>' +
      `<w:sdtContent>${run('REF-OLD')}</w:sdtContent></w:sdt>`
  ),
  paragraph(
    '10000005',
    run('Terms: ') +
      '<w:sdt><w:sdtPr><w:alias w:val="Terms"/><w:tag w:val="terms.standard"/><w:id w:val="104"/><w:lock w:val="contentLocked"/><w:text/></w:sdtPr>' +
      `<w:sdtContent>${run('Standard terms apply.')}</w:sdtContent></w:sdt>` +
      `<w:hyperlink r:id="rIdTerms">${run(' (details)', '<w:rStyle w:val="Hyperlink"/>')}</w:hyperlink>`
  ),
  '<w:sdt><w:sdtPr><w:alias w:val="Address"/><w:tag w:val="customer.address"/><w:id w:val="105"/><w:text w:multiLine="1"/></w:sdtPr><w:sdtContent>' +
    paragraph('10000006', `${run('1 Old Road')}<w:r><w:br/><w:t>Oldtown</w:t></w:r>`) +
    '</w:sdtContent></w:sdt>',
  paragraph(
    '10000007',
    run('Email: ') +
      `<w:sdt><w:sdtPr><w:alias w:val="Email"/><w:tag w:val="customer.email"/><w:id w:val="106"/><w:dataBinding w:xpath="/root[1]/email[1]" w:storeItemID="${STORE_ITEM}"/><w:text/></w:sdtPr>` +
      `<w:sdtContent>${run('ada@example.com')}</w:sdtContent></w:sdt>`
  ),
  paragraph(
    '10000008',
    `<w:bookmarkStart w:id="0" w:name="signature"/>${run('Signed by the parties.')}<w:bookmarkEnd w:id="0"/>`
  ),
  '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>',
].join('');

const header = paragraph(
  '20000001',
  run('Confidential: ') +
    '<w:sdt><w:sdtPr><w:alias w:val="Document title"/><w:tag w:val="document.title"/><w:id w:val="201"/><w:text/></w:sdtPr>' +
    `<w:sdtContent>${run('Untitled')}</w:sdtContent></w:sdt>`
);

const relationships = (entries: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

const parts: Record<string, string> = {
  '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/><Override PartName="/word/settings.xml" ContentType="${OFFICE}.settings+xml"/><Override PartName="/word/header1.xml" ContentType="${OFFICE}.header+xml"/><Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
  '_rels/.rels': relationships(
    `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/>`
  ),
  'docProps/core.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Content control template</dc:title><dc:creator>BetterOffice fixture</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>`,
  'docProps/app.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>BetterOffice fixture</Application></Properties>`,
  'word/_rels/document.xml.rels': relationships(
    `<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rIdSettings" Type="${REL}/settings" Target="settings.xml"/><Relationship Id="rIdHeader" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdCustomXml" Type="${REL}/customXml" Target="../customXml/item1.xml"/><Relationship Id="rIdTerms" Type="${REL}/hyperlink" Target="https://example.com/terms" TargetMode="External"/>`
  ),
  'word/document.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${NS}><w:body>${body}</w:body></w:document>`,
  'word/header1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr ${NS}>${header}</w:hdr>`,
  'word/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="40"/></w:rPr></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style><w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:rPr><w:color w:val="808080"/></w:rPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style></w:styles>`,
  'word/settings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="${W}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
  'customXml/item1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<root><email>ada@example.com</email></root>`,
  'customXml/itemProps1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="${STORE_ITEM}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>`,
  'customXml/_rels/item1.xml.rels': relationships(
    `<Relationship Id="rId1" Type="${REL}/customXmlProps" Target="itemProps1.xml"/>`
  ),
};

const zip = new JSZip();
for (const [name, contents] of Object.entries(parts)) {
  zip.file(name, contents, { date: zipDate, createFolders: false });
}
fs.mkdirSync(path.dirname(output), { recursive: true });
fs.writeFileSync(
  output,
  await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', platform: 'DOS' })
);
