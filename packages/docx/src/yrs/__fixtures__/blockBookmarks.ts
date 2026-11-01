import { rezipPartsToArrayBuffer, toBytes } from '../../docx/rezip/parts';
import { unzipContainer } from '../../wasm/opc';

export const blockContainers = ['cell', 'control', 'header', 'footer', 'footnote', 'endnote', 'comment', 'textbox'] as const;
export type BlockContainer = typeof blockContainers[number];
export const bookmarkStart = `<w:bookmarkStart w:name='Block&amp;Range' w:id = '&#55;' w:colFirst="0" w:colLast="2" />`;
export const bookmarkEnd = `<w:bookmarkEnd w:displacedByCustomXml='prev' w:id="7"></w:bookmarkEnd>`;
export const bookmarkParagraph = (id: string, text: string) => `<w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
export const bookmarkBlocks = `${bookmarkParagraph('00000001', 'Before')}${bookmarkStart}${bookmarkParagraph('00000002', 'Anchor')}${bookmarkEnd}${bookmarkParagraph('00000003', 'After')}`;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`;

export function blockBookmarkFixture(container: BlockContainer, blocks = bookmarkBlocks): Uint8Array<ArrayBuffer> {
  let body = bookmarkParagraph('00000010', 'Unrelated');
  let section = '';
  const parts = new Map<string, Uint8Array>();
  const relationships: Array<[string, string, string]> = [];
  const overrides: Array<[string, string]> = [];
  const story = (name: string, kind: string, xml: string) => {
    parts.set(`word/${name}.xml`, toBytes(xml));
    relationships.push([kind, kind, `${name}.xml`]);
    overrides.push([name, kind]);
  };
  switch (container) {
    case 'cell':
      body += `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${blocks}</w:tc></w:tr></w:tbl>`;
      break;
    case 'control':
      body += `<w:sdt><w:sdtPr><w:id w:val="42"/></w:sdtPr><w:sdtContent>${blocks}</w:sdtContent></w:sdt>`;
      break;
    case 'header':
    case 'footer': {
      const root = container === 'header' ? 'hdr' : 'ftr';
      story(`${container}1`, container, `<w:${root} ${NS}>${blocks}</w:${root}>`);
      section = `<w:${container}Reference w:type="default" r:id="${container}"/>`;
      break;
    }
    case 'footnote':
    case 'endnote':
      story(`${container}s`, `${container}s`, `<w:${container}s ${NS}><w:${container} w:id="1">${blocks}</w:${container}></w:${container}s>`);
      body += `<w:p><w:r><w:${container}Reference w:id="1"/></w:r></w:p>`;
      break;
    case 'comment':
      story('comments', 'comments', `<w:comments ${NS}><w:comment w:id="1" w:author="Reviewer">${blocks}</w:comment></w:comments>`);
      body += `<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Commented</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>`;
      break;
    case 'textbox':
      body += `<w:p w14:paraId="00000020"><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="20" name="Text Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvPr id="20"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent>${blocks}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  }
  parts.set('[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' + overrides.map(([name, kind]) => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`).join('') + '</Types>'));
  parts.set('_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`));
  parts.set('word/_rels/document.xml.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + relationships.map(([id, kind, target]) => `<Relationship Id="${id}" Type="${R}/${kind}" Target="${target}"/>`).join('') + '</Relationships>'));
  parts.set('word/document.xml', toBytes(`<w:document ${NS}><w:body>${body}<w:sectPr>${section}</w:sectPr></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

export function bookmarkPart(container: BlockContainer): string {
  if (container === 'header' || container === 'footer') return `word/${container}1.xml`;
  if (container === 'footnote' || container === 'endnote') return `word/${container}s.xml`;
  return container === 'comment' ? 'word/comments.xml' : 'word/document.xml';
}

export function drawingBookmarkFixture(kind: 'choice' | 'group' | 'vml'): Uint8Array<ArrayBuffer> {
  const parts = new Map(Object.entries(unzipContainer(blockBookmarkFixture('textbox'))));
  let xml = new TextDecoder().decode(parts.get('word/document.xml'));
  const drawing = /<w:drawing>[\s\S]*?<\/w:drawing>/.exec(xml)![0];
  const fallbackBlocks = bookmarkBlocks.replace(bookmarkStart, bookmarkStart.replace("'&#55;'", "'8'"))
    .replace(bookmarkEnd, bookmarkEnd.replace('"7"', '"8"'))
    .replaceAll('w14:paraId="0000000', 'w14:paraId="0000001');
  const fallback = `<w:pict><v:group><v:rect style="width:72pt;height:72pt"><v:textbox><w:txbxContent>${fallbackBlocks}</w:txbxContent></v:textbox></v:rect></v:group></w:pict>`;
  let choice = drawing;
  if (kind === 'group') {
    const shape = /<wps:wsp>[\s\S]*?<\/wps:wsp>/.exec(choice)![0];
    choice = choice.replace(shape, `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></wpg:grpSpPr>${shape}</wpg:wgp>`);
    choice = choice.replace('uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"', 'uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"');
  }
  const content = kind === 'vml' ? fallback
    : `<mc:AlternateContent><mc:Choice Requires="${kind === 'group' ? 'wpg' : 'wps'}">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
  xml = xml.replace(drawing, content).replace('<w:document ', '<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ');
  parts.set('word/document.xml', toBytes(xml));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}
