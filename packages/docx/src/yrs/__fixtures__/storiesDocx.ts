import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../../docx/rezip/parts';

const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const table = (cells: string[]) =>
  '<w:tbl><w:tblPr/><w:tblGrid>' +
  cells.map(() => '<w:gridCol w:w="2000"/>').join('') +
  '</w:tblGrid><w:tr>' +
  cells.map((cell) => `<w:tc>${cell}</w:tc>`).join('') +
  '</w:tr></w:tbl>';
const note = (kind: 'footnote' | 'endnote', id: number, text: string) =>
  `<w:${kind} w:id="${id}">${p(text)}</w:${kind}>`;
const separators = (kind: 'footnote' | 'endnote') =>
  `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
  `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;

/**
 * Two sections with headers, a footer the second inherits, a table in a header, a nested body
 * table, a block content control, a footnote and an endnote; `tables` adds one-row body tables.
 */
export function storiesDocx({ tables = 0 }: { tables?: number } = {}): Uint8Array {
  const filler = Array.from({ length: tables }, (_, index) =>
    table([p(`row ${index} a`), p(`row ${index} b`)])
  ).join('');
  const body =
    p('Intro {{title}}') +
    `<w:sdt><w:sdtPr><w:tag w:val="block"/></w:sdtPr><w:sdtContent>${p('Inside control')}</w:sdtContent></w:sdt>` +
    table([p('Cell {{a}}'), table([p('Nested {{b}}')])]) +
    '<w:p><w:r><w:t>Noted</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>' +
    '<w:p><w:pPr><w:sectPr>' +
    '<w:headerReference w:type="default" r:id="rIdH1"/><w:headerReference w:type="first" r:id="rIdH2"/>' +
    '<w:footerReference w:type="default" r:id="rIdF1"/><w:titlePg/>' +
    '</w:sectPr></w:pPr><w:r><w:t>End of one</w:t></w:r></w:p>' +
    filler +
    p('Second section') +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdH3"/></w:sectPr>';
  const parts: PartsMap = new Map([
    [
      '[Content_Types].xml',
      toBytes(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          `<Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/>` +
          ['header1', 'header2', 'header3']
            .map((name) => `<Override PartName="/word/${name}.xml" ContentType="${MAIN}.header+xml"/>`)
            .join('') +
          `<Override PartName="/word/footer1.xml" ContentType="${MAIN}.footer+xml"/>` +
          `<Override PartName="/word/footnotes.xml" ContentType="${MAIN}.footnotes+xml"/>` +
          `<Override PartName="/word/endnotes.xml" ContentType="${MAIN}.endnotes+xml"/>` +
          '</Types>'
      ),
    ],
    [
      '_rels/.rels',
      toBytes(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="rIdDoc" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
      ),
    ],
    [
      'word/_rels/document.xml.rels',
      toBytes(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/>` +
          `<Relationship Id="rIdH2" Type="${REL}/header" Target="header2.xml"/>` +
          `<Relationship Id="rIdH3" Type="${REL}/header" Target="header3.xml"/>` +
          `<Relationship Id="rIdF1" Type="${REL}/footer" Target="footer1.xml"/>` +
          `<Relationship Id="rIdFn" Type="${REL}/footnotes" Target="footnotes.xml"/>` +
          `<Relationship Id="rIdEn" Type="${REL}/endnotes" Target="endnotes.xml"/>` +
          '</Relationships>'
      ),
    ],
    ['word/document.xml', toBytes(`<w:document ${W}><w:body>${body}</w:body></w:document>`)],
    ['word/header1.xml', toBytes(`<w:hdr ${W}>${table([p('Header cell {{c}}')])}${p('Header one')}</w:hdr>`)],
    ['word/header2.xml', toBytes(`<w:hdr ${W}>${p('First page {{d}}')}</w:hdr>`)],
    ['word/header3.xml', toBytes(`<w:hdr ${W}>${p('Header three')}</w:hdr>`)],
    ['word/footer1.xml', toBytes(`<w:ftr ${W}>${p('Footer {{e}}')}</w:ftr>`)],
    [
      'word/footnotes.xml',
      toBytes(`<w:footnotes ${W}>${separators('footnote')}${note('footnote', 1, 'Footnote text')}</w:footnotes>`),
    ],
    [
      'word/endnotes.xml',
      toBytes(`<w:endnotes ${W}>${separators('endnote')}${note('endnote', 1, 'Endnote text')}</w:endnotes>`),
    ],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}
