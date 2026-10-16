import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const FONT = '<w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/></w:rPr>';

export async function pagedDocx(
  pages: number,
  paragraphsPerPage = 1,
  options: { trackedInsertion?: boolean | number; pageTotal?: boolean; headerFooter?: boolean } = {}
): Promise<ArrayBuffer> {
  const insertions = Number(options.trackedInsertion ?? 0);
  const body = Array.from({ length: pages }, (_, page) =>
    Array.from({ length: paragraphsPerPage }, (_, paragraph) => {
      const id = (page * paragraphsPerPage + paragraph + 1).toString(16).padStart(8, '0');
      const text =
        paragraphsPerPage === 1
          ? `Page ${page + 1}`
          : `Page ${page + 1} paragraph ${paragraph + 1}: Original text.`;
      const breakBefore =
        page > 0 && paragraph === 0 ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : '';
      const identity = paragraphsPerPage === 1 ? '' : ` w14:paraId="${id}"`;
      const font =
        paragraphsPerPage === 1
          ? ''
          : '<w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/></w:rPr>';
      const insertion =
        page >= pages - insertions && paragraph === paragraphsPerPage - 1
          ? `<w:ins w:id="${pages - page}" w:author="Document reviewer" w:date="2026-09-29T00:00:00Z">` +
            `<w:r>${font}<w:t xml:space="preserve"> Existing insertion.</w:t></w:r></w:ins>`
          : '';
      const total = options.pageTotal && page === 0 && paragraph === 0
        ? `<w:fldSimple w:instr=" NUMPAGES "><w:r>${font}<w:t>${pages}</w:t></w:r></w:fldSimple>` : '';
      return `<w:p${identity}>${breakBefore}<w:r>${font}<w:t>${text}</w:t></w:r>${total}${insertion}</w:p>`;
    }).join('')
  ).join('');
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      (options.headerFooter
        ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        : '') +
      '</Types>'
  );
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>'
  );
  const identityNamespace =
    paragraphsPerPage === 1 ? '' : ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
  const section = options.headerFooter
    ? '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr>'
    : '';
  if (options.headerFooter) {
    zip.file(
      'word/_rels/document.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdHeader" Type="${R}/header" Target="header1.xml"/>` +
        `<Relationship Id="rIdFooter" Type="${R}/footer" Target="footer1.xml"/>` +
        '</Relationships>'
    );
    for (const [part, root, paraId, text] of [['header1', 'hdr', '7F000001', 'Document header'], ['footer1', 'ftr', '7F000002', 'Document footer']]) {
      zip.file(`word/${part}.xml`,
        `<w:${root} xmlns:w="${W}" xmlns:w14="${W14}"><w:p w14:paraId="${paraId}"><w:r>${FONT}<w:t>${text}</w:t></w:r></w:p></w:${root}>`);
    }
  }
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${W}"${options.headerFooter ? ` xmlns:r="${R}"` : ''}${identityNamespace}><w:body>${body}${section}</w:body></w:document>`
  );
  zip.forEach((_, entry) => {
    entry.date = new Date('2026-09-30T00:00:00Z');
  });
  return zip.generateAsync({ type: 'arraybuffer' });
}
