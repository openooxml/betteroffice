import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export async function pagedDocx(
  pages: number,
  paragraphsPerPage = 1,
  options: { trackedInsertion?: boolean } = {}
): Promise<ArrayBuffer> {
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
        options.trackedInsertion && page === pages - 1 && paragraph === paragraphsPerPage - 1
          ? '<w:ins w:id="1" w:author="Document reviewer" w:date="2026-09-29T00:00:00Z">' +
            `<w:r>${font}<w:t xml:space="preserve"> Existing insertion.</w:t></w:r></w:ins>`
          : '';
      return `<w:p${identity}>${breakBefore}<w:r>${font}<w:t>${text}</w:t></w:r>${insertion}</w:p>`;
    }).join('')
  ).join('');
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
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
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${W}"${identityNamespace}><w:body>${body}</w:body></w:document>`
  );
  zip.forEach((_, entry) => {
    entry.date = new Date('2026-09-30T00:00:00Z');
  });
  return zip.generateAsync({ type: 'arraybuffer' });
}
