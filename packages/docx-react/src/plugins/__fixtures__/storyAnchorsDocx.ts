import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const PART = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const PAGE =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';

const font = '<w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/></w:rPr>';

function paragraph(paraId: string, text: string, breakBefore = false): string {
  const properties = breakBefore ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : '';
  return `<w:p w14:paraId="${paraId}">${properties}<w:r>${font}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

function table(cells: string[]): string {
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid>' +
    cells.map(() => `<w:gridCol w:w="${Math.floor(9000 / cells.length)}"/>`).join('') +
    '</w:tblGrid><w:tr>' +
    cells.map((cell) => `<w:tc><w:tcPr><w:tcW w:w="${Math.floor(9000 / cells.length)}" w:type="dxa"/></w:tcPr>${cell}</w:tc>`).join('') +
    '</w:tr></w:tbl>'
  );
}

function part(root: 'hdr' | 'ftr', content: string): string {
  return `<w:${root} xmlns:w="${W}" xmlns:w14="${W14}">${content}</w:${root}>`;
}

/**
 * Six pages under `titlePg` and `evenAndOddHeaders`: page 1 paints only the first-page footer
 * `rIdF1`, odd pages the default header `rIdH1` (with a table cell), even pages the even header
 * `rIdH2`. Body table 0 holds `{{cell}}` in `body:t0:r0c0`; table 1's single row splits across
 * pages 2 and 3.
 */
export async function storyAnchorsDocx(): Promise<Uint8Array> {
  const split = Array.from({ length: 70 }, (_, index) =>
    paragraph(`2${index.toString(16).padStart(7, '0')}`, `Split line ${index + 1}`)
  ).join('');
  const body =
    paragraph('00000001', 'Intro') +
    table([paragraph('00000002', 'Cell {{cell}} text'), paragraph('00000003', 'Neighbour')]) +
    paragraph('00000004', 'Before split', true) +
    table([split]) +
    paragraph('00000005', 'After split') +
    paragraph('00000006', 'Page four', true) +
    paragraph('00000007', 'Page five', true) +
    paragraph('00000008', 'Page six', true) +
    `<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:headerReference w:type="even" r:id="rIdH2"/>` +
    `<w:footerReference w:type="first" r:id="rIdF1"/>${PAGE}<w:titlePg/></w:sectPr>`;
  return packaged(body, {
    rIdH1: [
      'header',
      paragraph('10000001', 'Odd {{odd}} header') +
        table([paragraph('10000002', 'Header {{hcell}}')]) +
        paragraph('10000003', ''),
    ],
    rIdH2: ['header', paragraph('10000011', 'Even {{even}} header')],
    rIdF1: ['footer', paragraph('10000021', 'First {{first}} footer')],
  });
}

/**
 * Two sections under `evenAndOddHeaders`. Section 1 (three pages) references default `rIdA`, even
 * `rIdE` and footer `rIdF`; section 2 starts on an odd page, so a blank parity filler page
 * precedes it, inherits those and adds first `rIdB` under `titlePg` and even `rIdC`. Each part
 * holds one paragraph, `1000000N` in part order A, E, F, B, C.
 */
export async function sectionedDocx(): Promise<Uint8Array> {
  const body =
    paragraph('00000001', 'One') +
    paragraph('00000002', 'Two', true) +
    `<w:p w14:paraId="00000003"><w:pPr><w:pageBreakBefore/><w:sectPr>` +
    '<w:headerReference w:type="default" r:id="rIdA"/><w:headerReference w:type="even" r:id="rIdE"/>' +
    `<w:footerReference w:type="default" r:id="rIdF"/>${PAGE}</w:sectPr></w:pPr><w:r>${font}<w:t>Three</w:t></w:r></w:p>` +
    paragraph('00000004', 'Four') +
    paragraph('00000005', 'Five', true) +
    paragraph('00000006', 'Six', true) +
    '<w:sectPr><w:headerReference w:type="first" r:id="rIdB"/><w:headerReference w:type="even" r:id="rIdC"/>' +
    `<w:type w:val="oddPage"/>${PAGE}<w:titlePg/></w:sectPr>`;
  return packaged(
    body,
    Object.fromEntries(
      (['A', 'E', 'F', 'B', 'C'] as const).map((name, index) => [
        `rId${name}`,
        [name === 'F' ? 'footer' : 'header', paragraph(`1000000${index + 1}`, `Part ${name}`)],
      ])
    ) as Record<string, ['header' | 'footer', string]>
  );
}

/** A package with `evenAndOddHeaders`, `body` and one part per relationship id. */
async function packaged(body: string, parts: Record<string, ['header' | 'footer', string]>) {
  const entries = Object.entries(parts);
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${PART}.document.main+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="${PART}.settings+xml"/>` +
      entries
        .map(([rId, [kind]]) => `<Override PartName="/word/${rId}.xml" ContentType="${PART}.${kind}+xml"/>`)
        .join('') +
      '</Types>'
  );
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>` +
      '</Relationships>'
  );
  zip.file(
    'word/_rels/document.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdS" Type="${R}/settings" Target="settings.xml"/>` +
      entries
        .map(([rId, [kind]]) => `<Relationship Id="${rId}" Type="${R}/${kind}" Target="${rId}.xml"/>`)
        .join('') +
      '</Relationships>'
  );
  zip.file('word/settings.xml', `<w:settings xmlns:w="${W}"><w:evenAndOddHeaders/></w:settings>`);
  for (const [rId, [kind, content]] of entries) {
    zip.file(`word/${rId}.xml`, part(kind === 'header' ? 'hdr' : 'ftr', content));
  }
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}"><w:body>${body}</w:body></w:document>`
  );
  zip.forEach((_, entry) => {
    entry.date = new Date('2026-10-01T00:00:00Z');
  });
  return zip.generateAsync({ type: 'uint8array' });
}
