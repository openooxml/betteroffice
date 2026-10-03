import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../../docx/rezip/parts';

const BLOCKS = 320;
const WHOLE_BLOCKS = 120;
export const FLAVOURS = [
  'plain',
  'keepnext',
  'keepnext-style',
  'widows',
  'footnotes',
  'endnotes',
  'floats',
  'final-sect',
  'early-sect',
  'whole',
] as const;
export type Flavour = (typeof FLAVOURS)[number];

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const DRAW =
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const WORDS =
  'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau'.split(
    ' '
  );
const PNG = new Uint8Array(
  Buffer.from(
    '89504e470d0a1a0a0000000d4948445200000001000000010806000000' +
      '1f15c4890000000d49444154789c6360f8cfc0000003010100c9fe92ef0000000049454e44ae426082',
    'hex'
  )
);

function run(text: string): string {
  return `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
}

function field(instruction: string, cached: string): string {
  return (
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    `<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    `<w:r><w:t>${cached}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

function anchor(id: number): string {
  return (
    `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="1828800" cy="1371600"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${id}" name="Float ${id}"/>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="f${id}.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
    '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1371600"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>' +
    '</wp:anchor></w:drawing></w:r>'
  );
}

function section(landscape = false, title = false, start?: number): string {
  const size = landscape
    ? '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
    : '<w:pgSz w:w="12240" w:h="15840"/>';
  const margins = landscape
    ? '<w:pgMar w:top="1080" w:right="1800" w:bottom="1080" w:left="1800" w:header="540" w:footer="540" w:gutter="0"/>'
    : '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
  return (
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>' +
    '<w:footerReference w:type="default" r:id="rIdF1"/>' +
    (title ? '<w:headerReference w:type="first" r:id="rIdH2"/>' : '') +
    size +
    margins +
    (start === undefined ? '' : `<w:pgNumType w:start="${start}"/>`) +
    (title ? '<w:titlePg/>' : '') +
    '</w:sectPr>'
  );
}

export function syntheticDocx(
  flavour: Flavour,
  size: number,
  seed: number,
  options: { pageNumberRestarts?: [number, number] } = {}
): Uint8Array {
  let state = seed;
  const random = (limit: number) => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state % limit;
  };
  const text = (low: number, high: number) => {
    const words = Array.from(
      { length: low + random(high - low + 1) },
      () => WORDS[random(WORDS.length)]
    );
    const sentence = words.join(' ');
    return sentence[0].toUpperCase() + sentence.slice(1) + '.';
  };
  const parts: PartsMap = new Map();
  const set = (name: string, content: string) => parts.set(name, toBytes(content));
  const rels: Array<[string, string, string]> = [
    ['rIdS', 'styles', 'styles.xml'],
    ['rIdH1', 'header', 'header1.xml'],
    ['rIdH2', 'header', 'header2.xml'],
    ['rIdF1', 'footer', 'footer1.xml'],
  ];
  const overrides: Array<[string, string]> = [
    ['/word/styles.xml', 'styles'],
    ['/word/header1.xml', 'header'],
    ['/word/header2.xml', 'header'],
    ['/word/footer1.xml', 'footer'],
  ];
  const footnotes: number[] = [];
  const endnotes: number[] = [];
  const body: string[] = [];
  const blocks = flavour === 'whole' ? WHOLE_BLOCKS : BLOCKS;
  for (let i = 1; i <= blocks; i += 1) {
    let properties = '';
    let runs = run(text(3, 9));
    if (flavour === 'keepnext' && i >= 150 && i <= 230 && i % 40 < 30) {
      properties += '<w:keepNext/>';
    }
    if (flavour === 'keepnext-style' && i % 25 < 18) {
      properties += '<w:pStyle w:val="Keep"/>';
    }
    if (flavour === 'widows' && i % 7 === 0) {
      properties += '<w:widowControl/><w:keepLines/>';
      runs = run(text(40, 90));
    }
    if (flavour === 'footnotes' && i % 9 === 0) {
      footnotes.push(footnotes.length + 1);
      runs += `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${footnotes.length}"/></w:r>`;
    }
    if (flavour === 'endnotes' && i % 13 === 0) {
      endnotes.push(endnotes.length + 1);
      runs += `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="${endnotes.length}"/></w:r>`;
    }
    if (flavour === 'floats' && i % 45 === 0) runs = anchor(i) + runs;
    if (flavour === 'early-sect' && i === 300) properties += section(true, true, 7);
    if (options.pageNumberRestarts && i === Math.floor(blocks / 2)) {
      properties += section(false, false, options.pageNumberRestarts[0]);
    }
    body.push(
      `<w:p><w:pPr>${properties}<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` +
        `<w:rPr><w:sz w:val="${size}"/></w:rPr></w:pPr>${runs}</w:p>`
    );
  }
  set(
    'word/styles.xml',
    `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/>` +
      `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:rPrDefault>` +
      '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:styleId="Keep"><w:name w:val="Keep"/><w:pPr><w:keepNext/></w:pPr></w:style></w:styles>'
  );
  set('word/header1.xml', `<w:hdr ${W}><w:p>${run('Default header')}</w:p></w:hdr>`);
  set('word/header2.xml', `<w:hdr ${W}><w:p>${run('First page header')}</w:p></w:hdr>`);
  set(
    'word/footer1.xml',
    `<w:ftr ${W}><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${run('Page ')}${field('PAGE', '1')}` +
      (flavour === 'whole' ? `${run(' of ')}${field('NUMPAGES', '99')}` : '') +
      '</w:p></w:ftr>'
  );
  for (const [kind, notes, id] of [
    ['footnote', footnotes, 'rIdFn'],
    ['endnote', endnotes, 'rIdEn'],
  ] as const) {
    if (notes.length === 0) continue;
    const items = notes
      .map(
        (note) =>
          `<w:${kind} w:id="${note}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${kind}Ref/></w:r>` +
          `${run(' ' + text(8, 30))}</w:p></w:${kind}>`
      )
      .join('');
    set(
      `word/${kind}s.xml`,
      `<w:${kind}s ${W}><w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
        `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>` +
        `${items}</w:${kind}s>`
    );
    rels.push([id, `${kind}s`, `${kind}s.xml`]);
    overrides.push([`/word/${kind}s.xml`, `${kind}s`]);
  }
  if (flavour === 'floats') {
    parts.set('word/media/f.png', PNG);
    rels.push(['rIdImg', 'image', 'media/f.png']);
  }
  const final =
    flavour === 'final-sect'
      ? section(true, true, 5)
      : section(false, false, options.pageNumberRestarts?.[1]);
  set(
    'word/document.xml',
    `<w:document ${W} ${R} ${DRAW}><w:body>${body.join('')}${final}</w:body></w:document>`
  );
  const contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  set(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
      `<Override PartName="/word/document.xml" ContentType="${contentType}.document.main+xml"/>` +
      overrides
        .map(
          ([part, kind]) =>
            `<Override PartName="${part}" ContentType="${contentType}.${kind}+xml"/>`
        )
        .join('') +
      '</Types>'
  );
  set(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>'
  );
  set(
    'word/_rels/document.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      rels
        .map(
          ([id, kind, target]) =>
            `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${target}"/>`
        )
        .join('') +
      '</Relationships>'
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}
