import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const WORDS = (
  'account action agreement amount analysis annual approval area asset audit balance basis ' +
  'benefit budget capacity capital change claim clause client committee company condition ' +
  'contract control cost coverage credit customer data date decision delivery demand design ' +
  'detail device document duty effect element energy equipment estimate event evidence ' +
  'exception expense facility factor feature figure finance focus format function fund goal ' +
  'guidance health history impact income index industry input insurance interest issue item ' +
  'labour language level liability licence limit list load maintenance margin market measure ' +
  'method model module network notice number object obligation offer operation option order ' +
  'output owner package party payment period permit phase plan policy portion practice price ' +
  'process product profile program project property provider quality quantity range rate ' +
  'record region report request resource result review right risk role rule safety sample ' +
  'schedule scope section security segment service share site source standard statement ' +
  'status step storage structure supply support system target task term test threshold time ' +
  'total transfer unit update usage value vendor version volume warranty weight window work'
).split(' ');

function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function escape(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function run(text: string, props = ''): string {
  return `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${escape(text)}</w:t></w:r>`;
}

/**
 * A deterministic synthetic DOCX of roughly `pages` A4 pages: headings, wrapped body
 * paragraphs with bold and italic runs, bulleted lists and tables. No real content.
 */
export async function typingDocx(pages: number, seed = 1): Promise<Uint8Array> {
  const next = random(seed);
  const pick = () => WORDS[Math.floor(next() * WORDS.length)];
  const words = (min: number, max: number) =>
    Array.from({ length: min + Math.floor(next() * (max - min)) }, pick);
  const sentence = (min: number, max: number) => {
    const list = words(min, max);
    list[0] = list[0][0].toUpperCase() + list[0].slice(1);
    return list.join(' ') + '.';
  };
  const body: string[] = [];
  const paragraph = () => {
    const sentences = Array.from({ length: 3 + Math.floor(next() * 5) }, () => sentence(8, 20));
    const emphasis = Math.floor(next() * sentences.length);
    return (
      '<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr>' +
      sentences
        .map((text, index) =>
          index === emphasis
            ? run(text + ' ', next() < 0.5 ? '<w:b/>' : '<w:i/>')
            : run(text + ' ')
        )
        .join('') +
      '</w:p>'
    );
  };
  // A section fills about 0.85 of an A4 page at 11 pt with 1-inch margins.
  const sections = Math.round(pages / 0.85);
  for (let section = 1; section <= sections; section++) {
    body.push(
      `<w:p><w:pPr><w:pStyle w:val="${section % 4 === 1 ? 'Heading1' : 'Heading2'}"/></w:pPr>` +
        run(`${section}. ${sentence(3, 7).slice(0, -1)}`) +
        '</w:p>'
    );
    for (let index = 0; index < 4; index++) body.push(paragraph());
    if (section % 3 === 0) {
      const rows = Array.from({ length: 5 }, (_, row) =>
        '<w:tr>' +
        Array.from(
          { length: 3 },
          () =>
            '<w:tc><w:tcPr><w:tcW w:w="3009" w:type="dxa"/></w:tcPr><w:p>' +
            run(row === 0 ? words(1, 3).join(' ') : sentence(2, 9), row === 0 ? '<w:b/>' : '') +
            '</w:p></w:tc>'
        ).join('') +
        '</w:tr>'
    );
      body.push(
        '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9027" w:type="dxa"/></w:tblPr>' +
          '<w:tblGrid><w:gridCol w:w="3009"/><w:gridCol w:w="3009"/><w:gridCol w:w="3009"/></w:tblGrid>' +
          rows.join('') +
          '</w:tbl>'
      );
    } else {
      for (let index = 0; index < 3; index++) {
        body.push(
          '<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
            run(sentence(3, 7)) +
            '</w:p>'
        );
      }
    }
    body.push(paragraph());
  }
  const border = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="808080"/>`;
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
      '</Types>'
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>'
  );
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
      '</Relationships>'
  );
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}">` +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>' +
      '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:jc w:val="both"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
      '<w:next w:val="BodyText"/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>' +
      '<w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="1F3864"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>' +
      '<w:next w:val="BodyText"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr>' +
      '<w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="2F5496"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:spacing w:after="60"/><w:ind w:left="720" w:hanging="360"/></w:pPr></w:style>' +
      '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders>' +
      ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('') +
      '</w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>' +
      '</w:tblPr></w:style></w:styles>'
  );
  zip.file(
    'word/numbering.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}">` +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>' +
      '<w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>` +
      body.join('') +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
      '</w:sectPr></w:body></w:document>'
  );
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
