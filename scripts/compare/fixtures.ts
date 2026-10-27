import JSZip from 'jszip';

const WORDS = ('time year people way day man thing woman life child world school state family student group country problem hand part ' +
  'place case week company system program question work government number night point home water room mother area money story fact ' +
  'month lot right study book eye job word business issue side kind head house service friend father power hour game line end member ' +
  'law car city community name president team minute idea kid body information back parent face others level office door health person ' +
  'art war history party result change morning reason research girl guy moment air teacher force education foot boy age policy process ' +
  'music market sense nation plan college interest death experience effect use class control care field development role effort rate ' +
  'heart drug show leader light voice wife police mind price report decision son view relationship town road arm difference value ' +
  'building action model season society tax director position player record paper space ground form event official matter center couple ' +
  'site project activity star table need court oil situation cost industry figure street image phone data picture practice piece land').split(' ');

function random(seed: number) {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function sentences(next: () => number, chars: number) {
  const out: string[] = [];
  let length = 0;
  while (length < chars) {
    const count = 8 + Math.floor(next() * 12);
    const words = Array.from({ length: count }, () => WORDS[Math.floor(next() * WORDS.length)]!);
    const sentence = `${words[0]![0]!.toUpperCase()}${words.join(' ').slice(1)}.`;
    out.push(sentence);
    length += sentence.length + 1;
  }
  return out.join(' ');
}

const escape = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export interface ControlParagraph { kind: 'target' | 'heading' | 'body' | 'list' | 'cell'; text: string; pageBreakBefore?: boolean }

// The ipx richDocument generator, unchanged in output, that also returns the paragraph text in document order.
export async function richDocument(units: number) {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const next = random(23 + units);
  const paragraphs: ControlParagraph[] = [];
  const target = (text: string, breakBefore: boolean) => {
    paragraphs.push({ kind: 'target', text, pageBreakBefore: breakBefore });
    return `<w:p><w:pPr>${breakBefore ? '<w:pageBreakBefore/>' : ''}<w:spacing w:after="600"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  };
  const para = (chars: number) => {
    const text = sentences(next, chars);
    paragraphs.push({ kind: 'body', text });
    return `<w:p><w:r><w:t xml:space="preserve">${escape(text)}</w:t></w:r></w:p>`;
  };
  const cell = (text: string) => {
    paragraphs.push({ kind: 'cell', text });
    return `<w:tc><w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${escape(text)}</w:t></w:r></w:p></w:tc>`;
  };
  const border = '<w:top w:val="single" w:sz="4" w:color="808080"/><w:left w:val="single" w:sz="4" w:color="808080"/><w:bottom w:val="single" w:sz="4" w:color="808080"/><w:right w:val="single" w:sz="4" w:color="808080"/><w:insideH w:val="single" w:sz="4" w:color="808080"/><w:insideV w:val="single" w:sz="4" w:color="808080"/>';
  const body: string[] = [target('Probe', false)];
  for (let unit = 0; unit < units; unit++) {
    if (unit === Math.floor(units / 2)) body.push(target('Marker', true));
    const heading = `${unit + 1}. ${sentences(next, 30).slice(0, 44)}`;
    paragraphs.push({ kind: 'heading', text: heading });
    body.push(`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${escape(heading)}</w:t></w:r></w:p>`);
    for (let i = 0; i < 5; i++) body.push(para(360 + Math.floor(next() * 240)));
    for (let i = 0; i < 4; i++) {
      const item = sentences(next, 60 + Math.floor(next() * 80));
      paragraphs.push({ kind: 'list', text: item });
      body.push(`<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${i % 2}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">${escape(item)}</w:t></w:r></w:p>`);
    }
    const rows = Array.from({ length: 5 }, (_, r) => `<w:tr>${[0, 1, 2].map((c) => cell(r === 0 ? `Column ${c + 1}` : sentences(next, 20 + Math.floor(next() * 50)))).join('')}</w:tr>`);
    body.push(`<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblBorders>${border}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/></w:tblGrid>${rows.join('')}</w:tbl>`);
    for (let i = 0; i < 5; i++) body.push(para(360 + Math.floor(next() * 240)));
  }
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>');
  zip.file('word/numbering.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`);
  zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style></w:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body.join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  zip.forEach((_path, file) => { file.date = new Date('2020-01-01T00:00:00Z'); });
  return { docx: await zip.generateAsync({ type: 'uint8array' }), paragraphs };
}
