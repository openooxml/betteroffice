export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
export const V = 'urn:schemas-microsoft-com:vml';
export const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

export const WORD_2007_ROOTS = [
  { name: 've without Ignorable', root: `<w:document xmlns:w="${W}" xmlns:ve="${MC}">` },
  { name: 've with Ignorable and drawing aliases', root: `<w:document xmlns:w="${W}" xmlns:ve="${MC}" xmlns:draw="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:legacy="urn:schemas-microsoft-com:office:word" ve:Ignorable="draw legacy">` },
  { name: 'two MC aliases', root: `<w:document xmlns:w="${W}" xmlns:ve="${MC}" xmlns:mc="${MC}">` },
  { name: 'default document tasks namespace', root: `<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns="http://schemas.microsoft.com/office/tasks/2019/documenttasks">` },
];

const rich = (id: string, text: string) =>
  `<w:p w14:paraId="${id}" w:rsidR="00752A6D" w:rsidRDefault="00593C67"><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr><w:proofErr w:type="spellStart"/><w:r w:rsidR="00761B10"><w:rPr><w:color w:val="auto"/><w:lang w:val="en-US"/></w:rPr><w:t>${text}</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:br/></w:r><w:proofErr w:type="spellEnd"/></w:p>`;

export function word2007Body(): string {
  return `${rich('10000001', 'Replace target words')}${rich('10000002', 'Retained paragraph words')}`.replace(/ w14:paraId="[^"]*"/g, '');
}

const textbox = (id: string) =>
  `<w:pict><v:shape id="${id}" style="width:72pt;height:36pt"><v:textbox><w:txbxContent><w:p w:rsidR="00112233"><w:r><w:rPr><w:lang w:val="en-CA"/></w:rPr><w:t>Retained ${id}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict>`;
export const DRAWING = `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${textbox('choice')}</mc:Choice><mc:Fallback>${textbox('fallback')}</mc:Fallback></mc:AlternateContent></w:r>`;

export const ALTERNATE_TEXTBOX_RUN = '<w:r w:rsidR="00112233"><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="41" name="Retained text box"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p w14:paraId="40000009"><w:r><w:t>Box text</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="Retained text box" style="width:72pt;height:36pt" type="#_x0000_t202"><v:textbox><w:txbxContent><w:p w14:paraId="40000008"><w:r><w:t>Box text</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>';

export function fieldLabelBody(): string {
  const opening = '<w:p w14:paraId="20000000"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>';
  const entries = [1, 2, 3].map(index => `<w:p w14:paraId="2000000${index}"><w:hyperlink w:anchor="_Toc${index}" w:history="1"><w:r><w:t>Background ${index}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Toc${index} \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${index}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>`).join('');
  const end = '<w:p w14:paraId="20000009"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
  return `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>${opening}${entries}${end}</w:sdtContent></w:sdt>${rich('30000001', 'After contents')}`;
}
const protectedParagraph = rich('1000000F', 'Protected text').replace('</w:p>', `${DRAWING}</w:p>`);

export function bodyMarkerBody(): string {
  const first = rich('10000001', 'Bookmarked title').replace('<w:proofErr', '<w:bookmarkStart w:id="0" w:name="Range"/><w:proofErr');
  const anchor = '<w:p w14:paraId="10000002" w:rsidR="003E0150"><w:pPr><w:rPr><w:lang w:val="en-GB"/></w:rPr></w:pPr></w:p>';
  return `${first}${anchor}<w:bookmarkEnd w:id="0"/>${protectedParagraph}`;
}

export function tocSdtBody(): string {
  const first = rich('10000001', 'Before contents').replace('<w:proofErr', '<w:bookmarkStart w:id="0" w:name="Range"/><w:proofErr');
  const anchor = rich('10000002', 'Contents anchor').replace('</w:p>', '<w:r><w:rPr><w:lang w:val="en-CA"/></w:rPr><w:br w:type="page"/></w:r></w:p>');
  const entry = '<w:p w14:paraId="10000003" w:rsidR="00AABBCC"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:webHidden/><w:lang w:val="en-CA"/></w:rPr><w:t>Contents entry</w:t></w:r></w:hyperlink></w:p>';
  const end = '<w:p w14:paraId="10000004"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
  return `${first}<w:bookmarkEnd w:id="0"/>${anchor}<w:sdt><w:sdtPr><w:id w:val="42"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>${entry}${protectedParagraph}${end}</w:sdtContent></w:sdt>`;
}

export function orphanTableMarkersBody(): string {
  const first = rich('10000001', 'Bookmarked heading').replace('<w:proofErr', '<w:bookmarkStart w:id="0" w:name="_Toc1"/><w:bookmarkStart w:id="1" w:name="_Toc2"/><w:bookmarkStart w:id="2" w:name="_GoBack"/><w:bookmarkEnd w:id="2"/><w:bookmarkEnd w:id="0"/><w:bookmarkEnd w:id="1"/><w:proofErr');
  const anchor = rich('10000002', 'Table anchor');
  const cell = (properties: string, paragraph: string) => `<w:tc><w:tcPr>${properties}</w:tcPr>${paragraph}</w:tc>`;
  const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell('<w:vMerge w:val="restart"/>', rich('10000003', 'Merged start'))}${cell('', rich('10000004', 'Right cell'))}<w:bookmarkEnd w:id="3"/></w:tr><w:bookmarkEnd w:id="4"/><w:tr>${cell('<w:vMerge/>', rich('10000005', 'Hidden continuation'))}${cell('', rich('10000006', 'Last cell'))}</w:tr><w:bookmarkEnd w:id="5"/></w:tbl>`;
  return `${first}${anchor}${table}${protectedParagraph}`;
}

export function offsetTableBody(): string {
  const cell = (id: number, text: string, span = '') => `<w:tc><w:tcPr><w:tcW w:w="887" w:type="dxa"/>${span}<w:tcBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:color="auto"/></w:tcBorders></w:tcPr>${rich(`1000${String(id).padStart(4, '0')}`, text)}</w:tc>`;
  const offsets = '<w:trPr><w:gridAfter w:val="1"/><w:wAfter w:w="7" w:type="dxa"/></w:trPr>';
  const first = `<w:tr>${offsets}${cell(10, 'sample')}${Array.from({ length: 15 }, (_, at) => cell(20 + at, `H${at}`)).join('')}</w:tr>`;
  const spanned = `<w:tr>${cell(50, 'Spanned row', '<w:gridSpan w:val="17"/>')}</w:tr>`;
  const last = `<w:tr>${offsets}${Array.from({ length: 16 }, (_, at) => cell(60 + at, `S${at}`)).join('')}</w:tr>`;
  return `${protectedParagraph}<w:tbl><w:tblPr><w:tblStyle w:val="Tabellenraster"/><w:tblW w:w="14430" w:type="dxa"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="887"/>'.repeat(17)}</w:tblGrid>${first}${spanned}${last}</w:tbl>`;
}

export function plainTableBody(): string {
  const cell = (id: string, text: string) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${rich(id, text)}</w:tc>`;
  return `${protectedParagraph}<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell('20000001', 'First cell')}${cell('20000002', 'Second cell')}</w:tr><w:tr>${cell('20000003', 'Third cell')}${cell('20000004', 'Last cell')}</w:tr></w:tbl>`;
}

export function paragraphsOf(xml: string): string[] {
  const found: Array<[number, string]> = [];
  const open: number[] = [];
  for (const match of xml.matchAll(/<(\/?)([A-Za-z0-9:.-]+)((?:[^>"']|"[^"]*"|'[^']*')*)(\/?)>/g)) {
    if (match[2] !== 'w:p') continue;
    const at = match.index!;
    if (match[1]) {
      const start = open.pop();
      if (start !== undefined) found.push([start, xml.slice(start, at + match[0].length)]);
    } else if (match[4]) {
      found.push([at, match[0]]);
    } else {
      open.push(at);
    }
  }
  return found.sort((a, b) => a[0] - b[0]).map(([, xml]) => xml);
}

export function textNodes(xml: string): string[] {
  return [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map(match => match[1]!);
}
