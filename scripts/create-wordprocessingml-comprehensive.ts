/** Generate the corpus fixture: bun scripts/create-wordprocessingml-comprehensive.ts [output.docx]. */

import JSZip from 'jszip';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = process.argv[2]
  ? path.resolve(process.argv[2])
  : path.join(ROOT, 'crates/betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx');
const ZIP_DATE = new Date('2026-01-01T00:00:00Z');

const NS_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const NS_FULL = [
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:w10="urn:schemas-microsoft-com:office:word"',
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:bofx="urn:betteroffice-fixture-x"',
  'mc:Ignorable="w14 w15 bofx"',
].join(' ');

// 1x1 PNG.
const IMAGE_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
);

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/header3.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`;

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`;

const DOCUMENT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdHeader2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/><Relationship Id="rIdHeader3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header3.xml"/><Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://betteroffice.dev/" TargetMode="External"/></Relationships>`;

const CORE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>WordprocessingML comprehensive</dc:title><dc:creator>corpus fixture generator</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>`;

const SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings ${NS_W}><w:zoom w:percent="100"/><w:defaultTabStop w:val="708"/><w:evenAndOddHeaders/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${NS_W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Yu Mincho" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="80"/><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style><w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:qFormat/><w:rPr><w:i/><w:iCs/></w:rPr></w:style><w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style><w:style w:type="table" w:styleId="GridTable"><w:name w:val="Grid Table"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="666666"/><w:left w:val="single" w:sz="4" w:space="0" w:color="666666"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="666666"/><w:right w:val="single" w:sz="4" w:space="0" w:color="666666"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="666666"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="666666"/></w:tblBorders></w:tblPr><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr></w:style></w:styles>`;

const NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${NS_W}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="480"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="10"/></w:lvlOverride></w:num></w:numbering>`;

const FOOTNOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes ${NS_W}><w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> A footnote with </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve"> text.</w:t></w:r></w:p></w:footnote></w:footnotes>`;

const ENDNOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes ${NS_W}><w:endnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="1"><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> An endnote.</w:t></w:r></w:p></w:endnote></w:endnotes>`;

const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments ${NS_W}><w:comment w:id="0" w:author="Reviewer A" w:date="2026-01-01T00:00:00Z" w:initials="RA"><w:p><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>Please check this phrase.</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="Reviewer B" w:date="2026-01-01T01:00:00Z" w:initials="RB"><w:p><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>Checked, looks right.</w:t></w:r></w:p></w:comment></w:comments>`;

const HEADER1_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${NS_W}><w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0"/></w:pPr><w:r><w:t>WordprocessingML comprehensive — default header</w:t></w:r><w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`;

const HEADER2_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${NS_W}><w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>First-page header</w:t></w:r></w:p></w:hdr>`;

const HEADER3_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${NS_W}><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Even-page header</w:t></w:r></w:p></w:hdr>`;

const FOOTER1_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr ${NS_W}><w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>4</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`;

function heading(text: string, bookmark?: { id: number; name: string }): string {
  const open = bookmark ? `<w:bookmarkStart w:id="${bookmark.id}" w:name="${bookmark.name}"/>` : '';
  const close = bookmark ? `<w:bookmarkEnd w:id="${bookmark.id}"/>` : '';
  return `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${open}<w:r><w:t>${text}</w:t></w:r>${close}</w:p>`;
}

const TITLE = `<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>WordprocessingML comprehensive</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>One document, every corpus criterion this generator can author.</w:t></w:r></w:p>`;

const TOC = `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Contents</w:t></w:r></w:p>
<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-1" \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="sec1"><w:r><w:t>1. Character formatting</w:t></w:r></w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;

const S1 = `${heading('1. Character formatting', { id: 1, name: 'sec1' })}
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>single underline</w:t></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:rPr><w:u w:val="double"/></w:rPr><w:t>double</w:t></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:rPr><w:u w:val="wave" w:color="FF0000"/></w:rPr><w:t>red wave</w:t></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:rPr><w:strike/></w:rPr><w:t>strike</w:t></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:rPr><w:dstrike/></w:rPr><w:t>double strike</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:caps/></w:rPr><w:t>all caps</w:t></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:rPr><w:smallCaps/></w:rPr><w:t>Small Caps</w:t></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden text</w:t></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:rPr><w:color w:val="C00000"/></w:rPr><w:t>dark red</w:t></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>highlighted</w:t></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="DEEAF6"/></w:rPr><w:t>shaded run</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">x</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r><w:r><w:t xml:space="preserve"> and H</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>2</w:t></w:r><w:r><w:t xml:space="preserve">O; </w:t></w:r><w:r><w:rPr><w:spacing w:val="40"/></w:rPr><w:t>expanded spacing</w:t></w:r><w:r><w:t xml:space="preserve">; </w:t></w:r><w:r><w:rPr><w:w w:val="150"/></w:rPr><w:t>stretched</w:t></w:r><w:r><w:t xml:space="preserve">; </w:t></w:r><w:r><w:rPr><w:position w:val="6"/></w:rPr><w:t>raised</w:t></w:r><w:r><w:t xml:space="preserve">; </w:t></w:r><w:r><w:rPr><w:kern w:val="28"/><w:sz w:val="28"/></w:rPr><w:t>kerned</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>Character style</w:t></w:r><w:r><w:t xml:space="preserve">; monospace </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t>Courier New</w:t></w:r><w:r><w:t xml:space="preserve">; preserved   spaces; a line</w:t></w:r><w:r><w:br/></w:r><w:r><w:t xml:space="preserve">break; soft</w:t></w:r><w:r><w:softHyphen/></w:r><w:r><w:t xml:space="preserve">hyphen; non</w:t></w:r><w:r><w:noBreakHyphen/></w:r><w:r><w:t xml:space="preserve">breaking; symbol </w:t></w:r><w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r><w:r><w:t>.</w:t></w:r></w:p>`;

const S2 = `${heading('2. Paragraph formatting')}
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Centered paragraph.</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Right-aligned paragraph.</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>Justified paragraph with enough words that justification has something visible to distribute across the measure of the line when it wraps onto the next line of the page.</w:t></w:r></w:p>
<w:p><w:pPr><w:ind w:left="720" w:right="720" w:hanging="360"/></w:pPr><w:r><w:t>Hanging indent paragraph: the first line starts left of the following lines, with symmetric left and right indents applied to the whole paragraph body.</w:t></w:r></w:p>
<w:p><w:pPr><w:spacing w:before="240" w:after="240" w:line="480" w:lineRule="exact"/></w:pPr><w:r><w:t>Exact 24pt line height with 12pt space before and after; a second sentence keeps two lines on the page so the rule is visible.</w:t></w:r></w:p>
<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="8" w:space="4" w:color="4472C4"/><w:left w:val="single" w:sz="8" w:space="4" w:color="4472C4"/><w:bottom w:val="single" w:sz="8" w:space="4" w:color="4472C4"/><w:right w:val="single" w:sz="8" w:space="4" w:color="4472C4"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="F2F6FC"/></w:pPr><w:r><w:t>Boxed and shaded paragraph.</w:t></w:r></w:p>
<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="2160" w:leader="dot"/><w:tab w:val="right" w:pos="8640" w:leader="underscore"/></w:tabs></w:pPr><w:r><w:t>Start</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>dotted stop</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>right stop</w:t></w:r></w:p>
<w:p><w:pPr><w:keepNext/><w:keepLines/><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:t>Keep-with-next paragraph whose mark carries bold 12pt properties.</w:t></w:r></w:p>
<w:p><w:r><w:t>Its companion paragraph, kept on the same page.</w:t></w:r></w:p>`;

const S3 = `${heading('3. Lists')}
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>First numbered item</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Nested item one.one</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Nested item one.two</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Second numbered item</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Bulleted item</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Another bullet</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Restarted list beginning at ten</w:t></w:r></w:p>`;

const S4 = `${heading('4. Tables')}
<w:tbl><w:tblPr><w:tblStyle w:val="GridTable"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Header A</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Header B</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Header C</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="6240" w:type="dxa"/><w:gridSpan w:val="2"/><w:shd w:val="clear" w:color="auto" w:fill="FBE4D5"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Spans two columns, shaded</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3120" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Merged down, centered</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="3120" w:type="dxa"/><w:tcBorders><w:bottom w:val="nil"/></w:tcBorders></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Explicit nil border below</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3120" w:type="dxa"/><w:textDirection w:val="btLr"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Vertical text</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3120" w:type="dxa"/><w:vMerge/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p></w:tc></w:tr></w:tbl>
<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>
<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:left w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:right w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="999999"/></w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Outer cell with a nested table:</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="dashed" w:sz="4" w:space="0" w:color="C00000"/><w:left w:val="dashed" w:sz="4" w:space="0" w:color="C00000"/><w:bottom w:val="dashed" w:sz="4" w:space="0" w:color="C00000"/><w:right w:val="dashed" w:sz="4" w:space="0" w:color="C00000"/><w:insideH w:val="dashed" w:sz="4" w:space="0" w:color="C00000"/><w:insideV w:val="dashed" w:sz="4" w:space="0" w:color="C00000"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2160"/><w:gridCol w:w="2160"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>inner 1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>inner 2</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/><w:tcMar><w:top w:w="216" w:type="dxa"/><w:left w:w="216" w:type="dxa"/><w:bottom w:w="216" w:type="dxa"/><w:right w:w="216" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Cell with wide custom margins.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>`;

const SECT_BREAK_1 = `<w:p><w:pPr><w:sectPr><w:headerReference w:type="even" r:id="rIdHeader3"/><w:headerReference w:type="default" r:id="rIdHeader1"/><w:headerReference w:type="first" r:id="rIdHeader2"/><w:footerReference w:type="default" r:id="rIdFooter1"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:pgNumType w:start="1"/><w:cols w:space="708"/><w:titlePg/></w:sectPr></w:pPr></w:p>`;

const S5 = `${heading('5. Sections and columns')}
<w:p><w:r><w:t>This section is landscape with two columns. The text fills the first column and then continues into the second column after an explicit column break, which is the deterministic way to prove column flow without depending on measurement.</w:t></w:r></w:p>
<w:p><w:r><w:t>Still the first column.</w:t></w:r><w:r><w:br w:type="column"/></w:r><w:r><w:t>This text starts the second column.</w:t></w:r></w:p>`;

const SECT_BREAK_2 = `<w:p><w:pPr><w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:num="2" w:space="708"/></w:sectPr></w:pPr></w:p>`;

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const PIC = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

const INLINE_IMAGE = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="457200" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Inline image" descr="A tiny inline square"/><wp:cNvGraphicFramePr><a:graphicFrameLocks ${A} noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic ${A}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic ${PIC}><pic:nvPicPr><pic:cNvPr id="1" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;

const ANCHORED_IMAGE = `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="2" name="Anchored image" descr="A square anchored to the right margin"/><wp:cNvGraphicFramePr><a:graphicFrameLocks ${A} noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic ${A}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic ${PIC}><pic:nvPicPr><pic:cNvPr id="2" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>`;

const WPS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

const TEXT_BOX = `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658241" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2286000" cy="685800"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="3" name="Text box"/><wp:cNvGraphicFramePr/><a:graphic ${A}><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp ${WPS}><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="685800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="BF9000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Text box</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>A wrapped-square shape with its own story.</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;

const S6 = `${heading('6. Drawings and text boxes')}
<w:p><w:r><w:t xml:space="preserve">An inline image sits in this sentence </w:t></w:r><w:r>${INLINE_IMAGE}</w:r><w:r><w:t xml:space="preserve"> between two runs of text.</w:t></w:r></w:p>
<w:p><w:r>${ANCHORED_IMAGE}</w:r><w:r><w:t>This paragraph hosts an image anchored to the right margin with square wrapping, so its lines should shorten to flow around the picture while it stays pinned to the margin.</w:t></w:r></w:p>
<w:p><w:r>${TEXT_BOX}</w:r><w:r><w:t>This paragraph hosts a yellow text box anchored to the left margin; body text wraps around it on the right. The box carries two paragraphs of its own story.</w:t></w:r></w:p>`;

const S7 = `${heading('7. Footnotes and endnotes')}
<w:p><w:r><w:t xml:space="preserve">A claim that needs a source</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> and a remark deferred to the end</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="1"/></w:r><w:r><w:t>.</w:t></w:r></w:p>`;

const S8 = `${heading('8. Fields, links, and bookmarks')}
<w:p><w:r><w:t xml:space="preserve">External link: </w:t></w:r><w:hyperlink r:id="rIdLink" w:history="1"><w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>betteroffice.dev</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve">. Internal link: </w:t></w:r><w:hyperlink w:anchor="sec1"><w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>back to section 1</w:t></w:r></w:hyperlink><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Simple field: page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>3</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve">. Complex field: created in </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE \\@ "yyyy" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2026</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:bookmarkStart w:id="10" w:name="spanning"/><w:r><w:t>A bookmark opens in this paragraph</w:t></w:r></w:p>
<w:p><w:r><w:t>and closes in this one.</w:t></w:r><w:bookmarkEnd w:id="10"/></w:p>`;

const S9 = `${heading('9. Tracked changes')}
<w:p><w:r><w:t xml:space="preserve">This sentence has </w:t></w:r><w:ins w:id="100" w:author="Reviewer A" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">an inserted phrase </w:t></w:r></w:ins><w:r><w:t xml:space="preserve">and </w:t></w:r><w:del w:id="101" w:author="Reviewer A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve">a deleted one </w:delText></w:r></w:del><w:r><w:t xml:space="preserve">plus </w:t></w:r><w:r><w:rPr><w:b/><w:rPrChange w:id="102" w:author="Reviewer B" w:date="2026-01-01T01:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>a formatting change</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>`;

const S10 = `${heading('10. Comments')}
<w:p><w:r><w:t xml:space="preserve">This paragraph contains </w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>a commented phrase</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:t>a second one</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r><w:r><w:t>.</w:t></w:r></w:p>`;

const S11 = `${heading('11. Content controls')}
<w:sdt><w:sdtPr><w:alias w:val="Project name"/><w:tag w:val="project-name"/><w:id w:val="1001"/><w:lock w:val="sdtContentLocked"/><w:richText/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>A locked rich-text control.</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:p><w:r><w:t xml:space="preserve">Inline control: </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="Status"/><w:tag w:val="status"/><w:id w:val="1002"/><w:dropDownList><w:listItem w:displayText="Draft" w:value="draft"/><w:listItem w:displayText="Final" w:value="final"/></w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>Draft</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>.</w:t></w:r></w:p>`;

const S12 = `${heading('12. Math')}
<w:p><m:oMathPara><m:oMath><m:r><m:t>x=</m:t></m:r><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara></w:p>`;

const S13 = `${heading('13. Right-to-left and East Asian text')}
<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>&#1605;&#1585;&#1581;&#1576;&#1575; &#1576;&#1575;&#1604;&#1593;&#1575;&#1604;&#1605;</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>&#26085;&#26412;&#35486;&#12398;&#25991;&#31456;&#12391;&#12377;&#12290;&#31105;&#21063;&#12398;&#12486;&#12473;&#12488;&#12290;</w:t></w:r></w:p>`;

// The foreign namespace is declared at the root and listed in mc:Ignorable:
// Word refuses to open a document carrying non-ignorable unknown markup.
const S14 = `${heading('14. Unknown markup')}
<w:p><w:r><w:t xml:space="preserve">This paragraph carries </w:t></w:r><bofx:marker bofx:kind="inline"/><w:r><w:t xml:space="preserve">a foreign inline element and a foreign attribute on the next paragraph.</w:t></w:r></w:p>
<w:p bofx:origin="fixture"><w:r><w:t>Paragraph with a foreign attribute.</w:t></w:r></w:p>`;

const FINAL_SECT = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="708"/></w:sectPr>`;

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS_FULL}><w:body>${TITLE}
${TOC}
${S1}
${S2}
${S3}
${S4}
${SECT_BREAK_1}
${S5}
${SECT_BREAK_2}
${S6}
${S7}
${S8}
${S9}
${S10}
${S11}
${S12}
${S13}
${S14}
${FINAL_SECT}</w:body></w:document>`;

const zip = new JSZip();
const opts = { date: ZIP_DATE, createFolders: false };
zip.file('[Content_Types].xml', CONTENT_TYPES_XML, opts);
zip.file('_rels/.rels', RELS_XML, opts);
zip.file('docProps/core.xml', CORE_XML, opts);
zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS_XML, opts);
zip.file('word/document.xml', DOCUMENT_XML, opts);
zip.file('word/styles.xml', STYLES_XML, opts);
zip.file('word/numbering.xml', NUMBERING_XML, opts);
zip.file('word/settings.xml', SETTINGS_XML, opts);
zip.file('word/footnotes.xml', FOOTNOTES_XML, opts);
zip.file('word/endnotes.xml', ENDNOTES_XML, opts);
zip.file('word/comments.xml', COMMENTS_XML, opts);
zip.file('word/header1.xml', HEADER1_XML, opts);
zip.file('word/header2.xml', HEADER2_XML, opts);
zip.file('word/header3.xml', HEADER3_XML, opts);
zip.file('word/footer1.xml', FOOTER1_XML, opts);
zip.file('word/media/image1.png', IMAGE_PNG, opts);

const buffer = await zip.generateAsync({
  type: 'nodebuffer',
  compression: 'DEFLATE',
  compressionOptions: { level: 9 },
});

fs.mkdirSync(path.dirname(OUT), { recursive: true });
fs.writeFileSync(OUT, buffer);
console.log(`Created ${OUT} (${buffer.length} bytes)`);
