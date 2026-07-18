/**
 * Generate the BetterOffice first-open demo document.
 *
 * This is the sample DOCX the OpenOOXML editor loads on first open. It is an
 * original showcase authored here (no inherited content) whose purpose is to
 * exercise the editor's rendering breadth: Title/Heading style hierarchy,
 * character formatting (bold/italic/underline/color/highlight), bulleted and
 * numbered lists, a shaded-header table, a bordered/shaded blockquote, a
 * centered caption, and a header + footer.
 *
 * The OOXML parts are hand-authored, zipped with JSZip, and written to the
 * public/ folder of BOTH example apps.
 *
 * Run: bun scripts/create-demo-doc.mjs
 */

import JSZip from 'jszip';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT_PATHS = [path.join(ROOT, 'apps/demo/public/betteroffice-demo.docx')];
const ZIP_DATE = new Date('2026-01-01T00:00:00Z');

/** Escape a string for use as XML text or an attribute value. */
function esc(s: string): string {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ---------------------------------------------------------------------------
// Package-level parts
// ---------------------------------------------------------------------------

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const DOCUMENT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`;

const CORE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Welcome to BetterOffice</dc:title>
  <dc:subject>BetterOffice demo document</dc:subject>
  <dc:creator>OpenOOXML</dc:creator>
  <cp:keywords>docx, editor, openooxml, betteroffice</cp:keywords>
  <cp:lastModifiedBy>OpenOOXML</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr><w:spacing w:after="140" w:line="276" w:lineRule="auto"/></w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
      <w:contextualSpacing/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/>
      <w:color w:val="1F3864"/>
      <w:spacing w:val="-10"/>
      <w:kern w:val="28"/>
      <w:sz w:val="56"/>
      <w:szCs w:val="56"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="40" w:after="240"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/>
      <w:i/>
      <w:iCs/>
      <w:color w:val="595959"/>
      <w:spacing w:val="15"/>
      <w:sz w:val="30"/>
      <w:szCs w:val="30"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="280" w:after="80"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="2E74B5"/>
      <w:sz w:val="32"/>
      <w:szCs w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="200" w:after="40"/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="2E74B5"/>
      <w:sz w:val="26"/>
      <w:szCs w:val="26"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="60"/>
      <w:ind w:left="720"/>
      <w:contextualSpacing/>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Quote">
    <w:name w:val="Quote"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:i/>
      <w:iCs/>
      <w:color w:val="404040"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Caption">
    <w:name w:val="Caption"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="0" w:after="200"/>
    </w:pPr>
    <w:rPr>
      <w:i/>
      <w:iCs/>
      <w:color w:val="595959"/>
      <w:sz w:val="18"/>
      <w:szCs w:val="18"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:rPr>
      <w:color w:val="0563C1"/>
      <w:u w:val="single"/>
    </w:rPr>
  </w:style>
</w:styles>`;

// ---------------------------------------------------------------------------
// Numbering (one bullet list + one decimal list)
// ---------------------------------------------------------------------------

const NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#8226;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#9702;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

// ---------------------------------------------------------------------------
// Header / Footer
// ---------------------------------------------------------------------------

const HEADER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr>
      <w:pBdr><w:bottom w:val="single" w:sz="4" w:space="4" w:color="BFBFBF"/></w:pBdr>
      <w:spacing w:after="0"/>
      <w:jc w:val="right"/>
    </w:pPr>
    <w:r>
      <w:rPr><w:color w:val="808080"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
      <w:t>Welcome to BetterOffice</w:t>
    </w:r>
  </w:p>
</w:hdr>`;

const FOOTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr>
      <w:pBdr><w:top w:val="single" w:sz="4" w:space="4" w:color="BFBFBF"/></w:pBdr>
      <w:spacing w:before="0" w:after="0"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:r>
      <w:rPr><w:color w:val="808080"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
      <w:t xml:space="preserve">Created with BetterOffice </w:t>
    </w:r>
    <w:r>
      <w:rPr><w:color w:val="808080"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
      <w:t>&#8226;</w:t>
    </w:r>
    <w:r>
      <w:rPr><w:color w:val="808080"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
      <w:t xml:space="preserve"> openooxml.org</w:t>
    </w:r>
  </w:p>
</w:ftr>`;

// ---------------------------------------------------------------------------
// Document body builders
// ---------------------------------------------------------------------------

/** A run with optional run-properties. `props` is raw rPr child XML. */
function run(text: string, props = ''): string {
  const rPr = props ? `<w:rPr>${props}</w:rPr>` : '';
  return `<w:r>${rPr}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

/** A styled paragraph. `runs` is pre-built run XML. `pPrExtra` is raw pPr child XML. */
function para(styleId: string, runs: string, pPrExtra = ''): string {
  const style = styleId ? `<w:pStyle w:val="${styleId}"/>` : '';
  return `<w:p><w:pPr>${style}${pPrExtra}</w:pPr>${runs}</w:p>`;
}

/** A list item paragraph bound to numId at the given level. */
function listItem(numId: number, text: string, ilvl = 0): string {
  return para(
    'ListParagraph',
    run(text),
    `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`,
  );
}

/** A table cell. `shadeFill` optional hex; `runs` pre-built. */
function cell(width: number, runs: string, shadeFill?: string): string {
  const shd = shadeFill
    ? `<w:shd w:val="clear" w:color="auto" w:fill="${shadeFill}"/>`
    : '';
  const tcPr = `<w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${shd}<w:vAlign w:val="center"/></w:tcPr>`;
  const cellPara = `<w:p><w:pPr><w:spacing w:before="20" w:after="20" w:line="240" w:lineRule="auto"/></w:pPr>${runs}</w:p>`;
  return `<w:tc>${tcPr}${cellPara}</w:tc>`;
}

const COL = [2700, 4500, 2160]; // sums to 9360 (Letter usable width @ 1in margins)

// Build header/body cells explicitly so each column keeps its width.
function hCell(width: number, text: string): string {
  return cell(width, run(text, '<w:b/><w:bCs/><w:color w:val="FFFFFF"/>'), '2E74B5');
}
function bCell(width: number, text: string, opts: { bold?: boolean; color?: string } = {}): string {
  const props =
    (opts.bold ? '<w:b/><w:bCs/>' : '') + (opts.color ? `<w:color w:val="${opts.color}"/>` : '');
  return cell(width, run(text, props));
}

function tableRow(cells: string[]): string {
  return `<w:tr>${cells.join('')}</w:tr>`;
}

const TABLE = `<w:tbl>
  <w:tblPr>
    <w:tblW w:w="9360" w:type="dxa"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="40" w:type="dxa"/>
      <w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="40" w:type="dxa"/>
      <w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar>
    <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="${COL[0]}"/>
    <w:gridCol w:w="${COL[1]}"/>
    <w:gridCol w:w="${COL[2]}"/>
  </w:tblGrid>
  ${tableRow([hCell(COL[0], 'Stage'), hCell(COL[1], 'What happens'), hCell(COL[2], 'Where it runs')])}
  ${tableRow([
    bCell(COL[0], 'Parse & save', { bold: true }),
    bCell(COL[1], 'OOXML in, OOXML out — the parser and serializer round-trip your file byte-faithfully.'),
    bCell(COL[2], 'Rust (wasm)', { color: '2E7D32' }),
  ])}
  ${tableRow([
    bCell(COL[0], 'Editing core', { bold: true }),
    bCell(COL[1], 'Document state, selections, undo, and tracked changes live in a CRDT engine.'),
    bCell(COL[2], 'Rust (wasm)', { color: '2E7D32' }),
  ])}
  ${tableRow([
    bCell(COL[0], 'Text & layout', { bold: true }),
    bCell(COL[1], 'Font shaping, line breaking, and pagination — no browser layout involved.'),
    bCell(COL[2], 'Rust (wasm)', { color: '2E7D32' }),
  ])}
  ${tableRow([
    bCell(COL[0], 'Rendering', { bold: true }),
    bCell(COL[1], 'The engine emits a display list that is replayed onto page canvases.'),
    bCell(COL[2], 'Canvas + worker', { color: '2E7D32' }),
  ])}
</w:tbl>`;

// Mixed character-formatting paragraph.
const MIXED_RUNS =
  run('The Rust text shaper keeps character formatting intact: ') +
  run('bold', '<w:b/><w:bCs/>') +
  run(', ') +
  run('italic', '<w:i/><w:iCs/>') +
  run(', ') +
  run('underline', '<w:u w:val="single"/>') +
  run(', a ') +
  run('colored run', '<w:color w:val="C00000"/>') +
  run(', and even a ') +
  run('highlighted run', '<w:highlight w:val="yellow"/>') +
  run(' all survive the round trip.');

const BODY = [
  para('Title', run('Welcome to BetterOffice')),
  para('Subtitle', run('A faithful, open .docx editor — running a native Rust engine in your browser')),
  para(
    'Normal',
    run(
      'BetterOffice is a client-side editor for .docx documents with an unusual architecture: parsing, editing, text shaping, and page layout all run in Rust, compiled to WebAssembly. Your file never leaves your device — there is no upload and no server round trip — yet every page is laid out with the fidelity of a desktop word processor.',
    ),
  ),
  para(
    'Normal',
    run(
      'The page you are reading is not HTML pretending to be a document. A native layout engine measured every line and painted this page onto a canvas — the same way a desktop word processor would. Try selecting text, changing a heading, or editing a table cell: the engine reflows the page live.',
    ),
  ),

  para('Heading1', run('A native engine, compiled to WebAssembly')),
  para(
    'Normal',
    run(
      'Opening a file runs the whole pipeline in Rust: the parser reads the OOXML into a typed document model, the text engine shapes every run with real font bytes, the paginator breaks lines into pages, and the result is emitted as a display list that the browser simply replays onto canvas. The DOM is never asked to lay out a document — fidelity does not depend on the browser.',
    ),
  ),

  para('Heading2', run('Character formatting')),
  para('Normal', MIXED_RUNS),

  para('Heading2', run('What runs in Rust')),
  para('Normal', run('The engine owns the document end to end:')),
  listItem(1, 'An OOXML parser and serializer that round-trip .docx files byte-faithfully'),
  listItem(1, 'Text shaping and measurement with real font bytes — never the browser'),
  listItem(1, 'Pagination: line breaking, tables, floats, headers and footers'),
  listItem(1, 'A CRDT editing core, built for multiplayer and agent collaboration'),
  para('Normal', run('And this is what happens on every keystroke:')),
  listItem(2, 'Your input lands in the resident engine as a single WebAssembly call.'),
  listItem(2, 'The engine edits the document and re-measures only the dirty paragraphs.'),
  listItem(2, 'A binary frame delta repaints just the damaged part of the page.'),

  para('Heading1', run('Architecture at a glance')),
  para('Normal', run('Where each stage of the pipeline actually runs.')),
  TABLE,
  para('Caption', run('Table 1. The pipeline — everything before the canvas is native Rust.'), '<w:jc w:val="center"/>'),

  para('Heading2', run('Why it matters')),
  para(
    'Quote',
    run(
      'Fidelity needs a real engine. Because layout runs in Rust rather than the DOM, the same code can paint this page in every browser today — and in native apps tomorrow.',
    ),
    '<w:pBdr><w:left w:val="single" w:sz="18" w:space="10" w:color="2E74B5"/></w:pBdr>' +
      '<w:shd w:val="clear" w:color="auto" w:fill="F2F6FB"/>' +
      '<w:ind w:left="360" w:right="360"/>' +
      '<w:spacing w:before="120" w:after="120"/>',
  ),

  para(
    'Normal',
    run('BetterOffice is open source and maintained by the OpenOOXML project. Explore the code, file an issue, or just keep typing — everything you see is being computed by Rust, right here in your browser.'),
  ),
].join('\n    ');

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${BODY}
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId3"/>
      <w:footerReference w:type="default" r:id="rId4"/>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>`;

// ---------------------------------------------------------------------------
// Assemble + write
// ---------------------------------------------------------------------------

const zip = new JSZip();
const opts = { date: ZIP_DATE, createFolders: false };
zip.file('[Content_Types].xml', CONTENT_TYPES_XML, opts);
zip.file('_rels/.rels', RELS_XML, opts);
zip.file('docProps/core.xml', CORE_XML, opts);
zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS_XML, opts);
zip.file('word/document.xml', DOCUMENT_XML, opts);
zip.file('word/styles.xml', STYLES_XML, opts);
zip.file('word/numbering.xml', NUMBERING_XML, opts);
zip.file('word/header1.xml', HEADER_XML, opts);
zip.file('word/footer1.xml', FOOTER_XML, opts);

const buffer = await zip.generateAsync({
  type: 'nodebuffer',
  compression: 'DEFLATE',
  compressionOptions: { level: 9 },
});

for (const out of OUT_PATHS) {
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, buffer);
  console.log(`Created ${out} (${buffer.length} bytes)`);
}
