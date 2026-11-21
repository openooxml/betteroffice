import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { applyFrameDelta, decodeFrameDelta } from '../layout/render/frameDelta';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { preloadParseWasm } from '../wasm/parse';
import { createdListKey } from './listDefinitionSave';

const ROOT = resolve(import.meta.dir, '../../../..');
const FIXTURE = resolve(import.meta.dir, '__fixtures__/seed-parity/lists/word');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const STYLES = `${XML}<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/><w:spacing w:before="240" w:after="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style></w:styles>`;
const BULLET_NUMBERING = `${XML}<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
const NUMBERED_HEADING_STYLES = STYLES.replace(
  '<w:pPr><w:outlineLvl w:val="0"/>',
  '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="0"/>'
);
/** A Heading 1 style numbered by list 1 that gives its left indent in character units only. */
const CHARACTER_HEADING_STYLES = STYLES.replace(
  '<w:pPr><w:outlineLvl w:val="0"/>',
  '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:leftChars="200"/><w:outlineLvl w:val="0"/>'
);
const HEADING_NUMBERING = `${XML}<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
const levelXml = (ilvl: number, format: string, text: string) =>
  `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (ilvl + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
const abstractNumXml = (id: number, levels: string) =>
  `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${levels}</w:abstractNum>`;
const numXml = (numId: number, abstractNumId: number) => `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractNumId}"/></w:num>`;
const numberingXml = (body: string) => `${XML}<w:numbering xmlns:w="${W}">${body}</w:numbering>`;
/** List 1: bullets at levels 0 and 1, decimal `%3.` at level 2. */
const HYBRID_NUMBERING = numberingXml(
  abstractNumXml(0, levelXml(0, 'bullet', '•') + levelXml(1, 'bullet', 'o') + levelXml(2, 'decimal', '%3.')) + numXml(1, 0)
);
/** List 1: a bullet list defining level 0 only. List 2: decimal at levels 0 to 2. */
const SINGLE_LEVEL_NUMBERING = numberingXml(
  abstractNumXml(0, levelXml(0, 'bullet', '•')) +
  abstractNumXml(1, levelXml(0, 'decimal', '%1.') + levelXml(1, 'lowerLetter', '%2.') + levelXml(2, 'decimal', '%3.')) +
  numXml(1, 0) + numXml(2, 1)
);
/** Lists 2 and 7: decimal at level 0 and lowerLetter at level 1, over two different abstract definitions. */
const TWO_DECIMAL_NUMBERING = numberingXml(
  abstractNumXml(0, levelXml(0, 'decimal', '%1.') + levelXml(1, 'lowerLetter', '%2.')) +
  abstractNumXml(1, levelXml(0, 'decimal', '%1.') + levelXml(1, 'lowerLetter', '%2.')) +
  numXml(2, 0) + numXml(7, 1)
);
/** List 2: decimal at level 0 only. List 7: decimal, lowerLetter and lowerRoman at levels 0 to 2. */
const DEEP_SEVEN_NUMBERING = numberingXml(
  abstractNumXml(0, levelXml(0, 'decimal', '%1.')) +
  abstractNumXml(1, levelXml(0, 'decimal', '%1.') + levelXml(1, 'lowerLetter', '%2.') + levelXml(2, 'lowerRoman', '%3.')) +
  numXml(2, 0) + numXml(7, 1)
);
/** List 5: a bullet list whose level 0 is indented 1440 twips with a 360 twip hanging indent. */
const INDENTED_NUMBERING = numberingXml(
  abstractNumXml(0, '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>') +
  numXml(5, 0)
);
const LEFT_ONLY_NUMBERING = numberingXml(
  abstractNumXml(0, '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440"/></w:pPr></w:lvl>') +
  numXml(5, 0)
);
/** List 5: the indented bullet list of `INDENTED_NUMBERING`. List 6: a bullet list whose level 0 gives no indent. */
const CHARACTER_INDENT_NUMBERING = numberingXml(
  abstractNumXml(0, '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>') +
  abstractNumXml(1, '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/></w:lvl>') +
  numXml(5, 0) + numXml(6, 1)
);
/** List 1: decimal, its level 0 indented 1440 twips with a 360 twip hanging indent. List 6: decimal, its level 0 gives no indent. */
const CHARACTER_HEADING_NUMBERING = numberingXml(
  abstractNumXml(0, '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>') +
  abstractNumXml(1, '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>') +
  numXml(1, 0) + numXml(6, 1)
);

/** The paragraph the list commands are applied to; the last paragraph added to each document. */
export const PLAIN = 'Plain paragraph QA';
/** A paragraph on no list, put directly above `PLAIN` when no list may be beside it. */
export const GAP = 'Gap paragraph QA';
/** A paragraph on no list that can follow `PLAIN`. */
export const FOLLOWER = 'Follower paragraph QA';
/** The item of the bullet list 1 that opens document B. */
export const EXISTING_BULLET = 'Existing bullet item QA';
/** An item of the decimal list 1 of document A, directly above an item of the bullet list 2. */
export const DECIMAL_ITEM = 'Second numbered item';
/** The level 2 item of the decimal list 1 of document A directly above `DECIMAL_ITEM`. */
const NESTED_DECIMAL = 'Nested level three';
/** The item of the decimal list 3 of document A directly below a level 1 item of the bullet list 2. */
const RESTARTED_ITEM = 'Restarted list with explicit indent override';
/** The two headings of document D, numbered by their style. */
export const HEADING = 'Introduction heading QA';
const NEXT_HEADING = 'Second heading QA';
const SECOND_FOLLOWER = 'Second follower QA';
const THIRD_FOLLOWER = 'Third follower QA';
/** The level 0 bullet item and the level 2 numbered item below it, both on list 1 of document E. */
const HYBRID_BULLET = 'Hybrid bullet item QA';
const HYBRID_NUMBER = 'Hybrid numbered item QA';
/** The level 0 item of the one-level bullet list 1 of document F, and the level 2 item of its list 2 below. */
const SINGLE_BULLET = 'Single level bullet QA';
const DEEP_NUMBER = 'Deep numbered item QA';
/** The plain paragraph opening document G, and the items of its lists 2 and 7 below it. */
const SELECTION_LEAD = 'Selection lead QA';
const LIST_TWO_ITEM = 'List two item QA';
const LIST_SEVEN_ITEM = 'List seven item QA';
/** The item of the indented bullet list 5 of document H. */
const INDENTED_ITEM = 'Indented list item QA';
/** The item of list 5 in document J whose only indent of its own is a left indent in character units, and the item of list 6. */
const CHARACTER_ITEM = 'Character indent item QA';
const UNINDENTED_ITEM = 'Unindented list item QA';
export const LIST_LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});

/**
 * A: the seed-parity lists document, whose list 1 is decimal, list 2 a bullet list and list 3
 * decimal again, its last item directly above `PLAIN`. B: `EXISTING_BULLET` of bullet list 1,
 * directly above `PLAIN`, and no `w:num` 2. C: no numbering part. D: `PLAIN` between two Heading 1
 * paragraphs numbered by their style's list 1. E: `HYBRID_BULLET` and `HYBRID_NUMBER` of the hybrid
 * list 1. F: `SINGLE_BULLET` of the one-level bullet list 1 above `DEEP_NUMBER` of the decimal list 2.
 * G: `SELECTION_LEAD` above `LIST_TWO_ITEM` at level 0 of the numbered list 2 and `LIST_SEVEN_ITEM`
 * at level 1 of the numbered list 7. H: `INDENTED_ITEM` of the indented bullet list 5. I: as G, with
 * a list 2 that defines level 0 only and `LIST_SEVEN_ITEM` at level 2 of list 7. J: `CHARACTER_ITEM`
 * of the indented bullet list 5 and `UNINDENTED_ITEM` of the bullet list 6, whose level gives no indent.
 * K: `SINGLE_BULLET` of the one-level bullet list 1, directly above `PLAIN`. L: `HEADING`, numbered
 * by its style's indented list 1 with the style's left indent in character units, above
 * `UNINDENTED_ITEM` of the decimal list 6, whose level gives no indent. M: as H, with a list 5
 * whose level gives a left indent and no hanging indent. N: `HYBRID_BULLET` at the bullet level 1
 * of the hybrid list 1, directly above `PLAIN`.
 */
export type ListDocument = 'A' | 'B' | 'C' | 'D' | 'E' | 'F' | 'G' | 'H' | 'I' | 'J' | 'K' | 'L' | 'M' | 'N';
export type ListKind = 'bullet' | 'numbered';
export interface ListDocumentOptions {
  /** Puts `GAP` directly above `PLAIN`. */
  isolated?: boolean;
  /** Plain paragraphs after `PLAIN`. */
  after?: readonly string[];
  /** A saved package to open in place of the document built. */
  saved?: Uint8Array;
  /** A left indent of its own for `PLAIN`, in twips. */
  indent?: number;
}

/** The `w:numId`s of the numbering each document opens with. */
export const SOURCE_NUM_IDS: Readonly<Record<ListDocument, readonly string[]>> = {
  A: ['1', '2', '3'], B: ['1'], C: [], D: ['1'], E: ['1'], F: ['1', '2'], G: ['2', '7'], H: ['5'], I: ['2', '7'], J: ['5', '6'], K: ['1', '2'], L: ['1', '6'],
  M: ['5'], N: ['1'],
};

export async function preloadListWasm(): Promise<void> {
  const generated = (path: string) => new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated', path)));
  await preloadEditWasm(generated('edit/docx_edit_bg.wasm'));
  await preloadOpcWasm(generated('opc/ooxml_opc_bg.wasm'));
  await preloadParseWasm(generated('parse/docx_parse_bg.wasm'));
}

export function listFont(): Uint8Array {
  return new Uint8Array(readFileSync(resolve(ROOT, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')));
}

const plain = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const item = (text: string, numId: number, ilvl: number) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

function pack(document: string, numbering: string | null, relationships?: Uint8Array, styles = STYLES): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', toBytes(`${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/>${numbering ? `<Override PartName="/word/numbering.xml" ContentType="${OFFICE}.numbering+xml"/>` : ''}</Types>`));
  parts.set('_rels/.rels', toBytes(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`));
  parts.set('word/_rels/document.xml.rels', relationships ?? toBytes(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/>${numbering ? `<Relationship Id="rId2" Type="${R}/numbering" Target="numbering.xml"/>` : ''}</Relationships>`));
  parts.set('word/styles.xml', toBytes(styles));
  if (numbering) parts.set('word/numbering.xml', toBytes(numbering));
  parts.set('word/document.xml', toBytes(document));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

export function listDocument(doc: ListDocument, options: ListDocumentOptions = {}): Uint8Array {
  if (options.saved) return options.saved;
  const target = options.indent === undefined ? plain(PLAIN)
    : `<w:p><w:pPr><w:ind w:left="${options.indent}"/></w:pPr><w:r><w:t>${PLAIN}</w:t></w:r></w:p>`;
  const added = [...(options.isolated ? [GAP] : []).map(plain), target, ...(options.after ?? []).map(plain)].join('');
  if (doc === 'A') {
    const source = readFileSync(resolve(FIXTURE, 'document.xml'), 'utf8');
    const anchor = '    <w:sectPr>';
    expect(source.split(anchor)).toHaveLength(2);
    return pack(
      source.replace(anchor, `${added}\n${anchor}`),
      readFileSync(resolve(FIXTURE, 'numbering.xml'), 'utf8'),
      new Uint8Array(readFileSync(resolve(FIXTURE, '_rels/document.xml.rels')))
    );
  }
  if (doc === 'D') {
    const heading = (text: string) => `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    return pack(
      `${XML}<w:document xmlns:w="${W}"><w:body>${heading(HEADING)}${added}${heading(NEXT_HEADING)}${SECT}</w:body></w:document>`,
      HEADING_NUMBERING, undefined, NUMBERED_HEADING_STYLES
    );
  }
  if (doc === 'L') {
    const heading = `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${HEADING}</w:t></w:r></w:p>`;
    return pack(
      `${XML}<w:document xmlns:w="${W}"><w:body>${heading}${item(UNINDENTED_ITEM, 6, 0)}${added}${SECT}</w:body></w:document>`,
      CHARACTER_HEADING_NUMBERING, undefined, CHARACTER_HEADING_STYLES
    );
  }
  if (doc === 'E' || doc === 'F' || doc === 'G' || doc === 'H' || doc === 'I' || doc === 'J' || doc === 'K' || doc === 'M' || doc === 'N') {
    const opening = {
      E: item(HYBRID_BULLET, 1, 0) + item(HYBRID_NUMBER, 1, 2),
      F: item(SINGLE_BULLET, 1, 0) + item(DEEP_NUMBER, 2, 2),
      G: plain(SELECTION_LEAD) + item(LIST_TWO_ITEM, 2, 0) + item(LIST_SEVEN_ITEM, 7, 1),
      H: item(INDENTED_ITEM, 5, 0),
      I: plain(SELECTION_LEAD) + item(LIST_TWO_ITEM, 2, 0) + item(LIST_SEVEN_ITEM, 7, 2),
      J: `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr><w:ind w:leftChars="200"/></w:pPr><w:r><w:t>${CHARACTER_ITEM}</w:t></w:r></w:p>` +
        item(UNINDENTED_ITEM, 6, 0),
      K: item(SINGLE_BULLET, 1, 0),
      M: item(INDENTED_ITEM, 5, 0),
      N: item(HYBRID_BULLET, 1, 1),
    }[doc];
    const numbering = {
      E: HYBRID_NUMBERING, F: SINGLE_LEVEL_NUMBERING, G: TWO_DECIMAL_NUMBERING, H: INDENTED_NUMBERING, I: DEEP_SEVEN_NUMBERING,
      J: CHARACTER_INDENT_NUMBERING, K: SINGLE_LEVEL_NUMBERING, M: LEFT_ONLY_NUMBERING,
      N: HYBRID_NUMBERING,
    }[doc];
    return pack(`${XML}<w:document xmlns:w="${W}"><w:body>${opening}${added}${SECT}</w:body></w:document>`, numbering);
  }
  const lead = doc === 'B'
    ? `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${EXISTING_BULLET}</w:t></w:r></w:p>`
    : plain('Opening paragraph QA');
  return pack(`${XML}<w:document xmlns:w="${W}"><w:body>${lead}${added}${SECT}</w:body></w:document>`, doc === 'B' ? BULLET_NUMBERING : null);
}

function flatten(primitives: readonly unknown[], out: Record<string, unknown>[] = []): Record<string, unknown>[] {
  for (const primitive of primitives) {
    if (!primitive || typeof primitive !== 'object') continue;
    const record = primitive as Record<string, unknown>;
    out.push(record);
    if (Array.isArray(record.primitives)) flatten(record.primitives, out);
  }
  return out;
}

const blockKey = (primitive: Record<string, unknown>) =>
  primitive.blockKey !== undefined ? `k${String(primitive.blockKey)}`
    : primitive.blockId !== undefined ? `b${String(primitive.blockId)}`
      : primitive.paraId !== undefined ? `p${String(primitive.paraId)}` : undefined;

const unspaced = (value: string) => value.replace(/\s+/g, '');

/** A block of a display list: its text, its list marker, and the page x of the first glyph of each. */
interface ShownBlock {
  text: string;
  marker: string;
  textX?: number;
  markerX?: number;
}

const firstX = (primitive: Record<string, unknown>) => {
  const x = primitive.kind === 'text' ? primitive.x : (primitive.glyphs as ReadonlyArray<{ x?: unknown }> | undefined)?.[0]?.x;
  return typeof x === 'number' && Number.isFinite(x) ? x : undefined;
};

const least = (left: number | undefined, right: number | undefined) =>
  left === undefined ? right : right === undefined ? left : Math.min(left, right);

/** The one block of the display list `frame` reading `text`; null when no single block does. */
function shownBlock(frame: Uint8Array, text: string): ShownBlock | null {
  const displayList = applyFrameDelta(null, decodeFrameDelta(frame)).displayList as unknown as {
    pages: ReadonlyArray<{ primitives?: readonly unknown[] }>;
  };
  for (const page of displayList.pages) {
    const blocks = new Map<string, ShownBlock>();
    for (const primitive of flatten(page.primitives ?? [])) {
      if (primitive.kind !== 'text' && primitive.kind !== 'glyphRun') continue;
      const key = blockKey(primitive);
      if (key === undefined) continue;
      const block = blocks.get(key) ?? { text: '', marker: '' };
      blocks.set(key, block);
      if (primitive.listMarker === true) {
        block.marker += String(primitive.text ?? '');
        block.markerX = least(block.markerX, firstX(primitive));
      } else {
        block.text += String(primitive.text ?? '');
        block.textX = least(block.textX, firstX(primitive));
      }
    }
    const hits = [...blocks.values()].filter(block => unspaced(block.text) === unspaced(text));
    if (hits.length === 1) return hits[0]!;
    if (hits.length > 1) return null;
  }
  return null;
}

/**
 * The list marker the display list `frame` shows before the one block reading `text`, without
 * white space; '' when it shows none, null when no single block reads `text`.
 */
export function shownMarker(frame: Uint8Array, text: string): string | null {
  const block = shownBlock(frame, text);
  return block ? unspaced(block.marker) : null;
}

/** Where a block is laid out: the page x, in hundredths of a pixel, of the first glyph of its text and of its list marker. */
export interface ShownLeft {
  text: number | null;
  marker: number | null;
}

/** Where the display list `frame` lays out the one block reading `text`; nulls for what it does not show. */
export function shownLeft(frame: Uint8Array, text: string): ShownLeft {
  const block = shownBlock(frame, text);
  const hundredths = (x: number | undefined) => (x === undefined ? null : Math.round(x * 100));
  return { text: hundredths(block?.textX), marker: hundredths(block?.markerX) };
}

/** What a saved package says of the list of the paragraph reading `text`. */
export interface SavedList {
  /** The saved paragraphs reading `text`. */
  paragraphs: number;
  numId?: string;
  /** The `w:name` of the definition the paragraph's `w:num` points at. */
  name?: string;
  /** The `w:numFmt` of level 0 of that definition. */
  numFmt?: string;
  /** The `w:lvlText` of level 0 of that definition. */
  lvlText?: string;
}

export function savedList(bytes: Uint8Array, text: string): SavedList {
  const parts = unzipContainer(bytes);
  const decode = (path: string) => (parts[path] ? new TextDecoder().decode(parts[path]) : '');
  const textOf = (xml: string) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(match => match[1]).join('');
  const paragraphs = (decode('word/document.xml').match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? []).filter(paragraph => textOf(paragraph) === text);
  const saved: SavedList = { paragraphs: paragraphs.length };
  const numPr = paragraphs.length === 1 ? paragraphs[0]!.match(/<w:numPr\b[^>]*>([\s\S]*?)<\/w:numPr>/)?.[1] : undefined;
  saved.numId = numPr?.match(/<w:numId\b[^>]*\bw:val="(\d+)"/)?.[1];
  if (saved.numId === undefined) return saved;
  const numbering = decode('word/numbering.xml');
  const num = numbering.match(new RegExp(`<w:num\\b[^>]*\\bw:numId="${saved.numId}"[^>]*>([\\s\\S]*?)</w:num>`))?.[1];
  const abstractNumId = num?.match(/<w:abstractNumId\b[^>]*\bw:val="(\d+)"/)?.[1];
  if (abstractNumId === undefined) return saved;
  const abstract = numbering.match(new RegExp(`<w:abstractNum\\b[^>]*\\bw:abstractNumId="${abstractNumId}"[^>]*>([\\s\\S]*?)</w:abstractNum>`))?.[1] ?? '';
  const level = abstract.match(/<w:lvl\b[^>]*\bw:ilvl="0"[^>]*>([\s\S]*?)<\/w:lvl>/)?.[1] ?? '';
  saved.name = abstract.match(/<w:name\b[^>]*\bw:val="([^"]*)"/)?.[1];
  saved.numFmt = level.match(/<w:numFmt\b[^>]*\bw:val="([^"]*)"/)?.[1];
  saved.lvlText = level.match(/<w:lvlText\b[^>]*\bw:val="([^"]*)"/)?.[1];
  return saved;
}

/** What a saved package says of the list level of the one paragraph reading `text`. */
interface SavedLevel {
  numId?: string;
  ilvl?: string;
  /** The `w:numFmt` of that level of the definition the paragraph's `w:num` points at; undefined when it defines no such level. */
  numFmt?: string;
}

function savedLevel(bytes: Uint8Array, text: string): SavedLevel {
  const parts = unzipContainer(bytes);
  const decode = (path: string) => (parts[path] ? new TextDecoder().decode(parts[path]) : '');
  const textOf = (xml: string) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(match => match[1]).join('');
  const paragraphs = (decode('word/document.xml').match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? []).filter(paragraph => textOf(paragraph) === text);
  expect(paragraphs).toHaveLength(1);
  const numPr = paragraphs[0]!.match(/<w:numPr\b[^>]*>([\s\S]*?)<\/w:numPr>/)?.[1];
  const saved: SavedLevel = {
    numId: numPr?.match(/<w:numId\b[^>]*\bw:val="(\d+)"/)?.[1],
    ilvl: numPr?.match(/<w:ilvl\b[^>]*\bw:val="(\d+)"/)?.[1],
  };
  if (saved.numId === undefined) return saved;
  const numbering = decode('word/numbering.xml');
  const num = numbering.match(new RegExp(`<w:num\\b[^>]*\\bw:numId="${saved.numId}"[^>]*>([\\s\\S]*?)</w:num>`))?.[1];
  const abstractNumId = num?.match(/<w:abstractNumId\b[^>]*\bw:val="(\d+)"/)?.[1];
  if (abstractNumId === undefined) return saved;
  const abstract = numbering.match(new RegExp(`<w:abstractNum\\b[^>]*\\bw:abstractNumId="${abstractNumId}"[^>]*>([\\s\\S]*?)</w:abstractNum>`))?.[1] ?? '';
  const level = abstract.match(new RegExp(`<w:lvl\\b[^>]*\\bw:ilvl="${saved.ilvl ?? '0'}"[^>]*>([\\s\\S]*?)</w:lvl>`))?.[1];
  saved.numFmt = level?.match(/<w:numFmt\b[^>]*\bw:val="([^"]*)"/)?.[1];
  return saved;
}

/** The `w:pPr` of the one saved paragraph reading `text`; '' when it has none. */
function savedParagraphProperties(bytes: Uint8Array, text: string): string {
  const document = new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
  const textOf = (xml: string) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(match => match[1]).join('');
  const paragraphs = (document.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? []).filter(paragraph => textOf(paragraph) === text);
  expect(paragraphs).toHaveLength(1);
  return paragraphs[0]!.match(/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/)?.[0] ?? '';
}

/** The direct `w:ind` of the one saved paragraph reading `text`; undefined when it has none. */
const savedIndent = (bytes: Uint8Array, text: string) => savedParagraphProperties(bytes, text).match(/<w:ind\b[^>]*\/>/)?.[0];

/** The length in bytes of the numbering part of the saved package `bytes`; 0 when it has none. */
function savedNumberingLength(bytes: Uint8Array): number {
  return unzipContainer(bytes)['word/numbering.xml']?.length ?? 0;
}

/** One way the list commands reach a document: the worker's command engine or the toolbar on a session. */
export interface ListCommandDriver {
  /** Bullets or Numbering over the paragraphs from `from` to `to`, a caret in `from` when `to` is left out. */
  toggle(kind: ListKind, from: string, to?: string): void;
  indent(direction: 'increase' | 'decrease', at: string): void;
  undo(): void;
  redo(): void;
  canUndo(): boolean;
  paragraphs(): ReadonlyArray<{ paraId: string; text: string; properties: Record<string, unknown> }>;
  definitions(): ReadonlyArray<{ key: string; kind: string }>;
  marker(text: string): string | null;
  left(text: string): ShownLeft;
  save(): Promise<Uint8Array>;
}

export type OpenListCommandDriver = (doc: ListDocument, options?: ListDocumentOptions) => Promise<ListCommandDriver>;

function propertiesOf(driver: ListCommandDriver, text: string): Record<string, unknown> {
  const found = driver.paragraphs().filter(paragraph => paragraph.text === text);
  expect(found).toHaveLength(1);
  return found[0]!.properties;
}

const keyOf = (driver: ListCommandDriver, text: string) => createdListKey(propertiesOf(driver, text));

async function expectOnCreatedList(driver: ListCommandDriver, doc: ListDocument, text: string, kind: ListKind): Promise<string | undefined> {
  const key = keyOf(driver, text);
  expect(key).toBeDefined();
  expect(driver.definitions().find(definition => definition.key === key)?.kind).toBe(kind);
  expect(driver.marker(text)).toBe(kind === 'bullet' ? '•' : '1.');
  const saved = savedList(await driver.save(), text);
  expect(saved.numId ?? '').toMatch(/^\d+$/);
  expect(SOURCE_NUM_IDS[doc].includes(saved.numId ?? '')).toBe(false);
  expect(saved).toMatchObject({
    paragraphs: 1,
    name: `betteroffice-list:${key}`,
    numFmt: kind === 'bullet' ? 'bullet' : 'decimal',
    ...(kind === 'numbered' ? { lvlText: '%1.' } : {}),
  });
  return key;
}

function expectOffList(driver: ListCommandDriver, text: string): void {
  const properties = propertiesOf(driver, text);
  expect(properties.numPr ?? null).toBeNull();
  expect(properties.listDef ?? null).toBeNull();
}

const KINDS: readonly ListKind[] = ['bullet', 'numbered'];

/** The list command rows, run through `open`'s driver. */
export function listCommandScenarios(path: string, open: OpenListCommandDriver, timeout: number): void {
  for (const doc of ['A', 'B', 'C'] as const) {
    for (const kind of KINDS) {
      test(`${path}: ${kind} on a paragraph with no list beside it in document ${doc} creates a ${kind} list`, async () => {
        const driver = await open(doc, { isolated: true });
        driver.toggle(kind, PLAIN);
        await expectOnCreatedList(driver, doc, PLAIN, kind);
      }, timeout);
    }
  }

  test(`${path}: bullets directly below a source bullet item join its list`, async () => {
    const driver = await open('B');
    driver.toggle('bullet', PLAIN);
    const properties = propertiesOf(driver, PLAIN);
    expect(properties.numPr).toMatchObject({ numId: 1, ilvl: 0 });
    expect(properties.listDef ?? null).toBeNull();
    expect(driver.definitions()).toEqual([]);
    expect(driver.marker(PLAIN)).toBe('•');
    expect(savedList(await driver.save(), PLAIN)).toMatchObject({ paragraphs: 1, numId: '1', numFmt: 'bullet' });
  }, timeout);

  test(`${path}: numbering directly below a source numbered item joins its list`, async () => {
    const driver = await open('A');
    driver.toggle('numbered', PLAIN);
    const properties = propertiesOf(driver, PLAIN);
    expect(properties.numPr).toMatchObject({ numId: 3, ilvl: 0 });
    expect(properties.listDef ?? null).toBeNull();
    expect(driver.definitions()).toEqual([]);
    const above = driver.marker(RESTARTED_ITEM) ?? '';
    expect(above).toMatch(/^\d+\.$/);
    expect(driver.marker(PLAIN)).toBe(`${Number.parseInt(above, 10) + 1}.`);
    expect(savedList(await driver.save(), PLAIN)).toMatchObject({ paragraphs: 1, numId: '3', numFmt: 'decimal' });
  }, timeout);

  test(`${path}: numbering on the paragraph below a created numbered item continues that list`, async () => {
    const driver = await open('C', { after: [FOLLOWER] });
    driver.toggle('numbered', PLAIN);
    const key = keyOf(driver, PLAIN);
    expect(key).toBeDefined();
    driver.toggle('numbered', FOLLOWER);
    expect(keyOf(driver, FOLLOWER)).toBe(key);
    expect(driver.definitions()).toHaveLength(1);
    expect(driver.marker(PLAIN)).toBe('1.');
    expect(driver.marker(FOLLOWER)).toBe('2.');
    const bytes = await driver.save();
    expect(savedList(bytes, FOLLOWER).numId).toBe(savedList(bytes, PLAIN).numId);
  }, timeout);

  for (const kind of KINDS) {
    test(`${path}: ${kind} on a ${kind} item takes the list off, and on again puts it back on that list`, async () => {
      const driver = await open('C');
      driver.toggle(kind, PLAIN);
      const first = keyOf(driver, PLAIN);
      expect(first).toBeDefined();
      driver.toggle(kind, PLAIN);
      expectOffList(driver, PLAIN);
      expect(driver.marker(PLAIN)).toBe('');
      driver.toggle(kind, PLAIN);
      const again = keyOf(driver, PLAIN);
      expect(again).toBe(first);
      expect(driver.definitions()).toHaveLength(1);
    }, timeout);
  }

  test(`${path}: numbering on a created bullet item puts it on a numbered list`, async () => {
    const driver = await open('C');
    driver.toggle('bullet', PLAIN);
    const bullet = keyOf(driver, PLAIN);
    expect(bullet).toBeDefined();
    driver.toggle('numbered', PLAIN);
    expect(await expectOnCreatedList(driver, 'C', PLAIN, 'numbered')).not.toBe(bullet);
  }, timeout);

  test(`${path}: numbering on a source bullet item with no numbered list beside it creates a numbered list`, async () => {
    const driver = await open('B', { isolated: true });
    driver.toggle('numbered', EXISTING_BULLET);
    await expectOnCreatedList(driver, 'B', EXISTING_BULLET, 'numbered');
  }, timeout);

  test(`${path}: bullets on an item of a decimal list 1 join the bullet list below instead of taking the list off`, async () => {
    const driver = await open('A');
    driver.toggle('bullet', DECIMAL_ITEM);
    const properties = propertiesOf(driver, DECIMAL_ITEM);
    expect(properties.numPr).toMatchObject({ numId: 2, ilvl: 0 });
    expect(properties.listDef ?? null).toBeNull();
    expect(driver.marker(DECIMAL_ITEM)).toBe('•');
    expect(savedList(await driver.save(), DECIMAL_ITEM)).toMatchObject({ paragraphs: 1, numId: '2', numFmt: 'bullet' });
  }, timeout);

  test(`${path}: indent and outdent on a created list change only its level, and outdent at the top takes the list off`, async () => {
    const driver = await open('C');
    driver.toggle('numbered', PLAIN);
    const key = keyOf(driver, PLAIN);
    expect(key).toBeDefined();
    driver.indent('increase', PLAIN);
    expect(propertiesOf(driver, PLAIN).numPr).toMatchObject({ ilvl: 1 });
    expect(keyOf(driver, PLAIN)).toBe(key);
    expect(driver.marker(PLAIN)).toBe('a.');
    driver.indent('decrease', PLAIN);
    expect(propertiesOf(driver, PLAIN).numPr).toMatchObject({ ilvl: 0 });
    expect(keyOf(driver, PLAIN)).toBe(key);
    expect(driver.marker(PLAIN)).toBe('1.');
    driver.indent('decrease', PLAIN);
    expectOffList(driver, PLAIN);
  }, timeout);

  test(`${path}: one undo takes back the list and the definition bullets created`, async () => {
    const driver = await open('C');
    driver.toggle('bullet', PLAIN);
    const key = keyOf(driver, PLAIN);
    expect(key).toBeDefined();
    expect(driver.definitions()).toHaveLength(1);
    driver.undo();
    expectOffList(driver, PLAIN);
    expect(driver.definitions()).toEqual([]);
    expect(driver.canUndo()).toBe(false);
    driver.redo();
    expect(keyOf(driver, PLAIN)).toBe(key);
    expect(driver.definitions()).toHaveLength(1);
  }, timeout);

  test(`${path}: bullets on two selected paragraphs put both on one list in one undo step`, async () => {
    const driver = await open('C', { after: [FOLLOWER] });
    driver.toggle('bullet', PLAIN, FOLLOWER);
    const key = keyOf(driver, PLAIN);
    expect(key).toBeDefined();
    expect(keyOf(driver, FOLLOWER)).toBe(key);
    expect(driver.definitions()).toHaveLength(1);
    expect(driver.marker(PLAIN)).toBe('•');
    expect(driver.marker(FOLLOWER)).toBe('•');
    driver.undo();
    expectOffList(driver, PLAIN);
    expectOffList(driver, FOLLOWER);
    expect(driver.definitions()).toEqual([]);
    expect(driver.canUndo()).toBe(false);
  }, timeout);

  test(`${path}: between a created list above and a source list below of that kind a paragraph joins the one above`, async () => {
    const driver = await open('A');
    driver.toggle('bullet', NESTED_DECIMAL);
    const above = keyOf(driver, NESTED_DECIMAL);
    expect(above).toBeDefined();
    driver.toggle('bullet', DECIMAL_ITEM);
    expect(keyOf(driver, DECIMAL_ITEM)).toBe(above);
    expect(driver.definitions()).toHaveLength(1);
  }, timeout);

  test(`${path}: between a source list above and a created list below of that kind a paragraph joins the one above`, async () => {
    const driver = await open('B', { after: [FOLLOWER] });
    driver.toggle('bullet', FOLLOWER);
    expect(keyOf(driver, FOLLOWER)).toBeDefined();
    driver.toggle('bullet', PLAIN);
    const properties = propertiesOf(driver, PLAIN);
    expect(properties.numPr).toMatchObject({ numId: 1, ilvl: 0 });
    expect(properties.listDef ?? null).toBeNull();
    const bytes = await driver.save();
    expect(savedList(bytes, PLAIN)).toMatchObject({ paragraphs: 1, numId: '1', numFmt: 'bullet' });
    expect(savedList(bytes, FOLLOWER).numId).not.toBe('1');
  }, timeout);

  test(`${path}: a paragraph on no list joining a source list takes the level of the item above`, async () => {
    const driver = await open('A');
    driver.toggle('numbered', RESTARTED_ITEM);
    expectOffList(driver, RESTARTED_ITEM);
    driver.toggle('bullet', RESTARTED_ITEM);
    const properties = propertiesOf(driver, RESTARTED_ITEM);
    expect(properties.numPr).toMatchObject({ numId: 2, ilvl: 1 });
    expect(properties.listDef ?? null).toBeNull();
  }, timeout);

  test(`${path}: an item switched onto a source list of the other kind keeps its own level`, async () => {
    const driver = await open('A');
    driver.toggle('bullet', RESTARTED_ITEM);
    const properties = propertiesOf(driver, RESTARTED_ITEM);
    expect(properties.numPr).toMatchObject({ numId: 2, ilvl: 0 });
    expect(properties.listDef ?? null).toBeNull();
  }, timeout);

  test(`${path}: numbering over a bullet list keeps the level of every item`, async () => {
    const items = [PLAIN, FOLLOWER, SECOND_FOLLOWER, THIRD_FOLLOWER];
    const driver = await open('C', { after: items.slice(1) });
    driver.toggle('bullet', PLAIN, THIRD_FOLLOWER);
    const bullet = keyOf(driver, PLAIN);
    for (const text of [FOLLOWER, SECOND_FOLLOWER, THIRD_FOLLOWER, THIRD_FOLLOWER]) driver.indent('increase', text);
    driver.toggle('numbered', PLAIN, THIRD_FOLLOWER);
    expect(items.map(text => (propertiesOf(driver, text).numPr as { ilvl?: number } | undefined)?.ilvl)).toEqual([0, 1, 1, 2]);
    const key = keyOf(driver, PLAIN);
    expect(key).toBeDefined();
    expect(key).not.toBe(bullet);
    expect(items.map(text => keyOf(driver, text))).toEqual(items.map(() => key));
    expect(driver.definitions().find(definition => definition.key === key)?.kind).toBe('numbered');
    expect(items.map(text => driver.marker(text))).toEqual(['1.', 'a.', 'b.', 'i.']);
  }, timeout);

  test(`${path}: numbering on a level 2 bullet item keeps it at level 2`, async () => {
    const driver = await open('C');
    driver.toggle('bullet', PLAIN);
    driver.indent('increase', PLAIN);
    driver.indent('increase', PLAIN);
    driver.toggle('numbered', PLAIN);
    expect(propertiesOf(driver, PLAIN).numPr).toMatchObject({ ilvl: 2 });
    expect(keyOf(driver, PLAIN)).toBeDefined();
    expect(driver.marker(PLAIN)).toBe('i.');
  }, timeout);

  test(`${path}: numbering below a heading numbered by its style creates a list instead of joining the heading's`, async () => {
    const driver = await open('D');
    const heading = propertiesOf(driver, HEADING).numPr;
    expect(heading).toMatchObject({ numId: 1 });
    driver.toggle('numbered', PLAIN);
    await expectOnCreatedList(driver, 'D', PLAIN, 'numbered');
    expect(driver.definitions()).toHaveLength(1);
    expect(propertiesOf(driver, HEADING).numPr).toEqual(heading);
    expect(propertiesOf(driver, NEXT_HEADING).numPr).toEqual(heading);
  }, timeout);

  test(`${path}: bullets on a heading numbered by its style save a direct list of its own and leave the next heading on the style's list`, async () => {
    const driver = await open('D');
    driver.toggle('bullet', HEADING);
    const key = keyOf(driver, HEADING);
    expect(key).toBeDefined();
    expect(driver.marker(HEADING)).toBe('•');
    const bytes = await driver.save();
    const heading = savedList(bytes, HEADING);
    expect(heading.numId ?? '').toMatch(/^\d+$/);
    expect(SOURCE_NUM_IDS.D.includes(heading.numId ?? '')).toBe(false);
    expect(heading).toMatchObject({ paragraphs: 1, name: `betteroffice-list:${key}`, numFmt: 'bullet' });
    expect(savedList(bytes, NEXT_HEADING)).toEqual({ paragraphs: 1, numId: undefined });
  }, timeout);

  test(`${path}: a paragraph joining a created list takes the level of the item above`, async () => {
    const driver = await open('C', { after: [FOLLOWER] });
    driver.toggle('numbered', PLAIN);
    driver.indent('increase', PLAIN);
    driver.toggle('numbered', FOLLOWER);
    expect(propertiesOf(driver, FOLLOWER).numPr).toMatchObject({ ilvl: 1 });
    expect(keyOf(driver, FOLLOWER)).toBe(keyOf(driver, PLAIN));
    expect(driver.marker(PLAIN)).toBe('a.');
    expect(driver.marker(FOLLOWER)).toBe('b.');
  }, timeout);

  test(`${path}: a click on a selection whose first paragraph is on a list of that kind takes every paragraph off`, async () => {
    const driver = await open('C', { after: [FOLLOWER] });
    driver.toggle('bullet', PLAIN);
    expect(keyOf(driver, PLAIN)).toBeDefined();
    driver.toggle('bullet', PLAIN, FOLLOWER);
    expectOffList(driver, PLAIN);
    expectOffList(driver, FOLLOWER);
  }, timeout);

  test(`${path}: a click on a selection whose first paragraph is on no list puts every paragraph on one`, async () => {
    const driver = await open('C', { after: [FOLLOWER] });
    driver.toggle('bullet', FOLLOWER);
    const key = keyOf(driver, FOLLOWER);
    expect(key).toBeDefined();
    driver.toggle('bullet', PLAIN, FOLLOWER);
    expect(keyOf(driver, PLAIN)).toBe(key);
    expect(keyOf(driver, FOLLOWER)).toBe(key);
    expect(driver.definitions()).toHaveLength(1);
  }, timeout);

  test(`${path}: bullets on a numbered level of a hybrid list below its bullet item show a bullet at that level`, async () => {
    const driver = await open('E');
    const before = { properties: propertiesOf(driver, HYBRID_NUMBER), marker: driver.marker(HYBRID_NUMBER) };
    driver.toggle('bullet', HYBRID_NUMBER);
    expect(propertiesOf(driver, HYBRID_NUMBER)).toMatchObject({ listIsBullet: true, numPr: { ilvl: 2 } });
    expect(driver.marker(HYBRID_NUMBER) ?? '').not.toMatch(/\d/);
    expect(driver.marker(HYBRID_NUMBER) ?? '').not.toBe('');
    expect(savedLevel(await driver.save(), HYBRID_NUMBER)).toMatchObject({ ilvl: '2', numFmt: 'bullet' });
    expect(before.properties).toMatchObject({ listNumFmt: 'decimal', numPr: { numId: 1, ilvl: 2 } });
    expect(before.marker).toMatch(/^\d+\.$/);
    expect(propertiesOf(driver, HYBRID_BULLET)).toMatchObject({ listNumFmt: 'bullet', numPr: { numId: 1, ilvl: 0 } });
  }, timeout);

  test(`${path}: bullets on a level 2 numbered item below a one-level bullet list put it on a bullet list that has level 2`, async () => {
    const driver = await open('F');
    const before = propertiesOf(driver, DEEP_NUMBER);
    driver.toggle('bullet', DEEP_NUMBER);
    const key = keyOf(driver, DEEP_NUMBER);
    expect(key).toBeDefined();
    expect(driver.definitions().find(definition => definition.key === key)?.kind).toBe('bullet');
    expect(propertiesOf(driver, DEEP_NUMBER)).toMatchObject({ listIsBullet: true, numPr: { ilvl: 2 } });
    const bytes = await driver.save();
    const saved = savedLevel(bytes, DEEP_NUMBER);
    expect(saved).toMatchObject({ ilvl: '2', numFmt: 'bullet' });
    expect(SOURCE_NUM_IDS.F.includes(saved.numId ?? '')).toBe(false);
    expect(driver.marker(DEEP_NUMBER) ?? '').not.toBe('');
    const reopened = await open('F', { saved: bytes });
    expect(reopened.marker(DEEP_NUMBER) ?? '').not.toBe('');
    expect(reopened.marker(DEEP_NUMBER) ?? '').not.toMatch(/\d/);
    expect(before).toMatchObject({ listNumFmt: 'decimal', numPr: { numId: 2, ilvl: 2 } });
    expect(propertiesOf(driver, SINGLE_BULLET)).toMatchObject({ listNumFmt: 'bullet', numPr: { numId: 1, ilvl: 0 } });
  }, timeout);

  test(`${path}: numbering over a plain paragraph and the items of two numbered lists puts all three on the list of the first`, async () => {
    const driver = await open('G');
    const items = [SELECTION_LEAD, LIST_TWO_ITEM, LIST_SEVEN_ITEM];
    const numIdOf = (text: string) => (propertiesOf(driver, text).numPr as { numId?: number } | null | undefined)?.numId;
    const before = items.map(numIdOf);
    driver.toggle('numbered', SELECTION_LEAD, LIST_SEVEN_ITEM);
    expect(items.map(numIdOf)).toEqual([2, 2, 2]);
    expect(propertiesOf(driver, LIST_SEVEN_ITEM)).toMatchObject({ listNumFmt: 'lowerLetter', numPr: { numId: 2, ilvl: 1 } });
    const bytes = await driver.save();
    expect(items.map(text => savedList(bytes, text).numId)).toEqual(['2', '2', '2']);
    expect(savedLevel(bytes, LIST_SEVEN_ITEM)).toMatchObject({ numId: '2', ilvl: '1', numFmt: 'lowerLetter' });
    expect(items.map(text => keyOf(driver, text))).toEqual([undefined, undefined, undefined]);
    expect(before).toEqual([undefined, 2, 7]);
  }, timeout);

  test(`${path}: numbering off on a heading numbered by its style saves numId 0, and numbering on again puts it on a list`, async () => {
    const driver = await open('D');
    const before = { numPr: propertiesOf(driver, HEADING).numPr, marker: driver.marker(HEADING) };
    driver.toggle('numbered', HEADING);
    const saved = await driver.save();
    expect(savedList(saved, HEADING)).toMatchObject({ paragraphs: 1, numId: '0' });
    expect(driver.marker(HEADING) ?? '').toBe('');
    const reopened = await open('D', { saved });
    expect(reopened.marker(HEADING) ?? '').toBe('');
    reopened.toggle('numbered', HEADING);
    expect(reopened.marker(HEADING) ?? '').toMatch(/^\d+\.$/);
    driver.toggle('numbered', HEADING);
    expect(driver.marker(HEADING) ?? '').toMatch(/^\d+\.$/);
    expect(before.numPr).toMatchObject({ numId: 1 });
    expect(before.marker).toMatch(/^\d+\.$/);
  }, timeout);

  test(`${path}: bullets off on a directly numbered item save no numbering at all`, async () => {
    const driver = await open('B');
    const before = propertiesOf(driver, EXISTING_BULLET).numPr;
    driver.toggle('bullet', EXISTING_BULLET);
    expect(savedList(await driver.save(), EXISTING_BULLET)).toEqual({ paragraphs: 1, numId: undefined });
    expectOffList(driver, EXISTING_BULLET);
    expect(before).toMatchObject({ numId: 1, ilvl: 0 });
  }, timeout);

  test(`${path}: bullets on and off on a lone paragraph 200 times keep one definition and save the same numbering`, async () => {
    const driver = await open('C', { isolated: true });
    driver.toggle('bullet', PLAIN);
    const key = keyOf(driver, PLAIN);
    const first = savedNumberingLength(await driver.save());
    for (let click = 2; click <= 199; click += 1) driver.toggle('bullet', PLAIN);
    expect(driver.definitions()).toHaveLength(1);
    expect(keyOf(driver, PLAIN)).toBe(key);
    expect(savedNumberingLength(await driver.save())).toBe(first);
    driver.toggle('bullet', PLAIN);
    expectOffList(driver, PLAIN);
    expect(driver.definitions()).toHaveLength(1);
    expect(key).toBeDefined();
    expect(first).toBeGreaterThan(0);
  }, timeout);

  test(`${path}: bullets below an item of a list with its own indent lay the paragraph out where that list puts its items, as after reopening`, async () => {
    const driver = await open('H');
    const before = driver.left(PLAIN);
    driver.toggle('bullet', PLAIN);
    const item = driver.left(INDENTED_ITEM);
    const live = driver.left(PLAIN);
    expect(live).toEqual(item);
    const saved = await driver.save();
    const reopened = await open('H', { saved });
    expect(reopened.left(PLAIN)).toEqual(live);
    expect(reopened.left(INDENTED_ITEM)).toEqual(item);
    expect([PLAIN, INDENTED_ITEM].map(text => savedIndent(saved, text))).toEqual([undefined, undefined]);
    const resaved = await reopened.save();
    expect([PLAIN, INDENTED_ITEM].map(text => savedIndent(resaved, text))).toEqual([undefined, undefined]);
    expect(propertiesOf(driver, PLAIN).numPr).toMatchObject({ numId: 5, ilvl: 0 });
    expect(item.marker).not.toBeNull();
    expect(before).toMatchObject({ marker: null });
    expect(item.text ?? 0).toBeGreaterThan((before.text ?? Infinity) + 100);
  }, timeout);

  test(`${path}: an item taken off its list and put back on it is laid out and saved as it was loaded`, async () => {
    const driver = await open('H');
    const before = driver.left(INDENTED_ITEM);
    const original = savedParagraphProperties(await driver.save(), INDENTED_ITEM);
    driver.toggle('bullet', PLAIN);
    driver.toggle('bullet', INDENTED_ITEM);
    expectOffList(driver, INDENTED_ITEM);
    driver.toggle('bullet', INDENTED_ITEM);
    expect(propertiesOf(driver, INDENTED_ITEM).numPr).toMatchObject({ numId: 5, ilvl: 0 });
    expect(driver.left(INDENTED_ITEM)).toEqual(before);
    expect(savedParagraphProperties(await driver.save(), INDENTED_ITEM)).toBe(original);
    expect(before.marker).not.toBeNull();
  }, timeout);

  test(`${path}: bullets on and off again below an item of a list with its own indent lay the paragraph out where it was, as after reopening`, async () => {
    const driver = await open('H');
    const before = driver.left(PLAIN);
    const original = savedParagraphProperties(await driver.save(), PLAIN);
    driver.toggle('bullet', PLAIN);
    driver.toggle('bullet', PLAIN);
    expect(driver.left(PLAIN)).toEqual(before);
    const saved = await driver.save();
    expect(savedParagraphProperties(saved, PLAIN)).toBe(original);
    const reopened = await open('H', { saved });
    expect(reopened.left(PLAIN)).toEqual(before);
    expectOffList(driver, PLAIN);
    expect(before.text).not.toBeNull();
  }, timeout);

  test(`${path}: a paragraph with a left indent of its own keeps it on a list it joins and after the list is taken off`, async () => {
    const driver = await open('H', { indent: 2880 });
    const before = driver.left(PLAIN);
    const original = savedParagraphProperties(await driver.save(), PLAIN);
    driver.toggle('bullet', PLAIN);
    const item = driver.left(INDENTED_ITEM);
    const live = driver.left(PLAIN);
    const saved = await driver.save();
    const reopened = await open('H', { saved });
    expect(reopened.left(PLAIN)).toEqual(live);
    expect(live.text ?? 0).toBeGreaterThan((item.text ?? Infinity) + 100);
    expect(savedIndent(saved, PLAIN) ?? '').toContain('w:left="2880"');
    expect(propertiesOf(driver, PLAIN).numPr).toMatchObject({ numId: 5, ilvl: 0 });
    driver.toggle('bullet', PLAIN);
    expect(driver.left(PLAIN)).toEqual(before);
    const off = await driver.save();
    expect(savedIndent(off, PLAIN)).toBe(original.match(/<w:ind\b[^>]*\/>/)?.[0]);
    expect(savedParagraphProperties(off, PLAIN)).not.toContain('<w:numPr');
    expect((await open('H', { saved: off })).left(PLAIN)).toEqual(before);
    expect(original).toContain('w:left="2880"');
    expect(before).toMatchObject({ marker: null });
    expect(before.text ?? 0).toBeGreaterThan((item.text ?? Infinity) + 100);
  }, timeout);

  test(`${path}: numbering over a selection puts a paragraph whose level the list of the first lacks at the level of the first`, async () => {
    const driver = await open('I');
    const before = propertiesOf(driver, LIST_SEVEN_ITEM).numPr;
    driver.toggle('numbered', SELECTION_LEAD, LIST_SEVEN_ITEM);
    expect(propertiesOf(driver, LIST_SEVEN_ITEM).numPr).toMatchObject({ numId: 2, ilvl: 0 });
    expect(propertiesOf(driver, SELECTION_LEAD).numPr).toMatchObject({ numId: 2, ilvl: 0 });
    expect(savedLevel(await driver.save(), LIST_SEVEN_ITEM)).toMatchObject({ numId: '2', ilvl: '0', numFmt: 'decimal' });
    expect(before).toMatchObject({ numId: 7, ilvl: 2 });
  }, timeout);

  test(`${path}: outdent at the top level on a heading numbered by its style saves numId 0 and stays off after reopening`, async () => {
    const driver = await open('D');
    const before = { numPr: propertiesOf(driver, HEADING).numPr, marker: driver.marker(HEADING) };
    driver.indent('decrease', HEADING);
    const saved = await driver.save();
    expect(savedList(saved, HEADING)).toMatchObject({ paragraphs: 1, numId: '0' });
    expect(driver.marker(HEADING) ?? '').toBe('');
    const reopened = await open('D', { saved });
    expect(reopened.marker(HEADING) ?? '').toBe('');
    expect(before.numPr).toMatchObject({ numId: 1 });
    expect(before.marker).toMatch(/^\d+\.$/);
  }, timeout);

  test(`${path}: a list item whose own left indent is in character units is laid out as an item of a list that gives no indent`, async () => {
    const driver = await open('J');
    const character = driver.left(CHARACTER_ITEM);
    expect(character).toEqual(driver.left(UNINDENTED_ITEM));
    expect(character.text).not.toBeNull();
    expect(character.marker).not.toBeNull();
    expect(propertiesOf(driver, CHARACTER_ITEM).numPr).toMatchObject({ numId: 5, ilvl: 0 });
    expect(propertiesOf(driver, UNINDENTED_ITEM).numPr).toMatchObject({ numId: 6, ilvl: 0 });
  }, timeout);

  test(`${path}: bullets below a bullet item indented onto a numbered level of its list stay bullets after a save and a reopen`, async () => {
    const driver = await open('N');
    driver.indent('increase', HYBRID_BULLET);
    driver.toggle('bullet', PLAIN);
    const live = driver.marker(PLAIN) ?? '';
    const bytes = await driver.save();
    const reopened = await open('N', { saved: bytes });
    expect(savedLevel(bytes, PLAIN)).toMatchObject({ numFmt: 'bullet' });
    expect(reopened.marker(PLAIN) ?? '').toBe(live);
    expect(live).not.toMatch(/\d/);
    expect(live).not.toBe('');
    expect(propertiesOf(driver, HYBRID_BULLET).numPr).toMatchObject({ numId: 1, ilvl: 2 });
  }, timeout);

  test(`${path}: bullets below an item of a list whose level gives a left indent and no hanging indent lay the paragraph out as after reopening`, async () => {
    const driver = await open('M');
    driver.toggle('bullet', PLAIN);
    const live = driver.left(PLAIN);
    const item = driver.left(INDENTED_ITEM);
    const reopened = await open('M', { saved: await driver.save() });
    expect(live).toEqual(reopened.left(PLAIN));
    expect(reopened.left(INDENTED_ITEM)).toEqual(item);
    expect(propertiesOf(driver, PLAIN).numPr).toMatchObject({ numId: 5, ilvl: 0 });
    expect(live.marker).not.toBeNull();
  }, timeout);

  test(`${path}: bullets below a bullet item indented to a level its list does not define start a list instead of copying that level`, async () => {
    const driver = await open('K');
    driver.indent('increase', SINGLE_BULLET);
    driver.toggle('bullet', PLAIN);
    expect(keyOf(driver, PLAIN)).toBeDefined();
    const bytes = await driver.save();
    expect(savedLevel(bytes, PLAIN)).toMatchObject({ ilvl: '0', numFmt: 'bullet' });
    const reopened = await open('K', { saved: bytes });
    expect(reopened.marker(PLAIN) ?? '').not.toBe('');
    expect(propertiesOf(driver, SINGLE_BULLET).numPr).toMatchObject({ numId: 1, ilvl: 1 });
  }, timeout);

  test(`${path}: a heading whose style gives its left indent in character units is laid out as an item of a list that gives no indent`, async () => {
    const driver = await open('L');
    const heading = driver.left(HEADING);
    expect(heading).toEqual(driver.left(UNINDENTED_ITEM));
    expect(heading.text).not.toBeNull();
    expect(heading.marker).not.toBeNull();
    expect(propertiesOf(driver, HEADING).numPr).toMatchObject({ numId: 1 });
    expect(propertiesOf(driver, UNINDENTED_ITEM).numPr).toMatchObject({ numId: 6, ilvl: 0 });
  }, timeout);

  test(`${path}: numbering taken off a style-numbered heading a second time, after a save and a reopen, saves numId 0 again`, async () => {
    const first = await open('D');
    first.toggle('numbered', HEADING);
    const reopened = await open('D', { saved: await first.save() });
    reopened.toggle('numbered', HEADING);
    const on = reopened.marker(HEADING) ?? '';
    reopened.toggle('numbered', HEADING);
    const saved = await reopened.save();
    expect(savedList(saved, HEADING)).toMatchObject({ paragraphs: 1, numId: '0' });
    expect((await open('D', { saved })).marker(HEADING) ?? '').toBe('');
    expect(on).toMatch(/^\d+\.$/);
  }, timeout);

  for (const kind of KINDS) {
    for (const level of [0, 2]) {
      test(`${path}: a paragraph at level ${level} of a created ${kind} list is laid out the same after reopening`, async () => {
        const driver = await open('C', { isolated: true });
        driver.toggle(kind, PLAIN);
        for (let step = 0; step < level; step += 1) driver.indent('increase', PLAIN);
        const live = driver.left(PLAIN);
        const reopened = await open('C', { saved: await driver.save() });
        expect(reopened.left(PLAIN)).toEqual(live);
        expect(propertiesOf(driver, PLAIN).numPr).toMatchObject({ ilvl: level });
        expect(live.text).not.toBeNull();
        expect(live.marker).not.toBeNull();
      }, timeout);
    }
  }
}
