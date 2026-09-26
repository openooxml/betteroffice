import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import {
  compareDocx,
  type DocxCompareDiagnosticCode,
  type DocxCompareOptions,
  type DocxCompareResult,
} from '../core';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from '../yrs/index';
import { yrsToDocument } from '../yrs/yrsToDocument';
import { repackDocx } from './rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from './rezip/parts';
import { unzipContainer } from './wasm';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument';
const NS = `xmlns:w="${W}" xmlns:r="${REL}" xmlns:w14="${W14}"`;
const AUTHOR = 'Reviewer <R&D> "QA"';
const OPTIONS: DocxCompareOptions = { author: AUTHOR, date: '2024-05-06T07:08:09+02:00' };
const UTC = '2024-05-06T05:08:09Z';
const THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;

const run = (text: string, props = '') =>
  `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (...runs: string[]) => `<w:p>${runs.join('')}</w:p>`;
const text = (value: string) => p(run(value));
const table = (cell: string, revision = '') =>
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${revision}</w:tcPr>${text(cell)}</w:tc></w:tr></w:tbl>`;

interface Fixture {
  body: string;
  header?: string;
  theme?: string;
  /** Document defaults and styles ahead of the fixture's own. */
  styles?: string;
}

function fixture({ body, header = text('Header text'), theme = THEME, styles = '' }: Fixture): Uint8Array {
  const parts: PartsMap = new Map();
  const set = (name: string, content: string | Uint8Array) => parts.set(name, toBytes(content));
  set(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="${OFFICE}.wordprocessingml.header+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="${OFFICE}.theme+xml"/></Types>`
  );
  set(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  set(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rIdHeader" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdTheme" Type="${REL}/theme" Target="theme/theme1.xml"/><Relationship Id="rIdImage" Type="${REL}/image" Target="media/image1.png"/></Relationships>`
  );
  set(
    'word/styles.xml',
    `<w:styles ${NS}>${styles}<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="160"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:style></w:styles>`
  );
  set(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`
  );
  set('word/header1.xml', `<w:hdr ${NS}>${header}</w:hdr>`);
  set('word/theme/theme1.xml', theme);
  set('word/media/image1.png', new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 7]));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function parts(bytes: Uint8Array): Record<string, Uint8Array> {
  return unzipContainer(bytes) as Record<string, Uint8Array>;
}

const decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes);

interface XmlNode {
  ns: string | null;
  local: string;
  attributes: Map<string, string>;
  children: XmlNode[];
  text: string;
  start: number;
  end: number;
}

/** A small namespace-aware XML reader, independent of the engine's parser, with source spans. */
function parseXml(xml: string): XmlNode {
  const entity = (value: string) =>
    value
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, code: string) => String.fromCodePoint(Number(code)))
      .replace(/&#x([0-9a-f]+);/gi, (_, code: string) => String.fromCodePoint(parseInt(code, 16)))
      .replace(/&amp;/g, '&');
  const root: XmlNode = {
    ns: null,
    local: '#root',
    attributes: new Map(),
    children: [],
    text: '',
    start: 0,
    end: xml.length,
  };
  const stack: Array<{ node: XmlNode; scope: Map<string, string> }> = [
    { node: root, scope: new Map([['xml', 'http://www.w3.org/XML/1998/namespace']]) },
  ];
  const pattern =
    /<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<\/([^>\s]+)\s*>|<([^\s/>]+)((?:\s+[^\s=]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  for (const match of xml.matchAll(pattern)) {
    const top = stack[stack.length - 1]!;
    const end = match.index + match[0].length;
    if (match[5] !== undefined) {
      top.node.text += entity(match[5]);
      continue;
    }
    if (match[1] !== undefined) {
      top.node.end = end;
      stack.pop();
      continue;
    }
    if (match[2] === undefined) continue;
    const raw = [...(match[3] ?? '').matchAll(/([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)].map(
      (attribute) => [attribute[1]!, entity(attribute[2] ?? attribute[3] ?? '')] as const
    );
    const scope = new Map(top.scope);
    for (const [name, value] of raw) {
      if (name === 'xmlns') scope.set('', value);
      else if (name.startsWith('xmlns:')) scope.set(name.slice(6), value);
    }
    const qualify = (name: string, attribute: boolean) => {
      const [prefix, local] = name.includes(':') ? name.split(':') : ['', name];
      const ns = attribute && prefix === '' ? null : (scope.get(prefix!) ?? null);
      return { ns, local: local! };
    };
    const { ns, local } = qualify(match[2], false);
    const node: XmlNode = {
      ns,
      local,
      attributes: new Map(),
      children: [],
      text: '',
      start: match.index,
      end,
    };
    for (const [name, value] of raw) {
      if (name === 'xmlns' || name.startsWith('xmlns:')) continue;
      const attribute = qualify(name, true);
      node.attributes.set(
        attribute.ns ? `{${attribute.ns}}${attribute.local}` : attribute.local,
        value
      );
    }
    top.node.children.push(node);
    if (match[4] !== '/') stack.push({ node, scope });
  }
  return root.children[0]!;
}

const isW = (node: XmlNode, local: string) => node.ns === W && node.local === local;

function descendants(node: XmlNode, local: string): XmlNode[] {
  const found: XmlNode[] = [];
  const visit = (current: XmlNode) => {
    for (const child of current.children) {
      if (isW(child, local)) found.push(child);
      visit(child);
    }
  };
  visit(node);
  return found;
}

function at(root: XmlNode, path: number[]): XmlNode {
  return path.reduce((node, index) => node.children[index]!, root);
}

interface Unit {
  text: string;
  props: string[];
}

/** A paragraph's text units in one view, each with its direct run properties by attribute. */
function units(paragraph: XmlNode, view: 'plain' | 'original' | 'accepted'): Unit[] {
  const out: Unit[] = [];
  const readRun = (runNode: XmlNode) => {
    const rPr = runNode.children.find((child) => isW(child, 'rPr'));
    const props = (rPr?.children ?? []).flatMap((child) => [
      child.local,
      ...[...child.attributes].map(([name, value]) => `${child.local}@${name}=${value}`),
    ]);
    for (const child of runNode.children) {
      if (isW(child, 't') || isW(child, 'delText')) {
        for (let index = 0; index < child.text.length; index += 1) {
          out.push({ text: child.text[index]!, props });
        }
      } else if (isW(child, 'tab')) out.push({ text: '\t', props });
      else if (isW(child, 'softHyphen')) out.push({ text: '­', props });
      else if (isW(child, 'noBreakHyphen')) out.push({ text: '‑', props });
      else if (isW(child, 'br')) out.push({ text: '￼', props });
    }
  };
  for (const child of paragraph.children) {
    if (isW(child, 'r')) readRun(child);
    else if (isW(child, 'ins') && view === 'accepted') child.children.forEach(readRun);
    else if (isW(child, 'del') && view === 'original') child.children.forEach(readRun);
    else if ((isW(child, 'ins') || isW(child, 'del')) && view === 'plain') {
      throw new Error('a source paragraph holds a revision');
    }
  }
  return out;
}

function redline(result: DocxCompareResult): Extract<DocxCompareResult, { ok: true }> {
  if (!result.ok) throw new Error(JSON.stringify(result.diagnostics));
  return result;
}

let nextClientId = 88100;

async function open(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: nextClientId++ });
  session.openDocx(bytes, true);
  return session;
}

function texts(session: YrsSession, view: 'accepted' | 'original' = 'accepted'): string[] {
  const read = session.readParagraphs({ view });
  if (!read.ok) throw new Error(read.failure.message);
  return read.paragraphs.map((paragraph) => paragraph.text);
}

function segments(session: YrsSession): string[] {
  return session.storySegments('body').flatMap((segment) =>
    segment.kind === 'text' ? [`${segment.text}|${JSON.stringify(segment.attributes)}`] : []
  );
}

/** Resolves every tracked change in BetterOffice, saves, and reopens the result. */
async function resolveAll(docx: Uint8Array, accept: boolean): Promise<YrsSession> {
  const session = await open(docx);
  try {
    for (const revision of session.listRevisions()) {
      if (accept) session.acceptChange({ revisionId: revision.revisionId });
      else session.rejectChange({ revisionId: revision.revisionId });
    }
    expect(session.listRevisions()).toEqual([]);
    const saved = await repackDocx(yrsToDocument(session, session.materializeDocx()!));
    return await open(new Uint8Array(saved));
  } finally {
    session.destroy();
  }
}

/**
 * Compares through the public API and checks the redline independently: native revisions with
 * attribution and unique ids covering exactly the changed text, direct run properties kept on
 * both sides, every other part and byte preserved, and accept-all and reject-all in BetterOffice
 * reproducing the revised and original text and formatting after a save and reopen.
 */
async function expectRedline(
  original: Uint8Array,
  revised: Uint8Array,
  options: DocxCompareOptions = OPTIONS
): Promise<Extract<DocxCompareResult, { ok: true }>> {
  const result = redline(await compareDocx(original, revised, options));
  expect(result.changes.length).toBeGreaterThan(0);
  const [before, after, target] = [original, result.docx, revised].map(parts);
  expect(Object.keys(after!).sort()).toEqual(Object.keys(before!).sort());
  for (const name of Object.keys(before!)) {
    if (name !== 'word/document.xml') expect(after![name]).toEqual(before![name]);
  }
  const [sourceXml, savedXml, revisedXml] = [before, after, target].map((files) =>
    decode(files!['word/document.xml']!)
  );
  const [source, saved, wanted] = [sourceXml!, savedXml!, revisedXml!].map(parseXml);
  const changed = new Map<string, typeof result.changes>();
  for (const change of result.changes) {
    const key = JSON.stringify(change.original.path);
    changed.set(key, [...(changed.get(key) ?? []), change]);
  }
  const ids = new Set<string>();
  let sourceCursor = 0;
  let savedCursor = 0;
  for (const [key, changes] of [...changed].sort(
    ([left], [right]) => at(source!, JSON.parse(left)).start - at(source!, JSON.parse(right)).start
  )) {
    const path = JSON.parse(key) as number[];
    const [from, to] = [at(source!, path), at(saved!, path)];
    expect(savedXml!.slice(savedCursor, to.start)).toBe(sourceXml!.slice(sourceCursor, from.start));
    sourceCursor = from.end;
    savedCursor = to.end;
    for (const revision of [...descendants(to, 'ins'), ...descendants(to, 'del')]) {
      expect(revision.attributes.get(`{${W}}author`)).toBe(AUTHOR);
      expect(revision.attributes.get(`{${W}}date`)).toBe(UTC);
      const id = revision.attributes.get(`{${W}}id`)!;
      expect(id).toMatch(/^[1-9]\d*$/);
      expect(ids.has(id)).toBe(false);
      ids.add(id);
    }
    const covered = (local: 'ins' | 'del') =>
      descendants(to, local).map((wrapper) =>
        units(wrapper, local === 'ins' ? 'accepted' : 'original')
          .map((unit) => unit.text)
          .join('')
      );
    expect(covered('ins')).toEqual(changes.map((change) => change.revised.text).filter(Boolean));
    expect(covered('del')).toEqual(changes.map((change) => change.original.text).filter(Boolean));
    const revisedPath = changes[0]!.revised.path;
    for (const [view, expectedUnits] of [
      ['original', units(from, 'plain')],
      ['accepted', units(at(wanted!, revisedPath), 'plain')],
    ] as const) {
      const actual = units(to, view);
      expect(actual.map((unit) => unit.text).join('')).toBe(
        expectedUnits.map((unit) => unit.text).join('')
      );
      expectedUnits.forEach((unit, index) => {
        for (const prop of unit.props) expect(actual[index]!.props).toContain(prop);
      });
    }
  }
  expect(savedXml!.slice(savedCursor)).toBe(sourceXml!.slice(sourceCursor));

  const [expectedRevised, expectedOriginal, reopened] = await Promise.all([
    open(revised),
    open(original),
    open(result.docx),
  ]);
  const [accepted, rejected] = [
    await resolveAll(result.docx, true),
    await resolveAll(result.docx, false),
  ];
  try {
    expect(texts(reopened)).toEqual(texts(expectedRevised));
    expect(texts(reopened, 'original')).toEqual(texts(expectedOriginal));
    for (const revision of reopened.listRevisions()) {
      expect([revision.author, revision.date]).toEqual([AUTHOR, UTC]);
    }
    expect(texts(accepted)).toEqual(texts(expectedRevised));
    expect(texts(rejected)).toEqual(texts(expectedOriginal));
    expect(segments(accepted)).toEqual(segments(expectedRevised));
    expect(segments(rejected)).toEqual(segments(expectedOriginal));
  } finally {
    for (const session of [expectedRevised, expectedOriginal, reopened, accepted, rejected]) {
      session.destroy();
    }
  }
  return result;
}

async function expectRefusal(
  original: Uint8Array,
  revised: Uint8Array,
  code: DocxCompareDiagnosticCode,
  options: DocxCompareOptions = OPTIONS
): Promise<void> {
  for (const unsupported of ['fail', 'report'] as const) {
    const result = await compareDocx(original, revised, { ...options, unsupported });
    expect(result.ok).toBe(false);
    expect('docx' in result).toBe(false);
    expect(result.diagnostics.map((diagnostic) => diagnostic.code)).toContain(code);
    expect(result.diagnostics.every((diagnostic) => diagnostic.severity !== undefined)).toBe(true);
  }
}

describe('compareDocx', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  const original = fixture({
    body: [
      p('<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>', run('Service agreement')),
      text('The supplier delivers the goods within ten days.'),
      text('This paragraph stays exactly as it was.'),
      p(run('Payment is due '), run('soon', '<w:i/><w:iCs/>'), run(' on receipt.')),
      text('Remove the obsolete word here.'),
      text('Append'),
    ].join(''),
  });
  const revised = fixture({
    body: [
      p('<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>', run('Service agreement (draft)')),
      text('The supplier delivers the goods within five days.'),
      text('This paragraph stays exactly as it was.'),
      p(run('Payment is due '), run('at twelve', '<w:b/><w:bCs/>'), run(' on receipt.')),
      text('Remove the word here.'),
      text('Append more'),
    ].join(''),
  });

  it('records insertions, deletions and replacements as attributed native revisions', async () => {
    const result = await expectRedline(original, revised);
    expect(result.diagnostics).toEqual([]);
    expect(
      result.changes.map((change) => [change.id, change.kind, change.original.text, change.revised.text])
    ).toEqual([
      ['change-0', 'insertion', '', ' (draft)'],
      ['change-1', 'replacement', 'ten', 'five'],
      ['change-2', 'replacement', 'soon', 'at twelve'],
      ['change-3', 'deletion', 'obsolete ', ''],
      ['change-4', 'insertion', '', ' more'],
    ]);
    expect(result.changes[1]!.original).toMatchObject({
      part: 'word/document.xml',
      path: [0, 1],
      start: 39,
      end: 42,
    });
  });

  it('keeps each side formatting and a cleared paragraph mark', async () => {
    await expectRedline(
      fixture({ body: [p(run('Keep '), run('old', '<w:i/><w:iCs/>')), text('Gone text'), text('Stay')].join('') }),
      fixture({
        body: [
          p(
            run('Keep '),
            run('new', '<w:b/><w:bCs/>'),
            run(' hue', '<w:color w:val="FF0000"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>')
          ),
          '<w:p/>',
          text('Stay'),
        ].join(''),
      })
    );
  });

  it('keeps complex-script formatting as the source states or leaves it', async () => {
    const implicit =
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Carlito" w:hAnsi="Carlito"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>';
    const explicit =
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Carlito" w:hAnsi="Carlito" w:cs="Carlito"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>';
    const body = (word: string) =>
      [
        p(run('Bold '), run(word, '<w:b/>')),
        p(run('Italic '), run(word, '<w:i/>')),
        p(run('Script '), run(word, '<w:bCs/><w:iCs/>')),
        p(run('Both '), run(word, '<w:b/><w:bCs/>')),
        p(run('Sized '), run(word, '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="28"/>')),
      ].join('');
    for (const styles of [implicit, explicit]) {
      const { docx } = await expectRedline(
        fixture({ body: body('old'), styles }),
        fixture({ body: body('new'), styles })
      );
      if (styles === implicit) {
        const saved = decode(parts(docx)['word/document.xml']!);
        expect(saved).not.toContain('w:szCs');
        expect(saved).not.toContain('w:cs=');
        expect(saved.match(/<w:bCs\/>/g)).toHaveLength(4);
      }
    }
  });

  it('reads style toggles at every level of a basedOn chain', async () => {
    const styles =
      '<w:style w:type="paragraph" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Off"><w:name w:val="Off"/><w:basedOn w:val="Strong"/><w:rPr><w:b w:val="0"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Toggled"><w:name w:val="Toggled"/><w:basedOn w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>';
    const styled = (style: string, word: string) =>
      fixture({ body: p(`<w:pPr><w:pStyle w:val="${style}"/></w:pPr>`, run(`The ${word} clause`)), styles });
    await expectRedline(styled('Off', 'old'), styled('Off', 'new'));
    await expectRefusal(styled('Toggled', 'old'), styled('Toggled', 'new'), 'roundtrip-mismatch');
  });

  it('keeps Unicode, tabs and special hyphens exact in both granularities', async () => {
    const special = (tail: string) =>
      p(
        run('co'),
        '<w:r><w:softHyphen/></w:r>',
        run('op'),
        '<w:r><w:tab/></w:r>',
        run('x'),
        '<w:r><w:noBreakHyphen/></w:r>',
        run(tail)
      );
    const before = fixture({
      body: [
        text('Hi \u{1F600} there'),
        text('café au lait'),
        text('你好世界'),
        text('שלום world'),
        special('y'),
      ].join(''),
    });
    const after = fixture({
      body: [
        text('Hi \u{1F468}‍\u{1F469}‍\u{1F467} there'),
        text('café au lait'),
        text('你好朋友'),
        text('שלום there'),
        special('z\tw­v‑u'),
      ].join(''),
    });
    for (const granularity of ['word', 'char'] as const) {
      const { docx } = await expectRedline(before, after, { ...OPTIONS, granularity });
      const paragraph = parseXml(decode(parts(docx)['word/document.xml']!)).children[0]!.children[4]!;
      for (const special of ['tab', 'softHyphen', 'noBreakHyphen']) {
        expect(descendants(paragraph, 'ins').flatMap((ins) => descendants(ins, special))).toHaveLength(1);
      }
    }
  });

  it('patches only changed paragraphs around tables, opaque blocks and bookkeeping', async () => {
    const opaque = '<bofx:block xmlns:bofx="urn:fidelity" bofx:value="kept"/>';
    const noisy = (value: string) =>
      `<w:p w:rsidR="00AA11BB" w14:paraId="1A2B3C4D"><w:pPr><w:jc w:val="both"/></w:pPr><w:proofErr w:type="spellStart"/><w:r w:rsidR="00CC22DD"><w:t xml:space="preserve">${value}</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve"> tail</w:t></w:r></w:p>`;
    const body = (first: string, second: string) =>
      [text('Lead'), table('Cell'), noisy(first), opaque, text(second), table('Cell')].join('');
    const { changes } = await expectRedline(
      fixture({ body: body('Justified clause', 'After the block') }),
      fixture({ body: body('Justified clauses', 'After the opaque block') })
    );
    expect(changes.map((change) => [change.original.path, change.revised.text])).toEqual([
      [[0, 2], 'clauses'],
      [[0, 4], 'opaque '],
    ]);
  });

  it('addresses paragraphs by position when their ids repeat', async () => {
    const withId = (value: string) => text(value).replace('<w:p>', '<w:p w14:paraId="0000ABCD">');
    const result = await expectRedline(
      fixture({ body: [withId('First'), withId('Second')].join('') }),
      fixture({ body: [withId('First'), withId('Second edit')].join('') })
    );
    expect(result.diagnostics).toMatchObject([{ code: 'ambiguous-identity', severity: 'warning' }]);
  });

  it('returns the original bytes for identical documents within the output limit', async () => {
    const result = redline(await compareDocx(original, original, OPTIONS));
    expect(result.changes).toEqual([]);
    expect(result.docx).toEqual(original);
    expect(await compareDocx(original, original, { ...OPTIONS, limits: { maxOutputBytes: 1 } })).toMatchObject({
      ok: false,
      diagnostics: [{ code: 'limit-exceeded' }],
    });
  });

  it('refuses whole-paragraph, structural and unsupported changes without output', async () => {
    const base = fixture({ body: [text('One'), text('Two'), text('Three')].join('') });
    const cases: Array<[Uint8Array, Uint8Array, DocxCompareDiagnosticCode]> = [
      [base, fixture({ body: [text('One'), text('Two'), text('Inserted'), text('Three')].join('') }), 'paragraph-insertion'],
      [base, fixture({ body: [text('One'), text('Three')].join('') }), 'paragraph-deletion'],
      [base, fixture({ body: [text('One'), '<w:p/>', text('Two'), text('Three')].join('') }), 'paragraph-insertion'],
      [
        fixture({ body: `<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">${run('One')}</w:ins></w:p>` }),
        base,
        'existing-revisions',
      ],
      [
        fixture({ body: [text('One'), table('Cell', '<w:tcPrChange w:id="3" w:author="A"><w:tcPr/></w:tcPrChange>')].join('') }),
        fixture({ body: [text('One'), table('Cell')].join('') }),
        'existing-revisions',
      ],
      [base, fixture({ body: [text('One'), text('Two'), text('Three')].join(''), header: text('Changed header') }), 'out-of-scope-change'],
      [base, fixture({ body: [text('One'), text('Two'), text('Three')].join(''), theme: THEME.replace('Calibri"', 'Arial"') }), 'formatting-change'],
      [
        fixture({ body: [text('One'), table('Cell')].join('') }),
        fixture({ body: [text('One'), table('Cell edited')].join('') }),
        'table-change',
      ],
      [
        fixture({ body: `<w:sdt><w:sdtPr><w:tag w:val="clause"/></w:sdtPr><w:sdtContent>${text('Inside')}</w:sdtContent></w:sdt>` }),
        fixture({ body: `<w:sdt><w:sdtPr><w:tag w:val="clause"/></w:sdtPr><w:sdtContent>${text('Inside edited')}</w:sdtContent></w:sdt>` }),
        'content-control-change',
      ],
      [
        fixture({ body: p(run('Picture '), '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:docPr id="1" name="a"/></wp:inline></w:drawing></w:r>') }),
        fixture({ body: p(run('Picture '), '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:docPr id="1" name="b"/></wp:inline></w:drawing></w:r>') }),
        'object-change',
      ],
      [
        fixture({ body: p(run('Edited '), '<w:sdt><w:sdtContent><w:r><w:t>control</w:t></w:r></w:sdtContent></w:sdt>') }),
        fixture({ body: p(run('Changed '), '<w:sdt><w:sdtContent><w:r><w:t>control</w:t></w:r></w:sdtContent></w:sdt>') }),
        'unsupported-content',
      ],
      [
        fixture({ body: p(run('Keep '), run('old', '<w:shd w:val="pct20" w:color="auto" w:fill="FFFF00"/>')) }),
        fixture({ body: p(run('Keep '), run('new', '<w:shd w:val="pct20" w:color="auto" w:fill="FFFF00"/>')) }),
        'unsupported-formatting',
      ],
      [
        fixture({ body: text('the cat sat') }),
        fixture({ body: p(run('the '), run('car', '<w:b/><w:bCs/>'), run(' sat')) }),
        'formatting-change',
      ],
    ];
    for (const [before, after, code] of cases) await expectRefusal(before, after, code);
    await expectRefusal(
      fixture({ body: text('the cat sat') }),
      fixture({ body: p(run('the '), run('car', '<w:b/><w:bCs/>'), run(' sat')) }),
      'formatting-change',
      { ...OPTIONS, granularity: 'char' }
    );
  });

  it('refuses unusable options as data and throws for malformed ones', async () => {
    const base = fixture({ body: text('One') });
    for (const options of [
      { ...OPTIONS, author: '  ' },
      { ...OPTIONS, date: '2024-05-06' },
      { ...OPTIONS, limits: { maxChanges: 1000 } },
      { ...OPTIONS, limits: { maxResultBytes: 1023 } },
    ]) {
      expect(await compareDocx(base, base, options)).toMatchObject({
        ok: false,
        diagnostics: [{ code: 'invalid-options' }],
      });
    }
    expect(await compareDocx(new Uint8Array([1, 2, 3]), base, OPTIONS)).toMatchObject({
      ok: false,
      diagnostics: [{ code: 'invalid-docx' }],
    });
    await expect(
      compareDocx(base, base, { date: OPTIONS.date } as unknown as DocxCompareOptions)
    ).rejects.toThrow();
  });

  it('enforces change, output and diagnostic limits', async () => {
    for (const limits of [{ maxChanges: 2 }, { maxOutputBytes: 64 }]) {
      expect(await compareDocx(original, revised, { ...OPTIONS, limits })).toMatchObject({
        ok: false,
        diagnostics: [{ code: 'limit-exceeded' }],
      });
    }
    const many = fixture({ body: [text('a'), text('b'), text('c'), text('d')].join('') });
    const inserted = fixture({
      body: [text('a'), text('n1'), text('b'), text('n2'), text('c'), text('n3'), text('d')].join(''),
    });
    const truncated = await compareDocx(many, inserted, {
      ...OPTIONS,
      unsupported: 'report',
      limits: { maxDiagnostics: 2 },
    });
    expect(truncated).toMatchObject({
      ok: false,
      diagnostics: [{ code: 'paragraph-insertion' }, { code: 'diagnostics-truncated', severity: 'error' }],
    });
  });
});
