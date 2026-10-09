import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import type { LayoutBlock, Run as LayoutRun, ShapeBlock } from '../layout/pagination/types';
import type { ComplexField, Document, ParagraphContent, Run, SimpleField } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';
import { documentToYrs } from './documentToYrs';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURE = resolve(import.meta.dir, '../../../../apps/demo/public/betteroffice-demo.docx');
const EXISTING_ROOM_SEED = resolve(import.meta.dir, '../../../../apps/demo/public/seeds/docx.bin');

function sequenceHyperlinkPackage(instruction: string): Uint8Array<ArrayBuffer> {
  const hyperlink = `<w:hyperlink w:anchor="top"><w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>${instruction}<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>`;
  const body = `<w:p><w:fldSimple w:instr="QUOTE">${hyperlink}</w:fldSimple></w:p><w:p><w:fldSimple w:instr="SEQ Figure"><w:r><w:t>2</w:t></w:r></w:fldSimple></w:p><w:p><w:fldSimple w:instr="SEQ Table"><w:r><w:t>7</w:t></w:r></w:fldSimple></w:p>`;
  return sequencePackage(body);
}

function sequencePackage(body: string): Uint8Array<ArrayBuffer> {
  const parts = new Map([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'],
    ['_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'],
    ['word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body>${body}<w:sectPr/></w:body></w:document>`],
  ]);
  return new Uint8Array(
    rezipPartsToArrayBuffer(new Map([...parts].map(([name, xml]) => [name, toBytes(xml)])))
  );
}

function sequenceTextBox(content: string, hidden = false): string {
  return `<w:r>${hidden ? '<w:rPr><w:vanish/></w:rPr>' : ''}<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1600000" cy="228600"/><wp:wrapNone/><wp:docPr id="1" name="Box 1"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1600000" cy="228600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p>${content}</w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
}

function shapeSequenceRuns(shape: ShapeBlock): LayoutRun[] {
  return [
    ...(shape.innerText ?? []).flatMap((paragraph) => paragraph.runs),
    ...(shape.children ?? []).flatMap(shapeSequenceRuns),
  ];
}

function sequenceRuns(blocks: readonly LayoutBlock[]): LayoutRun[] {
  return blocks.flatMap((block) => {
    switch (block.kind) {
      case 'paragraph': return block.runs;
      case 'shape': return shapeSequenceRuns(block);
      case 'textBox': return block.content.flatMap((paragraph) => paragraph.runs);
      case 'table': return block.rows.flatMap((row) => row.cells.flatMap((cell) => sequenceRuns(cell.blocks)));
      default: return [];
    }
  });
}

async function seedSequenceSessions(
  bytes: Uint8Array<ArrayBuffer>,
  edit?: (session: YrsSession) => void
) {
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  const [engine, projected, plain, nativePeer, projectedPeer, hydrated] = await Promise.all([
    createYrsSession({ clientId: 47030 }),
    createYrsSession({ clientId: 47031 }),
    createYrsSession({ clientId: 47032 }),
    createYrsSession({ clientId: 47033 }),
    createYrsSession({ clientId: 47034 }),
    createYrsSession({ clientId: 47035 }),
  ]);
  const sessions = [engine, projected, plain, nativePeer, projectedPeer, hydrated];
  try {
    engine.seedFromDocx(bytes);
    documentToYrs(projected, parsed);
    documentToYrs(plain, withoutSourceOrdinals(parsed) as Document);
    for (const session of [engine, projected, plain]) edit?.(session);
    nativePeer.loadState(engine.encodeState());
    projectedPeer.loadState(projected.encodeState());
    hydrated.openDocx(bytes, false);
    hydrated.loadState(projected.encodeState());
    return { engine, projected, plain, nativePeer, projectedPeer, hydrated, sessions };
  } catch (error) {
    for (const session of sessions) session.destroy();
    throw error;
  }
}

/** The model without the source occurrences a parse for an editing session records. */
function withoutSourceOrdinals(value: unknown): unknown {
  if (value instanceof Map) {
    return new Map([...value].map(([key, entry]) => [key, withoutSourceOrdinals(entry)]));
  }
  if (Array.isArray(value)) return value.map(withoutSourceOrdinals);
  if (value !== null && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .filter(([key]) => key !== 'sourceOrdinal')
        .map(([key, entry]) => [key, withoutSourceOrdinals(entry)])
    );
  }
  return value;
}

function expectEquivalentStories(left: YrsSession, right: YrsSession): void {
  expect(left.storyIds()).toEqual(right.storyIds());
  for (const storyId of left.storyIds()) {
    expect(left.storySegments(storyId)).toEqual(right.storySegments(storyId));
  }
}

function seedClocks(session: YrsSession): Map<bigint, bigint> {
  const vector = session.encodeStateVector();
  let offset = 0;
  const read = (): bigint => {
    let value = 0n;
    let shift = 0n;
    while (offset < vector.length) {
      const byte = vector[offset++];
      value += BigInt(byte & 0x7f) << shift;
      if ((byte & 0x80) === 0) return value;
      shift += 7n;
    }
    throw new Error('Truncated state vector');
  };
  const clocks = new Map<bigint, bigint>();
  for (let remaining = read(); remaining > 0n; remaining--) {
    const client = read();
    const clock = read();
    if (client !== 0x1_0000_05e9n) clocks.set(client, clock);
  }
  expect(offset).toBe(vector.length);
  return clocks;
}

describe('DOCX engine seeding', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('seeds nested SEQ names without adding them to plain fields', async () => {
    const sequence = (cached: string): ComplexField => ({
      type: 'complexField',
      fieldType: 'SEQ',
      instruction: 'SEQ Figure',
      fieldCode: [],
      fieldResult: [{ type: 'run', content: [{ type: 'text', text: cached }] }],
    });
    const quote: ComplexField = {
      ...sequence('2'),
      fieldType: 'QUOTE',
      instruction: 'QUOTE "{ SEQ Figure }"',
      structuredCode: { inline: [sequence('2')] },
    };
    const document: Document = {
      package: {
        document: {
          content: [{ type: 'paragraph', content: [sequence('1'), quote, sequence('3')] }],
        },
      },
    };
    const session = await createYrsSession({ clientId: 47007 });
    try {
      documentToYrs(session, document);
      const payloads = session.storySegments('body').flatMap((segment) =>
        segment.kind === 'embed' && segment.embedKind === 'field' ? [segment.payload] : []
      );
      expect(payloads).toHaveLength(3);
      for (const payload of payloads) expect(payload).not.toHaveProperty('nestedSequences');
      expect(payloads[1]).toMatchObject({ displayText: '2' });
      const blocks = session.yrsBlocksForStory('body', {}) as LayoutBlock[];
      const results = sequenceRuns(blocks)
        .flatMap((run) => run.kind === 'field' && run.rawType === 'SEQ' ? [run.fallback] : []);
      expect(results).toEqual(['1', '3']);
    } finally {
      session.destroy();
    }
  });

  it.each(['field', 'hyperlink', 'sdt'])('collects nested SEQ names from field trees with Rust token and case semantics (%s)', async (placement) => {
    const run: Run = { type: 'run', content: [{ type: 'text', text: '2' }] };
    const simple = (instruction: string, cached = '2'): SimpleField => ({
      type: 'simpleField',
      fieldType: 'SEQ',
      instruction,
      content: [{ type: 'run', content: [{ type: 'text', text: cached }] }],
    });
    const quote: ComplexField = {
      type: 'complexField',
      fieldType: 'QUOTE',
      instruction: 'QUOTE "cached"',
      fieldCode: [],
      fieldResult: [run],
      fieldTree: {
        code: {
          inline: [
            simple(' sEq "Figure" \\r 1'),
            simple('SEQ FIGURE'),
            simple('SEQ "Table Caption"'),
            simple('SEQ \\r 1'),
            simple('SEQ'),
            simple('ſEQ Ignored'),
            simple('"" Ignored'),
            simple('"SEQ"Other'),
          ],
        },
        children: [{ result: { inline: [simple('SEQ FIGURE')] } }],
      },
      structuredCode: { inline: [simple('SEQ Other')] },
      structuredResult: {
        inline: [{
          type: 'complexField',
          fieldType: 'SEQ',
          instruction: 'SEQ Result',
          fieldCode: [],
          fieldResult: [run],
        }],
      },
    };
    const content: ParagraphContent = placement === 'hyperlink'
      ? { type: 'hyperlink', anchor: 'top', children: [], structuredChildren: [quote] }
      : placement === 'sdt'
        ? { type: 'inlineSdt', properties: { sdtType: 'richText' }, content: [quote] }
        : quote;
    const instructions = ['figure', 'table caption', 'other', 'result', 'ignored']
      .map((name) => `SEQ "${name}"`);
    const document: Document = {
      package: {
        document: {
          content: [
            { type: 'paragraph', content: [content] },
            ...instructions.map((instruction) => ({
              type: 'paragraph' as const,
              content: [simple(instruction, '9')],
            })),
          ],
        },
      },
    };
    const session = await createYrsSession({ clientId: 47008 });
    try {
      documentToYrs(session, document);
      const payloads = session.storySegments('body').flatMap((segment) =>
        segment.kind === 'embed' ? [segment.payload] : []
      );
      expect(JSON.stringify(payloads)).not.toContain('nestedSequences');
      const blocks = session.yrsBlocksForStory('body', {}) as LayoutBlock[];
      const results = sequenceRuns(blocks).flatMap((run) =>
        run.kind === 'field' && run.rawType === 'SEQ' && instructions.includes(run.instruction ?? '')
          ? [run.fallback] : []
      );
      expect(results).toEqual(['9', '9', '9', '9', '1']);
    } finally {
      session.destroy();
    }
  });

  it('keeps seed clocks unchanged by nested SEQ names', async () => {
    const seed = async (kind: string): Promise<Map<bigint, bigint>> => {
      const field = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${kind} Figure </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
      const nested = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> QUOTE "</w:instrText></w:r>${field}<w:r><w:instrText xml:space="preserve">" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
      const bytes = sequencePackage(`<w:p>${nested}</w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>`);
      const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
      const session = await createYrsSession({ clientId: 47039 });
      try {
        documentToYrs(session, parsed);
        return seedClocks(session);
      } finally {
        session.destroy();
      }
    };
    const sequence = await seed('SEQ');
    const control = await seed('XEQ');
    expect(sequence.size).toBeGreaterThan(0);
    expect(sequence).toEqual(control);
  });

  it('keeps seed clocks unchanged by SEQs in hidden text boxes', async () => {
    const seed = async (kind: string): Promise<Map<bigint, bigint>> => {
      const boxed = `<w:fldSimple w:instr="${kind} Figure"><w:r><w:t>1</w:t></w:r></w:fldSimple>`;
      const bytes = sequencePackage(`<w:p>${sequenceTextBox(boxed, true)}</w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>`);
      const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
      const session = await createYrsSession({ clientId: 47040 });
      try {
        documentToYrs(session, parsed);
        const drawing = session.storySegments('body').find((segment) =>
          segment.kind === 'embed' && segment.embedKind === 'shape'
        );
        expect(drawing).toBeDefined();
        expect(drawing?.attributes).not.toHaveProperty('hidden');
        return seedClocks(session);
      } finally {
        session.destroy();
      }
    };
    const sequence = await seed('SEQ');
    const control = await seed('XEQ');
    expect(sequence.size).toBeGreaterThan(0);
    expect(sequence).toEqual(control);
  });

  it.each(['hyperlink', 'simpleField'])('keeps projected SEQ owners opaque when seeding and hydrating (%s)', async (kind) => {
    const result = kind === 'hyperlink'
      ? '<w:hyperlink w:anchor="top"><w:r><w:t>1</w:t></w:r></w:hyperlink>'
      : '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>';
    const bytes = sequencePackage(`<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Figure </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`);
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const engine = await createYrsSession({ clientId: 47015 });
    const projected = await createYrsSession({ clientId: 47016 });
    const hydrated = await createYrsSession({ clientId: 47017 });
    try {
      engine.seedFromDocx(bytes);
      documentToYrs(projected, parsed);
      hydrated.openDocx(bytes, false);
      hydrated.loadState(projected.encodeState());
      for (const session of [engine, projected, hydrated]) {
        const owner = session.storySegments('body').find((segment) =>
          segment.kind === 'embed' &&
          segment.embedKind === 'field' &&
          segment.payload.fieldType === 'SEQ'
        );
        expect(owner).toBeDefined();
        expect(owner?.kind === 'embed' && owner.payload).not.toHaveProperty('nestedSequences');
        const blocks = session.yrsBlocksForStory('body', {}) as LayoutBlock[];
        const text = blocks.flatMap((block) => block.kind === 'paragraph' ? block.runs : [])
          .map((run) => run.kind === 'text' ? run.text : run.kind === 'field' ? run.fallback ?? '' : '')
          .join('');
        expect(text).toBe('1');
      }
    } finally {
      engine.destroy();
      projected.destroy();
      hydrated.destroy();
    }
  });

  it.each([
    ['native', 'hyperlink'],
    ['native', 'simpleField'],
    ['typescript', 'hyperlink'],
    ['typescript', 'simpleField'],
  ])('keeps legacy projected SEQ owners opaque (%s, %s)', async (seeder, kind) => {
    const result = kind === 'hyperlink'
      ? '<w:hyperlink w:anchor="top"><w:r><w:t>1</w:t></w:r></w:hyperlink>'
      : '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>';
    const bytes = sequencePackage(`<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Figure </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`);
    const source = await createYrsSession({ clientId: 47036 });
    const peer = await createYrsSession({ clientId: 47037 });
    const hydrated = await createYrsSession({ clientId: 47038 });
    try {
      if (seeder === 'native') source.seedFromDocx(bytes);
      else documentToYrs(source, await parseDocx(bytes.buffer, { preloadFonts: false }));
      const owners = source.storySegments('body').filter((segment) =>
        segment.kind === 'embed' && segment.payload.resultProjection
      );
      expect(owners).toHaveLength(1);
      peer.loadState(source.encodeState());
      hydrated.openDocx(bytes, false);
      hydrated.loadState(source.encodeState());
      for (const session of [source, peer, hydrated]) {
        const owner = session.storySegments('body').find((segment) =>
          segment.kind === 'embed' && segment.payload.resultProjection
        );
        expect(owner).toBeDefined();
        expect(owner?.kind === 'embed' && owner.payload).not.toHaveProperty('nestedSequences');
        for (const showHiddenText of [false, true]) {
          const blocks = session.yrsBlocksForStory('body', { showHiddenText }) as LayoutBlock[];
          const text = sequenceRuns(blocks)
            .map((run) => run.kind === 'text' ? run.text : run.kind === 'field' ? run.fallback ?? '' : '')
            .join('');
          expect(text).toBe('1');
        }
      }
    } finally {
      source.destroy();
      peer.destroy();
      hydrated.destroy();
    }
  });

  it('keeps the visible caption at 2 when hidden SEQs are seeded and hydrated', async () => {
    const bytes = sequencePackage('<w:p><w:fldSimple w:instr="SEQ Figure"><w:r><w:rPr><w:vanish/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p><w:p><w:fldSimple w:instr="SEQ Figure"><w:r><w:t>2</w:t></w:r></w:fldSimple></w:p>');
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const engine = await createYrsSession({ clientId: 47018 });
    const projected = await createYrsSession({ clientId: 47019 });
    const hydrated = await createYrsSession({ clientId: 47020 });
    try {
      engine.seedFromDocx(bytes);
      documentToYrs(projected, parsed);
      hydrated.openDocx(bytes, false);
      hydrated.loadState(projected.encodeState());
      for (const session of [engine, projected, hydrated]) {
        for (const showHiddenText of [false, true]) {
          const blocks = session.yrsBlocksForStory('body', { showHiddenText }) as LayoutBlock[];
          const results = blocks.flatMap((block) => block.kind === 'paragraph' ? block.runs : [])
            .flatMap((run) => run.kind === 'field' && run.rawType === 'SEQ' ? [run.fallback] : []);
          expect(results).toEqual(showHiddenText ? ['1', '2'] : ['2']);
        }
      }
    } finally {
      engine.destroy();
      projected.destroy();
      hydrated.destroy();
    }
  });

  it.each(['run', 'sdt', 'nested hyperlink'])('seeds and hydrates raw hyperlink SEQs nested in a simple field (%s)', async (wrapper) => {
    const instruction = '<w:r><w:instrText> SEQ Figure </w:instrText></w:r>';
    const content = wrapper === 'nested hyperlink'
      ? `<w:hyperlink w:anchor="top">${instruction}</w:hyperlink>`
      : instruction;
    const bytes = sequenceHyperlinkPackage(wrapper === 'run'
      ? content
      : `<w:sdt><w:sdtPr/><w:sdtContent>${content}</w:sdtContent></w:sdt>`);
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const projected = await createYrsSession({ clientId: 47009 });
    const hydrated = await createYrsSession({ clientId: 47010 });
    try {
      documentToYrs(projected, parsed);
      const quote = projected.storySegments('body').find((segment) =>
        segment.kind === 'embed' &&
        segment.embedKind === 'field' &&
        segment.payload.fieldType === 'QUOTE'
      );
      expect(quote).toBeDefined();
      expect(quote?.kind === 'embed' && quote.payload).not.toHaveProperty('nestedSequences');
      hydrated.openDocx(bytes, false);
      hydrated.loadState(projected.encodeState());
      for (const session of [projected, hydrated]) {
        const blocks = session.yrsBlocksForStory('body', {}) as LayoutBlock[];
        const results = blocks.flatMap((block) => block.kind === 'paragraph' ? block.runs : [])
          .flatMap((run) => run.kind === 'field' && run.rawType === 'SEQ' ? [run.fallback] : []);
        expect(results).toEqual(['2', '1']);
      }
    } finally {
      projected.destroy();
      hydrated.destroy();
    }
  });

  it('keeps cached body SEQs after a typed SEQ in a hyperlink SDT when seeding and hydrating', async () => {
    const field = (cached: string): string => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Figure </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${cached}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const bytes = sequencePackage(`<w:p><w:hyperlink w:anchor="top"><w:sdt><w:sdtPr/><w:sdtContent>${field('1')}</w:sdtContent></w:sdt></w:hyperlink></w:p><w:p>${field('2')}</w:p>`);
    const { engine, projected, plain, nativePeer, projectedPeer, hydrated, sessions } = await seedSequenceSessions(bytes);
    try {
      for (const session of [engine, projected, plain, nativePeer, projectedPeer, hydrated]) {
        const paragraph = session.storySegments('body').find((segment) => segment.kind === 'pilcrow');
        expect(paragraph?.kind === 'pilcrow' && paragraph.properties).not.toHaveProperty('opaqueSequences');
        for (const showHiddenText of [false, true]) {
          const blocks = session.yrsBlocksForStory('body', { showHiddenText }) as LayoutBlock[];
          const results = sequenceRuns(blocks)
            .flatMap((run) => run.kind === 'field' && run.rawType === 'SEQ' ? [run.fallback] : []);
          expect(results).toEqual(['2']);
        }
      }
    } finally {
      for (const session of sessions) session.destroy();
    }
  });

  it('keeps document opacity after deleting a hyperlink SEQ across seed and hydration paths', async () => {
    const hyperlink = '<w:hyperlink w:anchor="top"><w:fldSimple w:instr="SEQ Figure" w:fldLock="true"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:hyperlink>';
    const bytes = sequencePackage(`<w:p><w:r><w:t>x</w:t></w:r></w:p><w:p>${hyperlink}</w:p><w:p><w:fldSimple w:instr="SEQ Figure"><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>`);
    const { engine, hydrated, sessions } = await seedSequenceSessions(bytes, (session) => {
      const [first, linked] = session.paragraphs('body');
      session.deleteRange({
        story: 'body',
        start: { paraId: first.paraId, offset: 1 },
        end: { paraId: linked.paraId, offset: 1 },
      });
    });
    try {
      for (const session of sessions) {
        expect(session.paragraphs('body')[0].text).toBe('x');
        for (const showHiddenText of [false, true]) {
          const blocks = session.yrsBlocksForStory('body', { showHiddenText }) as LayoutBlock[];
          const results = sequenceRuns(blocks)
            .flatMap((run) => run.kind === 'field' && run.rawType === 'SEQ' ? [run.fallback] : []);
          expect(results).toEqual(['9']);
        }
      }
      for (const session of [engine, hydrated]) {
        const saved = await saveYrsDocx(session);
        for (const part of Object.values(unzipContainer(saved.bytes))) {
          expect(new TextDecoder().decode(part)).not.toContain('opaqueSequences');
        }
      }
    } finally {
      for (const session of sessions) session.destroy();
    }
  });

  it('keeps document opacity after merging a typed hyperlink SDT across seed and hydration paths', async () => {
    const field = (cached: string): string => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Figure </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${cached}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const bytes = sequencePackage(`<w:p><w:r><w:t>x</w:t></w:r></w:p><w:p><w:hyperlink w:anchor="top"><w:sdt><w:sdtPr/><w:sdtContent>${field('1')}</w:sdtContent></w:sdt></w:hyperlink></w:p><w:p>${field('2')}</w:p>`);
    const { sessions } = await seedSequenceSessions(bytes, (session) => {
      session.mergeParagraphs('body', session.paragraphs('body')[0].paraId);
    });
    try {
      for (const session of sessions) {
        const paragraphs = session.paragraphs('body');
        expect(paragraphs).toHaveLength(2);
        for (const paragraph of paragraphs) {
          expect(paragraph.properties).not.toHaveProperty('opaqueSequences');
        }
        for (const showHiddenText of [false, true]) {
          const blocks = session.yrsBlocksForStory('body', { showHiddenText }) as LayoutBlock[];
          const results = sequenceRuns(blocks)
            .flatMap((run) => run.kind === 'field' && run.rawType === 'SEQ' ? [run.fallback] : []);
          expect(results).toEqual(['2']);
        }
      }
    } finally {
      for (const session of sessions) session.destroy();
    }
  });

  it('preserves stale captions after hidden text boxes across seed and hydration paths', async () => {
    const boxed = '<w:fldSimple w:instr="SEQ Figure"><w:r><w:t>1</w:t></w:r></w:fldSimple>';
    const bytes = sequencePackage(`<w:p>${sequenceTextBox(boxed, true)}</w:p><w:p><w:fldSimple w:instr="SEQ Figure"><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>`);
    const { engine, projected, plain, nativePeer, projectedPeer, hydrated, sessions } = await seedSequenceSessions(bytes);
    try {
      for (const session of [engine, nativePeer]) {
        const drawing = session.storySegments('body').find((segment) =>
          segment.kind === 'embed' && segment.embedKind === 'shape'
        );
        expect(drawing?.attributes.hidden).toBe(true);
        for (const showHiddenText of [false, true]) {
          const blocks = session.yrsBlocksForStory('body', { showHiddenText }) as LayoutBlock[];
          const results = sequenceRuns(blocks)
            .flatMap((run) => run.kind === 'field' && run.rawType === 'SEQ' ? [run.fallback] : []);
          expect(results).toEqual(showHiddenText ? ['1', '2'] : ['9']);
          expect(blocks.some((block) => block.kind === 'shape')).toBe(showHiddenText);
        }
      }
      for (const session of [projected, plain, projectedPeer, hydrated]) {
        const drawing = session.storySegments('body').find((segment) =>
          segment.kind === 'embed' && segment.embedKind === 'shape'
        );
        expect(drawing).toBeDefined();
        expect(drawing?.attributes).not.toHaveProperty('hidden');
        for (const showHiddenText of [false, true]) {
          const blocks = session.yrsBlocksForStory('body', { showHiddenText }) as LayoutBlock[];
          const results = sequenceRuns(blocks)
            .flatMap((run) => run.kind === 'field' && run.rawType === 'SEQ' ? [run.fallback] : []);
          expect(results).toEqual(['1', '9']);
          expect(blocks.some((block) => block.kind === 'shape')).toBe(true);
        }
      }
    } finally {
      for (const session of sessions) session.destroy();
    }
  });

  it.each(['', '0'])('preserves visible cached shape field fragments (%s)', async (hiddenText) => {
    const hiddenResult = hiddenText ? `<w:t>${hiddenText}</w:t>` : '';
    const cached = `<w:fldSimple w:instr="SEQ Figure" w:fldLock="true"><w:r><w:rPr><w:vanish/></w:rPr>${hiddenResult}</w:r><w:r><w:t>1</w:t></w:r></w:fldSimple>`;
    const bytes = sequencePackage(`<w:p>${sequenceTextBox(cached)}</w:p><w:p><w:fldSimple w:instr="SEQ Figure"><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>`);
    const { sessions } = await seedSequenceSessions(bytes);
    try {
      for (const session of sessions) {
        for (const showHiddenText of [false, true]) {
          const blocks = session.yrsBlocksForStory('body', { showHiddenText }) as LayoutBlock[];
          const results = sequenceRuns(blocks)
            .flatMap((run) => run.kind === 'field' && run.rawType === 'SEQ' ? [run.fallback] : []);
          expect(results).toEqual([
            showHiddenText ? `${hiddenText}1` : '1',
            !showHiddenText && hiddenText ? '9' : '2',
          ]);
        }
      }
    } finally {
      for (const session of sessions) session.destroy();
    }
  });

  it('omits line breaks from hidden shape runs across seed and hydration paths', async () => {
    const content = '<w:r><w:t>Before</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:br/></w:r><w:r><w:t>After</w:t></w:r>';
    const bytes = sequencePackage(`<w:p>${sequenceTextBox(content)}</w:p>`);
    const { sessions } = await seedSequenceSessions(bytes);
    try {
      for (const session of sessions) {
        for (const showHiddenText of [false, true]) {
          const blocks = session.yrsBlocksForStory('body', { showHiddenText }) as LayoutBlock[];
          const runs = sequenceRuns(blocks);
          expect(runs.filter((run) => run.kind === 'lineBreak')).toHaveLength(showHiddenText ? 1 : 0);
          expect(runs.flatMap((run) => run.kind === 'text' ? [run.text] : []).join('')).toBe('BeforeAfter');
        }
      }
    } finally {
      for (const session of sessions) session.destroy();
    }
  });

  it('tolerates missing arrays and null children in nested hyperlinks', async () => {
    const quote = {
      type: 'simpleField',
      fieldType: 'QUOTE',
      instruction: 'QUOTE',
      content: [],
      structuredResult: {
        inline: [
          { type: 'hyperlink' },
          { type: 'hyperlink', children: {} },
          { type: 'hyperlink', structuredChildren: [
            null,
            { type: 'inlineSdt' },
            { type: 'simpleField' },
            { type: 'complexField' },
            { type: 'run' },
            { type: 'run', content: [
              null,
              { type: 'fieldChar', charType: 'begin' },
              { type: 'instrText' },
              { type: 'fieldChar', charType: 'end' },
            ] },
          ] },
        ],
      },
    } as unknown as SimpleField;
    const document: Document = {
      package: { document: { content: [{ type: 'paragraph', content: [quote] }] } },
    };
    const session = await createYrsSession({ clientId: 47014 });
    try {
      expect(() => documentToYrs(session, document)).not.toThrow();
    } finally {
      session.destroy();
    }
  });

  it('produces equivalent story structure and state updates', async () => {
    const bytes = Uint8Array.from(readFileSync(FIXTURE));
    const parsed = await parseDocx(bytes.buffer);
    const projected = await createYrsSession({ clientId: 47001 });
    const engine = await createYrsSession({ clientId: 47001 });
    try {
      documentToYrs(projected, parsed, { generation: 'parity' });
      engine.seedFromDocx(bytes, { generation: 'parity' });

      expectEquivalentStories(engine, projected);
      expect(engine.encodeStateVector()).toEqual(projected.encodeStateVector());
      expect(engine.encodeState()).toEqual(projected.encodeState());

      const projectedStatePeer = await createYrsSession({ clientId: 47002 });
      const engineStatePeer = await createYrsSession({ clientId: 47003 });
      try {
        projectedStatePeer.loadState(projected.encodeState());
        engineStatePeer.loadState(engine.encodeState());
        const firstParagraph = projectedStatePeer.paragraphs('body')[0];
        projectedStatePeer.insertText(
          { story: 'body', paraId: firstParagraph.paraId, offset: 1 },
          'legacy'
        );
        engineStatePeer.loadState(
          projectedStatePeer.encodeStateAsUpdate(engineStatePeer.encodeStateVector())
        );
        const secondParagraph = engineStatePeer.paragraphs('body')[1];
        engineStatePeer.insertText(
          { story: 'body', paraId: secondParagraph.paraId, offset: 1 },
          'engine'
        );
        projectedStatePeer.loadState(
          engineStatePeer.encodeStateAsUpdate(projectedStatePeer.encodeStateVector())
        );
        expectEquivalentStories(engineStatePeer, projectedStatePeer);
      } finally {
        projectedStatePeer.destroy();
        engineStatePeer.destroy();
      }
    } finally {
      projected.destroy();
      engine.destroy();
    }
  });

  it('preserves committed room story structure and state vector', async () => {
    const bytes = Uint8Array.from(readFileSync(FIXTURE));
    const existingRoom = await createYrsSession({ clientId: 47004 });
    const engine = await createYrsSession({ clientId: 1 });
    try {
      existingRoom.loadState(Uint8Array.from(readFileSync(EXISTING_ROOM_SEED)));
      engine.seedFromDocx(bytes);

      expectEquivalentStories(engine, existingRoom);
      expect(engine.encodeStateVector()).toEqual(existingRoom.encodeStateVector());
      const before = existingRoom.encodeStateVector();
      existingRoom.openDocx(bytes, false);
      expect(existingRoom.encodeStateVector()).toEqual(before);
    } finally {
      existingRoom.destroy();
      engine.destroy();
    }
  });

  it('returns thin host metadata and materializes the canonical package on demand', async () => {
    const bytes = Uint8Array.from(readFileSync(FIXTURE));
    const parsed = await parseDocx(bytes.buffer);
    const engine = await createYrsSession({ clientId: 47005 });
    const existingRoom = await createYrsSession({ clientId: 47006 });
    try {
      const host = engine.openDocx(bytes, false);

      expect(engine.storyIds()).toEqual([]);
      expect(host.document.package.document.content).toEqual([]);
      expect(
        host.document.package.document.sections?.every((section) => section.content.length === 0)
      ).toBe(true);
      expect(
        [...(host.document.package.headers?.values() ?? [])].every(
          (header) => header.content.length === 0
        )
      ).toBe(true);
      expect(host.document.package.media?.size).toBe(0);
      expect(host.document.package.charts?.size).toBe(0);
      expect(host.referencedFonts.length).toBeGreaterThan(0);

      const materialized = engine.materializeDocx();
      expect(materialized?.package.document.content[0]).toMatchObject({ sourceOrdinal: 0 });
      expect(withoutSourceOrdinals(materialized?.package.document.content)).toEqual(
        parsed.package.document.content
      );
      expect(withoutSourceOrdinals(materialized?.package.document.sections)).toEqual(
        parsed.package.document.sections
      );
      expect(withoutSourceOrdinals(materialized?.package.headers)).toEqual(parsed.package.headers);
      expect(withoutSourceOrdinals(materialized?.package.footers)).toEqual(parsed.package.footers);

      existingRoom.loadState(Uint8Array.from(readFileSync(EXISTING_ROOM_SEED)));
      engine.loadState(existingRoom.encodeState());
      expectEquivalentStories(engine, existingRoom);
    } finally {
      engine.destroy();
      existingRoom.destroy();
    }
  });
});
