import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import type { LayoutBlock } from '../layout/pagination/types';
import type { ComplexField, Document, Run, SimpleField } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
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
    ['word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`],
  ]);
  return new Uint8Array(
    rezipPartsToArrayBuffer(new Map([...parts].map(([name, xml]) => [name, toBytes(xml)])))
  );
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
      expect(payloads[0]).not.toHaveProperty('nestedSequences');
      expect(payloads[1]).toMatchObject({ displayText: '2', nestedSequences: ['figure'] });
      expect(payloads[2]).not.toHaveProperty('nestedSequences');
    } finally {
      session.destroy();
    }
  });

  it('collects nested SEQ names from field trees with Rust token and case semantics', async () => {
    const run: Run = { type: 'run', content: [{ type: 'text', text: '2' }] };
    const simple = (instruction: string): SimpleField => ({
      type: 'simpleField',
      fieldType: 'SEQ',
      instruction,
      content: [run],
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
    const document: Document = {
      package: { document: { content: [{ type: 'paragraph', content: [quote] }] } },
    };
    const session = await createYrsSession({ clientId: 47008 });
    try {
      documentToYrs(session, document);
      const outer = session.storySegments('body').find((segment) =>
        segment.kind === 'embed' &&
        segment.embedKind === 'field' &&
        segment.payload.fieldType === 'QUOTE'
      );
      expect(outer?.kind === 'embed' && outer.payload.nestedSequences).toEqual([
        'figure',
        'table caption',
        'other',
        'result',
      ]);
    } finally {
      session.destroy();
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
      expect(quote?.kind === 'embed' && quote.payload.nestedSequences).toEqual(['figure']);
      hydrated.openDocx(bytes, false);
      hydrated.loadState(projected.encodeState());
      const blocks = hydrated.yrsBlocksForStory('body', {}) as LayoutBlock[];
      const results = blocks.flatMap((block) => block.kind === 'paragraph' ? block.runs : [])
        .flatMap((run) => run.kind === 'field' && run.rawType === 'SEQ' ? [run.fallback] : []);
      expect(results).toEqual(['2', '1']);
    } finally {
      projected.destroy();
      hydrated.destroy();
    }
  });

  it('keeps cached body SEQs after a typed SEQ in a hyperlink SDT when seeding and hydrating', async () => {
    const field = (cached: string): string => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Figure </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${cached}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const bytes = sequencePackage(`<w:p><w:hyperlink w:anchor="top"><w:sdt><w:sdtPr/><w:sdtContent>${field('1')}</w:sdtContent></w:sdt></w:hyperlink></w:p><w:p>${field('2')}</w:p>`);
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const engine = await createYrsSession({ clientId: 47011 });
    const projected = await createYrsSession({ clientId: 47012 });
    const hydrated = await createYrsSession({ clientId: 47013 });
    try {
      engine.seedFromDocx(bytes);
      documentToYrs(projected, parsed);
      hydrated.openDocx(bytes, false);
      hydrated.loadState(projected.encodeState());
      for (const session of [engine, hydrated]) {
        const blocks = session.yrsBlocksForStory('body', {}) as LayoutBlock[];
        const results = blocks.flatMap((block) => block.kind === 'paragraph' ? block.runs : [])
          .flatMap((run) => run.kind === 'field' && run.rawType === 'SEQ' ? [run.fallback] : []);
        expect(results).toEqual(['2']);
      }
    } finally {
      engine.destroy();
      projected.destroy();
      hydrated.destroy();
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
