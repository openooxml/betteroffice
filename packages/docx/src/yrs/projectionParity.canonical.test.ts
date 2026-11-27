import { describe, expect, it } from 'bun:test';

import type { Paragraph } from '../types/content/paragraph';
import type { BlockContent, Document } from '../types/document';
import type { EditSession } from '../wasm/edit';
import type { YrsSession } from './index';
import { projectionAllowlist } from './projectionParity.allowlist';
import {
  canonicalProjection,
  canonicalStories,
  check,
  checkSourceOrdinals,
  checkStory,
  compareSnapshots,
  nativeProjection,
  type AllowanceUsage,
  type Deferral,
} from './projectionParity.support';

function wireDocument(content: readonly BlockContent[] = []) {
  return { package: {
    document: { content }, theme: {}, numbering: {}, settings: {}, fontTable: {},
    relationshipEntries: [], mediaEntries: [], chartEntries: [],
  } };
}

function fieldWithParagraph(paragraph: Paragraph, ...rest: Paragraph[]): Paragraph {
  return {
    type: 'paragraph',
    content: [{
      type: 'complexField',
      fieldType: 'IF',
      instruction: 'IF',
      fieldCode: [],
      fieldResult: [],
      structuredResult: { blocks: [paragraph, ...rest] },
    }],
  };
}

function textParagraph(text: string, sourceOrdinal: number): Paragraph {
  return { type: 'paragraph', sourceOrdinal, content: [{ type: 'run', content: [{ type: 'text', text }] }] };
}

function tableWithParagraph(paragraph: Paragraph): BlockContent {
  return { type: 'table', rows: [{ type: 'tableRow', cells: [{ type: 'tableCell', content: [paragraph] }] }] };
}

describe('DOCX projection canonical form', () => {
  it('validates the native document projection envelope version', () => {
    const raw = { project_document_json: () => JSON.stringify({
      wireVersion: 2,
      document: wireDocument(),
      embeddedFontParts: [],
      paragraphSessionKeys: [],
      sessionRevisionPlaceholders: [],
      createdListSites: [],
      deferred: [],
    }) } as unknown as EditSession;
    expect(() => nativeProjection(raw, new ArrayBuffer(0))).toThrow('wire.wireVersion must be 1');
  });

  it('reports every unlisted native deferral in equal documents and stories', () => {
    const session = {} as YrsSession;
    const document = wireDocument() as unknown as Document;
    const stories = canonicalStories(session, document);
    const deferred: Deferral[] = [
      { family: 'fields', kind: 'unlisted', storyId: 'body', paragraph: 0 },
      { family: 'fields', kind: 'table', storyId: 'body:t0:r0c0', paragraph: null },
    ];
    const native = { document, paragraphSessionKeys: [], sessionRevisionPlaceholders: [], createdListSites: [], deferred };
    const differences = check(session, document, native, stories);
    expect(differences).toHaveLength(deferred.length);
    expect(differences[0]).toContain('story body, paragraph 0, deferred: fields/unlisted');
    expect(differences[1]).toContain('story body:t0:r0c0, paragraph -, deferred: fields/table');
    expect(checkStory(stories, { ...native, storyId: 'body', blocks: [] }, stories)).toEqual(differences);
  });

  it('counts a named deferral and leaves unmatched difference allowances stale', () => {
    const session = {} as YrsSession;
    const document = wireDocument() as unknown as Document;
    const stories = canonicalStories(session, document);
    const table: Deferral = { family: 'tables', kind: 'table', storyId: 'body', paragraph: 0 };
    const native = { document, paragraphSessionKeys: [], sessionRevisionPlaceholders: [], createdListSites: [], deferred: [table] };
    const entries = projectionAllowlist.filter(({ family, kind }) => family === table.family && kind === table.kind);
    const used: AllowanceUsage = new Map();
    expect(check(session, document, native, stories, undefined, used)).toEqual([]);
    expect(checkStory(stories, { ...native, storyId: 'body', blocks: [] }, stories, used)).toEqual([]);
    expect([...used].map(([{ family, kind, path }, count]) => ({ family, kind, path, count })))
      .toEqual([{ family: 'tables', kind: 'table', path: 'deferred', count: 2 }]);
    expect(entries.filter((entry) => !used.has(entry))).toEqual(entries.filter(({ path }) => path !== 'deferred'));
  });

  it('checks the source ordinal of an edited nested field paragraph', () => {
    const source = fieldWithParagraph({ type: 'paragraph', sourceOrdinal: 3, content: [] });
    const edited = fieldWithParagraph({
      type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Changed' }] }],
    });
    expect(checkSourceOrdinals('body', [edited], [source]).join('\n'))
      .toContain('structuredResult.blocks[0].sourceOrdinal');
  });

  it('keeps nested position when a field paragraph is edited to match its neighbour', () => {
    const first: Paragraph = { type: 'paragraph', sourceOrdinal: 3, content: [] };
    const second: Paragraph = {
      type: 'paragraph', sourceOrdinal: 4,
      content: [{ type: 'run', content: [{ type: 'text', text: 'Same' }] }],
    };
    const source = fieldWithParagraph(first, second);
    const edited = { ...first, content: second.content };
    expect(checkSourceOrdinals('body', [fieldWithParagraph(edited, second)], [source])).toEqual([]);
    expect(checkSourceOrdinals('body', [fieldWithParagraph(second, second)], [source]).join('\n'))
      .toContain('structuredResult.blocks[0].sourceOrdinal: TS 3; native 4');
  });

  it('checks nested ordinals after a table or block control shifts past a new paragraph', () => {
    const first: Paragraph = { type: 'paragraph', sourceOrdinal: 3, content: [] };
    const second: Paragraph = { ...first, sourceOrdinal: 4 };
    const inserted: Paragraph = { type: 'paragraph', content: [] };
    const table = (content: Paragraph[]): BlockContent => ({
      type: 'table', rows: [{ type: 'tableRow', cells: [{ type: 'tableCell', content }] }],
    });
    const control = (content: Paragraph[]): BlockContent => ({
      type: 'blockSdt', properties: { sdtType: 'richText' }, content,
    });
    for (const [container, path] of [[table, 'rows[0].cells[0].content'], [control, 'content']] as const) {
      const source = [container([first, second])];
      expect(checkSourceOrdinals('body', [inserted, ...source], source, ['new'])).toEqual([]);
      const wrong = [inserted, container([second, second])];
      expect(checkSourceOrdinals('body', wrong, source, ['new']).join('\n'))
        .toContain(`blocks[1].${path}[0].sourceOrdinal: TS 3; native 4`);
    }
  });

  it('matches unaligned duplicate nested paragraphs in source order', () => {
    const first: Paragraph = { type: 'paragraph', sourceOrdinal: 3, content: [] };
    const second: Paragraph = { ...first, sourceOrdinal: 4 };
    const source = fieldWithParagraph(first, second);
    const moved = (paragraph: Paragraph, ...rest: Paragraph[]): Paragraph => ({
      ...source,
      content: [{ type: 'run', content: [] }, ...fieldWithParagraph(paragraph, ...rest).content],
    });
    expect(checkSourceOrdinals('body', [moved(first, second)], [source])).toEqual([]);
    expect(checkSourceOrdinals('body', [moved(second)], [source]).join('\n'))
      .toContain('content[1].structuredResult.blocks[0].sourceOrdinal: TS 3; native 4');
  });

  it('checks the surviving field paragraph ordinal after a nested deletion', () => {
    const first = textParagraph('A', 3);
    const second = textParagraph('B', 4);
    const source = [fieldWithParagraph(first, second)];
    const native = [fieldWithParagraph({ ...second, sourceOrdinal: 3 })];
    expect(checkSourceOrdinals('body', native, source).join('\n'))
      .toContain('structuredResult.blocks[0].sourceOrdinal: TS 4; native 3');
  });

  it('checks the source ordinals of swapped nested field paragraphs', () => {
    const first = textParagraph('A', 3);
    const second = textParagraph('B', 4);
    const source = [fieldWithParagraph(first, second)];
    const native = [fieldWithParagraph({ ...second, sourceOrdinal: 3 }, { ...first, sourceOrdinal: 4 })];
    const differences = checkSourceOrdinals('body', native, source).join('\n');
    expect(differences).toContain('structuredResult.blocks[0].sourceOrdinal: TS 4; native 3');
    expect(differences).toContain('structuredResult.blocks[1].sourceOrdinal: TS 3; native 4');
  });

  it('checks nested ordinals when a new table shifts an existing table', () => {
    const old = textParagraph('Old', 3);
    const source = [tableWithParagraph(old)];
    const native = [tableWithParagraph(textParagraph('New', 3)), tableWithParagraph({ ...old, sourceOrdinal: undefined })];
    const differences = checkSourceOrdinals('body', native, source).join('\n');
    expect(differences).toContain('blocks[0].rows[0].cells[0].content[0].sourceOrdinal: TS <missing>; native 3');
    expect(differences).toContain('blocks[1].rows[0].cells[0].content[0].sourceOrdinal: TS 3; native <missing>');
  });

  it('rejects a reused source ordinal on a replacement table paragraph with a new ID', () => {
    const source = [tableWithParagraph({ ...textParagraph('Old', 3), paraId: '11111111' })];
    const native = [tableWithParagraph({ ...textParagraph('New', 3), paraId: '22222222' })];
    expect(checkSourceOrdinals('body', native, source).join('\n'))
      .toContain('blocks[0].rows[0].cells[0].content[0].sourceOrdinal: TS <missing>; native 3');
  });

  it('rejects paragraph keys on a deferred auxiliary story', () => {
    const storyId = 'hf:header';
    const blocks: BlockContent[] = [];
    const stories = new Map([[storyId, { content: blocks, blocks, keys: [] }]]);
    const deferred: Deferral[] = [{ family: 'storiesAux', kind: 'headerFooterStory', storyId, paragraph: null }];
    const native = {
      storyId, blocks, paragraphSessionKeys: [], sessionRevisionPlaceholders: [], createdListSites: [],
      deferred,
    };
    expect(checkStory(stories, native, stories)).toEqual([]);
    expect(checkStory(stories, { ...native, paragraphSessionKeys: ['unexpected'] }, stories).join('\n'))
      .toContain('paragraphSessionKeys[0]');
  });

  it('detects a nested source ordinal change between snapshots', () => {
    const snapshot = (sourceOrdinal?: number) => wireDocument([
      fieldWithParagraph({ type: 'paragraph', content: [], sourceOrdinal }),
    ]) as unknown as Document;
    expect(compareSnapshots({} as YrsSession, snapshot(3), snapshot(), new Set(['body'])).join('\n'))
      .toContain('structuredResult.blocks[0].sourceOrdinal');
  });

  it('checks paragraph source ordinals against the source pilcrow key', () => {
    const paragraph: Paragraph = { type: 'paragraph', content: [] };
    const source = { ...paragraph, sourceOrdinal: 3 };
    const edited = { ...paragraph, content: [{ type: 'run', content: [] }] } satisfies Paragraph;
    expect(canonicalProjection(source)).toEqual(canonicalProjection(paragraph));
    expect(canonicalProjection(source)).not.toEqual(canonicalProjection(edited));
    expect(checkSourceOrdinals('body', [edited, source], [source], ['new', 'body:p0'])).toEqual([]);
    expect(checkSourceOrdinals('body', [{ ...source, sourceOrdinal: 4 }], [source], ['body:p0']))
      .toHaveLength(1);
    expect(canonicalProjection({ metadata: { sourceOrdinal: 3 } })).not
      .toEqual(canonicalProjection({ metadata: { sourceOrdinal: 4 } }));
  });

  it('equates a redundant whitespace hint while keeping a significant one', () => {
    const text = { type: 'text', text: ' some text' };
    const plain = { type: 'text', text: 'some text' };
    expect(canonicalProjection({ ...text, preserveSpace: true })).toEqual(canonicalProjection(text));
    expect(canonicalProjection({ ...plain, preserveSpace: true })).not.toEqual(canonicalProjection(plain));
  });

  it('equates empty run formatting with absence and preserves explicit formatting', () => {
    const run = { type: 'run', content: [] };
    expect(canonicalProjection({ ...run, formatting: {} })).toEqual(canonicalProjection(run));
    expect(canonicalProjection({ ...run, formatting: { bold: false } })).not.toEqual(canonicalProjection(run));
  });

  it('equates an empty table cascade with absence and preserves a selected style', () => {
    const table = { type: 'table', rows: [] };
    expect(canonicalProjection({ ...table, formatting: { styleCascade: {} } }))
      .toEqual(canonicalProjection(table));
    expect(canonicalProjection({ ...table, formatting: { styleCascade: { selectedStyleId: 'TableGrid' } } }))
      .not.toEqual(canonicalProjection(table));
  });

  it('equates omitted field display mode with result and preserves code mode', () => {
    const field = { type: 'simpleField', fieldTree: { version: 1 } };
    expect(canonicalProjection({ ...field, fieldTree: { ...field.fieldTree, displayMode: 'result' } }))
      .toEqual(canonicalProjection(field));
    expect(canonicalProjection({ ...field, fieldTree: { ...field.fieldTree, displayMode: 'code' } }))
      .not.toEqual(canonicalProjection(field));
  });

  it('equates absent and empty separator collections and preserves separator notes', () => {
    const pkg = { document: { content: [] } };
    for (const key of ['footnoteSeparators', 'endnoteSeparators']) {
      expect(canonicalProjection({ package: { ...pkg, [key]: [] } }))
        .toEqual(canonicalProjection({ package: pkg }));
      expect(canonicalProjection({ package: { ...pkg, [key]: [{ id: -1, content: [] }] } }))
        .not.toEqual(canonicalProjection({ package: pkg }));
    }
  });
});
