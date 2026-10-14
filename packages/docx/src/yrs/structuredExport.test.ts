import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import {
  DocxExportError,
  exportDocxMarkdown,
  exportDocxStructured,
  renderDocxMarkdown,
  type DocxExportBlock,
  type DocxExportOptions,
  type DocxStorySelection,
  type DocxStructuredContent,
} from '../core';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});
const FIXTURES = resolve(
  import.meta.dir,
  '../../../../crates/docx-edit/tests/fixtures/structured-export'
);
const ALL: DocxStorySelection[] = ['body', 'headers', 'footers', 'footnotes', 'endnotes', 'comments'];
const VIEWS = ['accepted', 'original', 'markup'] as const;

const principal = () => new Uint8Array(readFileSync(resolve(FIXTURES, 'principal.docx')));

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

/** A package whose body is `body`, with `comments` as its comments part when given. */
function docx(body: string, comments?: string): Uint8Array {
  const rels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const word = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const parts = new Map<string, Uint8Array>([
    [
      '[Content_Types].xml',
      toBytes(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${word}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${word}.comments+xml"/></Types>`
      ),
    ],
    [
      '_rels/.rels',
      toBytes(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/officeDocument" Target="word/document.xml"/></Relationships>`
      ),
    ],
    [
      'word/_rels/document.xml.rels',
      toBytes(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${comments === undefined ? '' : `<Relationship Id="rIdComments" Type="${rels}/comments" Target="comments.xml"/>`}</Relationships>`
      ),
    ],
    ['word/document.xml', toBytes(`<w:document ${W}><w:body>${body}</w:body></w:document>`)],
  ]);
  if (comments !== undefined) {
    parts.set('word/comments.xml', toBytes(`<w:comments ${W}>${comments}</w:comments>`));
  }
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}
const golden = (name: string) => readFileSync(resolve(FIXTURES, name), 'utf8');
const expected = (view: string) =>
  JSON.parse(golden(`principal.${view}.json`)) as DocxStructuredContent;

let nextClientId = 88100;

async function open(): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: nextClientId++ });
  session.openDocx(principal(), true);
  return session;
}

function paragraphText(block: DocxExportBlock | undefined): string {
  if (!block || !('paragraph' in block)) throw new Error('not a paragraph block');
  return block.paragraph.inlines
    .map((inline) => (inline.kind === 'text' ? inline.text : ''))
    .join('');
}

function blockFor(content: DocxStructuredContent, paraId: string): DocxExportBlock | undefined {
  return content.stories[0]!.blocks.find(
    (block) =>
      block.anchor.kind === 'paragraph' &&
      block.anchor.paraId === paraId &&
      block.kind !== 'sectionBreak'
  );
}

function replace(session: YrsSession, paraId: string, text: string) {
  const applied = session.applyEdits({
    expectVersion: session.version(),
    steps: [{ op: 'replaceText', target: { kind: 'paragraph', story: 'body', paraId }, text }],
  });
  if (!applied.ok) throw new Error(applied.failure.message);
}

describe('structured export', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('exports bytes headlessly exactly as the Rust engine golden files', async () => {
    for (const view of VIEWS) {
      const options: DocxExportOptions = { revisionView: view, stories: ALL };
      const content = await exportDocxStructured(principal(), options);
      expect(content).toEqual(expected(view));
      const markdown = await exportDocxMarkdown(principal(), options);
      expect(markdown.markdown).toBe(golden(`principal.${view}.md`));
      expect(await renderDocxMarkdown(content)).toEqual(markdown);
    }
  });

  it('refuses unusable options as typed errors and throws on malformed ones', async () => {
    const refused = await exportDocxStructured(principal(), {
      revisionView: 'accepted',
      maxBytes: 16,
    }).catch((error: unknown) => error);
    expect(refused).toBeInstanceOf(DocxExportError);
    expect((refused as DocxExportError).failure).toEqual({
      code: 'invalid-options',
      target: null,
      message: expect.any(String),
    });
    const tooMany = await exportDocxStructured(principal(), {
      revisionView: 'accepted',
      maxBlocks: 2_000_000,
    }).catch((error: unknown) => error);
    expect((tooMany as DocxExportError).failure.code).toBe('limit-exceeded');
    const malformed = await exportDocxStructured(principal(), {
      revisionView: 'accepted',
      pages: true,
    } as DocxExportOptions).catch((error: unknown) => error);
    expect(malformed).not.toBeInstanceOf(DocxExportError);
    expect(String(malformed)).toContain('pages');
    const unreadable = await exportDocxStructured(new Uint8Array([1, 2, 3]), {
      revisionView: 'accepted',
    }).catch((error: unknown) => error);
    expect(unreadable).not.toBeInstanceOf(DocxExportError);
    const empty = await createYrsSession({ clientId: nextClientId++ });
    try {
      expect(empty.exportStructured({ revisionView: 'accepted' })).toMatchObject({
        ok: false,
        failure: { code: 'unsupported', target: null },
      });
    } finally {
      empty.destroy();
    }
  });

  it('exports a live session with its version and leaves the session untouched', async () => {
    const session = await open();
    try {
      session.beginUndoCapture();
      replace(session, '00000017', 'First edit');
      replace(session, '00000013', 'Second edit');
      expect(session.undo()).toBe(true);
      session.setSelection({ story: 'body', paraId: '00000014', offset: 2 });
      session.registerFont(new Uint8Array(readFileSync(FONT)));
      session.layoutDocumentWithRegionsJson(LAYOUT);
      let updates = 0;
      const unsubscribe = session.onUpdate(() => {
        updates += 1;
      });
      const version = session.version();
      const state = session.encodeState();
      const vector = session.encodeStateVector();
      const selection = session.selection();
      const history = [session.canUndo(), session.canRedo(), session.historyStories()];
      const layout = session.residentWorkerProbe();
      expect(layout).not.toBeNull();
      const resident = session.residentWorkerSnapshot();
      for (const view of VIEWS) {
        const read = session.exportStructured({ revisionView: view, stories: ALL });
        if (!read.ok) throw new Error(read.failure.message);
        expect(read.version).toBe(version);
        expect(read.content.anchorScope).toBe('session');
        expect(session.exportMarkdown({ revisionView: view, stories: ALL }).ok).toBe(true);
        expect(session.exportStructured({ revisionView: view, maxBytes: 1_024 }).ok).toBe(true);
      }
      expect(session.version()).toBe(version);
      expect(session.encodeState()).toEqual(state);
      expect(session.encodeStateVector()).toEqual(vector);
      expect(session.selection()).toEqual(selection);
      expect([session.canUndo(), session.canRedo(), session.historyStories()]).toEqual(history);
      expect(session.residentWorkerProbe()).toEqual(layout);
      expect(session.residentWorkerSnapshot()).toEqual(resident);
      expect(updates).toBe(0);
      unsubscribe();
      expect(session.redo()).toBe(true);
      const after = session.exportStructured({ revisionView: 'accepted' });
      if (!after.ok) throw new Error(after.failure.message);
      expect(after.version).toBe(session.version());
      expect(paragraphText(blockFor(after.content, '00000013'))).toBe('Second edit');
      expect(session.exportStructured({ revisionView: 'accepted', maxBlocks: 0 })).toMatchObject({
        ok: false,
        version: session.version(),
        failure: { code: 'invalid-options', target: null },
      });
    } finally {
      session.destroy();
    }
  });

  it('exports an unedited session exactly as the golden files with session anchors', async () => {
    const session = await open();
    try {
      const read = session.exportStructured({ revisionView: 'markup', stories: ALL });
      if (!read.ok) throw new Error(read.failure.message);
      expect({ ...read.content, anchorScope: 'snapshot' }).toEqual(expected('markup'));
      const markdown = session.exportMarkdown({ revisionView: 'markup', stories: ALL });
      if (!markdown.ok) throw new Error(markdown.failure.message);
      expect(markdown.content.markdown).toBe(golden('principal.markup.md'));
    } finally {
      session.destroy();
    }
  });

  it('classifies session headings as the export does', async () => {
    const session = await open();
    try {
      const read = session.exportStructured({ revisionView: 'accepted' });
      if (!read.ok) throw new Error(read.failure.message);
      const exported = read.content.stories[0]!.blocks.flatMap((block) => {
        const heading =
          block.kind === 'heading' ? block.heading : block.kind === 'listItem' ? block.heading : null;
        return heading && block.anchor.kind === 'paragraph'
          ? [{ paraId: block.anchor.paraId, heading }]
          : [];
      });
      expect(exported.length).toBeGreaterThan(3);
      expect(session.headings('body')).toEqual(exported);
    } finally {
      session.destroy();
    }
  });

  it('opens and exports a simple field whose cached result holds a tab and a page break', async () => {
    const bytes = docx(
      '<w:p w14:paraId="0E900001"><w:fldSimple w:instr=" REF Summary "><w:r><w:t>First</w:t><w:tab/><w:br w:type="page"/><w:t>Second</w:t></w:r></w:fldSimple></w:p>'
    );
    const fieldOf = (content: DocxStructuredContent) =>
      content.stories[0]!.blocks.flatMap((block) =>
        'paragraph' in block ? block.paragraph.inlines.filter((inline) => inline.kind === 'field') : []
      );
    const exported = await exportDocxStructured(bytes, { revisionView: 'accepted' });
    expect(fieldOf(exported)).toHaveLength(1);
    const session = await createYrsSession({ clientId: nextClientId++ });
    try {
      session.openDocx(bytes, true);
      const read = session.exportStructured({ revisionView: 'accepted' });
      if (!read.ok) throw new Error(read.failure.message);
      expect(fieldOf(read.content)).toHaveLength(1);
    } finally {
      session.destroy();
    }
  });

  it('exports a comment whose cached field result has differently formatted runs', async () => {
    const bytes = docx(
      '<w:p w14:paraId="61000001"><w:commentRangeStart w:id="1"/><w:r><w:t>Annotated</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>',
      '<w:comment w:id="1" w:author="Ann"><w:p w14:paraId="61000002"><w:fldSimple w:instr=" REF Summary "><w:r><w:rPr><w:b/></w:rPr><w:t>First</w:t></w:r><w:r><w:t>Second</w:t></w:r></w:fldSimple></w:p></w:comment>'
    );
    const options: DocxExportOptions = { revisionView: 'accepted', stories: ['comments'] };
    const resultOf = (content: DocxStructuredContent) =>
      content.stories.flatMap((story) =>
        story.blocks.flatMap((block) =>
          'paragraph' in block
            ? block.paragraph.inlines.flatMap((inline) =>
                inline.kind === 'field' && inline.cachedResult.kind === 'inline'
                  ? inline.cachedResult.inlines.map((part) => (part.kind === 'text' ? part.text : ''))
                  : []
              )
            : []
        )
      );
    expect(resultOf(await exportDocxStructured(bytes, options))).toEqual(['First', 'Second']);
    const session = await createYrsSession({ clientId: nextClientId++ });
    try {
      session.openDocx(bytes, true);
      const read = session.exportStructured(options);
      if (!read.ok) throw new Error(read.failure.message);
      expect(resultOf(read.content)).toEqual(['First', 'Second']);
    } finally {
      session.destroy();
    }
  });

  it('anchors content without a location of its own as unlocated, with the reason', async () => {
    const bytes = docx(
      '<w:p w14:paraId="72000001"><w:commentRangeStart w:id="1"/><w:r><w:t>Annotated</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>' +
        `<w:p w14:paraId="72000002"><w:r><w:t>${'x'.repeat(2_000_000)}</w:t></w:r></w:p>`,
      '<w:comment w:id="1" w:author="Ann"><w:p w14:paraId="72000003"><w:r><w:t>Check</w:t></w:r></w:p></w:comment>'
    );
    const content = await exportDocxStructured(bytes, {
      revisionView: 'accepted',
      stories: ['comments'],
      maxBytes: 65_536,
    });
    expect(content.stories[0]!.comment!.anchors).toEqual([
      { kind: 'unlocated', story: 'body', reason: 'story-too-large' },
    ]);
  });

  it('truncates at whole blocks within the byte budget', async () => {
    const content = await exportDocxStructured(principal(), {
      revisionView: 'markup',
      stories: ALL,
      maxBytes: 4_096,
    });
    expect(content.truncated).toBe(true);
    expect(new TextEncoder().encode(JSON.stringify(content)).length).toBeLessThanOrEqual(4_096);
    expect(content.diagnostics.at(-1)?.code).toBe('truncated');
  });
});
