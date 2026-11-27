import { beforeAll, describe, expect, it } from 'bun:test';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { readDocxContainer } from '../docx/zipContainer';
import type { Document } from '../types/document';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';
import { documentToYrs } from './documentToYrs';
import { buildFixtureDocx, preloadSeedParityWasm, seedParityFixtures } from './projectionParity.support';
import { yrsToDocument } from './yrsToDocument';

/** Required parsed node kinds by fixture. */
const REQUIRED_NODES: Record<string, string[]> = {
  'bookmarks-fields': ['bookmarkStart', 'bookmarkEnd', 'simpleField', 'complexField', 'instrText'],
  comments: ['commentRangeStart', 'commentRangeEnd', 'commentReference'],
  'content-controls': ['blockSdt', 'inlineSdt'],
  'field-code-marks': ['complexField', 'instrText', 'table'],
  hyperlinks: ['hyperlink', 'bookmarkStart'],
  images: ['drawing', 'image', 'inline'],
  lists: ['paragraph'],
  math: ['mathEquation'],
  notes: ['footnoteRef', 'endnoteRef'],
  'opaque-drawings': ['opaqueDrawing', 'image'],
  'page-break-beside-field': ['complexField', 'break'],
  'rtl-cjk': ['tab', 'break', 'symbol', 'softHyphen', 'noBreakHyphen'],
  'sections-headers': ['paragraph'],
  shapes: ['shape', 'table'],
  tables: ['table', 'tableRow', 'tableCell'],
  'tracked-changes': ['insertion', 'deletion', 'paragraphPropertyChange', 'runPropertyChange'],
  'tracked-moves': ['moveFrom', 'moveTo'],
};

/** Package parts each fixture must carry once parsed. */
const REQUIRED_PARTS: Record<string, Array<(document: Document) => boolean>> = {
  comments: [(doc) => (doc.package.document.comments?.length ?? 0) === 4],
  images: [(doc) => (doc.package.media?.size ?? 0) > 0],
  lists: [(doc) => (doc.package.numbering?.abstractNums?.length ?? 0) === 2],
  notes: [
    (doc) => (doc.package.footnotes?.length ?? 0) === 2,
    (doc) => (doc.package.endnotes?.length ?? 0) === 1,
  ],
  'sections-headers': [
    (doc) => (doc.package.headers?.size ?? 0) === 3,
    (doc) => (doc.package.footers?.size ?? 0) === 2,
    (doc) => (doc.package.document.sections?.length ?? 0) === 2,
  ],
};

function nodeKinds(node: unknown, kinds = new Set<string>(), depth = 0): Set<string> {
  if (depth > 40 || node === null || typeof node !== 'object') return kinds;
  if (Array.isArray(node)) {
    for (const item of node) nodeKinds(item, kinds, depth + 1);
    return kinds;
  }
  if (node instanceof Map) {
    for (const value of node.values()) nodeKinds(value, kinds, depth + 1);
    return kinds;
  }
  const record = node as Record<string, unknown>;
  if (typeof record.type === 'string') kinds.add(record.type);
  for (const value of Object.values(record)) nodeKinds(value, kinds, depth + 1);
  return kinds;
}

function expectEquivalentStories(left: YrsSession, right: YrsSession): void {
  expect(left.storyIds()).toEqual(right.storyIds());
  for (const storyId of left.storyIds()) {
    expect(left.storySegments(storyId)).toEqual(right.storySegments(storyId));
  }
}

const fixtures = seedParityFixtures();

// Yrs formatting marker counts make seeded state vectors nondeterministic.
describe('DOCX seeding across document features', () => {
  beforeAll(preloadSeedParityWasm, 30_000);

  it('exports reanchored comments with metadata intact through three undo/redo cycles', async () => {
    const bytes = buildFixtureDocx('comments');
    const session = await createYrsSession({ clientId: 80005 });
    try {
      const { document } = session.seedFromDocx(bytes);
      const originalComments = document.package.document.comments;
      const plain = await createYrsSession({ clientId: 80009 });
      const exportedComments = plain.seedFromDocx(
        new Uint8Array(await repackDocx(yrsToDocument(session, document)))
      ).document.package.document.comments;
      plain.destroy();
      const paragraph = session.paragraphs('body')[0];
      const before = session.resolveComment('1');
      session.beginUndoCapture();
      session.replaceRange({
        story: 'body',
        start: { paraId: paragraph.paraId, offset: 0 },
        end: { paraId: paragraph.paraId, offset: paragraph.text.length },
      }, 'Achado preservado. Conclusão nova.');
      session.setCommentRanges('1', [{
        story: 'body',
        start: { paraId: paragraph.paraId, offset: 0 },
        end: { paraId: paragraph.paraId, offset: 17 },
      }]);
      const after = [{ story: 'body', start: 0, end: 17 }];
      const verifyExport = async () => {
        expect(session.resolveComment('1')).toEqual(after);
        const saved = yrsToDocument(session, document);
        expect(saved.package.document.comments).toEqual(originalComments);
        const exported = await repackDocx(saved);
        const parts = readDocxContainer(exported);
        const xml = parts.text('word/document.xml') ?? '';
        const start = xml.indexOf('<w:commentRangeStart w:id="1"');
        const end = xml.indexOf('<w:commentRangeEnd w:id="1"');
        expect(start).toBeGreaterThan(-1);
        expect(end).toBeGreaterThan(start);
        const highlighted = [...xml.slice(start, end).matchAll(/<w:t(?: [^>]*)?>(.*?)<\/w:t>/g)]
          .map((match) => match[1]).join('');
        expect(highlighted).toBe('Achado preservado');
        const reopened = await createYrsSession({ clientId: 80006 });
        try {
          const reopenedDoc = reopened.seedFromDocx(new Uint8Array(exported)).document;
          expect(reopenedDoc.package.document.comments).toEqual(exportedComments);
          expect(reopened.resolveComment('1')).toEqual(after);
        } finally {
          reopened.destroy();
        }
      };
      await verifyExport();
      for (let cycle = 0; cycle < 3; cycle += 1) {
        expect(session.undo()).toBe(true);
        expect(session.paragraphs('body')[0].text).toBe(paragraph.text);
        expect(session.resolveComment('1')).toEqual(before);
        expect(session.redo()).toBe(true);
        await verifyExport();
      }
    } finally {
      session.destroy();
    }
  });

  it('saves reanchored comments with their identity and thread, and saves the undone ranges', async () => {
    const sessions: YrsSession[] = [];
    const open = async (bytes: Uint8Array) => {
      const session = await createYrsSession({ clientId: 80010 + sessions.length });
      sessions.push(session);
      session.openDocx(bytes, true);
      return session;
    };
    const ids = ['1', '2', '3', '4'];
    const ranges = (session: YrsSession) => ids.map((id) => session.resolveComment(id));
    const threads = (session: YrsSession) =>
      session.materializeDocx()!.package.document.comments!.map(
        ({ id, author, initials, date, parentId, done, content }) => ({
          id, author, initials, date, parentId, done, content: JSON.stringify(content),
        })
      );
    const exported = (session: YrsSession) => {
      const result = session.exportStructured({ revisionView: 'accepted', stories: ['comments'] });
      if (!result.ok) throw new Error(JSON.stringify(result));
      return result.content.stories.map((story) => story.comment!);
    };
    const metadata = (session: YrsSession) =>
      exported(session).map(({ anchors: _, ...comment }) => comment);
    try {
      const opened = await open(buildFixtureDocx('comments'));
      const [lead, overlap] = opened.paragraphs('body');
      const prefix = (paragraph: typeof lead, text: string) => {
        expect(paragraph.text.startsWith(text)).toBe(true);
        return {
          start: { paraId: paragraph.paraId, offset: 0 },
          end: { paraId: paragraph.paraId, offset: text.length },
        };
      };
      const identity = {
        threads: threads(opened),
        listed: opened.listComments(),
        metadata: metadata(opened),
      };
      expect(identity.metadata.find(({ id }) => id === '3')).toMatchObject({ resolved: true });
      expect(identity.metadata.find(({ id }) => id === '4')).toMatchObject({ parentId: '3' });
      const before = ranges(opened);

      opened.setUndoCaptureMode('manual');
      const root = prefix(lead, 'Plain lead in');
      const resolved = prefix(overlap, 'Overlapping');
      opened.setCommentRanges('1', [{ story: 'body', ...root }]);
      opened.setCommentRanges('3', [{ story: 'body', ...resolved }]);
      opened.addUndoBoundary();
      const after = ranges(opened);
      expect(after[0]).not.toEqual(before[0]);
      expect(after[2]).not.toEqual(before[2]);
      const anchorsOf = (session: YrsSession, id: string) =>
        exported(session).find((comment) => comment.id === id)!.anchors;
      const reanchored = (session: YrsSession) => {
        expect(anchorsOf(session, '1')).toEqual([
          { kind: 'range', story: 'body', ...root, view: 'accepted' },
        ]);
        expect(anchorsOf(session, '3')).toEqual([
          { kind: 'range', story: 'body', ...resolved, view: 'accepted' },
        ]);
      };
      reanchored(opened);

      const expectSaved = async (expected: typeof before) => {
        const reopened = await open((await saveYrsDocx(opened)).bytes);
        expect(ranges(reopened)).toEqual(expected);
        expect(threads(reopened)).toEqual(identity.threads);
        expect(reopened.listComments()).toEqual(identity.listed);
        expect(metadata(reopened)).toEqual(identity.metadata);
        return reopened;
      };
      for (let cycle = 0; cycle < 2; cycle += 1) {
        expect(metadata(opened)).toEqual(identity.metadata);
        reanchored(await expectSaved(after));
        expect(opened.undo()).toBe(true);
        expect(opened.historyStories()).toContain('body');
        expect(ranges(opened)).toEqual(before);
        await expectSaved(before);
        expect(opened.redo()).toBe(true);
        expect(ranges(opened)).toEqual(after);
      }
    } finally {
      for (const session of sessions) session.destroy();
    }
  });

  it('seeds raw scheme colours for picture outlines in both seeders', async () => {
    const bytes = buildFixtureDocx('images');
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    const projected = await createYrsSession({ clientId: 48002 });
    const engine = await createYrsSession({ clientId: 48002 });
    try {
      documentToYrs(projected, parsed);
      engine.seedFromDocx(bytes);
      expectEquivalentStories(engine, projected);
      for (const session of [engine, projected]) {
        const images = session.storySegments('body').flatMap((segment) =>
          segment.kind === 'embed' &&
          segment.embedKind === 'image' &&
          segment.payload.alt === 'Scheme outline'
            ? [segment.payload]
            : []
        );
        expect(images).toHaveLength(1);
        expect(images[0].borderColorValue).toEqual({
          themeColor: 'background1',
          transforms: [{ kind: 'lumMod', value: 0.75 }],
        });
        expect(images[0]).not.toHaveProperty('borderWidth');
      }
    } finally {
      projected.destroy();
      engine.destroy();
    }
  });

  for (const name of fixtures) {
    it(`preserves ${name} stories, comments, and save output`, async () => {
      const bytes = buildFixtureDocx(name);
      const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });

      const kinds = nodeKinds(parsed.package.document);
      for (const required of REQUIRED_NODES[name] ?? []) expect([...kinds]).toContain(required);
      for (const [index, check] of (REQUIRED_PARTS[name] ?? []).entries()) {
        expect(check(parsed), `${name} package check ${index}`).toBe(true);
      }

      const projected = await createYrsSession({ clientId: 48001 });
      const engine = await createYrsSession({ clientId: 48001 });
      try {
        documentToYrs(projected, parsed);
        engine.seedFromDocx(bytes);

        expectEquivalentStories(engine, projected);
        for (const comment of parsed.package.document.comments ?? []) {
          const id = String(comment.id);
          expect(engine.resolveComment(id)).toEqual(projected.resolveComment(id));
        }
        expect(yrsToDocument(engine, parsed)).toEqual(yrsToDocument(projected, parsed));
      } finally {
        projected.destroy();
        engine.destroy();
      }
    });
  }
});
