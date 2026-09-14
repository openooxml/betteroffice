import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Document, HorizontalRuleContent, Paragraph, ParagraphContent } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));

const rule: HorizontalRuleContent = {
  type: 'horizontalRule',
  rule: {
    width: null,
    widthPercent: null,
    height: 19050,
    alignment: 'left',
    noShade: false,
    color: '#A0A0A0',
    xml: '<w:pict xmlns:w="w" xmlns:v="v" xmlns:o="o"><v:rect style="width:0pt;height:1.5pt" o:hr="t" o:hrstd="t"/></w:pict>',
  },
};

function fixture(content: ParagraphContent[] = [{ type: 'run', content: [rule] }]): Document {
  return { package: { document: { content: [{ type: 'paragraph', paraId: '00000001', content }] } } };
}

const range = {
  story: 'body',
  start: { paraId: '00000001', offset: 0 },
  end: { paraId: '00000001', offset: 1 },
};

test('horizontal rules remain editable embeds and preserve VML on save', async () => {
  const document = fixture();
  const session = await createYrsSession({ clientId: 58431 });
  try {
    documentToYrs(session, document);
    const saved = yrsToDocument(session, document);
    const paragraph = saved.package.document.content[0] as Paragraph;
    expect(paragraph.content).toEqual([{ type: 'run', content: [rule] }]);
    expect(session.storySegments('body').some((segment) =>
      segment.kind === 'embed' && segment.embedKind === 'horizontalRule'
    )).toBe(true);
    session.deleteRange(range);
    const deleted = yrsToDocument(session, document).package.document.content[0] as Paragraph;
    expect(deleted.content.every((child) =>
      child.type !== 'run' || child.content.every((entry) => entry.type !== 'horizontalRule')
    )).toBe(true);
  } finally {
    session.destroy();
  }
});

test('tracked rule deletion retains VML and run formatting for rejection', async () => {
  const document = fixture([{ type: 'run', formatting: { fontSize: 36, fontSizeCs: 36 }, content: [rule] }]);
  const session = await createYrsSession({ clientId: 58432 });
  const reopened = await createYrsSession({ clientId: 58433 });
  try {
    documentToYrs(session, document);
    session.deleteRange(range, { name: 'Reviewer', date: '2026-09-14T00:00:00Z' });
    const saved = yrsToDocument(session, document);
    const paragraph = saved.package.document.content[0] as Paragraph;
    const deletion = paragraph.content[0];
    expect(deletion.type).toBe('deletion');
    if (deletion.type !== 'deletion') throw new Error('expected tracked deletion');
    expect(deletion.content).toEqual([{ type: 'run', formatting: { fontSize: 36, fontSizeCs: 36 }, content: [rule] }]);
    documentToYrs(reopened, saved);
    reopened.rejectChange(range);
    const restored = yrsToDocument(reopened, saved).package.document.content[0] as Paragraph;
    expect(restored.content).toEqual([{ type: 'run', formatting: { fontSize: 36, fontSizeCs: 36 }, content: [rule] }]);
  } finally {
    session.destroy();
    reopened.destroy();
  }
});

test('bookmarks after a rule preserve their one-unit document positions', async () => {
  const document = fixture([
    { type: 'run', content: [rule] },
    { type: 'bookmarkStart', id: 7, name: 'afterRule' },
    { type: 'run', content: [{ type: 'text', text: 'A' }] },
    { type: 'bookmarkEnd', id: 7 },
  ]);
  const session = await createYrsSession({ clientId: 58434 });
  try {
    documentToYrs(session, document);
    const saved = yrsToDocument(session, document).package.document.content[0] as Paragraph;
    expect(saved.content.map((child) => child.type)).toEqual(['run', 'bookmarkStart', 'run', 'bookmarkEnd']);
    expect(saved.content[1]).toMatchObject({ type: 'bookmarkStart', id: 7, name: 'afterRule' });
    expect(saved.content[3]).toMatchObject({ type: 'bookmarkEnd', id: 7 });
  } finally {
    session.destroy();
  }
});


test('linked horizontal rules survive save reconstruction', async () => {
  const document = fixture([{
    type: 'hyperlink',
    href: 'https://example.com/rule',
    children: [{ type: 'run', content: [rule] }],
  }]);
  const session = await createYrsSession({ clientId: 58435 });
  try {
    documentToYrs(session, document);
    const saved = yrsToDocument(session, document).package.document.content[0] as Paragraph;
    expect(saved.content[0]).toMatchObject({
      type: 'hyperlink',
      href: 'https://example.com/rule',
      children: [{ type: 'run', content: [rule] }],
    });
  } finally {
    session.destroy();
  }
});
