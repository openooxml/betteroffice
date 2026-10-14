import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { readDocxContainer } from '../docx/zipContainer';
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

const payloadCases = JSON.parse(readFileSync(resolve(
  import.meta.dir, '../../../../crates/docx-edit/tests/fixtures/horizontal_rule_payloads.json'
), 'utf8')) as Array<{ name: string; valid: boolean; payload: Record<string, unknown> }>;

for (const entry of payloadCases) {
  test(`raw horizontal rule payload: ${entry.name}`, async () => {
    const document = fixture([{ type: 'run', content: [{ type: 'text', text: 'AB' }] }]);
    const session = await createYrsSession({ clientId: 58436 });
    try {
      documentToYrs(session, document);
      session.applyRawOps('body', [{ op: 'insertEmbed', index: 1, kind: 'horizontalRule', payload: entry.payload }]);
      if (!entry.valid) {
        expect(() => yrsToDocument(session, document)).toThrow('Malformed horizontalRule embed payload');
        return;
      }
      const saved = yrsToDocument(session, document).package.document.content[0] as Paragraph;
      const value = entry.payload.rule as HorizontalRuleContent['rule'];
      expect(saved.content).toEqual([
        { type: 'run', content: [{ type: 'text', text: 'A' }] },
        { type: 'run', content: [{ type: 'horizontalRule', rule: { ...value, width: value.width ?? null, widthPercent: value.widthPercent ?? null } }] },
        { type: 'run', content: [{ type: 'text', text: 'B' }] },
      ]);
    } finally {
      session.destroy();
    }
  });
}

for (const attributes of [
  {},
  { hyperlink: { href: 'https://example.com/rule' } },
  { ins: { id: 'rule-insertion', author: 'Reviewer', date: '2026-09-14T00:00:00Z' } },
]) {
  test(`malformed rule recovery preserves surrounding text with ${Object.keys(attributes).join() || 'plain'} content`, async () => {
    const document = fixture([{ type: 'run', content: [{ type: 'text', text: 'AB' }] }]);
    const session = await createYrsSession({ clientId: 58437 });
    try {
      documentToYrs(session, document);
      session.applyRawOps('body', [{ op: 'insertEmbed', index: 1, kind: 'horizontalRule', payload: {}, attrs: attributes }]);
      expect(() => yrsToDocument(session, document)).toThrow('Malformed horizontalRule embed payload');
      session.applyRawOps('body', [{ op: 'setEmbedAttr', index: 1, key: 'rule', value: rule.rule }]);
      const saved = yrsToDocument(session, document).package.document.content[0] as Paragraph;
      expect(saved.content[0]).toEqual({ type: 'run', content: [{ type: 'text', text: 'A' }] });
      expect(saved.content.at(-1)).toEqual({ type: 'run', content: [{ type: 'text', text: 'B' }] });
      const middle = saved.content[1];
      if (middle.type === 'hyperlink') expect(middle.children).toEqual([{ type: 'run', content: [rule] }]);
      else if (middle.type === 'insertion') expect(middle.content).toEqual([{ type: 'run', content: [rule] }]);
      else expect(middle).toEqual({ type: 'run', content: [rule] });
    } finally {
      session.destroy();
    }
  });
}

test('nullable and omitted rule widths preserve authored XML on save', async () => {
  const pict = '<w:pict><v:rect style="height:1.5pt" o:hr="t" o:hrstd="t"/></w:pict>';
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body><w:p><w:r><w:t>Before</w:t></w:r><w:r>${pict}</w:r><w:r><w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
  const parts = new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', toBytes(xml)],
  ]);
  const bytes = rezipPartsToArrayBuffer(parts);
  const parsed = await parseDocx(bytes, { preloadFonts: false });
  let nullableXml: string | null = null;
  for (const omitWidths of [false, true]) {
    const session = await createYrsSession({ clientId: 58438 });
    try {
      documentToYrs(session, parsed);
      if (omitWidths) {
        const segment = session.storySegments('body').find((entry) =>
          entry.kind === 'embed' && entry.embedKind === 'horizontalRule'
        );
        if (!segment || segment.kind !== 'embed') throw new Error('expected rule embed');
        const value = { ...(segment.payload.rule as HorizontalRuleContent['rule']) };
        delete (value as Partial<typeof value>).width;
        delete (value as Partial<typeof value>).widthPercent;
        session.applyRawOps('body', [{ op: 'setEmbedAttr', index: 6, key: 'rule', value }]);
      }
      const saved = await repackDocx(yrsToDocument(session, parsed));
      const savedXml = readDocxContainer(saved).text('word/document.xml');
      expect(savedXml).toContain(pict);
      if (omitWidths) expect(savedXml).toBe(nullableXml);
      else nullableXml = savedXml;
      const paragraph = session.paragraphs('body')[0];
      session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Edited ');
      const edited = readDocxContainer(await repackDocx(yrsToDocument(session, parsed)));
      expect(edited.text('word/document.xml')).toContain('Edited Before');
      expect(edited.text('word/document.xml')).toContain(pict);
    } finally {
      session.destroy();
    }
  }
});
