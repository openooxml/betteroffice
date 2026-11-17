import { afterEach, expect, test } from 'bun:test';
import type { YrsSession as Session } from '@betteroffice/docx/yrs';
import { engines } from '#agents-engines';
import { attachDocx, type DocxAgentDocument } from '../src';
import { editFixture } from './fixture';

const owned: Array<{ destroy(): void }> = [];
const attached: DocxAgentDocument[] = [];
afterEach(() => {
  attached.splice(0).forEach((doc) => doc.close());
  owned.splice(0).forEach((session) => session.destroy());
});

function exchange(a: Session, b: Session): void {
  const fromA = a.encodeState();
  const fromB = b.encodeState();
  b.applyUpdate(fromA);
  a.applyUpdate(fromB);
}

/** A document whose table two collaborators' crossing undos left without its payload. */
async function emptiedTable(): Promise<Session> {
  const bytes = await editFixture();
  const { module } = await engines.load('docx');
  const a = await module.yrs.createYrsSession({ clientId: 61 });
  const b = await module.yrs.createYrsSession({ clientId: 62 });
  owned.push(a, b);
  a.seedFromDocx(bytes);
  b.openDocx(bytes, false);
  b.applyUpdate(a.encodeState());
  const at = { story: 'body', tableIndex: 0, row: 0, column: 0 };
  a.insertRow(at, 'below');
  exchange(a, b);
  b.setColumnWidth(at, 2000);
  expect(a.undo()).toBe(true);
  exchange(a, b);
  expect(b.undo()).toBe(true);
  exchange(a, b);
  const payload = a.tablePayload('body', 0) ?? {};
  expect(['tblPr', 'grid', 'rows'].filter((key) => key in payload)).toEqual([]);
  return a;
}

async function caught(call: () => Promise<unknown>): Promise<Error & { code?: string }> {
  try {
    await call();
  } catch (error) {
    return error as Error & { code?: string };
  }
  throw new Error('expected a refusal');
}

test('an agent edit that reads run formatting is refused with the table-unavailable reason and changes nothing', async () => {
  const session = await emptiedTable();
  const doc = attachDocx(session);
  attached.push(doc);
  const before = session.encodeState();
  const target = doc.list({ limit: 100 }).items.find((item) => item.text.endsWith('Closing note.'))!;
  expect(target).toBeDefined();
  const error = await caught(() => doc.edit({ version: doc.list().version, edits: [{
    op: 'insert_paragraphs', before: target.ref, runFormat: 'anchor', paragraphs: [{ text: 'Inserted' }],
  }] }));
  expect(error.code).toBe('TABLE_UNAVAILABLE');
  expect(error.message).toContain('a table is unavailable after a concurrent undo');
  expect(session.encodeState()).toEqual(before);
  expect((await caught(() => doc.export())).code).toBe('table-unavailable');
});

test('reading the document keeps working and the structured export marks the table', async () => {
  const session = await emptiedTable();
  const doc = attachDocx(session);
  attached.push(doc);
  const texts = doc.list({ limit: 100 }).items.map((item) => item.text);
  expect(texts).toContain('Quarterly report');
  expect(texts.some((text) => text.endsWith('Closing note.'))).toBe(true);
  const exported = session.exportStructured({ revisionView: 'accepted' });
  if (!exported.ok) throw new Error('expected an export');
  expect(exported.content.diagnostics).toContainEqual(expect.objectContaining({
    code: 'unsupported-content',
    severity: 'error',
    anchor: { kind: 'table', story: 'body', tableIndex: 0 },
  }));
});
