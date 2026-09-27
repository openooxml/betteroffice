import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import {
  createYrsSession,
  saveYrsDocx,
  type DocxContentControl,
  type DocxContentControlsResult,
  type DocxContentControlsSnapshot,
  type DocxEditResult,
  type DocxParagraphIdentity,
  type DocxSourceStory,
  type YrsSession,
} from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const TEMPLATE = resolve(import.meta.dir, '__fixtures__/content-controls/template.docx');
const BODY: DocxSourceStory = { partUri: '/word/document.xml', kind: 'body' };
const ADDRESS = 'body:sdt0';

const SAVED_BODY = `import sys,io,zipfile,json,xml.etree.ElementTree as E
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W14='{http://schemas.microsoft.com/office/word/2010/wordml}paraId'
root=E.fromstring(zipfile.ZipFile(io.BytesIO(sys.stdin.buffer.read())).read('word/document.xml'))
def text(node):
  return ''.join((el.text or '') if el.tag==W+'t' else '\\n' for el in node.iter() if el.tag in (W+'t',W+'br'))
def val(pr,name):
  el=pr.find(W+name)
  return None if el is None else el.get(W+'val')
controls=[]
for sdt in root.iter(W+'sdt'):
  pr=sdt.find(W+'sdtPr')
  content=sdt.find(W+'sdtContent')
  controls.append({'tag':val(pr,'tag'),'id':val(pr,'id'),'lock':val(pr,'lock'),
    'placeholder':pr.find(W+'showingPlcHdr') is not None,
    'paragraphs':[[p.get(W14),text(p)] for p in content.findall(W+'p')],'text':text(content)})
print(json.dumps({'paragraphs':[[p.get(W14),text(p)] for p in root.iter(W+'p')],'controls':controls}))`;

interface SavedControl {
  tag: string | null;
  id: string | null;
  lock: string | null;
  placeholder: boolean;
  paragraphs: Array<[string | null, string]>;
  text: string;
}

interface SavedBody {
  paragraphs: Array<[string | null, string]>;
  controls: SavedControl[];
}

/** The main part's paragraphs and controls, read by an independent namespace-aware parser. */
function savedBody(bytes: Uint8Array): SavedBody {
  const result = spawnSync('python3', ['-c', SAVED_BODY], {
    input: Buffer.from(bytes),
    encoding: 'utf8',
  });
  expect(result.stderr).toBe('');
  expect(result.status).toBe(0);
  return JSON.parse(result.stdout) as SavedBody;
}

function control(body: SavedBody, tag: string): SavedControl {
  const found = body.controls.find((candidate) => candidate.tag === tag);
  if (!found) throw new Error(`no saved control tagged ${tag}`);
  return found;
}

const sessions: YrsSession[] = [];

async function open(bytes: Uint8Array, clientId: number): Promise<YrsSession> {
  const session = await createYrsSession({ clientId });
  sessions.push(session);
  session.openDocx(bytes, true);
  return session;
}

function snapshot(result: DocxContentControlsResult): DocxContentControlsSnapshot {
  if (!result.ok) throw new Error(`${result.failure.code}: ${result.failure.message}`);
  return result.content;
}

function byTag(content: DocxContentControlsSnapshot, tag: string): DocxContentControl {
  const found = content.controls.find((candidate) => candidate.tag === tag);
  if (!found) throw new Error(`no control tagged ${tag}`);
  return found;
}

function applied(result: DocxEditResult): Extract<DocxEditResult, { ok: true }> {
  if (!result.ok) throw new Error(`${result.failure.code}: ${result.failure.message}`);
  return result;
}

function identity(session: YrsSession, story: string, key: string): DocxParagraphIdentity {
  const found = session
    .paragraphIdentities()
    .paragraphs.find((entry) => entry.session?.story === story && entry.session.paraId === key);
  if (!found) throw new Error(`no paragraph ${key} in ${story}`);
  return found;
}

/** The Word paragraph ID a paragraph authored in the session is claimed. */
function authoredId(session: YrsSession, story: string, key: string): string {
  const authored = identity(session, story, key);
  expect(authored).toMatchObject({ origin: 'authored', idOrigin: 'authored' });
  expect(authored.ooxmlParaId).toMatch(/^[0-9A-F]{8}$/);
  return authored.ooxmlParaId!;
}

function listing(session: YrsSession): Array<[string, string | null, string | null]> {
  return snapshot(session.listContentControls()).controls.map((entry) => [
    entry.controlId,
    entry.tag,
    entry.ooxmlId,
  ]);
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

afterEach(() => {
  for (const created of sessions.splice(0)) created.destroy();
});

describe('content controls and paragraph identities', () => {
  it('keeps identities and controls through a batch with fills, a save, a reopen and an ooxmlId fill', async () => {
    const template = new Uint8Array(readFileSync(TEMPLATE));
    const source = savedBody(template);
    const opened = await open(template, 83001);
    const name = byTag(snapshot(opened.listContentControls()), 'customer.name');
    expect(name.controlId).toBe('body|10000002|0');
    const result = applied(
      opened.applyEdits({
        expectVersion: opened.version(),
        steps: [
          {
            op: 'insertParagraphs',
            target: { story: 'body', paraId: '10000008' },
            at: 'end',
            paragraphs: [{ text: 'Witnessed.' }],
          },
          { op: 'setContentControlText', target: { kind: 'id', controlId: name.controlId }, text: 'Ada Lovelace' },
          {
            op: 'setContentControlText',
            target: { kind: 'tag', tag: 'customer.address' },
            text: '12 Example Street\nLondon\nUK',
          },
        ],
      })
    );
    const [witnessed] = result.receipts[0]!.newParagraphs;
    const [london, uk] = result.receipts[2]!.newParagraphs;
    expect([witnessed!.story, london!.story, uk!.story]).toEqual(['body', ADDRESS, ADDRESS]);
    const ids = {
      witnessed: authoredId(opened, 'body', witnessed!.paraId),
      london: authoredId(opened, ADDRESS, london!.paraId),
      uk: authoredId(opened, ADDRESS, uk!.paraId),
    };
    expect(new Set([...Object.values(ids), ...source.paragraphs.map(([id]) => id)]).size).toBe(
      source.paragraphs.length + 3
    );
    expect(identity(opened, ADDRESS, '10000006')).toMatchObject({
      origin: 'source',
      ooxmlParaId: '10000006',
      idOrigin: 'source',
    });
    const listed = listing(opened);

    const saved = await saveYrsDocx(opened);
    expect(saved.conflicts).toEqual([]);
    const persisted = (key: string) =>
      saved.paragraphs.find((entry) => entry.session.paraId === key)?.persisted;
    for (const [key, id] of [
      [witnessed!.paraId, ids.witnessed],
      [london!.paraId, ids.london],
      [uk!.paraId, ids.uk],
      ['10000006', '10000006'],
      ['10000002', '10000002'],
    ] as const) {
      expect(persisted(key)).toEqual({ kind: 'persisted', story: BODY, paraId: id });
    }
    const body = savedBody(saved.bytes);
    expect(body.paragraphs.map(([id]) => id)).toEqual([
      '10000001',
      '10000002',
      '10000003',
      '10000004',
      '10000005',
      '10000006',
      ids.london,
      ids.uk,
      '10000007',
      '10000008',
      ids.witnessed,
    ]);
    expect(body.paragraphs.at(-1)).toEqual([ids.witnessed, 'Witnessed.']);
    const properties = ({ tag, id, lock }: SavedControl) => ({ tag, id, lock });
    expect(body.controls.map(properties)).toEqual(source.controls.map(properties));
    expect(control(body, 'customer.name')).toMatchObject({ placeholder: false, text: 'Ada Lovelace' });
    expect(control(body, 'customer.address').paragraphs).toEqual([
      ['10000006', '12 Example Street'],
      [ids.london, 'London'],
      [ids.uk, 'UK'],
    ]);

    const reopened = await open(saved.bytes, 83002);
    for (const { session, persisted: anchor } of saved.paragraphs) {
      const found = reopened.resolveParagraphAnchor(anchor);
      if (found.status !== 'found' || found.anchor.kind !== 'session') {
        throw new Error(`unresolved: ${JSON.stringify(found)}`);
      }
      const text = (target: YrsSession, story: string, key: string) =>
        target.paragraphs(story).find((entry) => entry.paraId === key)?.text;
      expect(text(reopened, found.anchor.story, found.anchor.paraId)).toBe(
        text(opened, session.story, session.paraId)
      );
    }
    expect(reopened.paragraphs(ADDRESS).map((entry) => entry.paraId)).toEqual([
      '10000006',
      ids.london,
      ids.uk,
    ]);
    expect(listing(reopened)).toEqual(listed);
    const again = snapshot(reopened.listContentControls());
    expect(byTag(again, 'customer.address').value).toEqual({
      kind: 'text',
      text: '12 Example Street\nLondon\nUK',
    });

    applied(
      reopened.applyEdits({
        expectVersion: reopened.version(),
        steps: [
          {
            op: 'setContentControlText',
            target: { kind: 'ooxmlId', ooxmlId: byTag(again, 'customer.address').ooxmlId! },
            text: 'Line A\nLine B',
          },
          {
            op: 'setContentControlText',
            target: { kind: 'ooxmlId', ooxmlId: byTag(again, 'customer.name').ooxmlId! },
            text: 'Grace Hopper',
          },
        ],
      })
    );
    expect(reopened.paragraphs(ADDRESS).map((entry) => [entry.paraId, entry.text])).toEqual([
      ['10000006', 'Line A'],
      [ids.london, 'Line B'],
    ]);
    const final = savedBody((await saveYrsDocx(reopened)).bytes);
    expect(final.controls.map(properties)).toEqual(source.controls.map(properties));
    expect(control(final, 'customer.name').text).toBe('Grace Hopper');
    expect(control(final, 'customer.address').paragraphs).toEqual([
      ['10000006', 'Line A'],
      [ids.london, 'Line B'],
    ]);
    expect(final.paragraphs.map(([id]) => id)).not.toContain(ids.uk);
    expect(final.paragraphs.at(-1)).toEqual([ids.witnessed, 'Witnessed.']);
  });
});
