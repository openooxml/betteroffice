import { beforeAll, describe, expect, it } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import {
  DocxContentControlsError,
  findDocxContentControls,
  listDocxContentControls,
} from '../docx/contentControls';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import {
  createYrsSession,
  type DocxContentControl,
  type DocxContentControlsResult,
  type DocxContentControlsSnapshot,
  type DocxEditResult,
  type YrsSession,
} from './index';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const TEMPLATE = resolve(import.meta.dir, '__fixtures__/content-controls/template.docx');
const LEGACY_STATE = resolve(import.meta.dir, '__fixtures__/content-controls/legacy-value.bin');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

let nextClientId = 82100;

function template(): Uint8Array {
  return new Uint8Array(readFileSync(TEMPLATE));
}

async function open(bytes = template()): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: nextClientId++ });
  session.openDocx(bytes, true);
  return session;
}

function snapshot(result: DocxContentControlsResult): DocxContentControlsSnapshot {
  if (!result.ok) throw new Error(`${result.failure.code}: ${result.failure.message}`);
  return result.content;
}

function byTag(content: DocxContentControlsSnapshot, tag: string): DocxContentControl {
  const control = content.controls.find((candidate) => candidate.tag === tag);
  if (!control) throw new Error(`no control tagged ${tag}`);
  return control;
}

function applied(result: DocxEditResult): Extract<DocxEditResult, { ok: true }> {
  if (!result.ok) throw new Error(`${result.failure.code}: ${result.failure.message}`);
  return result;
}

function part(bytes: ArrayBuffer | Uint8Array, name: string): Uint8Array {
  const entry = unzipContainer(new Uint8Array(bytes))[name];
  if (!(entry instanceof Uint8Array)) throw new Error(`missing part ${name}`);
  return entry;
}

interface SdtRecord {
  tag: string | null;
  alias: string | null;
  id: string | null;
  lock: string | null;
  placeholder: boolean;
  multiLine: string | null;
  binding: string | null;
  appearance: string | null;
  endProperties: boolean;
  paragraphs: string[];
  text: string;
}

/** Every `w:sdt` of a part, read with Python's namespace-aware XML parser. */
function controlsInXml(xml: Uint8Array): SdtRecord[] {
  const result = spawnSync(
    'python3',
    [
      '-c',
      `import sys,json,xml.etree.ElementTree as E
W='{${W}}'
W15='{http://schemas.microsoft.com/office/word/2012/wordml}'
root=E.fromstring(sys.stdin.buffer.read())
def val(pr,name):
  el=pr.find(W+name) if pr is not None else None
  return None if el is None else el.get(W+'val')
out=[]
for sdt in root.iter(W+'sdt'):
  pr=sdt.find(W+'sdtPr')
  content=sdt.find(W+'sdtContent')
  text_el=pr.find(W+'text') if pr is not None else None
  binding=pr.find(W+'dataBinding') if pr is not None else None
  appearance=pr.find(W15+'appearance') if pr is not None else None
  def text(node):
    parts=[]
    for el in node.iter():
      if el.tag==W+'t': parts.append(el.text or '')
      elif el.tag==W+'tab': parts.append('\\t')
      elif el.tag==W+'br': parts.append('\\n')
    return ''.join(parts)
  out.append({'tag':val(pr,'tag'),'alias':val(pr,'alias'),'id':val(pr,'id'),'lock':val(pr,'lock'),
    'placeholder':pr is not None and pr.find(W+'showingPlcHdr') is not None,
    'multiLine':None if text_el is None else text_el.get(W+'multiLine'),
    'binding':None if binding is None else binding.get(W+'xpath'),
    'appearance':None if appearance is None else appearance.get(W15+'val'),
    'endProperties':sdt.find(W+'sdtEndPr') is not None,
    'paragraphs':[text(p) for p in content.findall(W+'p')],
    'text':text(content)})
print(json.dumps(out))`,
    ],
    { input: Buffer.from(xml), encoding: 'utf8' }
  );
  expect(result.stderr).toBe('');
  expect(result.status).toBe(0);
  return JSON.parse(result.stdout) as SdtRecord[];
}

describe('content controls', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('lists and finds controls from a session and from bytes', async () => {
    const session = await open();
    try {
      const listed = session.listContentControls();
      expect(listed).toMatchObject({ ok: true, version: session.version() });
      const content = snapshot(listed);
      expect(content.anchorScope).toBe('session');
      expect(content.complete).toBe(true);
      expect(content.controls.map((control) => [control.controlId, control.tag, control.placement])).toEqual([
        ['body|10000002|0', 'customer.name', 'inline'],
        ['body|10000003|0', 'account.reference', 'inline'],
        ['body|10000004|0', 'account.reference', 'inline'],
        ['body|10000005|0', 'terms.standard', 'inline'],
        ['body:sdt0', 'customer.address', 'block'],
        ['body|10000007|0', 'customer.email', 'inline'],
        ['hf:rIdHeader|20000001|0', 'document.title', 'inline'],
      ]);
      expect(byTag(content, 'customer.name')).toMatchObject({
        ooxmlId: '101',
        controlType: 'plainText',
        showingPlaceholder: true,
        multiLine: false,
        value: { kind: 'text', text: 'Click to enter a name.' },
        effectiveLock: { content: false, control: false, known: true },
      });
      expect(byTag(content, 'customer.address').value).toEqual({
        kind: 'text',
        text: '1 Old Road\nOldtown',
      });
      const duplicates = snapshot(session.findContentControls({ kind: 'tag', tag: 'account.reference' }));
      expect(duplicates.controls.map((control) => control.controlId)).toEqual([
        'body|10000003|0',
        'body|10000004|0',
      ]);
      expect(snapshot(session.findContentControls({ kind: 'alias', alias: 'address' })).controls).toEqual([]);
      expect(session.listContentControls({ maxControls: 2 })).toMatchObject({
        ok: false,
        failure: { code: 'limit-exceeded' },
      });

      const bytes = await listDocxContentControls(template());
      expect(bytes.anchorScope).toBe('snapshot');
      expect(bytes.controls).toEqual(content.controls);
      const found = await findDocxContentControls(template(), { kind: 'tag', tag: 'document.title' });
      expect(found.controls).toHaveLength(1);
      await expect(listDocxContentControls(template(), { maxControls: 0 })).rejects.toBeInstanceOf(
        DocxContentControlsError
      );
    } finally {
      session.destroy();
    }
  });

  it('refuses missing, duplicate, locked and bound targets without changing anything', async () => {
    const session = await open();
    try {
      const version = session.version();
      const refuse = (target: { kind: 'tag'; tag: string } | { kind: 'id'; controlId: string }, text = 'x') =>
        session.applyEdits({ expectVersion: version, steps: [{ op: 'setContentControlText', target, text }] });
      expect(refuse({ kind: 'tag', tag: 'missing' })).toMatchObject({
        ok: false,
        failure: { code: 'missing-target', reason: 'missing-tag' },
      });
      expect(refuse({ kind: 'tag', tag: 'account.reference' })).toMatchObject({
        ok: false,
        failure: { code: 'ambiguous-target', reason: 'ambiguous-tag' },
      });
      expect(refuse({ kind: 'tag', tag: 'terms.standard' })).toMatchObject({
        ok: false,
        failure: { code: 'locked-target', reason: 'content-locked' },
      });
      expect(refuse({ kind: 'tag', tag: 'customer.email' })).toMatchObject({
        ok: false,
        failure: { code: 'unsupported', reason: 'bound-control' },
      });
      expect(
        session.applyEdits({
          expectVersion: version,
          steps: [
            { op: 'setContentControlText', target: { kind: 'id', controlId: 'body|10000002|0' }, text: 'Ada' },
            { op: 'setContentControlText', target: { kind: 'tag', tag: 'document.title' }, text: 'a\nb' },
          ],
        })
      ).toMatchObject({
        ok: false,
        failure: { code: 'invalid-step', reason: 'multiline-not-allowed', stepIndex: 1 },
      });
      expect(session.version()).toBe(version);
      expect(byTag(snapshot(session.listContentControls()), 'customer.name').showingPlaceholder).toBe(true);
    } finally {
      session.destroy();
    }
  });

  it('fills two controls by id, saves through the save projection and reopens them', async () => {
    const bytes = template();
    const session = await open(bytes);
    const reopened = await createYrsSession({ clientId: nextClientId++ });
    try {
      const base = session.materializeDocx();
      if (!base) throw new Error('the opened package must materialize');
      yrsToDocument(session, base);
      const discovery = session.listContentControls();
      if (!discovery.ok) throw new Error(discovery.failure.message);
      const customer = byTag(discovery.content, 'customer.name');
      const address = byTag(discovery.content, 'customer.address');
      const result = applied(
        session.applyEdits({
          expectVersion: discovery.version,
          steps: [
            { op: 'setContentControlText', target: { kind: 'id', controlId: customer.controlId }, text: 'Ada Lovelace' },
            {
              op: 'setContentControlText',
              target: { kind: 'id', controlId: address.controlId },
              text: '12 Example Street\nLondon',
              expect: { text: '1 Old Road\nOldtown' },
            },
          ],
        })
      );
      expect(result.changedStories).toEqual(['body', 'body:sdt0']);
      expect(result.receipts.map((receipt) => receipt.control?.controlId)).toEqual([
        customer.controlId,
        address.controlId,
      ]);
      expect(session.undo()).toBe(true);
      expect(byTag(snapshot(session.listContentControls()), 'customer.name').showingPlaceholder).toBe(true);
      expect(session.redo()).toBe(true);

      const saved = await repackDocx(yrsToDocument(session, base));
      for (const name of [
        'word/styles.xml',
        'word/settings.xml',
        'customXml/item1.xml',
        'customXml/itemProps1.xml',
        'customXml/_rels/item1.xml.rels',
        'docProps/app.xml',
        'word/_rels/document.xml.rels',
      ]) {
        expect(part(saved, name)).toEqual(part(bytes, name));
      }
      expect(controlsInXml(part(saved, 'word/header1.xml'))).toEqual(
        controlsInXml(part(bytes, 'word/header1.xml'))
      );
      const before = controlsInXml(part(bytes, 'word/document.xml'));
      const after = controlsInXml(part(saved, 'word/document.xml'));
      expect(after.map(({ tag, alias, id, lock, binding, multiLine, endProperties }) => ({
        tag, alias, id, lock, binding, multiLine, endProperties,
      }))).toEqual(before.map(({ tag, alias, id, lock, binding, multiLine, endProperties }) => ({
        tag, alias, id, lock, binding, multiLine, endProperties,
      })));
      const name = after.find((control) => control.tag === 'customer.name')!;
      expect(name).toMatchObject({ placeholder: false, appearance: 'tags', text: 'Ada Lovelace' });
      expect(after.find((control) => control.tag === 'customer.address')).toMatchObject({
        placeholder: false,
        paragraphs: ['12 Example Street', 'London'],
        multiLine: '1',
      });
      for (const tag of ['account.reference', 'terms.standard', 'customer.email']) {
        expect(after.find((control) => control.tag === tag)).toEqual(
          before.find((control) => control.tag === tag)
        );
      }

      reopened.openDocx(new Uint8Array(saved), true);
      const again = snapshot(reopened.listContentControls());
      expect(again.controls.map((control) => [control.tag, control.alias, control.ooxmlId, control.lock])).toEqual(
        snapshot(session.listContentControls()).controls.map((control) => [
          control.tag,
          control.alias,
          control.ooxmlId,
          control.lock,
        ])
      );
      expect(byTag(again, 'customer.name')).toMatchObject({
        showingPlaceholder: false,
        value: { kind: 'text', text: 'Ada Lovelace' },
      });
      expect(byTag(again, 'customer.address').value).toEqual({
        kind: 'text',
        text: '12 Example Street\nLondon',
      });
    } finally {
      reopened.destroy();
      session.destroy();
    }
  });

  it('loads a legacy text value as it is and saves the control content, as before', async () => {
    const bytes = template();
    const session = await createYrsSession({ clientId: nextClientId++ });
    try {
      session.openDocx(bytes, false);
      session.loadState(new Uint8Array(readFileSync(LEGACY_STATE)));
      expect(session.canUndo()).toBe(false);
      const base = session.materializeDocx();
      if (!base) throw new Error('the opened package must materialize');
      const content = snapshot(session.listContentControls());
      expect(content.controls.find((control) => control.ooxmlId === '102')?.value).toEqual({
        kind: 'text',
        text: 'REF-000',
      });
      expect(content.diagnostics.map((diagnostic) => diagnostic.code)).toContain('legacy-control-value');
      const saved = await repackDocx(yrsToDocument(session, base));
      const xml = part(saved, 'word/document.xml');
      expect(new TextDecoder().decode(xml)).not.toContain('REF-LEGACY');
      expect(controlsInXml(xml).map((control) => control.text)).toEqual(
        controlsInXml(part(bytes, 'word/document.xml')).map((control) => control.text)
      );
    } finally {
      session.destroy();
    }
  });

  it('keeps two peers of a legacy state in step through delete, undo and redo', async () => {
    const bytes = template();
    const state = new Uint8Array(readFileSync(LEGACY_STATE));
    const left = await createYrsSession({ clientId: nextClientId++ });
    const right = await createYrsSession({ clientId: nextClientId++ });
    try {
      const bases = [left, right].map((peer) => {
        peer.openDocx(bytes, false);
        peer.loadState(state);
        const base = peer.materializeDocx();
        if (!base) throw new Error('the opened package must materialize');
        return base;
      });
      const sync = (from: YrsSession, to: YrsSession) =>
        to.applyUpdate(from.encodeStateAsUpdate(to.encodeStateVector()));
      const agree = async (text: string | null) => {
        expect(right.encodeStateVector()).toEqual(left.encodeStateVector());
        const [one, two] = [snapshot(left.listContentControls()), snapshot(right.listContentControls())];
        expect(two).toEqual(one);
        const saves = await Promise.all(
          [left, right].map(async (peer, index) =>
            part(await repackDocx(yrsToDocument(peer, bases[index]!)), 'word/document.xml')
          )
        );
        expect(saves[1]).toEqual(saves[0]!);
        const xml = new TextDecoder().decode(saves[0]);
        expect(xml).not.toContain('REF-LEGACY');
        expect(controlsInXml(saves[0]!).find((control) => control.id === '102')?.text ?? null).toBe(text);
      };
      await agree('REF-000');
      left.deleteRange({
        story: 'body',
        start: { paraId: '10000003', offset: 11 },
        end: { paraId: '10000003', offset: 12 },
      });
      sync(left, right);
      await agree(null);
      expect(left.undo()).toBe(true);
      sync(left, right);
      await agree('REF-000');
      expect(
        snapshot(right.listContentControls()).diagnostics.map((diagnostic) => diagnostic.code)
      ).toContain('legacy-control-value');
      expect(left.redo()).toBe(true);
      sync(left, right);
      await agree(null);
      sync(right, left);
      await agree(null);
    } finally {
      left.destroy();
      right.destroy();
    }
  });

  it('refuses unpaired surrogates as invalid text in validation and application', async () => {
    const session = await open();
    try {
      const version = session.version();
      const request = {
        expectVersion: version,
        steps: [
          { op: 'setContentControlText', target: { kind: 'tag', tag: 'document.title' }, text: 'a\uD800b' },
        ],
      } as const;
      for (const result of [session.validateEdits(request), session.applyEdits(request)]) {
        expect(result).toMatchObject({
          ok: false,
          failure: { code: 'invalid-step', reason: 'invalid-text', stepIndex: 0 },
        });
      }
      expect(session.version()).toBe(version);
      expect(session.canUndo()).toBe(false);
    } finally {
      session.destroy();
    }
  });

  it('routes the legacy string setter through the fill and keeps typed setters', async () => {
    const checkbox = `<w:sdt><w:sdtPr><w:tag w:val="agree"/><w:id w:val="9"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>`;
    const text = `<w:sdt><w:sdtPr><w:tag w:val="name"/><w:id w:val="8"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Old</w:t></w:r></w:sdtContent></w:sdt>`;
    const parts: PartsMap = new Map();
    parts.set(
      '[Content_Types].xml',
      toBytes(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
      )
    );
    parts.set(
      '_rels/.rels',
      toBytes(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
      )
    );
    parts.set(
      'word/document.xml',
      toBytes(
        `<w:document xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p><w:r><w:t xml:space="preserve">Agree: </w:t></w:r>${checkbox}<w:r><w:t xml:space="preserve"> Name: </w:t></w:r>${text}</w:p></w:body></w:document>`
      )
    );
    const bytes = new Uint8Array(rezipPartsToArrayBuffer(parts));
    const session = await open(bytes);
    try {
      const base = session.materializeDocx();
      if (!base) throw new Error('the opened package must materialize');
      session.setContentControlValue('8', 'Grace Hopper');
      session.setContentControlValue('9', { kind: 'checkbox', checked: true });
      expect(() => session.setContentControlValue('8', { kind: 'checkbox', checked: true })).toThrow();
      const content = snapshot(session.listContentControls());
      expect(byTag(content, 'name').value).toEqual({ kind: 'text', text: 'Grace Hopper' });
      const saved = await repackDocx(yrsToDocument(session, base));
      const xml = new TextDecoder().decode(part(saved, 'word/document.xml'));
      expect(xml).toContain('w14:checked w14:val="1"');
      expect(controlsInXml(part(saved, 'word/document.xml')).map((control) => control.text)).toEqual([
        '☒',
        'Grace Hopper',
      ]);
    } finally {
      session.destroy();
    }
  });
});
