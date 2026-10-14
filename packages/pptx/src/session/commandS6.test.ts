import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createSessionClient } from '../../../../shared/office-session/client';
import { createSessionHost } from '../../../../shared/office-session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import * as wasm from '../wasm/loader';
import { registerPresentationCommandHandlers } from './commandHost';
import { PRESENTATION_COMMAND_METHODS, type PresentationCommandAck, type PresentationCommandEnvelope,
  type PresentationCommandEvents, type PresentationCommandIntent, type PresentationCommandMethods } from './commandProtocol';

const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function xml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

async function deck(texts = ['alpha', 'beta', 'gamma']): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${texts.map((_, index) => `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst>${texts.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${texts.map((_, index) => `<Relationship Id="rId${index + 1}" Type="${rel}/slide" Target="slides/slide${index + 1}.xml"/>`).join('')}</Relationships>`);
  for (const [index, text] of texts.entries()) {
    const paragraphs = text.split('\n').map((paragraph) => `<a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>${xml(paragraph)}</a:t></a:r></a:p>`).join('');
    zip.file(`ppt/slides/slide${index + 1}.xml`, `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  }
  return zip.generateAsync({ type: 'uint8array' });
}

const identity = { session: 's6', incarnation: 's6-worker' };
const input = { inputTime: 1, undoGroup: 's6' };

test('command_host_does_not_import_react_source', async () => {
  const source = await readFile(new URL('./commandHost.ts', import.meta.url), 'utf8');
  expect(source).not.toMatch(/from ['"][^'"]*pptx-react/);
});

async function open() {
  const pair = createInProcessPair();
  const registration = registerPresentationCommandHandlers({ initWasm: async () => {
    await wasm.initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  } });
  let handle: wasm.PresentationHandle;
  const original = wasm.openPresentation;
  const opened = spyOn(wasm, 'openPresentation').mockImplementation((...args) => {
    handle = original(...args);
    return handle;
  });
  createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>(pair.host, {
    handlers: registration.handlers, policies: registration.policies, context: null, onDispose: registration.dispose,
  });
  const rpc = createSessionClient<PresentationCommandMethods, PresentationCommandEvents>(pair.client,
    { methods: PRESENTATION_COMMAND_METHODS });
  try {
    const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    const state = await bounded(rpc.call.openCommandEditor(new Uint8Array(await bounded(deck(['alpha']))).buffer,
      { ...identity, fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }] }));
    const result = (await bounded(rpc.call.readEditor({ ...identity, seq: 0 }))).result;
    if (!result.ok) throw new Error(result.failure.message);
    return { state, story: result.stories[0]!, handle: () => handle!,
      command(state: PresentationCommandAck, seq: number, intent: unknown, afterSeq?: number) {
        return bounded(rpc.call.command({ ...identity, seq, basis: { epoch: state.epoch, version: state.version },
          ...(afterSeq === undefined ? {} : { afterSeq }), ...(intent as PresentationCommandIntent) } as PresentationCommandEnvelope));
      },
      async dispose() { try { await bounded(rpc.dispose()); } finally { opened.mockRestore(); registration.dispose(); pair.client.close(); } },
    };
  } catch (error) {
    opened.mockRestore();
    registration.dispose();
    pair.client.close();
    throw error;
  }
}

function bounded<T>(promise: Promise<T>): Promise<T> {
  let timer: ReturnType<typeof setTimeout>;
  return Promise.race([promise, new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error('S6 command timed out')), 5000);
  })]).finally(() => clearTimeout(timer));
}

test('stale_target_removed_or_replaced', async () => {
  for (const replacement of [false, true]) {
    const session = await open();
    let restoreSnapshot = () => {};
    try {
      const selected = await session.command(session.state, 1, { kind: 'selectObject', payload: {
        ...input, frame: session.state.frame, x: 11, y: 11, scale: 1,
      } });
      expect(selected.selection).toMatchObject({ kind: 'object', shapeId: session.story.shapeId });
      const object = selected.selection;
      if (object?.kind !== 'object') throw new Error('Object selection required');
      const snapshot = session.handle().snapshot();
      const shape = snapshot.slides[0]!.shapes[0]!;
      if (replacement) snapshot.slides[0]!.shapes[0] = { ...shape, sourceId: shape.sourceId + 100 };
      else snapshot.slides[0]!.shapes = [];
      const snapshotSpy = spyOn(session.handle(), 'snapshot').mockReturnValue(snapshot);
      restoreSnapshot = () => snapshotSpy.mockRestore();
      const move = spyOn(session.handle(), 'moveShape');
      try {
        const refused = await session.command(selected, 2, { kind: 'moveShape', payload: {
          ...input, frame: selected.frame, target: { slideId: object.slideId, shapeId: object.shapeId, token: object.token },
          x: 200000, y: 200000,
        } });
        expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'bad-target', committed: false });
        expect(move).not.toHaveBeenCalled();
      } finally { move.mockRestore(); }
    } finally { restoreSnapshot(); await session.dispose(); }
  }
});

test('shape_transform_read_only_and_frame_basis', async () => {
  const session = await open();
  try {
    const selected = await session.command(session.state, 1, { kind: 'selectObject', payload: {
      ...input, frame: session.state.frame, x: 11, y: 11, scale: 1,
    } });
    expect(selected.selection?.kind).toBe('object');
    const object = selected.selection;
    if (object?.kind !== 'object') throw new Error('Object selection required');
    const payload = { ...input, frame: selected.frame,
      target: { slideId: object.slideId, shapeId: object.shapeId, token: object.token }, x: 200000, y: 200000 };
    const stale = await session.command(selected, 2, { kind: 'moveShape',
      payload: { ...payload, frame: { ...selected.frame, incarnation: 'other' } } });
    expect(stale.outcomes[0]).toMatchObject({ reason: 'stale-basis', committed: false });
    const mode = await session.command(stale, 3, { kind: 'setReadOnly', payload: { ...input, readOnly: true } });
    const refused = await session.command(mode, 4, { kind: 'moveShape', payload });
    expect(refused.outcomes[0]).toMatchObject({ reason: 'read-only', committed: false });
  } finally { await session.dispose(); }
});

test('shape_identity_selection_checks_target_and_frame', async () => {
  const session = await open();
  try {
    const selected = await session.command(session.state, 1, { kind: 'selectObject', payload: {
      ...input, frame: session.state.frame, shapeId: session.story.shapeId,
    } });
    expect(selected.selection).toMatchObject({ kind: 'object', shapeId: session.story.shapeId });
    const missing = await session.command(selected, 2, { kind: 'selectObject', payload: {
      ...input, frame: selected.frame, shapeId: 'missing',
    } });
    expect(missing.outcomes[0]).toMatchObject({ status: 'refused', reason: 'bad-target', committed: false });
    const stale = await session.command(missing, 3, { kind: 'selectObject', payload: {
      ...input, frame: { ...missing.frame, incarnation: 'other' }, shapeId: session.story.shapeId,
    } });
    expect(stale.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-basis', committed: false });
  } finally { await session.dispose(); }
});
