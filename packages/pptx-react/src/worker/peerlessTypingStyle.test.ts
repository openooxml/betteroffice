import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { PresentationCommandEvents, PresentationCommandMethods, PresentationCommandSelection,
  ShapeDraft, TextBoxPrimitive } from '@betteroffice/pptx';
import { createSessionHost } from '../../../../shared/office-session/host';
import { isHostMessage } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { createPresentationCommandClient } from '../../../pptx/src/session/commandClient';
import { registerPresentationCommandHandlers } from '../../../pptx/src/session/commandHost';
import { DEFAULT_EDITOR_TEXT_STYLE } from '../../../pptx/src/textFormatting';
import * as wasm from '../../../pptx/src/wasm/loader';
import { effectiveStyleFromSelection } from '../textFormatting';
import { createPeerlessViewStore, peerlessInputReference } from './peerlessViewStore';

type Session = Awaited<ReturnType<typeof open>>;

const input = { inputTime: 1, undoGroup: 'typing-style' };
const draft: ShapeDraft = { name: 'New box',
  rect: { x: 200000, y: 3500000, width: 6000000, height: 1000000 },
  text: '', style: DEFAULT_EDITOR_TEXT_STYLE };
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
const tree = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';
const textBox = '<p:cNvSpPr txBox="1"/><p:nvPr/>';
const title = '<p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr>';

function rels(...targets: [kind: string, target: string][]): string {
  return `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${targets.map(([kind, target], index) =>
    `<Relationship Id="rId${index + 1}" Type="${rel}/${kind}" Target="${target}"/>`).join('')}</Relationships>`;
}

function shape(id: number, name: string, properties: string, y: number, paragraph: string): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/>${properties}</p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="${y}"/><a:ext cx="6000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paragraph}</p:txBody></p:sp>`;
}

async function deck(): Promise<Uint8Array> {
  const zip = new JSZip();
  const part = (path: string, type: string) => `<Override PartName="/ppt/${path}" ContentType="application/vnd.openxmlformats-officedocument.${type}+xml"/>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${part('presentation.xml', 'presentationml.presentation.main')}${part('slides/slide1.xml', 'presentationml.slide')}${part('slideLayouts/slideLayout1.xml', 'presentationml.slideLayout')}${part('slideMasters/slideMaster1.xml', 'presentationml.slideMaster')}${part('theme/theme1.xml', 'theme')}</Types>`);
  zip.file('_rels/.rels', rels(['officeDocument', 'ppt/presentation.xml']));
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId2"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:defaultTextStyle></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', rels(['slide', 'slides/slide1.xml'], ['slideMaster', 'slideMasters/slideMaster1.xml']));
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree>${tree}${
    shape(2, 'Existing box', textBox, 100000, '<a:p><a:r><a:rPr sz="3200"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>alpha</a:t></a:r></a:p>')}${
    shape(3, 'Title', title, 1200000, '<a:p/>')}${
    shape(4, 'Source box', textBox, 2300000, '<a:p><a:endParaRPr lang="en-US" sz="2800"/></a:p>')}</p:spTree></p:cSld></p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', rels(['slideLayout', '../slideLayouts/slideLayout1.xml']));
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout ${ns} type="title"><p:cSld><p:spTree>${tree}</p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels(['slideMaster', '../slideMasters/slideMaster1.xml']));
  zip.file('ppt/slideMasters/slideMaster1.xml', `<p:sldMaster ${ns}><p:cSld><p:spTree>${tree}${shape(2, 'Title', title, 1200000, '<a:p/>')}</p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle></p:txStyles></p:sldMaster>`);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels(['slideLayout', '../slideLayouts/slideLayout1.xml'], ['theme', '../theme/theme1.xml']));
  zip.file('ppt/theme/theme1.xml', '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Typing"><a:themeElements><a:clrScheme name="Typing"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>');
  return zip.generateAsync({ type: 'uint8array' });
}

async function open() {
  const bytes = await deck();
  const font = new Uint8Array(await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
  const pair = createInProcessPair();
  const registration = registerPresentationCommandHandlers({ initWasm: async () => {
    await wasm.initWasm(await readFile(new URL('../../../pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  } });
  let handle!: wasm.PresentationHandle;
  const original = wasm.openPresentation;
  const opened = spyOn(wasm, 'openPresentation').mockImplementation((...args) => {
    handle = original(...args);
    return handle;
  });
  let hold = false;
  const replies: (() => void)[] = [];
  createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>({ ...pair.host,
    post(message, transfer) {
      if (hold && isHostMessage(message) && message.kind === 'reply' && message.ok) {
        replies.push(() => pair.host.post(message, transfer));
      } else pair.host.post(message, transfer);
    },
  }, { handlers: registration.handlers, policies: registration.policies,
    context: null, onDispose: registration.dispose });
  try {
    const client = await createPresentationCommandClient(bytes, { session: 'typing-style', incarnation: 'typing-worker',
      fonts: [{ family: 'Arial', bytes: font.slice().buffer },
        { family: 'Liberation Sans', bytes: font.slice().buffer }] }, pair.client);
    opened.mockRestore();
    const legacy = original(bytes, { fonts: [{ family: 'Arial', bytes: font },
      { family: 'Liberation Sans', bytes: font }] });
    return { client, handle, legacy,
      hold() { hold = true; },
      release() { hold = false; for (const reply of replies.splice(0)) reply(); },
      async dispose() {
        hold = false;
        for (const reply of replies.splice(0)) reply();
        legacy.dispose();
        try { await client.dispose(); } finally { registration.dispose(); pair.client.close(); }
      } };
  } catch (error) {
    opened.mockRestore(); registration.dispose(); pair.client.close();
    throw error;
  }
}

function selected(session: Session): PresentationCommandSelection {
  const at = session.client.state.selection;
  if (!at || at.kind === 'object') throw new Error('Expected text selection');
  return at;
}

function storyNamed(handle: wasm.PresentationHandle, name: string, position = 0): PresentationCommandSelection {
  const slide = handle.snapshot().slides[0]!;
  const shape = slide.shapes.find((value) => value.name === name);
  if (!shape?.textStories[0]) throw new Error(`Expected story in ${name}`);
  return { slideId: slide.id, shapeId: shape.id, storyId: shape.textStories[0].id,
    anchor: position, focus: position, affinity: 'downstream' };
}

function toolbarSize(session: Session): number | null | undefined {
  return session.client.state.editState.toolbar.formatting?.fontSizePt;
}

async function select(session: Session, selection: PresentationCommandSelection): Promise<void> {
  await session.client.dispatch({ kind: 'setSelection', payload: { ...input, selection } });
}

async function type(session: Session, at: PresentationCommandSelection, text: string): Promise<void> {
  await session.client.dispatch({ kind: 'insertText', payload: { ...input, at: { source: 'confirmed', selection: at }, text } });
}

async function insertBox(session: Session) {
  const slideId = session.client.state.slides[0]!.id;
  await session.client.dispatch({ kind: 'insertTextBox', payload: { ...input, slideId, draft } });
  const legacySlide = session.legacy.snapshot().slides[0]!;
  const receipt = session.legacy.addTextBox(legacySlide.id, draft);
  const shape = session.legacy.snapshot().slides[0]!.shapes.find((shape) => shape.id === receipt.shapeId)!;
  return shape.textStories[0]!.id;
}

function legacyStyleAt(session: Session, storyId: string, position = 0) {
  return effectiveStyleFromSelection(session.legacy.story(storyId), position, position, DEFAULT_EDITOR_TEXT_STYLE);
}

async function savedRunSize(handle: wasm.PresentationHandle, marker: string): Promise<number> {
  const zip = await JSZip.loadAsync(handle.save());
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
  const run = [...slide.matchAll(/<a:r\b[^>]*>[\s\S]*?<\/a:r>/g)]
    .map((match) => match[0]).find((run) => run.includes(marker));
  const size = run?.match(/\bsz="([\d.]+)"/);
  if (!size) throw new Error('Expected an explicit saved run size');
  return Number(size[1]) / 100;
}

async function sizes(handle: wasm.PresentationHandle, at: PresentationCommandSelection, marker: string) {
  const run = wasm.presentationStoryFormatting(handle, at).paragraphs.flatMap((paragraph) => paragraph.runs)
    .find((run) => run.text.includes(marker));
  return { effective: run?.style.fontSizePt, saved: await savedRunSize(handle, marker) };
}

test('empty default draft stores no inherited run on either path', async () => {
  const session = await open();
  try {
    const legacyStory = await insertBox(session);
    const at = selected(session);
    for (const [handle, storyId] of [[session.handle, at.storyId], [session.legacy, legacyStory]] as const) {
      expect(handle.story(storyId).paragraphs.flatMap((paragraph) => paragraph.runs)).toEqual([]);
      expect(effectiveStyleFromSelection(handle.story(storyId), 0, 0, DEFAULT_EDITOR_TEXT_STYLE).fontSizePt).toBe(24);
      const zip = await JSZip.loadAsync(handle.save());
      const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
      const box = [...slide.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)]
        .map((match) => match[0]).find((shape) => shape.includes('name="New box"'));
      expect(box).toBeDefined();
      expect(box).not.toContain('sz="1800"');
    }
  } finally { await session.dispose(); }
});

for (const kind of ['insertText', 'replaceText', 'compositionCommit'] as const) {
  for (const scenario of ['new', 'existing', 'chosen'] as const) {
    test(`${kind} preserves the legacy run size in a ${scenario} box`, async () => {
      const session = await open();
      try {
        let legacyStory: string;
        let position = 0;
        if (scenario === 'existing') {
          position = 2;
          await select(session, storyNamed(session.handle, 'Existing box', position));
          legacyStory = storyNamed(session.legacy, 'Existing box').storyId;
        } else legacyStory = await insertBox(session);
        const at = selected(session);
        const legacyStyle = legacyStyleAt(session, legacyStory, position);
        if (scenario === 'chosen') {
          legacyStyle.fontSizePt = 40;
          await session.client.dispatch({ kind: 'formatText', payload: { ...input,
            at: { source: 'confirmed', selection: at }, patch: { fontSizePt: 40 } } });
        }
        if (kind === 'compositionCommit') await session.client.dispatch({ kind: 'compositionStart', payload: {
          ...input, at: { source: 'confirmed', selection: at }, compositionId: 'typing-ime' } });
        const expected = scenario === 'existing' ? 32 : scenario === 'chosen' ? 40 : 24;
        expect(toolbarSize(session)).toBe(expected);
        const marker = 'marker';
        const reference = { source: 'confirmed' as const, selection: at };
        session.legacy.insertText(legacyStory, position, marker, legacyStyle);
        await session.client.dispatch(kind === 'compositionCommit'
          ? { kind, payload: { ...input, at: reference, text: marker, compositionId: 'typing-ime' } }
          : { kind, payload: { ...input, at: reference, text: marker } });
        expect(await savedRunSize(session.legacy, marker)).toBe(expected);
        expect(await sizes(session.handle, at, marker)).toEqual({ effective: expected, saved: expected });
      } finally { await session.dispose(); }
    });
  }
}

for (const [name, points] of [['Title', 44], ['Source box', 28]] as const) {
  test(`typing into the empty ${name} keeps its inherited ${points} pt`, async () => {
    const session = await open();
    try {
      const at = storyNamed(session.handle, name);
      await select(session, at);
      expect(toolbarSize(session)).toBe(points);
      await type(session, at, 'marker');
      expect(await sizes(session.handle, at, 'marker')).toEqual({ effective: points, saved: points });
    } finally { await session.dispose(); }
  });
}

test('typing into an emptied new box falls back to the legacy editor size', async () => {
  const session = await open();
  try {
    const legacyStory = await insertBox(session);
    let at = selected(session);
    await session.client.dispatch({ kind: 'insertText', payload: { ...input,
      at: { source: 'confirmed', selection: at }, text: 'alpha', style: { fontSizePt: 32 } } });
    session.legacy.insertText(legacyStory, 0, 'alpha', { fontSizePt: 32 });
    at = { ...selected(session), anchor: 0, focus: 5 };
    await session.client.dispatch({ kind: 'replaceText', payload: { ...input,
      at: { source: 'confirmed', selection: at }, text: '' } });
    session.legacy.deleteText(legacyStory, 0, 5);
    at = selected(session);
    session.legacy.insertText(legacyStory, 0, 'marker', legacyStyleAt(session, legacyStory));
    expect(toolbarSize(session)).toBe(24);
    await type(session, at, 'marker');
    expect(await savedRunSize(session.legacy, 'marker')).toBe(24);
    expect(await sizes(session.handle, at, 'marker')).toEqual({ effective: 24, saved: 24 });
  } finally { await session.dispose(); }
});

test('leaving and re-entering the still-empty new box keeps the draft size', async () => {
  const session = await open();
  try {
    const legacyStory = await insertBox(session);
    const at = selected(session);
    await select(session, storyNamed(session.handle, 'Existing box', 1));
    expect(toolbarSize(session)).toBe(32);
    await select(session, at);
    expect(toolbarSize(session)).toBe(legacyStyleAt(session, legacyStory).fontSizePt);
    await type(session, at, 'marker');
    expect(await sizes(session.handle, at, 'marker')).toEqual({ effective: 24, saved: 24 });
  } finally { await session.dispose(); }
});

test('undo and redo of the insertion keep the draft size', async () => {
  const session = await open();
  try {
    await insertBox(session);
    const at = selected(session);
    await type(session, at, 'alpha');
    await session.client.dispatch({ kind: 'undo', payload: input });
    await select(session, at);
    expect(toolbarSize(session)).toBe(24);
    await session.client.dispatch({ kind: 'undo', payload: input });
    expect(session.handle.snapshot().slides[0]!.shapes.some((shape) => shape.id === at.shapeId)).toBe(false);
    await session.client.dispatch({ kind: 'redo', payload: input });
    await select(session, at);
    expect(toolbarSize(session)).toBe(24);
    await type(session, at, 'marker');
    expect(await sizes(session.handle, at, 'marker')).toEqual({ effective: 24, saved: 24 });
  } finally { await session.dispose(); }
});

for (const points of [24, 40]) {
  test(`pending local echo and committed new text both use ${points} pt`, async () => {
    const session = await open();
    let store: ReturnType<typeof createPeerlessViewStore> | undefined;
    try {
      await insertBox(session);
      const at = selected(session);
      if (points === 40) await session.client.dispatch({ kind: 'formatText', payload: { ...input,
        at: { source: 'confirmed', selection: at }, patch: { fontSizePt: points } } });
      store = createPeerlessViewStore(session.client, { paintFrame() {}, paintEcho() {}, restoreLine() {} },
        { echoEnabled: true });
      await store.dispatch({ kind: 'insertText', payload: { ...input, at: peerlessInputReference(store)!, text: 'alpha' } });
      session.hold();
      const pending = store.dispatch({ kind: 'insertText', payload: { ...input,
        at: peerlessInputReference(store)!, text: 'a' } });
      const echo = store.getSnapshot().echo;
      session.release();
      await pending;
      const view = store.getSnapshot();
      const box = view.confirmed!.displayList.primitives.find((primitive) =>
        primitive.kind === 'textBox' && primitive.storyId === at.storyId) as TextBoxPrimitive;
      expect(echo?.insertionStyle.fontSizePx).toBeCloseTo(points * 96 / 72, 5);
      expect(box.lines[0]!.runs[0]!.fontSizePx).toBe(echo!.insertionStyle.fontSizePx);
      expect(await sizes(session.handle, at, 'alphaa')).toEqual({ effective: points, saved: points });
      expect(view.echo).toBeNull();
    } finally { store?.dispose(); await session.dispose(); }
  });
}
