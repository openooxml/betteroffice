import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createSessionClient } from '../../../../shared/office-session/client';
import { createSessionHost } from '../../../../shared/office-session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import type { PptxStoryText } from '../edits';
import * as wasm from '../wasm/loader';
import { registerPresentationCommandHandlers } from './commandHost';
import { PRESENTATION_COMMAND_METHODS, type PresentationCommandAck, type PresentationCommandEnvelope,
  type PresentationCommandEvents, type PresentationCommandIntent, type PresentationCommandMethods,
  type PresentationCommandSelection } from './commandProtocol';

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

const identity = { session: 's3', incarnation: 's3-worker' };
const input = { inputTime: 1, undoGroup: 's3' };

function selection(story: PptxStoryText, anchor = 0, focus = anchor): PresentationCommandSelection {
  return { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
    anchor, focus, affinity: 'downstream' };
}

async function open(bytes?: Uint8Array) {
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
    { methods: PRESENTATION_COMMAND_METHODS, silenceMs: 2000 });
  try {
    const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    const state = await rpc.call.openCommandEditor(new Uint8Array(bytes ?? await deck(['alpha'])).buffer,
      { ...identity, fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }] });
    const result = (await rpc.call.readEditor({ ...identity, seq: 0 })).result;
    if (!result.ok) throw new Error(result.failure.message);
    return { state, story: result.stories[0]!, stories: result.stories, handle: () => handle!,
      command(state: PresentationCommandAck, seq: number, intent: unknown, afterSeq?: number) {
        return rpc.call.command({ ...identity, seq, basis: { epoch: state.epoch, version: state.version },
          ...(afterSeq === undefined ? {} : { afterSeq }), ...(intent as PresentationCommandIntent) } as PresentationCommandEnvelope);
      },
      async dispose() { try { await rpc.dispose(); } finally { opened.mockRestore(); registration.dispose(); pair.client.close(); } },
    };
  } catch (error) {
    opened.mockRestore();
    registration.dispose();
    pair.client.close();
    throw error;
  }
}

test('range_format', async () => {
  const session = await open();
  try {
    const at = { source: 'confirmed', selection: selection(session.story, 1, 4) };
    const ack = await session.command(session.state, 1, { kind: 'formatText', payload: { ...input, at,
      patch: { bold: true, italic: true, underline: 'sng' } } });
    expect(ack.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
    const runs = session.handle().story(session.story.storyId).paragraphs[0]!.runs;
    expect(runs.map((run) => [run.text, run.style.bold === true, run.style.italic === true, run.style.underline === 'sng']))
      .toEqual([['a', false, false, false], ['lph', true, true, true], ['a', false, false, false]]);
    expect(ack.selection).toEqual(at.selection);
    expect(ack.editState.toolbar).toMatchObject({ textContext: 'range' });
    expect(ack.dirtySlideIds).toEqual([session.story.slideId]);
    expect(ack.frame?.slideRevision).toBe(1);
    const aligned = await session.command(ack, 2, { kind: 'setParagraphAlignment', payload: {
      ...input, at: { source: 'afterSeq', seq: 1 }, alignment: 'ctr' } }, 1);
    expect(aligned.editState.toolbar.alignment).toBe('ctr');
    expect(aligned.frame?.slideRevision).toBe(2);
  } finally { await session.dispose(); }
});

test('caret_marks', async () => {
  const session = await open();
  try {
    const marked = await session.command(session.state, 1, { kind: 'formatText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(session.story, 2) },
      patch: { bold: true, italic: true, underline: 'sng' } } });
    expect(marked.editState.toolbar.storedMarks).toEqual({ bold: true, italic: true, underline: 'sng' });
    expect(marked.editState.toolbar).toMatchObject({ textContext: 'caret' });
    expect(marked.editState.toolbar.formatting).toMatchObject({ bold: true, italic: true, underline: 'sng' });
    expect(marked.dirtySlideIds).toEqual([]);
    const inserted = await session.command(marked, 2, { kind: 'insertText', payload: {
      ...input, at: { source: 'afterSeq', seq: 1 }, text: 'Z' } }, 1);
    const run = session.handle().story(session.story.storyId).paragraphs[0]!.runs.find((run) => run.text === 'Z');
    expect(run?.style).toMatchObject({ bold: true, italic: true, underline: 'sng' });
    const moved = await session.command(inserted, 3, { kind: 'setSelection',
      payload: { ...input, selection: selection(session.story, 0) } });
    expect(moved.editState.toolbar.storedMarks).toEqual({});
    const typed = await session.command(moved, 4, { kind: 'insertText', payload: {
      ...input, at: { source: 'afterSeq', seq: 3 }, text: 'Q' } }, 3);
    expect(typed.outcomes[0]?.committed).toBe(true);
    expect(session.handle().story(session.story.storyId).paragraphs[0]!.runs[0]!.style.bold).not.toBe(true);
  } finally { await session.dispose(); }
});

for (const kind of ['insertText', 'compositionCommit'] as const) {
  test(`inherited_caret_marks_preserve_cascade_on_${kind}`, async () => {
    const zip = await JSZip.loadAsync(await deck(['alpha']));
    const path = 'ppt/slides/slide1.xml';
    zip.file(path, (await zip.file(path)!.async('string'))
      .replace('<a:lstStyle/>', '<a:lstStyle><a:lvl1pPr><a:defRPr b="1" i="1" u="sng" sz="3600"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></a:lstStyle>')
      .replace(/ sz="2400"/g, '')
      .replace(/<a:latin typeface="Liberation Sans"\/>/g, ''));
    const session = await open(await zip.generateAsync({ type: 'uint8array' }));
    try {
      const selected = await session.command(session.state, 1, { kind: 'setSelection', payload: {
        ...input, selection: selection(session.story, 2) } });
      const expected = { bold: false, italic: true, underline: 'sng', fontSizePt: 36,
        fontFamily: 'Arial', color: '#123456' };
      const marked = await session.command(selected, 2, { kind: 'formatText', payload: { ...input,
        at: { source: 'afterSeq', seq: 1 }, patch: {}, toggle: 'bold' } }, 1);
      expect(marked.editState.toolbar.storedMarks).toEqual({ bold: false });
      expect(marked.editState.toolbar.formatting).toMatchObject(expected);
      let state = marked;
      if (kind === 'compositionCommit') {
        state = await session.command(marked, 3, { kind: 'compositionStart', payload: { ...input,
          at: { source: 'afterSeq', seq: 2 }, compositionId: 'inherited-ime' } }, 2);
        expect(state.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
      }
      const text = kind === 'insertText' ? 'Z' : '日本語';
      const inserted = await session.command(state, state.throughSeq + 1, { kind, payload: { ...input,
        at: { source: 'afterSeq', seq: state.throughSeq }, text,
        ...(kind === 'compositionCommit' ? { compositionId: 'inherited-ime' } : {}) } }, state.throughSeq);
      expect(inserted.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
      const runs = session.handle().story(session.story.storyId).paragraphs[0]!.runs;
      expect(runs.find((run) => run.text === text)?.style).toMatchObject(expected);
      expect(inserted.editState.toolbar.formatting).toMatchObject(expected);
    } finally { await session.dispose(); }
  });
}

test('mixed_style', async () => {
  const session = await open();
  try {
    const formatted = await session.command(session.state, 1, { kind: 'formatText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(session.story, 0, 2) },
      patch: { bold: true, italic: true, underline: 'sng', fontSizePt: 36, color: '#FF0000' } } });
    const selected = await session.command(formatted, 2, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 0, 5) } });
    expect(selected.editState.toolbar.formatting).toMatchObject({
      bold: null, italic: null, underline: null, fontSizePt: null, color: null,
    });
    const reversed = await session.command(selected, 3, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 5, 0) } });
    expect(reversed.editState.toolbar.formatting).toEqual(selected.editState.toolbar.formatting);
    const shape = await session.command(reversed, 4, { kind: 'setSelection', payload: {
      ...input, selection: { ...selection(session.story), kind: 'shape' } } });
    expect(shape.editState.toolbar).toMatchObject({ textContext: 'shape', formatting: { bold: null, underline: null } });
  } finally { await session.dispose(); }
});

test('inherited_style_projection_and_toggles', async () => {
  const zip = await JSZip.loadAsync(await deck(['alpha\nbeta\n']));
  const path = 'ppt/slides/slide1.xml';
  const slide = (await zip.file(path)!.async('string'))
    .replace('<a:lstStyle/>', '<a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr b="1" i="1" u="sng" sz="3600"/></a:lvl1pPr></a:lstStyle>')
    .replace(/ sz="2400"/g, '')
    .replace('<a:p><a:r><a:rPr lang="en-US"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>beta',
      '<a:p><a:pPr algn="l"><a:defRPr b="0" i="0" u="none" sz="1000"/></a:pPr><a:r><a:rPr lang="en-US"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>beta');
  zip.file(path, slide);
  const session = await open(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const selected = await session.command(session.state, 1, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 1) } });
    expect(selected.editState.toolbar.formatting).toMatchObject({ bold: true, italic: true, underline: 'sng', fontSizePt: 36 });
    expect(selected.editState.toolbar.alignment).toBe('ctr');
    const bold = await session.command(selected, 2, { kind: 'formatText', payload: { ...input,
      at: { source: 'afterSeq', seq: 1 }, patch: {}, toggle: 'bold' } }, 1);
    expect(bold.editState.toolbar.storedMarks).toEqual({ bold: false });
    const underline = await session.command(bold, 3, { kind: 'formatText', payload: { ...input,
      at: { source: 'afterSeq', seq: 2 }, patch: {}, toggle: 'underline' } }, 2);
    expect(underline.editState.toolbar.storedMarks).toEqual({ bold: false, underline: 'none' });
    const mixed = await session.command(underline, 4, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 0, 10) } });
    expect(mixed.editState.toolbar.formatting).toMatchObject({ bold: null, italic: null, underline: null, fontSizePt: null });
    expect(mixed.editState.toolbar.alignment).toBeNull();
    const empty = await session.command(mixed, 5, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 11) } });
    expect(empty.editState.toolbar.formatting).toMatchObject({ bold: true, italic: true, underline: 'sng', fontSizePt: 36 });
  } finally { await session.dispose(); }
});

test('empty_shape_projection_matches_inherited_caret_and_size_step', async () => {
  const zip = await JSZip.loadAsync(await deck(['']));
  const path = 'ppt/slides/slide1.xml';
  zip.file(path, (await zip.file(path)!.async('string'))
    .replace('<a:lstStyle/>', '<a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr b="1" i="1" u="sng" sz="3600"><a:latin typeface="Calibri"/></a:defRPr></a:lvl1pPr></a:lstStyle>')
    .replace('<a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t></a:t></a:r>', ''));
  const session = await open(await zip.generateAsync({ type: 'uint8array' }));
  const formatted = spyOn(session.handle(), 'formatText');
  try {
    const caret = await session.command(session.state, 1, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story) } });
    expect(caret.editState.toolbar.formatting).toMatchObject({ bold: true, italic: true, underline: 'sng',
      fontSizePt: 36, fontFamily: 'Calibri' });
    const shape = await session.command(caret, 2, { kind: 'setSelection', payload: {
      ...input, selection: { ...selection(session.story), kind: 'shape' } } });
    expect(shape.editState.toolbar.formatting).toEqual(caret.editState.toolbar.formatting);
    expect(shape.editState.toolbar.alignment).toBe(caret.editState.toolbar.alignment);
    const stepped = await session.command(shape, 3, { kind: 'formatText', payload: { ...input,
      at: { source: 'afterSeq', seq: 2 }, patch: {}, fontSizeStep: { direction: 'increase', sizes: [24, 28, 36, 48] } } }, 2);
    expect(stepped.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(formatted).toHaveBeenCalledWith(session.story.storyId, 0, 1, { fontSizePt: 48 });
    const moved = await session.command(stepped, 4, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story) } });
    expect(moved.editState.toolbar.formatting).toEqual(stepped.editState.toolbar.formatting);
  } finally { formatted.mockRestore(); await session.dispose(); }
});

test('empty_paragraph_range_projection_matches_inherited_caret', async () => {
  const zip = await JSZip.loadAsync(await deck(['alpha\n\nbeta']));
  const path = 'ppt/slides/slide1.xml';
  zip.file(path, (await zip.file(path)!.async('string'))
    .replace('<a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t></a:t></a:r></a:p>',
      '<a:p><a:pPr algn="ctr"><a:defRPr b="1" i="1" u="sng" sz="3600"><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p>'));
  const session = await open(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const caret = await session.command(session.state, 1, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 6) } });
    expect(caret.editState.toolbar.formatting).toMatchObject({ bold: true, italic: true, underline: 'sng',
      fontSizePt: 36, fontFamily: 'Calibri' });
    const range = await session.command(caret, 2, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 6, 7) } });
    expect(range.editState.toolbar.formatting).toEqual(caret.editState.toolbar.formatting);
    expect(range.editState.toolbar.alignment).toBe(caret.editState.toolbar.alignment);
  } finally { await session.dispose(); }
});

for (const kind of ['range', 'shape'] as const) {
  test(`mixed_empty_paragraph_styles_${kind}`, async () => {
    const zip = await JSZip.loadAsync(await deck(['alpha\n\nbeta']));
    const path = 'ppt/slides/slide1.xml';
    zip.file(path, (await zip.file(path)!.async('string'))
      .replace('<a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t></a:t></a:r></a:p>',
        '<a:p><a:pPr><a:defRPr b="1" i="1" u="sng" sz="3600"><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p>'));
    const session = await open(await zip.generateAsync({ type: 'uint8array' }));
    try {
      const selected = await session.command(session.state, 1, { kind: 'setSelection', payload: {
        ...input, selection: kind === 'shape' ? { ...selection(session.story), kind: 'shape' }
          : selection(session.story, 0, 7) } });
      expect(selected.outcomes[0]?.status).toBe('applied');
      expect(selected.editState.toolbar.formatting).toMatchObject({ fontSizePt: null, fontFamily: null,
        bold: null, italic: null, underline: null });
      const textOnly = await session.command(selected, 2, { kind: 'setSelection', payload: {
        ...input, selection: selection(session.story, 0, 6) } });
      expect(textOnly.editState.toolbar.formatting?.fontSizePt).toBe(24);
      const reversed = await session.command(textOnly, 3, { kind: 'setSelection', payload: {
        ...input, selection: selection(session.story, 7, 0) } });
      expect(reversed.editState.toolbar.formatting).toEqual(selected.editState.toolbar.formatting);
    } finally { await session.dispose(); }
  });
}

test('font_size_step_resolves_predecessor_target', async () => {
  const session = await open();
  try {
    const large = await session.command(session.state, 1, { kind: 'formatText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(session.story, 2, 5) }, patch: { fontSizePt: 36 } } });
    const small = await session.command(large, 2, { kind: 'formatText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(session.story, 0, 2) }, patch: { fontSizePt: 10 } } });
    const moved = await session.command(small, 3, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 3) } });
    const stepped = await session.command(small, 4, { kind: 'formatText', payload: { ...input,
      at: { source: 'afterSeq', seq: 3 }, patch: {}, fontSizeStep: { direction: 'increase', sizes: [10, 11, 36, 48] } } }, 3);
    expect(moved.editState.toolbar.formatting?.fontSizePt).toBe(36);
    expect(stepped.outcomes[0]?.status).toBe('applied');
    expect(stepped.editState.toolbar.storedMarks).toEqual({ fontSizePt: 48 });
    const decreased = await session.command(stepped, 5, { kind: 'formatText', payload: { ...input,
      at: { source: 'afterSeq', seq: 4 }, patch: {}, fontSizeStep: { direction: 'decrease', sizes: [10, 11, 36, 48] } } }, 4);
    expect(decreased.editState.toolbar.storedMarks).toEqual({ fontSizePt: 36 });
    let state = decreased;
    for (const fontSizeStep of [{ direction: 'other', sizes: [] }, { direction: 'increase', sizes: [NaN] },
      { direction: 'decrease', sizes: [36, 10] }]) {
      const seq = state.throughSeq + 1;
      state = await session.command(state, seq, { kind: 'formatText', payload: { ...input,
        at: { source: 'afterSeq', seq: 5 }, patch: {}, fontSizeStep } }, 5);
      expect(state.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
      expect(state.editState.toolbar.storedMarks).toEqual({ fontSizePt: 36 });
    }
  } finally { await session.dispose(); }
});

test('autofit_font_size_steps_use_unscaled_hundredths', async () => {
  const zip = await JSZip.loadAsync(await deck(['alpha\nbeta']));
  const path = 'ppt/slides/slide1.xml';
  zip.file(path, (await zip.file(path)!.async('string'))
    .replace('<a:bodyPr/>', '<a:bodyPr><a:normAutofit fontScale="50000"/></a:bodyPr>')
    .replace(/ sz="2400"/g, '')
    .replace('<a:lstStyle/>', '<a:lstStyle><a:lvl1pPr><a:defRPr sz="3600"/></a:lvl1pPr></a:lstStyle>')
    .replace('<a:p><a:r><a:rPr lang="en-US"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>beta',
      '<a:p><a:pPr><a:defRPr sz="1000"/></a:pPr><a:r><a:rPr lang="en-US"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>beta'));
  const session = await open(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const selected = await session.command(session.state, 1, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 0, 5) } });
    expect(selected.editState.toolbar.formatting?.fontSizePt).toBe(36);
    const increased = await session.command(selected, 2, { kind: 'formatText', payload: { ...input,
      at: { source: 'afterSeq', seq: 1 }, patch: {}, fontSizeStep: { direction: 'increase', sizes: [10, 20, 36, 48] } } }, 1);
    expect(increased.editState.toolbar.formatting?.fontSizePt).toBe(48);
    expect(session.handle().story(session.story.storyId).paragraphs[0]!.runs[0]!.style.fontSizePt).toBe(48);
    const small = await session.command(increased, 3, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 7) } });
    expect(small.editState.toolbar.formatting?.fontSizePt).toBe(10);
    const stepped = await session.command(small, 4, { kind: 'formatText', payload: { ...input,
      at: { source: 'afterSeq', seq: 3 }, patch: {}, fontSizeStep: { direction: 'increase', sizes: [10.004, 11] } } }, 3);
    expect(stepped.editState.toolbar.storedMarks.fontSizePt).toBe(11);
    const fractional = await session.command(stepped, 5, { kind: 'formatText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(session.story, 6, 8) }, patch: { fontSizePt: 10.004 } } });
    const uniform = await session.command(fractional, 6, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 6, 10) } });
    expect(uniform.editState.toolbar.formatting?.fontSizePt).toBe(10);
  } finally { await session.dispose(); }
});

test('substituted_authored_families_remain_mixed_and_keep_font_choice', async () => {
  const zip = await JSZip.loadAsync(await deck(['alpha\nbeta']));
  const path = 'ppt/slides/slide1.xml';
  zip.file(path, (await zip.file(path)!.async('string'))
    .replace('typeface="Liberation Sans"', 'typeface="Arial"')
    .replace('typeface="Liberation Sans"', 'typeface="Calibri"'));
  const session = await open(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const selected = await session.command(session.state, 1, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 0, 5) } });
    expect(selected.editState.toolbar.formatting?.fontFamily).toBe('Arial');
    const mixed = await session.command(selected, 2, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 0, 10) } });
    expect(mixed.editState.toolbar.formatting?.fontFamily).toBeNull();
    const formatted = await session.command(mixed, 3, { kind: 'formatText', payload: { ...input,
      at: { source: 'afterSeq', seq: 2 }, patch: { fontFamily: 'Calibri' } } }, 2);
    expect(formatted.editState.toolbar.formatting?.fontFamily).toBe('Calibri');
  } finally { await session.dispose(); }
});

test('theme_font_references_project_before_substitution', async () => {
  const zip = await JSZip.loadAsync(await deck(['alpha\nbeta']));
  const path = 'ppt/slides/slide1.xml';
  zip.file(path, (await zip.file(path)!.async('string'))
    .replace('typeface="Liberation Sans"', 'typeface="+mj-lt"')
    .replace('typeface="Liberation Sans"', 'typeface="+mn-lt"'));
  zip.file('ppt/theme/theme1.xml', `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Test"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`);
  const session = await open(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const major = await session.command(session.state, 1, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 1) } });
    expect(major.editState.toolbar.formatting?.fontFamily).toBe('Arial');
    const minor = await session.command(major, 2, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 7) } });
    expect(minor.editState.toolbar.formatting?.fontFamily).toBe('Calibri');
  } finally { await session.dispose(); }
});

test('font_size_steps_above_400_are_monotonic_and_share_validation_bound', async () => {
  const session = await open();
  try {
    let state: PresentationCommandAck = session.state;
    for (const [size, expected] of [[450, 451], [3999.5, 4000], [4000, 4000]]) {
      const seq = state.throughSeq + 1;
      state = await session.command(state, seq, { kind: 'formatText', payload: { ...input,
        at: { source: 'confirmed', selection: selection(session.story, 0, 5) }, patch: { fontSizePt: size } } });
      state = await session.command(state, seq + 1, { kind: 'formatText', payload: { ...input,
        at: { source: 'afterSeq', seq }, patch: {}, fontSizeStep: { direction: 'increase', sizes: [] } } }, seq);
      expect(state.editState.toolbar.formatting?.fontSizePt).toBe(expected);
      expect(state.outcomes[0]?.status).toBe('applied');
    }
    const seq = state.throughSeq + 1;
    const refused = await session.command(state, seq, { kind: 'formatText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(session.story, 1) }, patch: { fontSizePt: 4000.01 } } });
    expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
  } finally { await session.dispose(); }
});

test('typing_does_not_add_a_full_slide_formatting_pass', async () => {
  const session = await open();
  const layout = spyOn(session.handle(), 'layoutSlide');
  const frames = spyOn(wasm, 'presentationDisplayListJson');
  try {
    const selected = await session.command(session.state, 1, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 2) } });
    layout.mockClear();
    frames.mockClear();
    const typed = await session.command(selected, 2, { kind: 'insertText', payload: { ...input,
      at: { source: 'afterSeq', seq: 1 }, text: 'Z' } }, 1);
    expect(typed.outcomes[0]?.committed).toBe(true);
    expect(layout).not.toHaveBeenCalled();
    expect(frames).toHaveBeenCalledTimes(1);
    expect(typed.editState.toolbar.formatting?.fontSizePt).toBe(24);
  } finally { layout.mockRestore(); frames.mockRestore(); await session.dispose(); }
});

test('formatting_projection_invalidates_only_the_changed_story', async () => {
  const session = await open(await deck(['alpha', 'beta']));
  const formatting = spyOn(wasm, 'presentationStoryFormatting');
  try {
    const selected = await session.command(session.state, 1, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 2) } });
    expect(formatting).toHaveBeenCalledTimes(1);
    const repeated = await session.command(selected, 2, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 3) } });
    expect(formatting).toHaveBeenCalledTimes(1);
    const other = await session.command(repeated, 3, { kind: 'formatText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(session.stories[1]!, 0, 4) }, patch: { fontSizePt: 36 } } });
    expect(formatting).toHaveBeenCalledTimes(2);
    const returned = await session.command(other, 4, { kind: 'setSelection', payload: {
      ...input, selection: selection(session.story, 3) } });
    expect(formatting).toHaveBeenCalledTimes(2);
    const typed = await session.command(returned, 5, { kind: 'insertText', payload: { ...input,
      at: { source: 'afterSeq', seq: 4 }, text: 'Z' } }, 4);
    expect(formatting).toHaveBeenCalledTimes(3);
    expect(typed.editState.toolbar.formatting?.fontSizePt).toBe(24);
    await session.command(typed, 6, { kind: 'formatText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(session.story, 0, 6) }, patch: { fontSizePt: 48 } } });
    expect(formatting).toHaveBeenCalledTimes(4);
  } finally { formatting.mockRestore(); await session.dispose(); }
});

test('undefined_format_patch_preserves_marks_and_validates_keys', async () => {
  const session = await open();
  try {
    const marked = await session.command(session.state, 1, { kind: 'formatText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(session.story, 2) }, patch: { bold: true } } });
    const ignored = await session.command(marked, 2, { kind: 'formatText', payload: { ...input,
      at: { source: 'afterSeq', seq: 1 }, patch: { bold: undefined, italic: true } } }, 1);
    expect(ignored.editState.toolbar.storedMarks).toEqual({ bold: true, italic: true });
    expect(Object.values(ignored.editState.toolbar.storedMarks)).not.toContain(undefined);
    const inserted = await session.command(ignored, 3, { kind: 'insertText', payload: { ...input,
      at: { source: 'afterSeq', seq: 2 }, text: 'Z' } }, 2);
    expect(session.handle().story(session.story.storyId).paragraphs[0]!.runs.find((run) => run.text === 'Z')?.style)
      .toMatchObject({ bold: true, italic: true });
    const range = await session.command(inserted, 4, { kind: 'formatText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(session.story, 0, 2) }, patch: { bold: true } } });
    const rangeIgnored = await session.command(range, 5, { kind: 'formatText', payload: { ...input,
      at: { source: 'afterSeq', seq: 4 }, patch: { bold: undefined, italic: true } } }, 4);
    expect(rangeIgnored.editState.toolbar.formatting).toMatchObject({ bold: true, italic: true });
    let state = rangeIgnored;
    for (const patch of [{ unknown: undefined }, { bold: 'yes' }, { underline: null }]) {
      const seq = state.throughSeq + 1;
      state = await session.command(state, seq, { kind: 'formatText', payload: { ...input,
        at: { source: 'afterSeq', seq: 5 }, patch } }, 5);
      expect(state.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
    }
  } finally { await session.dispose(); }
});
