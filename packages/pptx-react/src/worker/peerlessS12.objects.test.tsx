import { expect, mock, test } from 'bun:test';
import { act, fireEvent, waitFor } from '@testing-library/react';
import JSZip from 'jszip';
import { zhCN } from '@betteroffice/pptx-i18n';
import { resizeCursor } from '../interactions';
import { open } from './peerlessTestHarness';
import * as wasm from '../../../pptx/src/wasm/loader';

test('deleting every slide keeps document commands and history available', async () => {
  const saved = mock((_bytes: Uint8Array) => {});
  const value = await open({ onSave: saved });
  const access = await value.api.handleAsync();
  const slides = [...value.client.state.slides];
  for (const slide of slides) {
    await act(async () => { await access.deleteSlide(slide.id); });
  }
  expect(value.client.state.slides).toHaveLength(0);
  expect(value.view.getByTestId('pptx-worker-status').textContent).toBe('No slides');
  expect(value.view.queryByTestId('pptx-notes-textarea')).toBeNull();
  for (const command of ['undo', 'insertSlide', 'save'] as const) {
    expect(value.api.commands.getState(command).enabled).toBe(true);
  }
  for (const command of ['insertImage', 'exportPng', 'slideshow'] as const) {
    expect(value.api.commands.getState(command)).toMatchObject({ enabled: false, disabledReason: { code: 'slide-required' } });
  }
  await act(async () => { expect(await value.api.commands.execute('save', null)).toMatchObject({ ok: true }); });
  expect(saved).toHaveBeenCalledTimes(1);
  const empty = wasm.openPresentation(saved.mock.calls[0]![0]);
  try { expect(empty.snapshot().slides).toHaveLength(0); }
  finally { empty.dispose(); }
  await act(async () => { expect(await value.api.commands.execute('undo', null)).toMatchObject({ ok: true }); });
  expect(value.client.state.slides.map((slide) => slide.id)).toEqual([slides[1]!.id]);
  expect(value.api.commands.getState('redo').enabled).toBe(true);
  await act(async () => { expect(await value.api.commands.execute('redo', null)).toMatchObject({ ok: true }); });
  expect(value.client.state.slides).toHaveLength(0);
  await act(async () => { expect(await value.api.commands.execute('insertSlide', {})).toMatchObject({ ok: true }); });
  expect(value.client.state.slides).toHaveLength(1);
  expect(value.client.state.activeSlideId).toBe(value.client.state.slides[0]!.id);
  expect(value.view.queryByTestId('pptx-worker-status')).toBeNull();
  const recovered = wasm.openPresentation(await value.api.saveAsync());
  try { expect(recovered.snapshot().slides).toHaveLength(1); }
  finally { recovered.dispose(); }
});

test('object_slide_history', async () => {
  const value = await open();
  await act(async () => {
    await value.api.commands!.execute('tool', { value: 'shape:rect' });
  });
  await act(async () => {
    fireEvent.pointerDown(value.canvas, value.point(320, 200));
    fireEvent.pointerUp(value.stage, value.point(320, 200));
    await value.api.flushPendingInput();
  });
  expect(value.client.state.selection?.kind).toBe('object');
  await act(async () => {
    expect(await value.api.commands!.execute('shapeFill', { color: '#abcdef' })).toMatchObject({ ok: true });
    expect(await value.api.commands!.execute('shapeStrokeWidth', { points: 3 })).toMatchObject({ ok: true });
    expect(await value.api.commands!.execute('zOrder', { value: 'back' })).toMatchObject({ ok: true });
  });
  const access = await value.api.handleAsync();
  const snapshot = await access.snapshot();
  const shape = snapshot.slides[0]!.shapes.find((shape) => shape.id === value.client.state.selection?.shapeId)!;
  expect(shape.resolvedFillColor?.toLowerCase()).toBe('#abcdef');
  await act(async () => { expect(await value.api.commands!.execute('insertSlide', { layoutPartPath: null })).toMatchObject({ ok: true }); });
  expect(value.client.state.slides).toHaveLength(3);
  expect(value.client.state.activeSlideId).toBe(value.client.state.slides[1]!.id);
  await act(async () => { await value.api.commands!.execute('undo', null); });
  expect(value.client.state.slides).toHaveLength(2);
  await act(async () => { await value.api.commands!.execute('redo', null); });
  expect(value.client.state.slides).toHaveLength(3);
  const inserted = value.client.state.activeSlideId!;
  await act(async () => { await access.moveSlide(inserted, 2); });
  expect(value.client.state.slides[2]!.id).toBe(inserted);
  await act(async () => { await access.deleteSlide(inserted); });
  expect(value.client.state.slides).toHaveLength(2);
  expect(value.client.state.activeSlideId).toBe(value.client.state.slides[1]!.id);
  await act(async () => { await access.removeShape(snapshot.slides[0]!.id, shape.id); });
  expect((await access.snapshot()).slides[0]!.shapes.some((candidate) => candidate.id === shape.id)).toBe(false);
  await act(async () => { await value.api.commands!.execute('undo', null); });
  expect((await access.snapshot()).slides[0]!.shapes.some((candidate) => candidate.id === shape.id)).toBe(true);
  const saved = await value.api.saveAsync();
  expect(saved).toBeInstanceOf(Uint8Array);
  await act(async () => { value.view.rerender(value.ui(true)); await value.api.flushPendingInput(); });
  await expect(access.deleteSlide(value.client.state.slides[0]!.id)).rejects.toThrow();
  expect(value.view.queryByRole('button', { name: 'Delete slide' })).toBeNull();
});

test('hover cursor follows the confirmed geometry under the pointer', async () => {
  const value = await open();
  const line = value.box.lines[0]!;
  const word = { x: line.caretStops[2]!.x, y: line.y + line.height / 2 };
  const move = (x: number, y: number, buttons = 0, pointerId = 7) => {
    act(() => { fireEvent.pointerMove(value.stage, { clientX: x, clientY: y, pointerId, isPrimary: true, buttons }); });
    return value.canvas.style.cursor;
  };
  expect(value.canvas.style.cursor).toBe('default');
  expect(move(word.x, word.y)).toBe('text');
  expect(value.stage.style.cursor).toBe('');
  expect(move(value.box.x + 1, value.box.y + value.box.h / 2)).toBe('move');
  expect(move(value.list.width - 2, value.list.height - 2)).toBe('default');
  expect(move(word.x, word.y, 1)).toBe('default');
  const last = value.box.lines[value.box.lines.length - 1]!;
  expect(move(word.x, (last.y + last.height + value.box.y + value.box.h) / 2)).toBe('move');
  act(() => { fireEvent.pointerDown(value.canvas, value.point()); });
  expect(move(word.x, word.y, 0, 8)).toBe('move');
  act(() => { fireEvent.pointerUp(value.stage, value.point()); });
  expect(move(value.list.width - 2, value.list.height - 2)).toBe('default');
  await act(async () => { await value.api.commands!.execute('tool', { value: 'shape:rect' }); });
  expect(value.canvas.style.cursor).toBe('crosshair');
  expect(move(word.x, word.y)).toBe('crosshair');
});

async function deckWithMasterAndContentLayout(): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-officedocument.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slideMasters/slideMaster1.xml', `<p:sldMaster ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst></p:sldMaster>`);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${rel}/slideLayout" Target="../slideLayouts/slideLayout2.xml"/></Relationships>`);
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout ${ns} type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file('ppt/slideLayouts/slideLayout2.xml', `<p:sldLayout ${ns} type="obj"><p:cSld name="Title and Content"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="800000" y="450000"/><a:ext cx="10000000" cy="1100000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="800000" y="1750000"/><a:ext cx="10000000" cy="4400000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="t"/><a:lstStyle/></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout2.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`);
  return zip.generateAsync({ type: 'uint8array' });
}

test('insertSlide picks Title and Content layout, shows placeholder guides, supports typing, and undoes in one step', async () => {
  const bytes = await deckWithMasterAndContentLayout();
  const en = await open({ file: bytes });
  expect(en.api.commands.getState('insertSlide').options?.map((o) => o.label)).toEqual(['Blank', 'Title and Content']);
  await act(async () => { expect(await en.api.commands.execute('insertSlide', {})).toMatchObject({ ok: true }); });
  expect(en.client.state.slides[1]?.layoutPartPath).toBe('ppt/slideLayouts/slideLayout2.xml');
  const enAccess = await en.api.handleAsync();
  const enSnap = await enAccess.snapshot();
  const inserted = enSnap.slides[1]!;
  expect(inserted.shapes).toHaveLength(2);
  expect(inserted.shapes[0]!.placeholder).toMatchObject({ placeholderType: 'title' });
  expect(inserted.shapes[1]!.placeholder).toMatchObject({ index: 1 });

  await waitFor(() => expect(en.view.getAllByTestId('pptx-placeholder-guide')).toHaveLength(2));
  expect(en.view.getAllByTestId('pptx-placeholder-guide').map((el) => el.textContent))
    .toEqual(['Click to add title', 'Click to add text']);

  const titleStory = inserted.shapes[0]!.textStories[0]!;
  const input = en.view.getByRole('textbox', { name: 'Slide text' });
  await act(async () => {
    await en.api.selectText({ slide: 2, shapeId: inserted.shapes[0]!.id, storyId: titleStory.id, start: 0, end: 0 });
    input.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertText', data: 'Quarterly Review', bubbles: true, cancelable: true }));
    await en.api.flushPendingInput();
  });
  await waitFor(() => expect(en.view.getAllByTestId('pptx-placeholder-guide').map((el) => el.textContent))
    .toEqual(['Click to add text']));

  const saved = wasm.openPresentation(await en.api.saveAsync());
  try {
    expect(saved.snapshot().slides[1]!.shapes[0]!.placeholder).toMatchObject({ placeholderType: 'title' });
    expect(saved.snapshot().slides[1]!.shapes[0]!.textStories[0]!.paragraphs[0]!.runs.map((r) => r.text).join('')).toBe('Quarterly Review');
  } finally { saved.dispose(); }

  await act(async () => { expect(await en.api.commands.execute('undo', null)).toMatchObject({ ok: true }); });
  await act(async () => { expect(await en.api.commands.execute('undo', null)).toMatchObject({ ok: true }); });
  expect(en.client.state.slides).toHaveLength(1);

  const zh = await open({ file: bytes, i18n: zhCN });
  expect(zh.api.commands.getState('insertSlide').options?.map((o) => o.label)).toEqual(['空白', '标题和内容']);
  await act(async () => { expect(await zh.api.commands.execute('insertSlide', {})).toMatchObject({ ok: true }); });
  await waitFor(() => expect(zh.view.getAllByTestId('pptx-placeholder-guide')).toHaveLength(2));
  expect(zh.view.getAllByTestId('pptx-placeholder-guide').map((el) => el.textContent))
    .toEqual(['单击此处添加标题', '单击此处添加文本']);
});

test('the capturing stage shows the press cursor until the press ends', async () => {
  const value = await open();
  const line = value.box.lines[0]!;
  const word = value.point(line.caretStops[2]!.x, line.y + line.height / 2);
  act(() => { fireEvent.pointerDown(value.canvas, { ...word, buttons: 1 }); });
  expect(value.stage.setPointerCapture).toHaveBeenCalled();
  expect(value.stage.style.cursor).toBe('text');
  act(() => { fireEvent.pointerUp(value.stage, word); });
  expect(value.stage.style.cursor).toBe('');
  expect(value.canvas.style.cursor).toBe('text');
  await value.select();
  const corner = value.point(value.box.x + value.box.w, value.box.y + value.box.h);
  act(() => { fireEvent.pointerDown(value.view.getByTestId('pptx-resize-se'), { ...corner, buttons: 1 }); });
  expect(value.stage.style.cursor).toBe(resizeCursor('se'));
  act(() => { fireEvent.pointerCancel(value.stage, corner); });
  expect(value.stage.style.cursor).toBe('');
  await act(async () => { await value.api.commands!.execute('tool', { value: 'shape:rect' }); });
  act(() => { fireEvent.pointerDown(value.view.getByTestId('pptx-resize-se'), { ...corner, buttons: 1 }); });
  expect(value.stage.style.cursor).toBe('crosshair');
  act(() => { fireEvent.lostPointerCapture(value.stage, corner); });
  expect(value.stage.style.cursor).toBe('');
});

test('moving over the stage clears the press cursor after Escape releases capture', async () => {
  const value = await open();
  const body = value.point();
  act(() => { fireEvent.pointerDown(value.canvas, { ...body, buttons: 1 }); });
  expect(value.stage.hasPointerCapture(7)).toBe(true);
  expect(value.stage.style.cursor).toBe('move');
  act(() => { fireEvent.pointerMove(value.stage, { ...body, buttons: 1 }); });
  expect(value.stage.style.cursor).toBe('move');
  act(() => { fireEvent.keyDown(value.stage, { key: 'Escape' }); });
  expect(value.stage.hasPointerCapture(7)).toBe(false);
  expect(value.stage.style.cursor).toBe('move');
  act(() => { fireEvent.pointerMove(value.stage, { ...body, buttons: 0 }); });
  expect(value.stage.style.cursor).toBe('');
});

test('leaving the stage clears the press cursor after capture is released', async () => {
  const value = await open();
  const body = value.point();
  act(() => { fireEvent.pointerDown(value.canvas, { ...body, buttons: 1 }); });
  expect(value.stage.hasPointerCapture(7)).toBe(true);
  value.stage.releasePointerCapture(7);
  expect(value.stage.hasPointerCapture(7)).toBe(false);
  expect(value.stage.style.cursor).toBe('move');
  act(() => { fireEvent.pointerLeave(value.stage, { ...body, buttons: 0 }); });
  expect(value.stage.style.cursor).toBe('');
});

test('a press resolves its cursor at the press point', async () => {
  const value = await open();
  const line = value.box.lines[0]!;
  const word = value.point(line.caretStops[2]!.x, line.y + line.height / 2);
  act(() => { fireEvent.pointerMove(value.canvas, { ...word, buttons: 0 }); });
  expect(value.canvas.style.cursor).toBe('text');
  const body = value.point(value.box.x + 1, value.box.y + value.box.h / 2);
  act(() => { fireEvent.pointerDown(value.canvas, { ...body, buttons: 1 }); });
  expect(value.stage.style.cursor).toBe('move');
  act(() => { fireEvent.pointerUp(value.stage, body); });
});

test('right-clicking text leaves the stage cursor unset', async () => {
  const value = await open();
  const line = value.box.lines[0]!;
  const word = value.point(line.caretStops[2]!.x, line.y + line.height / 2);
  act(() => { fireEvent.pointerDown(value.canvas, { ...word, button: 2, buttons: 2 }); });
  expect(value.stage.style.cursor).toBe('');
});

test('pressing empty slide space leaves the stage cursor unset', async () => {
  const value = await open();
  const empty = value.point(value.list.width - 2, value.list.height - 2);
  act(() => { fireEvent.pointerDown(value.canvas, { ...empty, buttons: 1 }); });
  expect(value.stage.setPointerCapture).not.toHaveBeenCalled();
  expect(value.stage.style.cursor).toBe('');
});

test('another pointer leaves the captured press cursor in place', async () => {
  const value = await open();
  const line = value.box.lines[0]!;
  const word = value.point(line.caretStops[2]!.x, line.y + line.height / 2);
  act(() => { fireEvent.pointerDown(value.canvas, { ...word, buttons: 1 }); });
  expect(value.stage.hasPointerCapture(7)).toBe(true);
  expect(value.stage.style.cursor).toBe('text');
  const second = { ...value.point(value.list.width - 2, value.list.height - 2), pointerId: 8, isPrimary: false };
  act(() => { fireEvent.pointerDown(value.stage, { ...second, buttons: 1 }); });
  expect(value.stage.style.cursor).toBe('text');
  act(() => { fireEvent.pointerUp(value.stage, { ...second, buttons: 0 }); });
  expect(value.stage.style.cursor).toBe('text');
  act(() => { fireEvent.pointerUp(value.stage, { ...word, buttons: 0 }); });
  expect(value.stage.style.cursor).toBe('');
});

test('a resize handle keeps the resize cursor off the canvas', async () => {
  const value = await open();
  await value.select();
  const corner = value.point(value.box.x + value.box.w, value.box.y + value.box.h);
  const handle = value.view.getByTestId('pptx-resize-se');
  act(() => { fireEvent.pointerMove(handle, { ...corner, buttons: 0 }); });
  expect(value.canvas.style.cursor).not.toContain('resize');
  act(() => { fireEvent.pointerDown(handle, { ...corner, buttons: 1 }); });
  expect(value.stage.style.cursor).toBe(resizeCursor('se'));
  act(() => { fireEvent.pointerCancel(value.stage, { ...corner, buttons: 0 }); });
  expect(value.canvas.style.cursor).not.toContain('resize');
});
