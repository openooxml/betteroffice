import { expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import { decodeCommandDisplayList, presentationSharedInputs } from '@betteroffice/pptx/peerless';
import type { ShapeRect, TextBoxPrimitive } from '@betteroffice/pptx';
import { open, replyControl } from './peerlessTestHarness';
import { peerlessPainterFactories } from './usePeerlessPresentation';

const { act, fireEvent, waitFor } = await import('@testing-library/react');

function outline(view: Awaited<ReturnType<typeof open>>['view']) {
  const elements = view.getAllByTestId('pptx-text-edit-outline');
  expect(elements).toHaveLength(1);
  expect(view.queryAllByTestId(/pptx-resize-/)).toHaveLength(0);
  const element = elements[0]!;
  expect(element.style.border).toBe('2px dashed #6b7280');
  expect(element.style.boxShadow).toBe('0 0 0 1px rgba(255, 255, 255, 0.9)');
  return element;
}

function bounds(element: HTMLElement, rect: ShapeRect, scale: number) {
  for (const [property, value] of [['left', rect.x], ['top', rect.y],
    ['width', rect.width], ['height', rect.height]] as const) {
    expect(parseFloat(element.style[property])).toBeCloseTo(value * scale, 3);
  }
}

async function editExisting(h: Awaited<ReturnType<typeof open>>) {
  const line = h.box.lines[0]!;
  const at = h.point(line.caretStops[2]!.x, line.y + line.height / 2);
  act(() => { fireEvent.pointerDown(h.canvas, at); fireEvent.pointerUp(h.stage, at); });
  await waitFor(() => expect(h.client.state.selection).toMatchObject({ shapeId: h.box.shapeId, anchor: 2, focus: 2 }));
  await waitFor(() => outline(h.view));
}

for (const [start, end] of [[[200.123, 150.456], [400.789, 220.123]],
  [[900.123, 500.456], [980.789, 550.123]], [[200, 150], [201, 151]]] as const) {
  test(`textbox outline survives pointerup at ${start.join(',')}`, async () => {
    const h = await open();
    const scale = parseFloat(h.canvas.style.width) / h.list.width;
    let rect: ShapeRect | undefined;
    const input = presentationSharedInputs.get(h.client)!;
    const submit = input.submit.bind(input);
    const capture = spyOn(input, 'submit').mockImplementation((intent, basis, after) => {
      if (intent.kind === 'insertTextBox') rect = intent.payload.draft.rect;
      return submit(intent, basis, after);
    });
    try {
      fireEvent.click(h.view.getByRole('button', { name: 'Text box' }));
      act(() => {
        fireEvent.pointerDown(h.canvas, h.point(...start));
        fireEvent.pointerMove(h.stage, h.point(...end));
      });
      expect(h.view.getByTestId('pptx-insertion-preview')).toBeDefined();
      replyControl.hold();
      act(() => { fireEvent.pointerUp(h.stage, h.point(...end)); });
      await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
      expect(h.view.queryByTestId('pptx-insertion-preview') === null).toBe(true);
      expect(rect).toBeDefined();
      const echo = outline(h.view);
      bounds(echo, { x: rect!.x * h.list.width / h.client.state.size.width,
        y: rect!.y * h.list.height / h.client.state.size.height,
        width: rect!.width * h.list.width / h.client.state.size.width,
        height: rect!.height * h.list.height / h.client.state.size.height }, scale);
      const echoStyle = echo.getAttribute('style');
      await act(async () => { replyControl.release(); });
      await waitFor(() => expect(h.client.state.selection).toMatchObject({ anchor: 0, focus: 0 }));
      await waitFor(() => {
        const selected = h.client.state.selection!;
        const box = decodeCommandDisplayList(h.client.state.frame!).primitives.find((p) =>
          p.kind === 'textBox' && p.shapeId === selected.shapeId) as TextBoxPrimitive;
        expect(box).toBeDefined();
        bounds(outline(h.view), { x: box.x, y: box.y, width: box.w, height: box.h }, scale);
        expect(outline(h.view).getAttribute('style')).toBe(echoStyle);
        const caret = h.view.getByTestId('pptx-peerless-caret');
        expect(caret.style.visibility).toBe('visible');
        expect(parseFloat(caret.style.left)).toBeGreaterThanOrEqual(box.x);
        expect(parseFloat(caret.style.left)).toBeLessThan(box.x + box.w);
      });
      fireEvent.click(h.view.getByTestId('pptx-undo'));
      await waitFor(() => expect(h.view.queryByTestId('pptx-text-edit-outline') === null).toBe(true));
    } finally { capture.mockRestore(); replyControl.release(); }
  });
}

test('existing text editing survives Escape and leaves with object selection and slide navigation', async () => {
  const h = await open();
  await editExisting(h);
  bounds(outline(h.view), { x: h.box.x, y: h.box.y, width: h.box.w, height: h.box.h },
    parseFloat(h.canvas.style.width) / h.list.width);
  await h.select();
  expect(h.view.queryByTestId('pptx-text-edit-outline') === null).toBe(true);
  await editExisting(h);
  const selection = { ...h.client.state.selection! };
  const style = outline(h.view).getAttribute('style');
  const submit = spyOn(presentationSharedInputs.get(h.client)!, 'submit');
  try {
    await act(async () => {
      fireEvent.keyDown(h.view.getByRole('textbox', { name: 'Slide text' }), { key: 'Escape' });
      await new Promise<void>((resolve) => setImmediate(resolve));
    });
    expect(submit).not.toHaveBeenCalled();
    expect(h.client.state.selection).toEqual(selection);
    expect(outline(h.view).getAttribute('style')).toBe(style);
  } finally { submit.mockRestore(); }
  fireEvent.click(h.view.getByRole('button', { name: '2' }));
  await waitFor(() => expect(h.client.state.activeSlideId).toBe(h.client.state.slides[1]!.id));
  await waitFor(() => expect(h.view.queryByTestId('pptx-text-edit-outline') === null).toBe(true));
});

test('Escape cancels a textbox drag preview without inserting', async () => {
  const h = await open();
  const submit = spyOn(presentationSharedInputs.get(h.client)!, 'submit');
  try {
    fireEvent.click(h.view.getByRole('button', { name: 'Text box' }));
    act(() => {
      fireEvent.pointerDown(h.canvas, h.point(200, 150));
      fireEvent.pointerMove(h.stage, h.point(400, 220));
    });
    expect(h.view.queryByTestId('pptx-insertion-preview') !== null).toBe(true);
    fireEvent.keyDown(h.stage, { key: 'Escape' });
    expect(h.view.queryByTestId('pptx-insertion-preview') === null).toBe(true);
    expect(h.view.queryByTestId('pptx-text-edit-outline') === null).toBe(true);
    act(() => { fireEvent.pointerUp(h.stage, h.point(400, 220)); });
    expect(submit).not.toHaveBeenCalled();
  } finally { submit.mockRestore(); }
});

for (const change of ['failure', 'tool', 'slide', 'readOnly', 'replacement', 'Escape'] as const) {
  test(`pending textbox outline clears on ${change}`, async () => {
    const h = await open();
    const input = presentationSharedInputs.get(h.client)!;
    const submit = input.submit.bind(input);
    const refusal = spyOn(input, 'submit').mockImplementation((intent, basis, after) => submit(
      change === 'failure' && intent.kind === 'insertTextBox'
        ? { ...intent, payload: { ...intent.payload, draft: { ...intent.payload.draft, text: '\u0000' } } } : intent,
      basis, after));
    try {
      fireEvent.click(h.view.getByRole('button', { name: 'Text box' }));
      replyControl.hold();
      act(() => {
        fireEvent.pointerDown(h.canvas, h.point(200, 150));
        fireEvent.pointerMove(h.stage, h.point(400, 220));
        fireEvent.pointerUp(h.stage, h.point(400, 220));
      });
      await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
      outline(h.view);
      if (change === 'tool') await act(async () => { expect(await h.api.commands.execute('tool', { value: 'textBox' })).toMatchObject({ ok: true }); });
      if (change === 'slide') fireEvent.click(h.view.getByRole('button', { name: '2' }));
      if (change === 'readOnly') h.view.rerender(h.ui(true));
      if (change === 'replacement') h.view.rerender(<div />);
      if (change === 'Escape') fireEvent.keyDown(h.view.getByRole('textbox', { name: 'Slide text' }), { key: 'Escape' });
      if (change === 'tool' || change === 'readOnly' || change === 'replacement' || change === 'Escape') {
        await waitFor(() => expect(h.view.queryByTestId('pptx-text-edit-outline') === null).toBe(true));
      }
      if (change === 'Escape') expect(h.view.queryByTestId('pptx-insertion-preview') === null).toBe(true);
      await act(async () => { replyControl.release(); });
      if (change === 'failure') await waitFor(() => expect(h.client.state.outcomes).toEqual(
        expect.arrayContaining([expect.objectContaining({ status: 'refused' })])));
      if (change === 'failure' || change === 'slide') {
        await waitFor(() => expect(h.view.queryByTestId('pptx-text-edit-outline') === null).toBe(true));
      }
      if (change === 'Escape') await waitFor(() => {
        expect(h.client.state.selection).toMatchObject({ anchor: 0, focus: 0 });
        outline(h.view);
      });
    } finally { refusal.mockRestore(); replyControl.release(); }
  });
}

test('existing rotated and flipped text uses the displayed shape bounds', async () => {
  const seed = await open();
  const zip = await JSZip.loadAsync(seed.bytes);
  const source = await zip.file('ppt/slides/slide1.xml')!.async('string');
  zip.file('ppt/slides/slide1.xml', source.replace('<a:xfrm>', '<a:xfrm rot="1800000" flipH="1">'));
  seed.view.unmount();
  await seed.client.dispose();
  const h = await open({ file: await zip.generateAsync({ type: 'uint8array' }) });
  await act(async () => { await h.client.dispatch({ kind: 'setSelection', payload: {
    selection: { slideId: h.client.state.slides[0]!.id, shapeId: h.box.shapeId!, storyId: h.box.storyId!,
      anchor: 2, focus: 2, affinity: 'downstream' }, inputTime: 1, undoGroup: 'selection' } }); });
  await waitFor(() => expect(h.client.state.selection).toMatchObject({ anchor: 2, focus: 2 }));
  const angle = Math.PI / 6;
  const width = Math.abs(h.box.w * Math.cos(angle)) + Math.abs(h.box.h * Math.sin(angle));
  const height = Math.abs(h.box.w * Math.sin(angle)) + Math.abs(h.box.h * Math.cos(angle));
  await waitFor(() => bounds(outline(h.view), { x: h.box.x + (h.box.w - width) / 2,
    y: h.box.y + (h.box.h - height) / 2, width, height }, parseFloat(h.canvas.style.width) / h.list.width));
});

test('textbox echo waits for display and yields to changed confirmed bounds', async () => {
  const factory = peerlessPainterFactories.frame;
  let holding = false;
  let paint: (() => void) | undefined;
  const painter = spyOn(peerlessPainterFactories, 'frame').mockImplementation((canvas, options) => {
    const value = factory(canvas, options);
    return { ...value, paintFrame(frame, view, media) {
      if (holding && canvas.dataset.testid === 'pptx-slide-canvas') {
        paint = () => value.paintFrame(frame, view, media);
      } else value.paintFrame(frame, view, media);
    } };
  });
  let restore = () => {};
  try {
    const h = await open();
    const input = presentationSharedInputs.get(h.client)!;
    const submit = input.submit.bind(input);
    const capture = spyOn(input, 'submit').mockImplementation((intent, basis, after) => submit(
      intent.kind === 'insertTextBox' ? { ...intent, payload: { ...intent.payload, draft: { ...intent.payload.draft,
        rect: { ...intent.payload.draft.rect, width: intent.payload.draft.rect.width + 952500 } } } } : intent, basis, after));
    restore = () => capture.mockRestore();
    holding = true;
    fireEvent.click(h.view.getByRole('button', { name: 'Text box' }));
    act(() => {
      fireEvent.pointerDown(h.canvas, h.point(200, 150));
      fireEvent.pointerMove(h.stage, h.point(400, 220));
      fireEvent.pointerUp(h.stage, h.point(400, 220));
    });
    await waitFor(() => expect(h.client.state.selection).toMatchObject({ anchor: 0, focus: 0 }));
    await waitFor(() => expect(paint).toBeDefined());
    const scale = parseFloat(h.canvas.style.width) / h.list.width;
    bounds(outline(h.view), { x: 200, y: 150, width: 200, height: 70 }, scale);
    holding = false;
    await act(async () => { paint!(); });
    await waitFor(() => bounds(outline(h.view), { x: 200, y: 150, width: 300, height: 70 }, scale));
  } finally { restore(); painter.mockRestore(); }
});
