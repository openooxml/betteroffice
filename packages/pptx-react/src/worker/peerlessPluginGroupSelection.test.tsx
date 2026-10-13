import { deck, open } from './peerlessTestHarness';
import JSZip from 'jszip';
import { expect, test } from 'bun:test';
import { definePptxPlugin } from '../plugins/definePptxPlugin';
import type { PptxPluginContext } from '../plugins/types';

const { act, waitFor } = await import('@testing-library/react');

test('plugin navigation selects a group child without enabling parent mutations', async () => {
  const zip = await JSZip.loadAsync(await deck());
  const path = 'ppt/slides/slide1.xml';
  const original = await zip.file(path)!.async('string');
  const shape = original.match(/<p:sp>.*?<\/p:sp>/)![0];
  const child = shape.replace('id="2" name="Text"', 'id="4" name="Child"');
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="5143500"/><a:chOff x="0" y="0"/><a:chExt cx="9144000" cy="5143500"/></a:xfrm></p:grpSpPr>${child}</p:grpSp>`;
  zip.file(path, original.replace('</p:spTree>', `${group}</p:spTree>`));
  let context: PptxPluginContext<null> | undefined;
  const plugin = definePptxPlugin({ id: 'child', createState: () => null, onEvent(next) { context = next; } });
  const value = await open({ file: await zip.generateAsync({ type: 'uint8array' }), plugins: [plugin] });
  await waitFor(() => expect(context).toBeDefined());
  const snapshot = await (await value.api.handleAsync()).snapshot();
  const slide = snapshot.slides[0]!;
  const grouped = slide.shapes.find((shape) => shape.name === 'Group')!;
  const target = { slideId: slide.id, shapeId: grouped.children[0]!.id };
  await act(async () => {
    expect(await context!.navigation.selectShape(target, { expectVersion: await value.api.version() })).toMatchObject({ ok: true });
  });
  expect(value.client.state.selection).toMatchObject({ kind: 'object', shapeId: target.shapeId });
  expect(value.client.state.selection).toMatchObject({ canMove: false, canResize: false });
  expect(value.client.state.editState.object).toBeNull();
  expect(value.api.commands!.getState('shapeFill', { color: '#ff0000' }).enabled).toBe(false);
  expect(value.api.commands!.getState('zOrder', { value: 'front' }).enabled).toBe(false);
});
