import { afterEach, expect, spyOn, test } from 'bun:test';
import { applyFrameDelta, decodeFrameDelta } from '../layout/render/frameDelta';
import type { YrsStorySegment } from './index';
import { closeCommandEngines, openCommandEngine } from './residentCommandTestSupport';
import { readControlHandles } from './residentControls';
import { YrsPositionProjection } from './yrsPositionProjection';
afterEach(closeCommandEngines);

const checkbox = (id: number, tag: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:tag w:val="${tag}"/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>`;

test('duplicate authored widgets expose a native location for each physical handle', async () => {
  const widget = checkbox(7, 'same');
  const { engine, resident } = await openCommandEngine([], { bodyXml: '<w:p w14:paraId="00000001">'+widget+'<w:r><w:t>gap</w:t></w:r>'+widget+'</w:p>' });
  const snapshot = readControlHandles(resident);
  expect(snapshot.controls).toHaveLength(2);
  expect(new Set(snapshot.controls.map(control => control.handle)).size).toBe(2);
  expect(snapshot.controls.map(control => control.control.ooxmlId)).toEqual(['7', '7']);
  expect(snapshot.controls.map(target => target.location)).toEqual([
    { story: 'body', index: 0, rootStory: 'body', position: 1 },
    { story: 'body', index: 4, rootStory: 'body', position: 7 },
  ]);
  expect(snapshot.controls.map(target => target.widget?.checked)).toEqual([false, false]);
  const frame = applyFrameDelta(null, decodeFrameDelta(engine.buildDisplayListFrame('{}', 0)));
  const published = new Set(frame.pages.flatMap(({ page }) => page.primitives.flatMap(primitive =>
    [primitive.inlineSdtWidget?.groupId, primitive.sdt?.groupId, ...(primitive.sdtPath ?? []).map(sdt => sdt.groupId)])
    .filter(group => group?.startsWith('sdt@'))));
  const painted = readControlHandles(resident, frame, 3);
  expect(painted.frameEpoch).toBe(frame.frameEpoch);
  expect(published).toEqual(new Set(painted.controls.map(target => `sdt@${target.location.position}`)));
  const rects = painted.controls.map(target => target.rects);
  expect(rects.map(entries => entries.length)).toEqual([1, 1]);
  expect(rects[0]![0]!.x).toBeLessThan(rects[1]![0]!.x);
}, 120000);

test('controls after other inline controls read the published sdt groups in paragraphs, blocks and cells', async () => {
  const { engine, resident } = await openCommandEngine([], { bodyXml:
    '<w:p w14:paraId="00000001">' + checkbox(1, 'first') + '<w:r><w:t>a</w:t></w:r></w:p>' +
    '<w:p w14:paraId="00000002">' + checkbox(2, 'second') + '<w:r><w:t>b</w:t></w:r></w:p>' +
    '<w:sdt><w:sdtPr><w:id w:val="3"/><w:tag w:val="block"/></w:sdtPr><w:sdtContent><w:p w14:paraId="00000003"><w:r><w:t>block</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
    '<w:p w14:paraId="00000004"><w:r><w:t>tail</w:t></w:r>' + checkbox(4, 'last') + '</w:p>' +
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="00000005">' + checkbox(5, 'cell') +
    '</w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="00000006"/>' });
  const frame = applyFrameDelta(null, decodeFrameDelta(engine.buildDisplayListFrame('{}', 0)));
  const published = new Set(frame.pages.flatMap(({ page }) => page.primitives.flatMap(primitive =>
    [primitive.inlineSdtWidget?.groupId, primitive.sdt?.groupId, ...(primitive.sdtPath ?? []).map(sdt => sdt.groupId)])
    .filter(group => group?.startsWith('sdt@'))));
  const controls = readControlHandles(resident, frame, 1).controls;
  expect(controls.map(target => target.control.ooxmlId)).toEqual(['1', '2', '3', '4', '5']);
  expect(published).toEqual(new Set(controls.map(target => `sdt@${target.location.position}`)));
  expect(controls.filter(target => target.widget).map(target => target.rects.length)).toEqual([1, 1, 1, 1]);
}, 120000);

test.todo('engine resident geometry counts an inline control at its layout size and lists block control stories', async () => {
  const { resident } = await openCommandEngine([], { bodyXml:
    '<w:p w14:paraId="00000001">' + checkbox(1, 'first') + '</w:p><w:p w14:paraId="00000002">' + checkbox(2, 'second') + '</w:p>' +
    '<w:sdt><w:sdtPr><w:id w:val="3"/></w:sdtPr><w:sdtContent><w:p w14:paraId="00000003"><w:r><w:t>block</w:t></w:r></w:p></w:sdtContent></w:sdt>' });
  const controls = readControlHandles(resident).controls;
  const second = resident.geometryReader.paragraphSpans('body')[1]!.paraId;
  const geometry = resident.storyGeometry('body', [[0, 1_000]]);
  expect(geometry.paragraphs.find(paragraph => paragraph.paraId === second)?.from).toBe(controls[1]!.location.position!);
  const block = resident.geometryReader.storyIds().find(story => story.startsWith('body:sdt'));
  expect(geometry.paragraphs.some(paragraph => paragraph.story === block)).toBe(true);
});

test('block control locations come from one walk over the projected nodes', async () => {
  const count = 24;
  const blocks = Array.from({ length: count }, (_, index) =>
    `<w:sdt><w:sdtPr><w:id w:val="${index + 1}"/><w:tag w:val="block-${index}"/></w:sdtPr><w:sdtContent>` +
    `<w:p w14:paraId="${(index + 16).toString(16).padStart(8, '0')}"><w:r><w:t>Block ${index}</w:t></w:r></w:p></w:sdtContent></w:sdt>`).join('');
  const { resident } = await openCommandEngine([], { bodyXml: blocks + '<w:p w14:paraId="00000001"><w:r><w:t>Tail</w:t></w:r></w:p>' });
  const nodeAt = spyOn(YrsPositionProjection.prototype, 'nodeAt');
  try {
    const positions = readControlHandles(resident).controls.map(target => target.location.position);
    expect(positions).toHaveLength(count);
    expect(positions.every((position, index) => position !== null && (index === 0 || position > positions[index - 1]!))).toBe(true);
    expect(nodeAt.mock.calls.length).toBeLessThanOrEqual(2 * count + 4);
  } finally { nodeAt.mockRestore(); }
}, 120000);

test('a malformed control reads as null while the other controls resolve', async () => {
  const { resident } = await openCommandEngine([], {
    bodyXml: '<w:p w14:paraId="00000001">' + checkbox(1, 'good') + '<w:r><w:t>gap</w:t></w:r>' + checkbox(2, 'json') + checkbox(3, 'deep') + '</w:p>',
  });
  let deep: unknown[] = [{ kind: 'text', text: '☐' }];
  for (let depth = 0; depth < 300; depth++) deep = [{ kind: 'sdt', payload: { content: deep } }];
  const malformed: Array<(payload: Record<string, unknown>) => Record<string, unknown>> = [
    payload => payload,
    payload => ({ ...payload, propertiesJson: '{' }),
    payload => ({ ...payload, content: deep }),
  ];
  const segments = resident.geometryReader.storySegments;
  const read = spyOn(resident.geometryReader, 'storySegments').mockImplementation(story => {
    let ordinal = 0;
    return segments(story).map((segment): YrsStorySegment => segment.kind === 'embed' && segment.embedKind === 'sdt'
      ? { ...segment, payload: (malformed[ordinal++] ?? (payload => payload))(segment.payload) } : segment);
  });
  try {
    const controls = readControlHandles(resident).controls;
    expect(controls.map(target => target.widget)).toEqual([{ kind: 'checkbox', checked: false }, null, null]);
    expect(controls.map(target => target.location.position)).toEqual([1, 7, null]);
  } finally { read.mockRestore(); }
}, 120000);
