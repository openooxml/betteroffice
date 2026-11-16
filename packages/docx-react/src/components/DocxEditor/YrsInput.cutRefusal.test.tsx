import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { createRef } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { YrsInput, type YrsInputRef } from './YrsInput';
import type { DocxCutRefusedError } from './internals/cutRefused';
import { yrsCellStory } from './yrsCommands';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
const sessions: YrsSession[] = [];
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const BODY = '<w:p><w:r><w:t>ALPHA</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t xml:space="preserve">BRAVO before </w:t></w:r><w:r><w:br w:type="column"/></w:r><w:r><w:t xml:space="preserve"> after BRAVO</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>CHARLIE</w:t></w:r></w:p>';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(
  import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
)))), 30_000);
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
  delete (navigator as { clipboard?: unknown }).clipboard;
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function mount() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${BODY}<w:sectPr><w:cols w:num="2" w:space="720"/></w:sectPr></w:body></w:document>`);
  const session = await createYrsSession();
  sessions.push(session);
  session.openDocx(await zip.generateAsync({ type: 'uint8array' }), true);
  const written: string[] = [];
  Object.defineProperty(navigator, 'clipboard', {
    configurable: true,
    value: { writeText: async (text: string) => { written.push(text); } },
  });
  const refusals: DocxCutRefusedError[] = [];
  const input = createRef<YrsInputRef>();
  const map = (story = 'body') => createYrsInputPositionMap(story, session.paragraphSpans(story));
  const view = render(
    <YrsInput
      ref={input}
      enabled
      readOnly={false}
      seedSelection={false}
      session={session}
      inputPositionMap={map}
      displayPositionToLoc={(position, story) => displayPositionToYrsLoc(map(story), position)}
      locToDisplayPosition={loc => yrsLocToDisplayPosition(map(loc.story), loc)}
      onStateChange={() => {}}
      onDirectInput={() => {}}
      onCutRefused={error => refusals.push(error as DocxCutRefusedError)}
    />
  );
  const textarea = view.getByTestId('yrs-input');
  const flush = () => act(async () => { await input.current!.flushPendingInput(); });
  const extendThenShiftDelete = async () => {
    fireEvent.keyDown(textarea, { key: 'ArrowRight', shiftKey: true });
    fireEvent.keyDown(textarea, { key: 'Delete', shiftKey: true });
    await flush();
  };
  return { session, written, refusals, textarea, flush, extendThenShiftDelete };
}

test('a Shift+Delete pressed before a queued selection onto a column break paragraph settles is refused and deletes nothing', async () => {
  const { session, written, refusals, extendThenShiftDelete } = await mount();
  const bravo = session.paragraphs('body').find(paragraph => paragraph.text.startsWith('BRAVO'))!;
  session.setSelection({ story: 'body', paraId: bravo.paraId, offset: 'BRAVO before '.length });
  const before = session.storySegments('body');
  await extendThenShiftDelete();
  expect(session.storySegments('body')).toEqual(before);
  expect(refusals).toHaveLength(1);
  expect(refusals[0]!.leftOut).toContain('columnBreak');
  expect(written).toEqual([]);
});

test('a Shift+Delete pressed before a queued selection onto plain text settles cuts that selection', async () => {
  const { session, written, refusals, extendThenShiftDelete } = await mount();
  const charlie = session.paragraphs('body').find(paragraph => paragraph.text === 'CHARLIE')!;
  session.setSelection({ story: 'body', paraId: charlie.paraId, offset: 1 });
  await extendThenShiftDelete();
  expect(written).toEqual(['C']);
  expect(session.paragraphs('body').find(paragraph => paragraph.paraId === charlie.paraId)?.text).toBe('HARLIE');
  expect(refusals).toEqual([]);
});

test('a Shift+Delete whose queued selection holds only the hidden column-break unit writes nothing and deletes nothing', async () => {
  const { session, written, refusals, extendThenShiftDelete } = await mount();
  const charlie = session.paragraphs('body').find(paragraph => paragraph.text === 'CHARLIE')!;
  session.setSelection({ story: 'body', paraId: charlie.paraId, offset: 0 });
  const before = session.storySegments('body');
  await extendThenShiftDelete();
  expect(written).toEqual([]);
  expect(session.storySegments('body')).toEqual(before);
  expect(refusals).toEqual([]);
});

test('a Shift+Delete pressed in a table cell while input is pending deletes the next character', async () => {
  const { session, textarea, flush } = await mount();
  const alpha = session.paragraphs('body').find(paragraph => paragraph.text === 'ALPHA')!;
  const { table } = session.insertTable({ story: 'body', paraId: alpha.paraId, offset: 'ALPHA'.length }, 1, 2);
  const first = yrsCellStory(session, { ...table, row: 0, column: 0 })!;
  const second = yrsCellStory(session, { ...table, row: 0, column: 1 })!;
  session.insertText({ story: second, paraId: session.paragraphs(second)[0]!.paraId, offset: 0 }, 'ABC');
  session.setSelection({ story: first, paraId: session.paragraphs(first)[0]!.paraId, offset: 0 });
  fireEvent.keyDown(textarea, { key: 'Tab' });
  fireEvent.keyDown(textarea, { key: 'ArrowRight' });
  fireEvent.keyDown(textarea, { key: 'Delete', shiftKey: true });
  await flush();
  expect(session.paragraphs(second)[0]!.text).toBe('AC');
});

test('a Shift+Delete over a range with no clipboard available deletes nothing', async () => {
  const { session, written, refusals, extendThenShiftDelete } = await mount();
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined });
  const charlie = session.paragraphs('body').find(paragraph => paragraph.text === 'CHARLIE')!;
  session.setSelection({ story: 'body', paraId: charlie.paraId, offset: 0 });
  const before = session.storySegments('body');
  await extendThenShiftDelete();
  expect(session.storySegments('body')).toEqual(before);
  expect(written).toEqual([]);
  expect(refusals).toEqual([]);
});
