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

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
const sessions: YrsSession[] = [];
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const table = (prefix: string) => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${[1, 2].map(row => `<w:tr>${['A', 'B'].map(column => `<w:tc>${paragraph(`${prefix}-${column}${row}`)}</w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>`;

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(
  import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
)))));
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function mount(between: boolean) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  const body = paragraph('Paragraph before the table.') + table('T1') +
    (between ? paragraph('') + table('T2') : '') + paragraph('Paragraph after the table.');
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const session = await createYrsSession();
  sessions.push(session);
  session.openDocx(await zip.generateAsync({ type: 'uint8array' }), true);
  const target = session.paragraphs('body')[1]!;
  session.setSelection({ story: 'body', paraId: target.paraId, offset: 1 });
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
    />
  );
  const textarea = view.getByTestId('yrs-input');
  const press = async (key: string) => {
    fireEvent.keyDown(textarea, { key });
    await act(async () => { await input.current!.flushPendingInput(); });
  };
  return { session, textarea, input, press };
}

for (const between of [false, true]) {
  test(`in-thread backspace ${between ? 'between tables' : 'after a table'} preserves both the story and cell contents`, async () => {
    const { session, textarea, input, press } = await mount(between);
    const before = session.storySegments('body');
    const cellStory = session.storyIds().find(story =>
      session.paragraphs(story).some(para => para.text === 'T1-B2'))!;
    const cell = session.paragraphs(cellStory)[0]!;
    await press('Backspace');
    expect(session.storySegments('body')).toEqual(before);
    expect(session.selection()?.head).toEqual({ story: cellStory, paraId: cell.paraId, offset: 5 });
    fireEvent.input(textarea, { target: { value: 'Hello' } });
    await act(async () => { await input.current!.flushPendingInput(); });
    expect(session.paragraphs(cellStory)[0]!.text).toBe('T1-B2Hello');
    expect(session.storySegments('body')).toEqual(before);
    expect(() => session.yrsBlocksForStory('body', {})).not.toThrow();
    for (const text of ['T1-A1', 'T1-B1', 'T1-A2', ...(between ? ['T2-A1', 'T2-B1', 'T2-A2', 'T2-B2'] : [])]) {
      expect(session.storyIds().some(story => session.paragraphs(story).some(para => para.text === text))).toBe(true);
    }
  });
}

test('in-thread forward delete before a table preserves the boundary and table', async () => {
  const { session, press } = await mount(false);
  const before = session.storySegments('body');
  const first = session.paragraphSpans('body')[0]!;
  const caret = { story: 'body', paraId: first.paraId, offset: first.length };
  session.setSelection(caret);
  await press('Delete');
  expect(session.storySegments('body')).toEqual(before);
  expect(session.selection()?.head).toEqual(caret);
  expect(() => session.yrsBlocksForStory('body', {})).not.toThrow();
});
