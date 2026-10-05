import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { unzipContainer } from '@betteroffice/docx/docx/wasm';
import { DocxEditor, type DocxEditorRef } from '../../index';

const { act, cleanup, render } = await import('@testing-library/react');
const quiet = { error: console.error, warn: console.warn };

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {}, ready: Promise.resolve() },
      configurable: true,
    });
  }
  await preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'))
    )
  );
  console.error = () => {};
  console.warn = () => {};
});
afterEach(cleanup);
afterAll(async () => {
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom) await GlobalRegistrator.unregister();
});

const FIXTURE = resolve(
  import.meta.dir,
  '../../../../docx/src/yrs/__fixtures__/comment-ranges/structure.docx'
);

async function until(done: () => boolean) {
  for (let attempt = 0; attempt < 300 && !done(); attempt += 1) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10));
    });
  }
  expect(done()).toBe(true);
}

const documentXml = (buffer: ArrayBuffer) =>
  new TextDecoder().decode(unzipContainer(new Uint8Array(buffer))['word/document.xml']);
const paragraphs = (xml: string) => xml.match(/<w:p[ >/]/g)?.length ?? 0;
const markers = (xml: string, id: number) =>
  [...xml.matchAll(new RegExp(`<w:comment(RangeStart|RangeEnd|Reference) w:id="${id}"/>`, 'g'))].map(
    (match) => match[1]
  );
const covered = (xml: string, id: number) =>
  xml
    .split(`<w:commentRangeStart w:id="${id}"/>`)[1]
    ?.split(`<w:commentRangeEnd w:id="${id}"/>`)[0]
    ?.replace(/<[^>]+>/g, '');

for (const typing of [false, true]) {
  test(`three editor saves ${typing ? 'with' : 'without'} typing keep imported comment ranges and paragraphs`, async () => {
    let buffer = new Uint8Array(readFileSync(FIXTURE)).buffer;
    const source = paragraphs(documentXml(buffer));
    for (let cycle = 0; cycle < 3; cycle += 1) {
      const ref = createRef<DocxEditorRef>();
      const view = render(<DocxEditor ref={ref} documentBuffer={buffer} />);
      await until(() => ref.current?.commands.getState('save').enabled === true);
      if (typing) {
        const session = ref.current!.getEditorRef()!.getYrsSession()!;
        const [title] = session.paragraphs('body');
        await act(async () => {
          session.insertText({ story: 'body', paraId: title!.paraId, offset: 0 }, `QA${cycle} `);
        });
      }
      let saved: ArrayBuffer | null = null;
      await act(async () => {
        saved = await ref.current!.save();
      });
      view.unmount();
      expect(saved).not.toBeNull();
      buffer = saved!;
      const xml = documentXml(buffer);
      expect([cycle, paragraphs(xml)]).toEqual([cycle, source]);
      for (const [id, text] of [
        [0, 'Achado QA preservado. '],
        [1, 'Preservar comentário na célula'],
      ] as const) {
        expect([cycle, id, markers(xml, id), covered(xml, id)]).toEqual([
          cycle,
          id,
          ['RangeStart', 'RangeEnd', 'Reference'],
          text,
        ]);
      }
    }
  });
}
