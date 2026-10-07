import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { unzipContainer } from '@betteroffice/docx/docx/wasm';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import { takePreloadedResidentEngineWorker } from '@betteroffice/docx/yrs';
import { residentWorkerFactory, type InProcessResidentWorker } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { DocxEditor, type DocxEditorRef } from '../../index';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from './internals/engineChoice';

const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
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
      const view = render(<DocxEditor ref={ref} experimentalWorkerOpen={false} documentBuffer={buffer} />);
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

for (const typing of [false, true]) {
  test(`three editor saves ${typing ? 'with' : 'without'} typing keep imported comment ranges and paragraphs on the default worker`, async () => {
    const originalWorker = globalThis.Worker;
    const startWorker = await residentWorkerFactory();
    const workers: InProcessResidentWorker[] = [];
    const compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(new WebAssembly.Module(
      new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
    ));
    setMissingWorkerCapabilitiesForTests([]);
    globalThis.Worker = class {
      constructor() {
        const worker = startWorker();
        workers.push(worker);
        return worker;
      }
    } as unknown as typeof Worker;
    try {
      let buffer = new Uint8Array(readFileSync(FIXTURE)).buffer;
      const source = paragraphs(documentXml(buffer));
      for (let cycle = 0; cycle < 3; cycle += 1) {
        const opens = workers.flatMap((worker) => worker.requests).filter((type) => type === 'open').length;
        const ref = createRef<DocxEditorRef>();
        const view = render(<DocxEditor ref={ref} documentBuffer={buffer} />);
        await until(() => ref.current?.commands.getState('save').enabled === true);
        await act(async () => { await ref.current!.flushPendingInput(); });
        expect(workers.length).toBeGreaterThan(0);
        expect(workers.flatMap((worker) => worker.requests).filter((type) => type === 'open').length).toBeGreaterThan(opens);
        if (typing) {
          await act(async () => {
            await ref.current!.whenLayoutComplete({ timeoutMs: 3_000 });
            const read = await ref.current!.readParagraphs({ view: 'accepted' });
            expect(read.ok).toBe(true);
            if (!read.ok) throw new Error(read.failure.message);
            expect(await ref.current!.scrollToParagraph(read.paragraphs[0]!.paraId)).toBe(true);
            const input = view.getByTestId('yrs-input') as HTMLTextAreaElement;
            input.focus();
            await ref.current!.flushPendingInput();
            fireEvent.input(input, { target: { value: `QA${cycle} ` } });
            await ref.current!.flushPendingInput();
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
        if (typing) {
          expect(xml.replace(/<[^>]+>/g, '')).toContain(
            Array.from({ length: cycle + 1 }, (_, index) => `QA${cycle - index} `).join('')
          );
        }
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
    } finally {
      cleanup();
      takePreloadedResidentEngineWorker()?.destroy();
      await act(async () => {});
      for (const worker of workers) worker.terminate();
      compileModule.mockRestore();
      resetEngineChoiceForTests();
      globalThis.Worker = originalWorker;
    }
  }, 20_000);
}
