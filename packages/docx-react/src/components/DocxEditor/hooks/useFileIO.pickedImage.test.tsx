import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import { parseDocx } from '@betteroffice/docx/docx';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { unzipContainer } from '@betteroffice/docx/docx/wasm';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import { takePreloadedResidentEngineWorker } from '@betteroffice/docx/yrs';
import { residentWorkerFactory, type InProcessResidentWorker } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from '../internals/engineChoice';
import { DocxEditor, type DocxEditorRef } from '../../DocxEditor';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
const originalImage = globalThis.Image;
const originalWarn = console.warn;
const originalFonts = Object.getOwnPropertyDescriptor(window.document, 'fonts');
let warn: ReturnType<typeof spyOn<typeof console, 'warn'>>;
let addedFonts = false;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
const PNG = Uint8Array.from(atob(PNG_BASE64), (character) => character.charCodeAt(0));

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: Object.assign(new EventTarget(), { ready: Promise.resolve() }),
      configurable: true,
    });
    addedFonts = true;
  }
  await preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  );
  warn = spyOn(console, 'warn').mockImplementation((message, ...args) => {
    if (typeof message === 'string' && message.startsWith('[DocxEditor] experimentalWorkerOpen={false} selects')) return;
    originalWarn(message, ...args);
  });
});
afterEach(() => {
  cleanup();
  globalThis.Image = originalImage;
});
afterAll(async () => {
  warn.mockRestore();
  if (addedFonts) {
    if (originalFonts) Object.defineProperty(window.document, 'fonts', originalFonts);
    else Reflect.deleteProperty(window.document, 'fonts');
  }
  if (ownsDom) await GlobalRegistrator.unregister();
});

function fixture(): ArrayBuffer {
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      `<Types xmlns="${CONTENT_TYPES}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      `<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    )
  );
  parts.set('word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="${RELS}"/>`));
  parts.set(
    'word/document.xml',
    toBytes(
      `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Picture</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
    )
  );
  return rezipPartsToArrayBuffer(parts);
}

async function until(done: () => boolean) {
  for (let attempt = 0; attempt < 300 && !done(); attempt += 1) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10));
    });
  }
  expect(done()).toBe(true);
}

function elements(bytes: Uint8Array, namespace: string, localName: string): Element[] {
  const xml = new DOMParser().parseFromString(new TextDecoder().decode(bytes), 'application/xml');
  return Array.from(xml.getElementsByTagName('*')).filter(
    (element) => element.namespaceURI === namespace && element.localName === localName
  );
}

async function assertSavedImage(buffer: ArrayBuffer) {
  const parts = unzipContainer(new Uint8Array(buffer));
  const blips = elements(parts['word/document.xml']!, A, 'blip');
  expect(blips).toHaveLength(1);
  const rId = blips[0]!.getAttribute('r:embed');
  expect(rId).toBeTruthy();
  const relationship = elements(parts['word/_rels/document.xml.rels']!, RELS, 'Relationship').find(
    (element) => element.getAttribute('Id') === rId
  );
  expect(relationship).toBeDefined();
  expect(relationship!.getAttribute('Type')).toBe(`${R}/image`);
  expect(relationship!.getAttribute('TargetMode')).not.toBe('External');
  const target = relationship!.getAttribute('Target');
  expect(target).toBeTruthy();
  const mediaPath = new URL(target!, 'https://docx.test/word/document.xml').pathname.slice(1);
  expect(mediaPath.endsWith('.png')).toBe(true);
  expect(parts[mediaPath]).toEqual(PNG);
  const contentTypes = parts['[Content_Types].xml']!;
  const covered =
    elements(contentTypes, CONTENT_TYPES, 'Default').some(
      (element) =>
        element.getAttribute('Extension') === 'png' &&
        element.getAttribute('ContentType') === 'image/png'
    ) ||
    elements(contentTypes, CONTENT_TYPES, 'Override').some(
      (element) =>
        element.getAttribute('PartName') === `/${mediaPath}` &&
        element.getAttribute('ContentType') === 'image/png'
    );
  expect(covered).toBe(true);

  const reopened = await parseDocx(new Uint8Array(buffer), { preloadFonts: false });
  const images = reopened.package.document.content.flatMap((block) =>
    block.type === 'paragraph'
      ? block.content.flatMap((inline) =>
          inline.type === 'run'
            ? inline.content.flatMap((content) => (content.type === 'drawing' ? [content.image] : []))
            : []
        )
      : []
  );
  expect(images).toHaveLength(1);
  expect(images[0]!.rId).toBe(rId!);
  expect(images[0]!.src).toBe(`data:image/png;base64,${PNG_BASE64}`);
}

test('an image inserted through the picker keeps its media after two saves and reopen', async () => {
  class LoadedImage {
    naturalWidth = 1;
    naturalHeight = 1;
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    set src(_value: string) {
      queueMicrotask(() => this.onload?.());
    }
  }
  globalThis.Image = LoadedImage as never;

  const ref = createRef<DocxEditorRef>();
  const errors: Error[] = [];
  const view = render(
    <DocxEditor ref={ref} experimentalWorkerOpen={false} documentBuffer={fixture()} onError={(error) => errors.push(error)} />
  );
  await until(() => ref.current?.commands.getState('save').enabled === true);
  const editor = ref.current!.getEditorRef()!;
  const session = editor.getYrsSession()!;
  await act(async () => {
    session.setSelection({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 0 });
    editor.syncYrsInputState(false);
  });
  await until(() => ref.current?.commands.getState('insertImage').enabled === true);
  await act(async () => {
    expect(await ref.current!.commands.execute('insertImage', null)).toEqual({
      ok: true,
      status: 'opened',
    });
  });
  const input = view.container.querySelector<HTMLInputElement>(
    'input[type="file"][accept="image/*"]'
  );
  expect(input).not.toBeNull();
  act(() => {
    fireEvent.change(input!, {
      target: { files: [new File([PNG], 'pixel.png', { type: 'image/png' })] },
    });
  });
  await until(() =>
    session
      .storySegments('body')
      .some((segment) => segment.kind === 'embed' && segment.embedKind === 'image')
  );
  const image = session.storySegments('body').find(
    (segment) => segment.kind === 'embed' && segment.embedKind === 'image'
  );
  expect(image?.kind).toBe('embed');
  if (image?.kind !== 'embed') throw new Error('The picked image was not inserted');
  expect(image.payload.rId ?? '').toBe('');
  expect(image.payload.src).toBe(`data:image/png;base64,${PNG_BASE64}`);
  expect(input!.value).toBe('');

  for (let save = 0; save < 2; save += 1) {
    let buffer: ArrayBuffer | null = null;
    await act(async () => {
      await ref.current!.flushPendingInput();
      buffer = await ref.current!.save();
    });
    expect(buffer).not.toBeNull();
    await assertSavedImage(buffer!);
  }
  expect(errors).toEqual([]);
}, 20_000);

test('an image inserted through the picker keeps its media after two saves and reopen on the default worker', async () => {
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
  class LoadedImage {
    naturalWidth = 1;
    naturalHeight = 1;
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    set src(_value: string) {
      queueMicrotask(() => this.onload?.());
    }
  }
  globalThis.Image = LoadedImage as never;
  try {
    const ref = createRef<DocxEditorRef>();
    const errors: Error[] = [];
    const view = render(
      <DocxEditor ref={ref} documentBuffer={fixture()} onError={(error) => errors.push(error)} />
    );
    await until(() => ref.current?.commands.getState('save').enabled === true);
    await act(async () => {
      await ref.current!.flushPendingInput();
      await ref.current!.whenLayoutComplete({ timeoutMs: 3_000 });
      const read = await ref.current!.readParagraphs({ view: 'accepted' });
      expect(read.ok).toBe(true);
      if (!read.ok) throw new Error(read.failure.message);
      expect(await ref.current!.scrollToParagraph(read.paragraphs[0]!.paraId)).toBe(true);
    });
    expect(workers.length).toBeGreaterThan(0);
    expect(workers.some((worker) => worker.requests.includes('open') && worker.sessions.length > 0)).toBe(true);
    await until(() => ref.current?.commands.getState('insertImage').enabled === true);
    await act(async () => {
      expect(await ref.current!.commands.execute('insertImage', null)).toEqual({
        ok: true,
        status: 'opened',
      });
    });
    const input = view.container.querySelector<HTMLInputElement>(
      'input[type="file"][accept="image/*"]'
    );
    expect(input).not.toBeNull();
    act(() => {
      fireEvent.change(input!, {
        target: { files: [new File([PNG], 'pixel.png', { type: 'image/png' })] },
      });
    });
    const images = () => ref.current!.getDocument()?.package.document.content.flatMap((block) =>
      block.type === 'paragraph'
        ? block.content.flatMap((inline) => inline.type === 'run'
          ? inline.content.flatMap((content) => content.type === 'drawing' ? [content.image] : [])
          : [])
        : []
    ) ?? [];
    await until(() => images().length === 1);
    const [image] = images();
    expect(image).toBeDefined();
    expect(image!.rId ?? '').toBe('');
    expect(image!.src).toBe(`data:image/png;base64,${PNG_BASE64}`);
    expect(input!.value).toBe('');

    for (let save = 0; save < 2; save += 1) {
      let buffer: ArrayBuffer | null = null;
      await act(async () => {
        await ref.current!.flushPendingInput();
        buffer = await ref.current!.save();
      });
      expect(buffer).not.toBeNull();
      await assertSavedImage(buffer!);
    }
    expect(errors).toEqual([]);
  } finally {
    cleanup();
    takePreloadedResidentEngineWorker()?.destroy();
    await act(async () => {});
    for (const worker of workers) worker.terminate();
    compileModule.mockRestore();
    resetEngineChoiceForTests();
    globalThis.Worker = originalWorker;
    globalThis.Image = originalImage;
  }
}, 20_000);
