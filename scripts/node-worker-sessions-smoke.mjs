import assert from 'node:assert/strict';
import { fork } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { setTimeout as delay } from 'node:timers/promises';
import { fileURLToPath } from 'node:url';

const fixture = path => new Uint8Array(readFileSync(new URL(`../${path}`, import.meta.url)));
const xlsxBytes = fixture('crates/ooxml-opc/tests/fixtures/sample.xlsx');

if (process.argv.includes('--keep-alive-child')) {
  const { openWorkbookSession } = await import('@betteroffice/xlsx-native/worker');
  const session = await openWorkbookSession(xlsxBytes);
  const dispose = new Promise(resolve => process.once('message', resolve));
  await new Promise((resolve, reject) => process.send('ready', error => error ? reject(error) : resolve()));
  process.channel.unref();
  assert.equal(await dispose, 'dispose');
  await session.dispose();
} else if (process.argv.includes('--keep-alive')) {
  const child = fork(fileURLToPath(import.meta.url), ['--keep-alive-child'], {
    stdio: ['ignore', 'inherit', 'inherit', 'ipc'],
  });
  let timer;
  try {
    const exited = new Promise((resolve, reject) => {
      child.once('error', reject);
      child.once('exit', (code, signal) => resolve({ code, signal }));
    });
    const deadline = new Promise((_, reject) => {
      timer = setTimeout(() => reject(new Error('Worker session liveness smoke timed out')), 60_000);
    });
    const ready = new Promise(resolve => child.once('message', resolve));
    assert.equal(await Promise.race([ready, exited, deadline]), 'ready');
    await delay(250);
    assert.equal(child.exitCode, null, 'An open session must keep Node alive');
    assert.equal(child.signalCode, null);
    child.send('dispose');
    assert.deepEqual(await Promise.race([exited, deadline]), { code: 0, signal: null });
    console.log('Node worker session keeps the process alive until disposal.');
  } finally {
    clearTimeout(timer);
    if (child.exitCode === null && child.signalCode === null) child.kill();
  }
} else {
  const pptxBytes = fixture('apps/demo/public/betteroffice-demo.pptx');
  const docxBytes = fixture('apps/demo/public/betteroffice-demo.docx');
  const runtime = globalThis.Deno ? 'Deno' : globalThis.Bun ? 'Bun' : 'Node';
  assert.equal(typeof Worker === 'function', runtime !== 'Node', `${runtime} worker sessions use the wrong worker kind`);
  const started = [];
  const builtin = globalThis.process?.getBuiltinModule?.bind(globalThis.process);
  if (builtin) {
    globalThis.process.getBuiltinModule = (id) => {
      const module = builtin(id);
      if (id !== 'node:worker_threads' || !module) return module;
      return { ...module, Worker: class extends module.Worker {
        constructor(url, options) { super(url, options); started.push(['worker_threads', String(url)]); }
      } };
    };
  }
  if (runtime !== 'Node') {
    globalThis.Worker = class extends globalThis.Worker {
      constructor(url, options) { super(url, options); started.push(['web', String(url)]); }
    };
  }
  for (const suffix of ['', '-native']) {
    for (const format of ['xlsx', 'pptx', 'docx']) {
      const entry = import.meta.resolve(`@betteroffice/${format}${suffix}/worker`);
      const dist = new URL('.', entry);
      for (const worker of format === 'docx' ? ['docxWorker', 'yrs/docxWorker'] : [`${format}SessionWorker`]) {
        for (const file of [`${worker}.mjs`, `${worker}.node.mjs`]) {
          assert.ok(existsSync(new URL(file, dist)), `Missing worker entry: ${new URL(file, dist).href}`);
        }
      }
      if (runtime !== 'Node') continue;
      const adapter = createRequire(entry).resolve('#session-worker');
      assert.match(adapter, /sessionWorker\.node\.(?:js|mjs)$/);
      assert.ok(existsSync(adapter), `Missing conditional worker adapter: ${adapter}`);
    }
    const { openWorkbookSession } = await import(`@betteroffice/xlsx${suffix}/worker`);
    const workbook = await openWorkbookSession(xlsxBytes);
    let reopenedWorkbook;
    try {
      const target = { sheetId: (await workbook.call.sheets())[0].id, range: { kind: 'a1', a1: 'B2' } };
      assert.equal((await workbook.call.readCells({ ranges: [target] })).ok, true);
      const edited = await workbook.call.applyEdits({
        expectVersion: await workbook.call.version(),
        steps: [{ op: 'setCellInputs', target, inputs: [['=1+2']] }],
      });
      assert.equal(edited.ok, true);
      assert.equal(edited.applied, true);
      reopenedWorkbook = await openWorkbookSession(await workbook.save());
      const reopenedTarget = { ...target, sheetId: (await reopenedWorkbook.call.sheets())[0].id };
      const read = await reopenedWorkbook.call.readCells({ ranges: [reopenedTarget] });
      assert.equal(read.ok, true);
      assert.deepEqual(read.ranges[0].cells[0][0].value, { kind: 'number', value: 3 });
    } finally {
      await reopenedWorkbook?.dispose();
      await workbook.dispose();
    }

    const { openPresentationSession } = await import(`@betteroffice/pptx${suffix}/worker`);
    const presentation = await openPresentationSession(pptxBytes);
    let reopenedPresentation;
    try {
      const content = await presentation.call.readContent();
      assert.equal(content.ok, true);
      const story = content.stories.find(story => story.paragraphs[0]?.editable);
      assert.ok(story);
      const edited = await presentation.call.applyEdits({
        expectVersion: content.version,
        steps: [{ op: 'insertText', at: 'start', text: 'Node worker smoke ', target: {
          kind: 'range', slideId: story.slideId, shapeId: story.shapeId,
          storyId: story.storyId, start: 0, end: 0,
        } }],
      });
      assert.equal(edited.ok, true);
      assert.equal(edited.applied, true);
      const updated = await presentation.call.readContent();
      assert.equal(updated.ok, true);
      reopenedPresentation = await openPresentationSession(await presentation.save());
      const reopened = await reopenedPresentation.call.readContent();
      assert.equal(reopened.ok, true);
      assert.deepEqual(reopened.stories.map(story => story.text), updated.stories.map(story => story.text));
      assert.ok(reopened.stories.some(story => story.text.startsWith('Node worker smoke ')));
    } finally {
      await reopenedPresentation?.dispose();
      await presentation.dispose();
    }

    const { DocxSourceWorkerClient } = await import(`@betteroffice/docx${suffix}/worker`);
    const document = new DocxSourceWorkerClient();
    let reopenedDocument;
    try {
      await document.openSource({ bytes: docxBytes });
      const before = await document.documentRead({ kind: 'readParagraphs', request: { view: 'accepted' } });
      assert.equal(before.value.ok, true);
      assert.ok(before.value.paragraphs.length > 0);
      const after = before;
      const saved = await document.save({ comments: [] });
      reopenedDocument = new DocxSourceWorkerClient();
      await reopenedDocument.openSource({ bytes: new Uint8Array(saved.bytes) });
      const reopened = await reopenedDocument.documentRead({ kind: 'readParagraphs', request: { view: 'accepted' } });
      assert.equal(reopened.value.ok, true);
      assert.deepEqual(reopened.value.paragraphs.map(paragraph => paragraph.text),
        after.value.paragraphs.map(paragraph => paragraph.text));
    } finally {
      reopenedDocument?.destroy();
      document.destroy();
    }
    console.log(`${runtime} XLSX/PPTX${suffix} worker read, edit and save/reopen and DOCX${suffix} worker read and save/reopen checks passed.`);
  }
  const kind = runtime === 'Node' ? 'worker_threads' : 'web';
  const file = runtime === 'Node' ? /\/(?:xlsxSessionWorker|pptxSessionWorker|docxWorker)\.node\.mjs$/ : /\/(?:xlsxSessionWorker|pptxSessionWorker|docxWorker)\.mjs$/;
  assert.ok(started.length >= 12, `${runtime} started ${started.length} session workers: ${JSON.stringify(started)}`);
  for (const [path, url] of started) {
    assert.equal(path, kind, `${runtime} started a ${path} worker for ${url}`);
    assert.match(url, file, `${runtime} started the wrong worker file ${url}`);
  }
  for (const worker of ['xlsxSessionWorker', 'pptxSessionWorker', 'docxWorker']) {
    assert.ok(started.some(([, url]) => url.includes(`/${worker}.`)), `${runtime} never started ${worker}`);
  }
  console.log(`${runtime} sessions ran on ${kind} workers from the ${runtime === 'Node' ? 'Node' : 'browser'} worker files.`);
}
