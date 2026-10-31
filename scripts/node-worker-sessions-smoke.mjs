import assert from 'node:assert/strict';
import { fork } from 'node:child_process';
import { readFileSync } from 'node:fs';
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
  for (const suffix of ['', '-native']) {
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
      const point = { story: 'body', paraId: before.value.paragraphs[0].paraId, offset: 0 };
      const selection = { anchor: point, focus: point };
      const fontChains = { 'liberation sans|0|0': [0], 'calibri|0|0': [0] };
      const initial = await document.commandInit({
        session: 'node-worker-smoke', incarnation: 1, epoch: 1, readOnly: false, selection,
        fonts: [fixture('crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')],
        fontRegistrationSlots: true,
        layoutInput: JSON.stringify({
          bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
          measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains },
          renderEnv: {},
        }),
        layoutWithRegions: true, layoutRevision: 1, extras: JSON.stringify({ fontChains }),
        displayWindow: [0, 1], retainBuiltPages: false,
      });
      const edited = await document.edit({
        session: initial.session, incarnation: initial.incarnation, seq: 1,
        basis: { epoch: initial.epoch, version: initial.version },
        kind: 'insertText', payload: { text: 'Node worker smoke ', selection },
      });
      assert.equal(edited.outcomes[0]?.status, 'applied');
      const after = await document.documentRead({ kind: 'readParagraphs', request: { view: 'accepted' } });
      assert.equal(after.value.ok, true);
      assert.ok(after.value.paragraphs[0].text.startsWith('Node worker smoke '));
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
    console.log(`Node XLSX/PPTX/DOCX${suffix} worker read, edit and save/reopen checks passed.`);
  }
}
