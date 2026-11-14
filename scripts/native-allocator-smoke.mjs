import assert from 'node:assert/strict';
import { copyFileSync, mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Worker, isMainThread, parentPort, workerData } from 'node:worker_threads';

const root = fileURLToPath(new URL('..', import.meta.url));
const fixtures = {
  docx: readFileSync(join(root, 'apps/demo/public/betteroffice-demo.docx')),
  pptx: readFileSync(join(root, 'apps/demo/public/betteroffice-demo.pptx')),
};

function load(path) {
  const module = { exports: {} };
  process.dlopen(module, path);
  return module.exports;
}

function exercise(format, binding, rounds) {
  const bytes = new Uint8Array(fixtures[format]);
  for (let round = 0; round < rounds; round += 1) {
    if (format === 'docx') {
      const session = new binding.EditSession(round + 1);
      session.open_docx(bytes, true);
      assert.ok(session.encode_state().byteLength > 0);
      assert.ok(Array.isArray(JSON.parse(binding.font_preflight(bytes)).faces));
      const rezipped = binding.rezip_docx(binding.unzip_docx(bytes));
      assert.ok(rezipped.byteLength > 0);
      session.free();
    } else {
      const document = binding.PptxDocument.openReplayBaseline(bytes);
      const slides = JSON.parse(document.snapshotJson()).slides.length;
      assert.ok(slides > 0);
      const reopened = binding.PptxDocument.openReplayBaseline(document.saveBytes());
      assert.equal(JSON.parse(reopened.snapshotJson()).slides.length, slides);
    }
  }
}

function images(formats, directory) {
  return formats.flatMap((format) => {
    const folder = join(root, 'bindings', `${format}-native`);
    const built = readdirSync(folder).filter((name) => name.endsWith('.node'));
    assert.equal(built.length, 1, `expected one built ${format}-native binary, found ${built.length}`);
    const original = join(folder, built[0]);
    const copy = join(directory, `${format}-second-image.node`);
    copyFileSync(original, copy);
    return [
      { format, path: original },
      { format, path: copy },
    ];
  });
}

if (isMainThread) {
  const reverse = process.argv.includes('--reverse');
  const formats = process.argv.slice(2).filter((argument) => !argument.startsWith('--'));
  assert.ok(formats.length > 0 && formats.every((format) => format in fixtures), 'usage: native-allocator-smoke.mjs <docx|pptx>... [--reverse]');
  const directory = mkdtempSync(join(tmpdir(), 'native-allocator-'));
  try {
    const loaded = images(formats, directory);
    const order = reverse ? [...loaded].reverse() : loaded;
    const bindings = order.map((image) => ({ ...image, binding: load(image.path) }));
    for (const { format, binding } of bindings) exercise(format, binding, 3);
    for (const { format, binding } of [...bindings].reverse()) exercise(format, binding, 3);
    const worker = new Worker(fileURLToPath(import.meta.url), { workerData: order });
    const message = await new Promise((resolve, reject) => {
      worker.once('message', resolve);
      worker.once('error', reject);
    });
    assert.equal(message, 'worker ok');
    await worker.terminate();
    for (const { format, binding } of bindings) exercise(format, binding, 2);
    globalThis.gc?.();
    console.log(`allocator smoke ok: ${order.map((image) => `${image.format}:${image.path === join(directory, `${image.format}-second-image.node`) ? 'copy' : 'original'}`).join(' ')}`);
  } finally {
    try {
      rmSync(directory, { recursive: true, force: true });
    } catch {}
  }
} else {
  for (const image of workerData) exercise(image.format, load(image.path), 2);
  parentPort.postMessage('worker ok');
}
