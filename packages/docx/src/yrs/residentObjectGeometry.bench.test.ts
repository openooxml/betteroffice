import { expect, spyOn, test } from 'bun:test';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { relative, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { readResidentFrameGeometry } from './residentObjectGeometry';
import { readResidentFrameGeometry as referenceGeometry } from './residentObjectGeometry.reference';
import { applyFrameDeltaOwned, decodeFrameDelta } from '../layout/render/frameDelta';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
const testing = process.env.RESIDENT_GEOMETRY_BENCH === '1' ? await import('./residentCommandWorker.test') : null;

test.skipIf(process.env.RESIDENT_GEOMETRY_BENCH !== '1')('geometry benchmark compares resident direct reads and command acknowledgements', async () => {
  const { openWorker, intent } = testing!;
  const base = process.env.GEOMETRY_BASELINE_REF!;
  const edit = await import('./wasm/index');
  const baselineEdit = process.env.GEOMETRY_BASELINE_EDIT ? await import(pathToFileURL(resolve(process.env.GEOMETRY_BASELINE_EDIT)).href) : null;
  if (baselineEdit) {
    await baselineEdit.default({ module_or_path: new Uint8Array(readFileSync(resolve(process.env.GEOMETRY_BASELINE_EDIT!).replace(/\.js$/, '_bg.wasm'))) });
    expect(baselineEdit.EditSession.prototype.resident_geometry_json).toBeUndefined();
  }
  const originals = new Map<string, string>();
  const paths = ['residentObjectGeometry.ts', 'residentEngineSession.ts', 'residentEngineWorker.ts',
    'residentCommandActions.ts', 'residentCommandEditState.ts', 'residentSections.ts', 'residentReviewCommands.ts'];
  for (const path of paths) originals.set(path, execFileSync('git', ['show', `${base}:packages/docx/src/yrs/${path}`], { encoding: 'utf8' }));
  const build = Bun.build;
  const patch = spyOn(Bun, 'build').mockImplementation(options => build({ ...options, plugins: [{
    name: 'baseline-engine', setup(builder) {
      builder.onLoad({ filter: /\/yrs\/[^/]+\.ts$/ }, args => {
        const path = relative(process.cwd() + '/packages/docx/src/yrs', args.path);
        const contents = originals.get(path);
        return contents === undefined ? undefined : { contents, loader: 'ts' };
      });
    },
  }, ...(options.plugins ?? [])] }));
  let baseline: Awaited<ReturnType<typeof residentWorkerFactory>>;
  try { baseline = await residentWorkerFactory(); } finally { patch.mockRestore(); }
  const head = await residentWorkerFactory();
  const omit = spyOn(Bun, 'build').mockImplementation(options => build({ ...options, plugins: [{
    name: 'omit-frame-geometry', setup(builder) {
      builder.onLoad({ filter: /\/yrs\/residentObjectGeometry\.ts$/ }, args => ({ loader: 'ts',
        contents: readFileSync(args.path, 'utf8').replace('): Promise<ResidentFrameGeometry> {',
          '): Promise<ResidentFrameGeometry> { return { version: engine.proposalEngine.version(), frameEpoch: frame.frameEpoch, layoutRevision, selection: [], objects: [], unbuiltPages: frame.pages.filter(page => page.page.unbuilt).map(page => page.pageIndex) };'),
      }));
    },
  }, ...(options.plugins ?? [])] }));
  let withoutGeometry: typeof head;
  try { withoutGeometry = await residentWorkerFactory(); } finally { omit.mockRestore(); }
  const arms = [];
  for (const count of [100, 1000, 4000]) {
    const paragraph = (index: number) => `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0')}"><w:r><w:t>Seed paragraph ${index}</w:t></w:r></w:p>`;
    const table = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    const bodyXml = Array.from({ length: count }, (_, index) => paragraph(index) + (index === 2 || index === Math.floor(count / 2) || index === count - 2 ? table : '')).join('');
    for (const [label, factory, geometry] of [['base', baseline, referenceGeometry], ['head', head, readResidentFrameGeometry],
      ['withoutGeometry', withoutGeometry, readResidentFrameGeometry]] as const) {
      const native = label === 'base' && baselineEdit ? spyOn(edit, 'createEditSession').mockImplementation(clientId => new baselineEdit.EditSession(clientId)) : null;
      const limit = label === 'base' && baselineEdit ? spyOn(edit, 'setEditWasmHeapLimit').mockImplementation(bytes => baselineEdit.set_wasm_heap_limit(bytes ?? Infinity)) : null;
      let arm: Awaited<ReturnType<typeof openWorker>>;
      try {
        arm = await openWorker({ bodyXml, factory, displayWindow: [0, 1] });
        if (native) expect(native.mock.calls.length).toBeGreaterThan(0);
      } finally { native?.mockRestore(); limit?.mockRestore(); }
      arms.push({ ...arm, count, label, geometry, seq: 0, version: 0,
        frame: applyFrameDeltaOwned(null, decodeFrameDelta(arm.initial.frame.payload.frame)), direct: [] as number[], ack: [] as number[] });
    }
  }
  try {
    for (let round = -2; round < 5; round += 1) {
      const order = round % 2 === 0 ? arms : [...arms].reverse();
      for (const arm of order) {
        const started = performance.now();
        await arm.geometry(arm.engine, arm.frame, null, 1);
        const direct = performance.now() - started;
        const sent = performance.now();
        const response = await arm.client.edit(intent(++arm.seq, arm.version, 'insertText', { text: '+' }));
        const ack = performance.now() - sent;
        expect(response.outcomes[0]?.status).toBe('applied');
        arm.version = response.version;
        for (const frame of [...response.frame!.payload.pageFrames ?? [], response.frame!.payload.frame]) {
          arm.frame = applyFrameDeltaOwned(arm.frame, decodeFrameDelta(frame));
        }
        expect(arm.frame.pages.filter(page => !page.page.unbuilt)).toHaveLength(1);
        if (round >= 0) { arm.direct.push(direct); arm.ack.push(ack); }
      }
    }
    const median = (values: number[]) => [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)]!;
    for (const arm of arms) console.log('GEOMETRY_BENCH', JSON.stringify({ count: arm.count, version: arm.label, baselineNative: baselineEdit ? base : 'shared',
      directMs: median(arm.direct), ackMs: median(arm.ack), directSamples: arm.direct, ackSamples: arm.ack }));
  } finally { for (const arm of arms) arm.client.destroy(); }
}, 120_000);
