import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createEditSession } from '@betteroffice/docx/wasm/edit';
import {
  applyFrameDeltaOwned,
  decodeFrameDelta,
  type RetainedFrame,
} from '@betteroffice/docx/layout/render';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';

/** A worker that answers from a real engine session. */
export class EngineWorker {
  static engine: ReturnType<typeof createEditSession> | null = null;
  static last: EngineWorker | null = null;
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror = null;
  posted: ResidentEngineWorkerRequest[] = [];
  releaseRequests = 0;
  releasedIndices: number[] = [];
  supersedeNextRelease = false;
  private frame: RetainedFrame | null = null;
  private engineFrame: RetainedFrame | null = null;
  private releasedEpochs = 0;
  holdPageBuilds = false;
  slicePageBuilds = false;
  heldPageBuilds: (() => void)[] = [];
  holdInputReplies = false;
  heldInputReplies: (() => void)[] = [];
  constructor() {
    EngineWorker.last = this;
    EngineWorker.spawned += 1;
  }
  static failPageBuilds = false;
  static outOfMemoryPageBuilds = false;
  static spawned = 0;
  terminated = false;
  postMessage(request: ResidentEngineWorkerRequest): void {
    this.posted.push(request);
    const engine = EngineWorker.engine!;
    if (request.type === 'releasePages') {
      this.releaseRequests += 1;
      if (
        this.supersedeNextRelease ||
        !this.frame ||
        request.expectedFrameEpoch !== this.frame.frameEpoch ||
        request.pages.some(
          ({ index, pageId }) => this.frame!.pages[index]?.pageId.toString() !== pageId
        )
      ) {
        this.supersedeNextRelease = false;
        queueMicrotask(() =>
          this.onmessage?.({
            data: { id: request.id, ok: true, superseded: true },
          } as MessageEvent<ResidentEngineWorkerResponse>)
        );
        return;
      }
      const indices = request.pages.map(({ index }) => index);
      this.releasedIndices.push(...indices);
      const pages = indices.map((index) => ({
        ...this.frame!.pages[index]!,
        primitiveIds: new BigUint64Array(),
        page: { ...this.frame!.displayList.pages[index]!, primitives: [], unbuilt: true },
      }));
      const frame = pageFrame(this.frame, pages);
      this.frame = applyFrameDeltaOwned(this.frame, decodeFrameDelta(frame));
      this.releasedEpochs += 1;
      queueMicrotask(this.response(request.id, frame));
      return;
    }
    if (request.type === 'buildPages' && EngineWorker.outOfMemoryPageBuilds) {
      queueMicrotask(() =>
        this.onmessage?.({
          data: {
            id: request.id,
            ok: false,
            error: 'Resident engine worker ran out of memory allocating 64 bytes: unreachable',
            terminal: true,
            outOfMemory: true,
          },
        } as MessageEvent<ResidentEngineWorkerResponse>)
      );
      return;
    }
    if (request.type === 'buildPages' && EngineWorker.failPageBuilds) {
      queueMicrotask(() =>
        this.onmessage?.({
          data: { id: request.id, ok: false, error: 'page build failed' },
        } as MessageEvent<ResidentEngineWorkerResponse>)
      );
      return;
    }
    if (
      request.type !== 'bootstrap' &&
      request.type !== 'buildFrame' &&
      request.type !== 'buildPages' &&
      request.type !== 'applyInput'
    ) {
      return;
    }
    let frame: Uint8Array;
    let pageFrames: Uint8Array[] | undefined;
    const expectedEpoch = request.expectedFrameEpoch - this.releasedEpochs;
    engine.set_windowed_incremental_builds(
      'displayWindow' in request && request.displayWindow !== undefined
    );
    if ('displayWindow' in request && request.displayWindow) {
      engine.set_display_window(...request.displayWindow);
      engine.set_display_retain_built_pages(request.retainBuiltPages === true);
    }
    if (request.type === 'bootstrap') engine.reset_frame_base();
    if (request.type === 'bootstrap' || request.type === 'buildFrame') {
      frame = engine.build_display_list_frame(request.extras, expectedEpoch);
    } else if (request.type === 'buildPages') {
      if (this.slicePageBuilds && request.background) {
        pageFrames = [];
        let epoch = expectedEpoch;
        for (let offset = 0; offset < request.pages.length; offset += 4) {
          const built = engine.build_display_pages_frame(
            Uint32Array.from(request.pages.slice(offset, offset + 4)), epoch
          );
          pageFrames.push(built);
          epoch = JSON.parse(engine.resident_caret_snapshot_json()).frameEpoch;
        }
        frame = pageFrames[pageFrames.length - 1]!;
      } else {
        frame = engine.build_display_pages_frame(
          Uint32Array.from(request.pages), expectedEpoch
        );
      }
    } else {
      const { anchor, head } = request.selection;
      engine.set_selection(anchor.story, anchor.paraId, anchor.offset, head.paraId, head.offset);
      frame = engine.apply_input(request.text, expectedEpoch);
    }
    const frames = (pageFrames ?? [frame]).map((bytes, slice) => {
      this.engineFrame = applyFrameDeltaOwned(this.engineFrame, decodeFrameDelta(bytes));
      if (this.releasedEpochs > 0) {
        if (request.type === 'buildPages') {
          const indices = pageFrames ? request.pages.slice(slice * 4, slice * 4 + 4) : request.pages;
          bytes = pageFrame(this.frame!, indices.map((index) => this.engineFrame!.pages[index]!));
        } else {
          const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
          view.setBigUint64(32, BigInt(this.engineFrame.frameEpoch + this.releasedEpochs), true);
          const base = view.getBigUint64(40, true);
          if (base > 0n) view.setBigUint64(40, base + BigInt(this.releasedEpochs), true);
        }
      }
      this.frame = applyFrameDeltaOwned(this.frame, decodeFrameDelta(bytes));
      return bytes;
    });
    frame = frames[frames.length - 1]!;
    if (pageFrames) pageFrames = frames;
    const respond = this.response(request.id, frame, pageFrames);
    if (request.type === 'buildPages' && this.holdPageBuilds) this.heldPageBuilds.push(respond);
    else if (request.type === 'applyInput' && this.holdInputReplies)
      this.heldInputReplies.push(respond);
    else queueMicrotask(respond);
  }
  private response(id: number, frame: Uint8Array, pageFrames?: Uint8Array[]): () => void {
    const engine = EngineWorker.engine!;
    const caret = JSON.parse(engine.resident_caret_snapshot_json());
    caret.frameEpoch = this.frame!.frameEpoch;
    const selection = JSON.parse(engine.selection());
    return () =>
      this.onmessage?.({
        data: {
          id,
          ok: true,
          frame: frame.slice().buffer,
          ...(pageFrames ? { pageFrames: pageFrames.map((built) => built.slice().buffer) } : {}),
          caret,
          selection,
          layoutRevision: 1,
        },
      } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  releasePageBuilds(): void {
    for (const respond of this.heldPageBuilds.splice(0)) queueMicrotask(respond);
  }
  releaseInputReplies(): void {
    for (const respond of this.heldInputReplies.splice(0)) queueMicrotask(respond);
  }
  terminate(): void {
    this.terminated = true;
  }
}

function pageFrame(base: RetainedFrame, pages: RetainedFrame['pages']): Uint8Array {
  const strings: string[] = [];
  const u32 = (out: number[], value: number): void => {
    for (let shift = 0; shift < 32; shift += 8) out.push((value >>> shift) & 0xff);
  };
  const stringId = (value: string): number => {
    let index = strings.indexOf(value);
    if (index < 0) index = strings.push(value) - 1;
    return index;
  };
  const encode = (out: number[], value: unknown): void => {
    if (value === null) out.push(0);
    else if (typeof value === 'boolean') out.push(value ? 2 : 1);
    else if (typeof value === 'number') {
      out.push(5);
      const bytes = new Uint8Array(8);
      new DataView(bytes.buffer).setFloat64(0, value, true);
      out.push(...bytes);
    } else if (typeof value === 'string') {
      out.push(6);
      u32(out, stringId(value));
    } else {
      const body: number[] = [];
      const array = Array.isArray(value);
      const items = array
        ? value
        : Object.entries(value as object).filter(([, entry]) => entry !== undefined);
      for (const item of items) {
        if (array) encode(body, item);
        else {
          const [key, entry] = item;
          u32(body, stringId(key));
          encode(body, entry);
        }
      }
      out.push(array ? 7 : 8);
      u32(out, body.length);
      u32(out, items.length);
      out.push(...body);
    }
  };
  const payloads = pages.map(({ page }) => {
    const payload: number[] = [];
    encode(payload, page);
    return Uint8Array.from(payload);
  });
  const table: number[] = [];
  u32(table, strings.length);
  for (const value of strings) {
    const bytes = new TextEncoder().encode(value);
    u32(table, bytes.length);
    table.push(...bytes);
  }
  const align = (value: number): number => Math.ceil(value / 8) * 8;
  const stringsOffset = 80 + pages.length * 48;
  const dataOffset = align(stringsOffset + table.length);
  let total = dataOffset;
  const offsets = pages.map(({ primitiveIds }, index) => {
    const ids = align(total);
    const payload = ids + primitiveIds.length * 8;
    total = payload + payloads[index]!.length;
    return { ids, payload };
  });
  const bytes = new Uint8Array(total);
  const view = new DataView(bytes.buffer);
  bytes.set([0x46, 0x44, 0x56, 0x31]);
  view.setUint16(4, 1, true);
  view.setUint16(6, 80, true);
  view.setUint32(8, total, true);
  view.setBigUint64(16, BigInt(base.docEpoch), true);
  view.setBigUint64(24, BigInt(base.layoutEpoch), true);
  view.setBigUint64(32, BigInt(base.frameEpoch + 1), true);
  view.setBigUint64(40, BigInt(base.frameEpoch), true);
  view.setUint32(48, base.pages.length, true);
  view.setUint32(52, pages.length, true);
  view.setUint32(56, 80, true);
  view.setUint32(60, stringsOffset, true);
  view.setUint32(64, table.length, true);
  view.setUint32(68, dataOffset, true);
  view.setUint32(72, base.contractVersion ?? 0, true);
  bytes.set(table, stringsOffset);
  pages.forEach(({ pageIndex, pageId, fingerprint, primitiveIds }, index) => {
    const offset = 80 + index * 48;
    const { ids, payload } = offsets[index]!;
    bytes[offset] = 1;
    view.setUint32(offset + 4, pageIndex, true);
    view.setBigUint64(offset + 8, pageId, true);
    view.setBigUint64(offset + 16, fingerprint, true);
    view.setUint32(offset + 24, primitiveIds.length, true);
    view.setUint32(offset + 28, ids, true);
    view.setUint32(offset + 32, payload, true);
    view.setUint32(offset + 36, payloads[index]!.length, true);
    primitiveIds.forEach((id, position) => view.setBigUint64(ids + position * 8, id, true));
    bytes.set(payloads[index]!, payload);
  });
  return bytes;
}

export const PREVIEW = { r1: 'accepted' } as const;

export function lazyFixture(paragraphLength?: number) {
  const engine = createEditSession(9401);
  const text = 'Lazy pages. '.repeat(400);
  let { paraId } = JSON.parse(engine.create_story('body', text, 'Normal', 'left'));
  if (paragraphLength) {
    for (let remaining = text.length; remaining > paragraphLength; remaining -= paragraphLength) {
      paraId = JSON.parse(engine.split_paragraph('body', paraId, paragraphLength)).secondParaId;
      engine.set_paragraph_attr(paraId, 'pageBreakBefore', 'true');
    }
  }
  const fontId = engine.register_measure_font(
    new Uint8Array(
      readFileSync(
        resolve(
          import.meta.dir,
          '../../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
        )
      )
    )
  );
  const inputs = JSON.parse(
    engine.layout_document_with_regions_json(
      JSON.stringify({
        bodyStory: 'body',
        regions: {
          sections: [
            {
              sectionId: 'main',
              properties: {
                pageWidth: 4320,
                pageHeight: 2880,
                marginTop: 300,
                marginRight: 300,
                marginBottom: 300,
                marginLeft: 300,
              },
            },
          ],
        },
        measurement: {
          fontChains: { 'calibri|0|0': [fontId] },
          defaults: { fontSize: 11, fontFamily: 'Calibri' },
          authoritativeShaping: true,
        },
        renderEnv: {},
      })
    )
  );
  EngineWorker.engine = engine;
  EngineWorker.failPageBuilds = false;
  EngineWorker.outOfMemoryPageBuilds = false;
  EngineWorker.spawned = 0;
  globalThis.Worker = EngineWorker as unknown as typeof Worker;
  const host = {
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({
      state: new Uint8Array(),
      fonts: [],
      fontsRevision: 0,
      layoutRevision: 1,
      layoutInput: JSON.stringify({ renderEnv: { revisionPreview: PREVIEW } }),
    }),
    resetFrameBase: () => {},
    encodeStateVector: () => new Uint8Array(),
    onUpdate: () => () => {},
    selection: () => JSON.parse(engine.selection()),
    applyUpdate: () => null,
  } as unknown as YrsSession;
  return { engine, inputs, host };
}
