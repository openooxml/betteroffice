import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createEditSession } from '@betteroffice/docx/wasm/edit';
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
  holdPageBuilds = false;
  heldPageBuilds: (() => void)[] = [];
  constructor() {
    EngineWorker.last = this;
  }
  static failPageBuilds = false;
  postMessage(request: ResidentEngineWorkerRequest): void {
    this.posted.push(request);
    const engine = EngineWorker.engine!;
    if (request.type === 'buildPages' && EngineWorker.failPageBuilds) {
      queueMicrotask(() =>
        this.onmessage?.({
          data: { id: request.id, ok: false, error: 'page build failed' },
        } as MessageEvent<ResidentEngineWorkerResponse>)
      );
      return;
    }
    let frame: Uint8Array;
    if (request.type === 'bootstrap') {
      if (request.displayWindow) engine.set_display_window(...request.displayWindow);
      frame = engine.build_display_list_frame(request.extras, 0);
    } else if (request.type === 'buildPages') {
      frame = engine.build_display_pages_frame(
        Uint32Array.from(request.pages),
        request.expectedFrameEpoch
      );
    } else if (request.type === 'applyInput') {
      if (request.displayWindow) engine.set_display_window(...request.displayWindow);
      const { anchor, head } = request.selection;
      engine.set_selection(anchor.story, anchor.paraId, anchor.offset, head.paraId, head.offset);
      frame = engine.apply_input(request.text, request.expectedFrameEpoch);
    } else {
      return;
    }
    const caret = JSON.parse(engine.resident_caret_snapshot_json());
    const respond = () =>
      this.onmessage?.({
        data: {
          id: request.id,
          ok: true,
          frame: frame.slice().buffer,
          caret,
          selection: JSON.parse(engine.selection()),
          layoutRevision: 1,
        },
      } as MessageEvent<ResidentEngineWorkerResponse>);
    if (request.type === 'buildPages' && this.holdPageBuilds) this.heldPageBuilds.push(respond);
    else queueMicrotask(respond);
  }
  releasePageBuilds(): void {
    for (const respond of this.heldPageBuilds.splice(0)) queueMicrotask(respond);
  }
  terminate(): void {}
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
