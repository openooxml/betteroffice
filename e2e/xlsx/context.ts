/** Shared plumbing for the xlsx scenarios: wasm init, handle lifetime, stage mapping. */

import { expect } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { initWasm, openWorkbook } from '../../packages/xlsx/src/wasm/loader';
import type {
  DisplayListProfile,
  EditProfile,
  OpenWorkbookOptions,
  ProfiledEditResult,
  Viewport,
  WorkbookHandle,
} from '../../packages/xlsx/src/wasm/loader';
import type { DisplayList } from '../../packages/xlsx/src/display-list/types';
import type { PinnedSample } from '../corpus';
import type {
  ActorRecorder,
  ScenarioRecorder,
  StageProfile,
  Timer,
} from '../harness';
import type { Scenario } from '../suite';

const WASM = resolve(
  import.meta.dir,
  '../../packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm'
);
export const VIEWPORT: Viewport = { x: 0, y: 0, width: 1280, height: 800 };
export const SHEET = 0;

export interface XlsxCtx {
  sample: PinnedSample;
  bytes: Uint8Array;
  recorder: ScenarioRecorder;
  /** Opens a handle that is disposed with the scenario. */
  open(bytes?: Uint8Array, options?: OpenWorkbookOptions): WorkbookHandle;
  dispose(): void;
}

export type XlsxScenario = Scenario<XlsxCtx>;

export function setup(): Promise<void> {
  return initWasm(new Uint8Array(readFileSync(WASM)));
}

export function context(
  sample: PinnedSample,
  bytes: Uint8Array,
  recorder: ScenarioRecorder
): XlsxCtx {
  const handles: WorkbookHandle[] = [];
  return {
    sample,
    bytes,
    recorder,
    open(source = bytes, options) {
      const handle = openWorkbook(source, options);
      handles.push(handle);
      return handle;
    },
    dispose() {
      for (const handle of handles) {
        try {
          handle.dispose();
        } catch {}
      }
    },
  };
}

export function editStages(profile: EditProfile): StageProfile {
  return {
    validate: profile.validateMs,
    apply: profile.applyMs,
    recalc: profile.recalcMs,
    result: profile.resultMs,
  };
}

export function displayStages(profile: DisplayListProfile): StageProfile {
  return { build: profile.buildMs, encode: profile.encodeMs };
}

export function profiledEdit(
  handle: WorkbookHandle,
  timer: Timer,
  op: string,
  row: number,
  col: number,
  input: string,
  sheet = SHEET
): ProfiledEditResult {
  let profile: EditProfile | undefined;
  return timer.op(
    op,
    () => {
      const result = handle.editCellProfiled(sheet, row, col, input);
      profile = result.profile;
      return result;
    },
    () => editStages(profile!)
  );
}

export function profiledOps(
  handle: WorkbookHandle,
  timer: Timer,
  op: string,
  ops: unknown[]
): ProfiledEditResult {
  let profile: EditProfile | undefined;
  return timer.op(
    op,
    () => {
      const result = handle.applyOpsProfiled(ops);
      profile = result.profile;
      return result;
    },
    () => editStages(profile!)
  );
}

export function profiledDisplayList(
  handle: WorkbookHandle,
  timer: Timer,
  op: string,
  viewport = VIEWPORT
): DisplayList {
  let profile: DisplayListProfile | undefined;
  return timer.op(
    op,
    () => {
      const result = handle.displayListProfiled(viewport);
      profile = result.profile;
      return result.displayList;
    },
    () => displayStages(profile!)
  );
}

/** The formatted text the sheet shows at `row`/`col`, found through text search. */
export function displayedText(
  handle: WorkbookHandle,
  timer: Timer,
  op: string,
  query: string,
  row: number,
  col: number,
  sheet = SHEET
): string {
  const matches = timer.op(op, () => handle.searchText(query));
  const match = matches.find(
    (m) => m.sheet === sheet && m.row === row && m.col === col
  );
  expect(match, `${query} at row ${row} col ${col}`).toBeDefined();
  return match!.text;
}

/** One collaborative replica with the update bytes its own edits produced. */
export interface Replica {
  name: string;
  handle: WorkbookHandle;
  timer: ActorRecorder;
  outbox: Uint8Array[];
}

export function replicas(ctx: XlsxCtx, names: string[]): Replica[] {
  return names.map((name, index) => {
    const timer = ctx.recorder.as(name);
    const handle = timer.load(() =>
      ctx.open(ctx.bytes, { collaborative: true, clientId: index + 1 })
    );
    const outbox: Uint8Array[] = [];
    handle.onUpdate((update, origin) => {
      if (origin === 'local') outbox.push(update);
    });
    return { name, handle, timer, outbox };
  });
}

/** Delivers every pending update to every other replica, one recorded op each. */
export function exchange(peers: Replica[], op = 'applyUpdate'): number {
  let bytes = 0;
  for (const from of peers) {
    const pending = from.outbox.splice(0);
    for (const update of pending) {
      bytes += update.byteLength;
      for (const to of peers) {
        if (to === from) continue;
        const applied = to.timer.op(
          op,
          () => to.handle.applyUpdate(update),
          undefined,
          { from: from.name, bytes: update.byteLength }
        );
        expect(applied.applied).toBe(true);
      }
    }
  }
  return bytes;
}

/** Inputs, formatting, and rendered values in the selected range. */
export function fingerprint(
  handle: WorkbookHandle,
  range: string,
  sheet = SHEET
): string {
  const cells = handle.rangeCells(sheet, range).flat();
  return JSON.stringify({
    cells,
    formatting: cells.map((cell) => handle.selectionFormatting(sheet, cell.a1)),
    display: handle.printDisplayList(sheet, range, PRINT_METRICS, false),
  });
}

export const PRINT_METRICS = {
  dpi: 96,
  maxDigitWidth: 7,
  defaultRowHeightPt: 15,
  fontSizePt: 11,
  fontFamily: 'Calibri',
  fontAscent: 0.9,
  fontDescent: 0.2,
};
