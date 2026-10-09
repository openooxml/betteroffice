import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import * as vsdx from '@betteroffice/vsdx';
import type { CellWriteQuery, DiagramSnapshot, ShapeSnapshot } from '@betteroffice/vsdx';
import { GATED_CELLS, type ShapeWriteProbes } from './commands';

const root = resolve(import.meta.dir, '../../../../..');
let blank: Uint8Array | null = null;

export async function initEngineProbe(): Promise<void> {
  if (blank) return;
  await vsdx.initWasm(await readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')));
  blank = await readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/foundation.vsdx'));
}

/** Test-only: asks the real engine about a shape built from these formulas. */
export function probeCells(cells: ReadonlyArray<{ cellName: string; formula: string }>, extra: readonly CellWriteQuery[] = []): ShapeWriteProbes | null {
  if (!blank) throw new Error('call initEngineProbe first');
  const handle = vsdx.openDiagram(blank);
  try {
    const pageId = handle.snapshot().pages[0].id;
    const unique = [...new Map(cells.map((entry) => [entry.cellName, entry])).values()];
    const added = handle.addShape(pageId, { name: 'probe', cells: unique.map(({ cellName, formula }) => ({ locator: { cellName }, formula })) }) as unknown as { shapeId: string };
    // Gated queries carry the gesture the UI performs, so they must win over the bare ones.
    const queries = [...unique.map(({ cellName }) => ({ cellName })), ...extra, ...GATED_CELLS];
    const answers = handle.probeCellWrites(pageId, added.shapeId, queries);
    return new Map(answers.map((probe) => [probe.cellName, probe]));
  } finally { handle.dispose(); }
}

/** Test-only: the same, for a shape a fabricated snapshot describes. */
export function probeSnapshotShape(shape: ShapeSnapshot): ShapeWriteProbes | null {
  return probeCells(shape.cells
    .filter((cell) => cell.locator.section === null && (cell.formula ?? cell.value) !== null)
    .map((cell) => ({ cellName: cell.name, formula: (cell.formula ?? cell.value) as string })));
}

/** Test-only: the same, for every shape in a snapshot, keyed by id. */
export function probeSnapshot(state: DiagramSnapshot): Map<string, ShapeWriteProbes> {
  const probes = new Map<string, ShapeWriteProbes>();
  const walk = (shapes: readonly ShapeSnapshot[]) => {
    for (const shape of shapes) {
      const cells = shape.cells
        .filter((cell) => cell.locator.section === null && (cell.formula ?? cell.value) !== null)
        .map((cell) => ({ cellName: cell.name, formula: (cell.formula ?? cell.value) as string }));
      const answered = probeCells(cells);
      if (answered) probes.set(shape.id, answered);
      walk(shape.children);
    }
  };
  for (const page of state.pages) walk(page.shapes);
  return probes;
}
