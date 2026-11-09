import { expect } from 'bun:test';
import type { Layout } from '../packages/docx/src/layout/pagination';
import type { WorkbookHandle } from '../packages/xlsx/src/wasm/loader';

export function assertWorkbookMirror(
  handle: Pick<WorkbookHandle, 'cell' | 'searchText'>,
  mirror: {
    values: Record<string, unknown>;
    formulas: Record<string, string | null>;
  },
  cells: { row: number; col: number; value: string | number }[]
): void {
  for (const { row, col, value } of cells) {
    const cell = handle.cell(0, row, col);
    expect(mirror.values[cell.a1], cell.a1 + ' Python value').toBe(value);
    expect(mirror.formulas[cell.a1], cell.a1 + ' formula').toBe(
      cell.isFormula ? cell.input.replace(/^=/, '') : null
    );
    const displayed = handle
      .searchText(String(value))
      .find(
        (match) => match.sheet === 0 && match.row === row && match.col === col
      );
    expect(displayed?.text, cell.a1 + ' WASM value').toBe(String(value));
  }
}

export function layoutProjection(layout: Layout) {
  return {
    pageSize: layout.pageSize,
    pages: layout.pages.map((page) => ({
      pageSize: page.size,
      fragments: page.fragments.map((fragment) => ({
        kind: fragment.kind,
        blockId: fragment.blockId,
        x: fragment.x,
        y: fragment.y,
        width: fragment.width,
        height: fragment.height,
        ...(fragment.kind === 'paragraph'
          ? { fromLine: fragment.fromLine, toLine: fragment.toLine }
          : {}),
        ...(fragment.kind === 'table'
          ? { rowStart: fragment.rowStart, rowEnd: fragment.rowEnd }
          : {}),
      })),
    })),
  };
}

export function assertLayoutParity(wasm: Layout, native: Layout): void {
  const compare = (left: unknown, right: unknown): void => {
    if (typeof left === 'number') {
      expect(typeof right).toBe('number');
      expect(right as number).toBeCloseTo(left, 4);
    } else if (Array.isArray(left)) {
      expect(Array.isArray(right)).toBe(true);
      expect((right as unknown[]).length).toBe(left.length);
      left.forEach((value, i) => compare(value, (right as unknown[])[i]));
    } else if (left !== null && typeof left === 'object') {
      expect(right).not.toBeNull();
      expect(typeof right).toBe('object');
      expect(Object.keys(right as object).sort()).toEqual(
        Object.keys(left).sort()
      );
      for (const [key, value] of Object.entries(left))
        compare(value, (right as Record<string, unknown>)[key]);
    } else expect(right).toEqual(left);
  };
  compare(layoutProjection(wasm), layoutProjection(native));
}

/** A Yrs v1 state vector as sorted [client, clock] pairs, so equal state compares equal whatever the encoding order. */
export function stateVectorEntries(bytes: Uint8Array): [bigint, bigint][] {
  let at = 0;
  const next = (): bigint => {
    let value = 0n;
    for (let shift = 0n; ; shift += 7n) {
      if (at >= bytes.length) throw new Error('truncated state vector');
      const byte = bytes[at++]!;
      value |= BigInt(byte & 0x7f) << shift;
      if (byte < 0x80) return value;
    }
  };
  const entries = Array.from({ length: Number(next()) }, () => [next(), next()] as [bigint, bigint]);
  if (at !== bytes.length) throw new Error('trailing bytes after state vector');
  return entries.sort((left, right) => (left[0] < right[0] ? -1 : left[0] > right[0] ? 1 : 0));
}
