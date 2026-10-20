export function caretColumnOffset(lines: readonly { columnOffset?: number }[], index: number): number {
  return (lines[index]?.columnOffset ?? 0) - (lines[0]?.columnOffset ?? 0);
}
