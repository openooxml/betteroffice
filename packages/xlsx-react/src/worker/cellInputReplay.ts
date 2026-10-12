export type CellInputOperation = { kind: 'insert'; text: string } | { kind: 'selectAll' } |
  { kind: 'move'; direction: 'start' | 'end' | 'left' | 'right'; extend?: boolean } |
  { kind: 'delete'; direction: 'backward' | 'forward'; unit: 'character' | 'word' | 'line' };

export interface CellInputState {
  value: string;
  anchor: number;
  focus: number;
}

function characterOffset(value: string, at: number, backward: boolean): number {
  if (!backward) return Math.min(value.length, at + ((value.codePointAt(at) ?? 0) > 0xffff ? 2 : 1));
  const last = value.charCodeAt(at - 1);
  const previous = value.charCodeAt(at - 2);
  return Math.max(0, at - (last >= 0xdc00 && last <= 0xdfff && previous >= 0xd800 && previous <= 0xdbff ? 2 : 1));
}

function wordOffset(value: string, at: number, backward: boolean): number {
  const next = () => characterOffset(value, at, backward);
  const character = () => value.slice(Math.min(at, next()), Math.max(at, next()));
  while (/\s/u.test(character())) at = next();
  const word = /[\p{L}\p{N}_]/u.test(character());
  while (character() && !/\s/u.test(character()) && /[\p{L}\p{N}_]/u.test(character()) === word) at = next();
  return at;
}

export function replayCellInput(value: string, operations: readonly CellInputOperation[]): CellInputState {
  let anchor = value.length;
  let focus = anchor;
  for (const operation of operations) {
    let start = Math.min(anchor, focus);
    let end = Math.max(anchor, focus);
    if (operation.kind === 'selectAll') { anchor = 0; focus = value.length; continue; }
    if (operation.kind === 'move') {
      const backward = operation.direction === 'left';
      const position = operation.direction === 'start' ? 0 : operation.direction === 'end' ? value.length :
        start !== end && !operation.extend ? backward ? start : end : characterOffset(value, focus, backward);
      if (!operation.extend) anchor = position;
      focus = position;
      continue;
    }
    if (operation.kind === 'delete' && start === end) {
      const backward = operation.direction === 'backward';
      const position = operation.unit === 'line' ? backward ? 0 : value.length :
        operation.unit === 'word' ? wordOffset(value, focus, backward) : characterOffset(value, focus, backward);
      start = Math.min(focus, position);
      end = Math.max(focus, position);
    }
    const text = operation.kind === 'insert' ? operation.text : '';
    value = value.slice(0, start) + text + value.slice(end);
    anchor = focus = start + text.length;
  }
  return { value, anchor, focus };
}
