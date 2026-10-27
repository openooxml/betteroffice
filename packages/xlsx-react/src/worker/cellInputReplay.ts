export type CellInputOperation = { kind: 'insert'; text: string } | { kind: 'selectAll' } | { kind: 'select'; anchor: number; focus: number } |
  { kind: 'move'; direction: 'start' | 'end' | 'left' | 'right'; extend?: boolean; unit?: 'word' } |
  { kind: 'delete'; direction: 'backward' | 'forward'; unit: 'character' | 'word' | 'line' };

export interface CellInputState {
  value: string;
  anchor: number;
  focus: number;
}

type CellInputEvent = Pick<InputEvent, 'inputType' | 'data' | 'dataTransfer' | 'isComposing'> |
  { key: string; ctrlKey: boolean; metaKey: boolean; altKey: boolean; shiftKey: boolean };

export function cellInputOperation(event: CellInputEvent, mac: boolean): CellInputOperation | null {
  if ('inputType' in event) {
    if (event.isComposing || event.inputType === 'insertCompositionText') return null;
    if (['insertText', 'insertReplacementText', 'insertFromPaste', 'insertFromDrop'].includes(event.inputType)) {
      return { kind: 'insert', text: event.data ?? event.dataTransfer?.getData('text/plain') ?? '' };
    }
    const match = /^delete(Content|Word|SoftLine|HardLine)(Backward|Forward)$/.exec(event.inputType);
    return match ? { kind: 'delete', direction: match[2] === 'Backward' ? 'backward' : 'forward',
      unit: match[1] === 'Content' ? 'character' : match[1] === 'Word' ? 'word' : 'line' } : null;
  }
  const mod = event.ctrlKey || event.metaKey;
  const word = mac ? event.altKey && !event.metaKey : event.ctrlKey && !event.altKey;
  const line = mac && event.metaKey && !event.altKey;
  if (mod && !event.altKey && event.key.toLowerCase() === 'a') return { kind: 'selectAll' };
  if (!mod && event.key.length === 1) return { kind: 'insert', text: event.key };
  if (!mac && event.altKey) return null;
  const direction = event.key === 'Home' || line && event.key === 'ArrowLeft' ? 'start' :
    event.key === 'End' || line && event.key === 'ArrowRight' ? 'end' :
    event.key === 'ArrowLeft' ? 'left' : event.key === 'ArrowRight' ? 'right' : null;
  if (direction) return { kind: 'move', direction, extend: event.shiftKey, ...(word ? { unit: 'word' } : {}) };
  if (event.key === 'Backspace' || event.key === 'Delete') return { kind: 'delete',
    direction: event.key === 'Backspace' ? 'backward' : 'forward', unit: line ? 'line' : word ? 'word' : 'character' };
  return null;
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

export function replayCellInput(value: string, operations: readonly CellInputOperation[],
  selection?: Pick<CellInputState, 'anchor' | 'focus'>): CellInputState {
  let anchor = Math.max(0, Math.min(value.length, selection?.anchor ?? value.length));
  let focus = Math.max(0, Math.min(value.length, selection?.focus ?? value.length));
  for (const operation of operations) {
    let start = Math.min(anchor, focus);
    let end = Math.max(anchor, focus);
    if (operation.kind === 'select') {
      anchor = Math.max(0, Math.min(value.length, operation.anchor));
      focus = Math.max(0, Math.min(value.length, operation.focus));
      continue;
    }
    if (operation.kind === 'selectAll') { anchor = 0; focus = value.length; continue; }
    if (operation.kind === 'move') {
      const backward = operation.direction === 'left';
      const position = operation.direction === 'start' ? 0 : operation.direction === 'end' ? value.length :
        start !== end && !operation.extend ? backward ? start : end :
          operation.unit === 'word' ? wordOffset(value, focus, backward) : characterOffset(value, focus, backward);
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
