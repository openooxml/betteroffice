import type { PluginSelectionSettleEvent } from '../host-contracts/plugins';

type Source = PluginSelectionSettleEvent<unknown>['source'];

export interface SelectionSettleOptions {
  root: Element | (() => Element | null);
  isEditorTarget?(target: EventTarget | null): boolean;
  settleMs?: number;
  /** Re-notify the same key when this becomes true. */
  canSettle?(): boolean;
  onUnsettle(): void;
  onSettle(source: Source): void;
  onResettle?(source: Source): void;
}

function isTypingKey(event: KeyboardEvent): boolean {
  if (event.altKey && event.key === 'F10') return false;
  if (event.ctrlKey || event.metaKey) return ['v', 'x'].includes(event.key.toLowerCase());
  if (event.altKey) return false;
  return event.key.length === 1 || ['Enter', 'Backspace', 'Delete'].includes(event.key);
}

export function createSelectionSettle({
  root, isEditorTarget, settleMs = 350, canSettle = () => true, onUnsettle, onSettle, onResettle,
}: SelectionSettleOptions) {
  const rootOf = () => typeof root === 'function' ? root() : root;
  const document = rootOf()?.ownerDocument ?? globalThis.document;
  let active = true;
  let pointerDown = false;
  let waiting = false;
  let source: Source = 'keyboard';
  let key: string | undefined;
  let settledKey: string | undefined;
  let timer: ReturnType<typeof setTimeout> | null = null;
  const inside = (target: EventTarget | null) =>
    (target instanceof Node && !!rootOf()?.contains(target)) || !!isEditorTarget?.(target);
  const clear = () => {
    if (timer !== null) clearTimeout(timer);
    timer = null;
  };
  const settle = () => {
    if (!active || pointerDown || !canSettle()) return;
    waiting = false;
    if (key === settledKey) onResettle?.(source);
    else {
      settledKey = key;
      onSettle(source);
    }
  };
  const schedule = () => {
    clear();
    waiting = false;
    if (!active || pointerDown) return;
    timer = setTimeout(() => {
      timer = null;
      waiting = true;
      settle();
    }, settleMs);
  };
  const down = (event: PointerEvent) => {
    if (!inside(event.target)) return;
    pointerDown = true;
    waiting = false;
    source = event.pointerType === 'touch' ? 'touch' : 'pointer';
    clear();
    onUnsettle();
  };
  const up = () => {
    if (!pointerDown) return;
    pointerDown = false;
    schedule();
  };
  const keydown = (event: KeyboardEvent) => {
    if (!inside(event.target)) return;
    source = 'keyboard';
    if (isTypingKey(event) && (timer !== null || waiting)) schedule();
  };
  const window = document.defaultView;
  document.addEventListener('pointerdown', down, true);
  document.addEventListener('pointerup', up, true);
  document.addEventListener('pointercancel', up, true);
  window?.addEventListener('pointerup', up, true);
  window?.addEventListener('pointercancel', up, true);
  window?.addEventListener('blur', up);
  document.addEventListener('lostpointercapture', up, true);
  document.addEventListener('keydown', keydown, true);
  return {
    notifySelectionChange(next: string) {
      if (!active) return;
      if (next === key) {
        if (waiting) settle();
        return;
      }
      waiting = false;
      key = next;
      onUnsettle();
      if (!pointerDown) {
        if (timer === null || (source !== 'pointer' && source !== 'touch')) source = 'keyboard';
        schedule();
      }
    },
    dispose() {
      active = false;
      clear();
      document.removeEventListener('pointerdown', down, true);
      document.removeEventListener('pointerup', up, true);
      document.removeEventListener('pointercancel', up, true);
      window?.removeEventListener('pointerup', up, true);
      window?.removeEventListener('pointercancel', up, true);
      window?.removeEventListener('blur', up);
      document.removeEventListener('lostpointercapture', up, true);
      document.removeEventListener('keydown', keydown, true);
    },
  };
}
