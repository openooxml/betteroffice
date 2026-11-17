/** How the input treats the clipboard event a keystroke is about to cause. */
export type ClipboardGesture = 'plain' | 'ignore';

/** Longest wait between a keydown and the clipboard event it causes. */
const GESTURE_WINDOW_MS = 300;

interface Armed {
  mode: ClipboardGesture;
  disarm(): void;
}

const released = new WeakSet<Event>();
const armed = new WeakMap<EventTarget, Armed>();

/** Tells the text input that `event` no longer means its built-in action. */
export function releaseKeyFromInput(event: Event): void {
  released.add(event);
}

/** Whether the text input must leave `event` alone. */
export function isKeyReleasedFromInput(event: Event): boolean {
  return released.has(event);
}

export function disarmClipboardGesture(target: EventTarget): void {
  armed.get(target)?.disarm();
}

/**
 * Applies `mode` to the next clipboard event on `target` if it follows this keystroke. The flag
 * ends at key release, blur or a short timeout, so a keystroke the host swallowed, or one that
 * produced no clipboard event, can never change a later paste.
 */
export function armClipboardGesture(target: EventTarget, mode: ClipboardGesture): void {
  disarmClipboardGesture(target);
  const disarm = () => {
    clearTimeout(timer);
    target.removeEventListener('keyup', disarm);
    target.removeEventListener('blur', disarm);
    armed.delete(target);
  };
  const timer = setTimeout(disarm, GESTURE_WINDOW_MS);
  target.addEventListener('keyup', disarm);
  target.addEventListener('blur', disarm);
  armed.set(target, { mode, disarm });
}

/** The gesture armed for this clipboard event, if any; reading it disarms. */
export function takeClipboardGesture(target: EventTarget): ClipboardGesture | null {
  const entry = armed.get(target);
  entry?.disarm();
  return entry?.mode ?? null;
}
