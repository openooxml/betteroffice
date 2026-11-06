import { describe, expect, test } from 'bun:test';
import {
  armClipboardGesture,
  isKeyReleasedFromInput,
  releaseKeyFromInput,
  takeClipboardGesture,
} from './nativeGesture';


describe('clipboard gestures', () => {
  test('a gesture applies once', () => {
    const target = new EventTarget();
    armClipboardGesture(target, 'plain');
    expect(takeClipboardGesture(target)).toBe('plain');
    expect(takeClipboardGesture(target)).toBeNull();
  });

  test('key release and blur end a gesture no clipboard event followed', () => {
    for (const type of ['keyup', 'blur']) {
      const target = new EventTarget();
      armClipboardGesture(target, 'ignore');
      target.dispatchEvent(new Event(type));
      expect(takeClipboardGesture(target)).toBeNull();
    }
  });

  test('a timeout ends a gesture when no key release arrives', async () => {
    const target = new EventTarget();
    armClipboardGesture(target, 'ignore');
    await Bun.sleep(350);
    expect(takeClipboardGesture(target)).toBeNull();
  });

  test('gestures stay with their own input', () => {
    const first = new EventTarget();
    const second = new EventTarget();
    armClipboardGesture(first, 'plain');
    expect(takeClipboardGesture(second)).toBeNull();
    expect(takeClipboardGesture(first)).toBe('plain');
  });

  test('releases are per keystroke', () => {
    const event = new Event('keydown');
    releaseKeyFromInput(event);
    expect(isKeyReleasedFromInput(event)).toBe(true);
    expect(isKeyReleasedFromInput(new Event('keydown'))).toBe(false);
  });
});
