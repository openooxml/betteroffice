import { expect, test } from 'bun:test';
import { canvasKeyboardIntent, isOwnedBrowserShortcut } from './interactions';

test('Ctrl+S saves the diagram and is owned so the browser never saves the page', () => {
  expect(canvasKeyboardIntent({ key: 's', ctrlKey: true }, 1)).toEqual({ kind: 'save' });
  expect(canvasKeyboardIntent({ key: 'S', metaKey: true }, 1)).toEqual({ kind: 'save' });
  expect(isOwnedBrowserShortcut({ key: 's', ctrlKey: true })).toBe(true);
  expect(isOwnedBrowserShortcut({ key: 'S', metaKey: true })).toBe(true);
  expect(canvasKeyboardIntent({ key: 's', ctrlKey: true, altKey: true }, 1)).toBeNull();
  expect(isOwnedBrowserShortcut({ key: 's', ctrlKey: true, altKey: true })).toBe(false);
});

test('reload, address bar and the other browser shortcuts stay with the browser', () => {
  for (const key of ['r', 'l', 'f', 'p', 'w', 't']) {
    expect(canvasKeyboardIntent({ key, ctrlKey: true }, 1)).toBeNull();
    expect(canvasKeyboardIntent({ key, ctrlKey: true, shiftKey: true }, 1)).toBeNull();
    expect(isOwnedBrowserShortcut({ key, ctrlKey: true })).toBe(false);
  }
});

test('an editable target yields no intent, and Ctrl+S stays owned there so a draft survives', () => {
  const input = { tagName: 'INPUT' };
  expect(canvasKeyboardIntent({ key: 's', ctrlKey: true, target: input }, 1)).toBeNull();
  expect(canvasKeyboardIntent({ key: 'z', ctrlKey: true, target: input }, 1)).toBeNull();
  expect(isOwnedBrowserShortcut({ key: 's', ctrlKey: true, target: input })).toBe(true);
});
