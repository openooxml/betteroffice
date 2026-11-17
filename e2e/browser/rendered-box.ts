import { expect, type Locator } from 'playwright/test';

type Box = NonNullable<Awaited<ReturnType<Locator['boundingBox']>>>;

/** The box of an element the editor may re-mount between frames, read once it is rendered. */
export async function renderedBox(locator: Locator, timeout?: number): Promise<Box> {
  let box: Box | null = null;
  await expect.poll(async () => (box = await locator.boundingBox()), { timeout }).not.toBeNull();
  return box!;
}

export async function caretBoxAtRunStart(caret: Locator, run: Locator, press: () => Promise<void>): Promise<Box> {
  await press();
  let box: Box | null = null;
  await expect.poll(async () => {
    const target = await run.boundingBox();
    const current = await caret.boundingBox();
    if (!target || !current) return null;
    const aligned = Math.abs(current.x - target.x) < 2 &&
      Math.abs(current.y + current.height / 2 - target.y - target.height / 2) < target.height / 2;
    if (aligned) box = current;
    return aligned ? current : null;
  }).not.toBeNull();
  return box!;
}

export async function movedCaretBox(caret: Locator, press: () => Promise<void>): Promise<Box> {
  const before = await renderedBox(caret);
  await press();
  let box: Box | null = null;
  await expect.poll(async () => {
    box = await caret.boundingBox();
    return box && (box.x !== before.x || box.y !== before.y) ? box : null;
  }).not.toBeNull();
  return box!;
}
