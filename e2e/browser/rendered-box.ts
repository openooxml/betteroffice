import { expect, type Locator } from 'playwright/test';

type Box = NonNullable<Awaited<ReturnType<Locator['boundingBox']>>>;

/** The box of an element the editor may re-mount between frames, read once it is rendered. */
export async function renderedBox(locator: Locator, timeout?: number): Promise<Box> {
  let box: Box | null = null;
  await expect.poll(async () => (box = await locator.boundingBox()), { timeout }).not.toBeNull();
  return box!;
}
