import { expect, test } from 'bun:test';
import { TextMeasureFontRegistry } from './fontRegistry';

test("an unknown family asks the provider for Word's last-resort face", async () => {
  const offices: unknown[] = [];
  const registry = new TextMeasureFontRegistry(
    { registerFont: () => 3 },
    {
      bundled: {
        resolve: () => undefined,
        resolveLastResort: (_family, _bold, _italic, office) => {
          offices.push(office);
          return async () => new ArrayBuffer(16);
        },
      },
    }
  );
  expect(await registry.getFontIdChain('Lato', false, false)).toEqual([3]);
  expect(offices).toEqual(['word']);
});
