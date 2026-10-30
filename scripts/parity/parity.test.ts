import { expect, test } from 'bun:test';

import { compareOrWrite, formats, readExpected, run } from './harness';

for (const [format, target] of Object.entries(formats)) {
  test(`${format} scripted API parity matches the shared golden`, async () => {
    if (process.env.BETTEROFFICE_ENGINE !== undefined) {
      expect(target.engineBackend()).toBe(process.env.BETTEROFFICE_ENGINE);
    }
    const results = await run(format);
    compareOrWrite(format, results, process.env.BETTEROFFICE_PARITY_WRITE === '1');
    expect(results).toEqual(readExpected(format));
  });
}
