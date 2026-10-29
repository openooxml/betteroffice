import { expect, spyOn, test } from 'bun:test';
import { writeStartupDiagnostics } from '../src/diagnostics';

for (const backend of ['native', 'wasm'] as const) {
  test(`startup reports the ${backend} XLSX engine only on stderr`, () => {
    const stderr = spyOn(process.stderr, 'write').mockReturnValue(true);
    const stdout = spyOn(process.stdout, 'write').mockReturnValue(true);
    try {
      writeStartupDiagnostics(backend);
      expect(stderr).toHaveBeenCalledTimes(1);
      expect(stderr).toHaveBeenCalledWith(`betteroffice-agents: xlsx engine ${backend}\n`);
      expect(stdout).not.toHaveBeenCalled();
    } finally {
      stderr.mockRestore();
      stdout.mockRestore();
    }
  });
}
