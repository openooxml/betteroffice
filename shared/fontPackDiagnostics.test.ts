import { afterEach, expect, spyOn, test } from 'bun:test';
import { warnFontPackDiagnostics } from './fontPackDiagnostics';

const WARNED = Symbol.for('@betteroffice/font-pack-warnings');

afterEach(() => { delete (globalThis as { [WARNED]?: unknown })[WARNED]; });

test('warns once per pack and process with a code server hosts can filter', () => {
  const emitted = spyOn(process, 'emitWarning').mockImplementation(() => {});
  try {
    const cjk = { pack: 'cjk', reason: 'missing-file', detail: '@betteroffice/fonts-cjk is missing a; its faces are substituted' } as const;
    const intl = { pack: 'intl', reason: 'version-mismatch', detail: '@betteroffice/fonts-intl 0.0.1 does not match' } as const;
    warnFontPackDiagnostics([cjk]);
    warnFontPackDiagnostics([cjk, intl]);
    warnFontPackDiagnostics([]);
    expect(emitted.mock.calls).toEqual([
      [cjk.detail, { type: 'BetterOfficeFontPackWarning', code: 'BETTEROFFICE_FONT_PACK' }],
      [intl.detail, { type: 'BetterOfficeFontPackWarning', code: 'BETTEROFFICE_FONT_PACK' }],
    ]);
  } finally {
    emitted.mockRestore();
  }
});
