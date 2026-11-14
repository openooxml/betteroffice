import { expect, test } from 'bun:test';
import { embeddedFontViolations, lastResortFamilyMatches } from './check-embedded-fonts.mjs';

const embedded = 'const FONT: &[u8] = include_bytes!("../assets/font.ttf");';

test('production font embeds are restricted to the shared crate and migration allowlist', () => {
  expect(embeddedFontViolations('crates/docx-edit/src/fonts.rs', embedded)).toHaveLength(1);
  expect(embeddedFontViolations('crates/base-latin/src/lib.rs',
    'include_bytes!("../fonts/BetterOfficeBaseLatin-Regular.ttf")')).toEqual([]);
  expect(embeddedFontViolations('crates/docx-edit/src/fonts.rs',
    'include_bytes!("../../base-latin/fonts/BetterOfficeBaseLatin-Regular.ttf")')).toHaveLength(1);
  expect(embeddedFontViolations('crates/xlsx-raster/src/font.rs',
    'include_bytes!("../assets/Carlito-Regular.ttf")')).toEqual([]);
  expect(embeddedFontViolations('crates/xlsx-raster/src/font.rs', embedded)).toHaveLength(1);
});

test('test, example and benchmark fixtures are excluded', () => {
  for (const file of ['crates/a/tests/font.rs', 'apps/a/benches/font.rs',
    'crates/a/examples/font.rs', 'crates/a/src/font_tests.rs', 'crates/a/src/session/tests.rs']) {
    expect(embeddedFontViolations(file, embedded)).toEqual([]);
  }
  expect(embeddedFontViolations('crates/a/src/lib.rs', `#[cfg(test)]\nmod tests { ${embedded} }`)).toEqual([]);
  expect(embeddedFontViolations('crates/a/src/lib.rs', `#[cfg(all(test, feature = "wasm"))]\nmod tests { ${embedded} }`)).toEqual([]);
  expect(embeddedFontViolations('crates/a/src/lib.rs', `${embedded}\n#[cfg(test)]\nmod tests {}`)).toHaveLength(1);
});

test('the manifest last-resort name must match the embedded family', () => {
  const rust = 'pub const FAMILY: &str = "BetterOffice Base Latin";';
  expect(lastResortFamilyMatches(rust, "export const BASE_LATIN_FAMILY = 'BetterOffice Base Latin';")).toBe(true);
  expect(lastResortFamilyMatches(rust, "export const BASE_LATIN_FAMILY = 'Other';")).toBe(false);
  expect(lastResortFamilyMatches('', '')).toBe(false);
});
