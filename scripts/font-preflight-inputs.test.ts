import { expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

// Guards crates/ooxml-font-preflight/src/engine_inputs.rs. A loader input is
// seen only as a string literal in a non-test engine source: a relationship
// type ending in `relationships/<kind>` or `"/<kind>"`, or a part path
// `"ppt|xl|word/<path>.xml"`. Code after the first `#[cfg(test)]` line of a file
// is skipped. Paths built with format! are not seen; the table lists them by hand.
const root = fileURLToPath(new URL('../', import.meta.url));

const ENGINE_CRATES = {
  PPTX: ['pptx-parse', 'pptx-render', 'pptx-edit', 'ooxml-drawingml'],
  XLSX: ['xlsx-parse', 'betteroffice-xlsx', 'xlsx-model', 'xlsx-render'],
  DOCX: ['docx-parse', 'docx-edit', 'docx-layout'],
} as const;

const TEST_FILES = [
  'crates/ooxml-font-preflight/src/tests.rs',
  'crates/ooxml-font-preflight/src/pptx/tests.rs',
  'crates/ooxml-font-preflight/src/xlsx/tests.rs',
  'crates/ooxml-font-preflight/src/engine_inputs.rs',
  'crates/docx-edit/tests/font_preflight_parity.rs',
  'crates/docx-edit/src/engine.rs',
] as const;

const testNames = new Set(TEST_FILES.flatMap((path) =>
  [...readFileSync(join(root, path), 'utf8').matchAll(/#\[test\]\s*(?:pub(?:\(crate\))?\s+)?fn\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(/g)]
    .map((match) => match[1]!),
));

const INPUT = /"[^"]*relationships\/([A-Za-z]+)"|"\/([a-z][A-Za-z]+)"|"((?:ppt|xl|word)\/[A-Za-z_/]+\.xml)"/g;

function sources(directory: string): string[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) return entry.name === 'tests' ? [] : sources(path);
    return entry.name.endsWith('.rs') && !entry.name.includes('test') ? [path] : [];
  });
}

function engineInputs(crates: readonly string[]): Set<string> {
  const found = new Set<string>();
  for (const crate of crates) {
    for (const path of sources(join(root, 'crates', crate, 'src'))) {
      const text = readFileSync(path, 'utf8').split(/^#\[cfg\(test\)\]/m)[0]!;
      for (const match of text.matchAll(INPUT)) found.add((match[1] ?? match[2] ?? match[3])!);
    }
  }
  return found;
}

function table(name: keyof typeof ENGINE_CRATES) {
  const source = readFileSync(join(root, 'crates/ooxml-font-preflight/src/engine_inputs.rs'), 'utf8');
  const body = source.split(`pub(crate) const ${name}: &[(&str, Answer, &[&str])] = &[`)[1]?.split('];')[0];
  expect(body).toBeDefined();
  return [...body!.matchAll(/\(\s*"([^"]+)",\s*([A-Za-z]+),\s*&\[([^\]]*)\]\s*,?\s*\)/g)]
    .map((match) => ({
      input: match[1]!,
      answer: match[2]!,
      tests: [...match[3]!.matchAll(/"([^"]+)"/g)].map((test) => test[1]!),
    }));
}

for (const name of Object.keys(ENGINE_CRATES) as (keyof typeof ENGINE_CRATES)[]) {
  test(`font preflight answers every ${name} input the engine loaders name`, () => {
    const found = engineInputs(ENGINE_CRATES[name]);
    expect(found.size).toBeGreaterThan(10);
    const listed = new Set(table(name).map((row) => row.input));
    expect([...found].filter((input) => !listed.has(input)).sort()).toEqual([]);
  });

  test(`font preflight backs each ${name} answer with existing tests`, () => {
    const rows = table(name);
    expect(rows.filter((row) =>
      ['Modelled', 'Incomplete', 'Refused'].includes(row.answer) && row.tests.length === 0,
    ).map((row) => row.input)).toEqual([]);
    expect(rows.flatMap((row) => row.tests
      .filter((test) => !testNames.has(test))
      .map((test) => `${row.input}: ${test}`),
    )).toEqual([]);
    expect(rows.filter((row) => row.answer === 'Unverified' && row.tests.length > 0)
      .map((row) => row.input)).toEqual([]);
  });
}
