import { expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));

const ENGINE_CRATES = {
  PPTX: ['pptx-parse', 'pptx-render', 'pptx-edit', 'ooxml-drawingml'],
  XLSX: ['xlsx-parse', 'betteroffice-xlsx', 'xlsx-model', 'xlsx-render'],
  DOCX: ['docx-parse', 'docx-edit', 'docx-layout'],
} as const;

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

function table(name: keyof typeof ENGINE_CRATES): Set<string> {
  const source = readFileSync(join(root, 'crates/ooxml-font-preflight/src/engine_inputs.rs'), 'utf8');
  const body = source.split(`pub(crate) const ${name}: &[(&str, Answer)] = &[`)[1]?.split('];')[0];
  expect(body).toBeDefined();
  return new Set([...body!.matchAll(/\("([^"]+)", [A-Za-z]+\)/g)].map((match) => match[1]!));
}

for (const name of Object.keys(ENGINE_CRATES) as (keyof typeof ENGINE_CRATES)[]) {
  test(`font preflight answers every ${name} input the engine loaders name`, () => {
    const found = engineInputs(ENGINE_CRATES[name]);
    expect(found.size).toBeGreaterThan(10);
    const listed = table(name);
    expect([...found].filter((input) => !listed.has(input)).sort()).toEqual([]);
  });
}
