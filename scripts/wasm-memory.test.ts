import { describe, expect, test } from 'bun:test';
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const ROOT = resolve(import.meta.dir, '..');
const PAGE = 64 * 1024;
const GIB_PAGES = (1024 * 1024 * 1024) / PAGE;
// CI builds the modules before this file runs and sets this so a missing module fails.
const REQUIRED = process.env.BETTEROFFICE_REQUIRE_WASM === '1';

interface Limits {
  min: number;
  max?: number;
}

/** The limits of a module's own linear memory, read from its memory section. */
function memoryLimits(bytes: Uint8Array): Limits {
  let at = 8;
  const leb = () => {
    let result = 0;
    let shift = 0;
    let byte: number;
    do {
      byte = bytes[at++];
      result += (byte & 0x7f) * 2 ** shift;
      shift += 7;
    } while (byte & 0x80);
    return result;
  };
  while (at < bytes.length) {
    const id = bytes[at++];
    const size = leb();
    const end = at + size;
    if (id === 5) {
      leb();
      const flags = bytes[at++];
      const min = leb();
      return flags & 1 ? { min, max: leb() } : { min };
    }
    at = end;
  }
  throw new Error('module defines no memory');
}

const ENGINES = [
  { name: 'docx edit', wasm: 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm', glue: 'packages/docx/src/wasm/generated/edit/docx_edit.js' },
  { name: 'docx opc', wasm: 'packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm' },
  { name: 'docx parse', wasm: 'packages/docx/src/wasm/generated/parse/docx_parse_bg.wasm' },
  { name: 'docx preflight', wasm: 'packages/docx/src/wasm/generated/preflight/docx_preflight_bg.wasm' },
  { name: 'docx query', wasm: 'packages/docx/src/wasm/generated/query/docx_query_bg.wasm' },
  { name: 'pptx', wasm: 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm', glue: 'packages/pptx/src/wasm/generated/pptx_wasm.js' },
  { name: 'xlsx', wasm: 'packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', glue: 'packages/xlsx/src/wasm/generated/xlsx_wasm.js' },
  { name: 'xlsx lean', wasm: 'packages/xlsx/src/wasm/generated/xlsx_wasm_lean_bg.wasm' },
  { name: 'vsdx', wasm: 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm', glue: 'packages/vsdx/src/wasm/generated/vsdx_wasm.js' },
];
const DECODER = 'packages/fonts/src/decoder/fontpack.wasm';

function built(path: string): boolean {
  return REQUIRED || existsSync(resolve(ROOT, path));
}

describe('wasm memory limits', () => {
  test('memoryLimits reads a declared maximum and its absence', () => {
    const header = [0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00];
    expect(memoryLimits(new Uint8Array([...header, 0x05, 0x06, 0x01, 0x01, 0x11, 0x80, 0x80, 0x01])))
      .toEqual({ min: 17, max: GIB_PAGES });
    expect(memoryLimits(new Uint8Array([...header, 0x05, 0x03, 0x01, 0x00, 0x25]))).toEqual({ min: 37 });
  });

  for (const engine of ENGINES) {
    test.skipIf(!built(engine.wasm))(`${engine.name} engine memory is not capped below 4 GiB`, () => {
      const { max } = memoryLimits(new Uint8Array(readFileSync(resolve(ROOT, engine.wasm))));
      if (max !== undefined) expect(max).toBeGreaterThanOrEqual(4 * GIB_PAGES);
    });
  }

  for (const engine of ENGINES.filter((entry) => entry.glue)) {
    test.skipIf(!built(engine.wasm))(`${engine.name} engine memory grows past 1 GiB`, async () => {
      const glue = await import(pathToFileURL(resolve(ROOT, engine.glue!)).href);
      const { memory } = glue.initSync({ module: readFileSync(resolve(ROOT, engine.wasm)) }) as { memory: WebAssembly.Memory };
      memory.grow(GIB_PAGES + 16 - memory.buffer.byteLength / PAGE);
      expect(memory.buffer.byteLength).toBeGreaterThan(GIB_PAGES * PAGE);
    });
  }

  test.skipIf(!built(DECODER))('the standalone font decoder keeps its 1 GiB cap', () => {
    expect(memoryLimits(new Uint8Array(readFileSync(resolve(ROOT, DECODER)))).max).toBe(GIB_PAGES);
  });
});
