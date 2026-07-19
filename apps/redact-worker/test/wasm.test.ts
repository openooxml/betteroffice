import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";

import { initSync, sanitizeOoxml } from "../src/wasm/generated/ooxml_opc.js";

const wasmBytes = readBytes(
  new URL("../src/wasm/generated/ooxml_opc_bg.wasm", import.meta.url),
);
initSync({ module: new WebAssembly.Module(wasmBytes) });

const cases = [
  ["docx", "../../../apps/demo/public/betteroffice-demo.docx"],
  ["xlsx", "../../../apps/demo/public/sample.xlsx"],
  ["pptx", "../../../apps/demo/public/betteroffice-demo.pptx"],
] as const;

test("generated OPC WASM sanitizes all real demo formats", () => {
  for (const [format, relativePath] of cases) {
    const source = readBytes(new URL(relativePath, import.meta.url));
    const sanitized = sanitizeOoxml(source, format);
    expect(sanitized.byteLength).toBeGreaterThan(0);
    expect(Array.from(sanitized.subarray(0, 2))).toEqual([0x50, 0x4b]);
    expect(sanitized).not.toEqual(source);
  }
});

test("generated OPC WASM rejects a mismatched format", () => {
  const source = readBytes(new URL(cases[1][1], import.meta.url));
  expect(() => sanitizeOoxml(source, "docx")).toThrow();
});

function readBytes(path: string | URL): Uint8Array<ArrayBuffer> {
  const source = readFileSync(path);
  const bytes = new Uint8Array(source.byteLength);
  bytes.set(source);
  return bytes;
}
