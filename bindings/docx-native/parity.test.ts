import { readFileSync } from 'node:fs';
import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import {
  fontPreflight as wasmFontPreflight,
  preloadPreflightWasm,
} from '../../packages/docx/src/wasm/preflight';
import { fontPreflight } from './index.js';

test('native font preflight matches the standalone wasm on synthetic bytes', async () => {
  const zip = new JSZip();
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:b/><w:rFonts w:ascii="Run" w:cs="Complex"/><w:iCs/></w:rPr><w:t>Text العربية</w:t></w:r></w:p></w:body></w:document>');
  zip.file('word/styles.xml', '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Base" w:hAnsi="Base"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>');
  const bytes = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  await preloadPreflightWasm(readFileSync(new URL(
    '../../packages/docx/src/wasm/generated/preflight/docx_preflight_bg.wasm',
    import.meta.url
  )));
  expect(fontPreflight(bytes)).toBe(JSON.stringify(await wasmFontPreflight(bytes)));
});
