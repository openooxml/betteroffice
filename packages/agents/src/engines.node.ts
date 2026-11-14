import process from 'node:process';
import { DocumentToolError } from './types';
import { readFile } from 'node:fs/promises';
import { createEngineLoader, type EngineModules } from './engine-loader';
import { wasm } from './engine-importers';

const nativePackage = <T>(specifier: string): Promise<T> => import(specifier);
const native = {
  docx: async () => {
    const [yrs, editor, layout, utils, render] = await Promise.all([nativePackage<EngineModules['docx']['yrs']>('@betteroffice/docx-native/yrs'), nativePackage<EngineModules['docx']['editor']>('@betteroffice/docx-native/editor'), nativePackage<EngineModules['docx']['layout']>('@betteroffice/docx-native/layout'), nativePackage<EngineModules['docx']['utils']>('@betteroffice/docx-native/utils'), nativePackage<EngineModules['docx']['render']>('@betteroffice/docx-native/layout/render')]);
    return { yrs, editor, layout, utils, render };
  },
  xlsx: () => nativePackage<EngineModules['xlsx']>('@betteroffice/xlsx-native'),
  pptx: () => nativePackage<EngineModules['pptx']>('@betteroffice/pptx-native'),
};
export const engines = createEngineLoader(native, wasm, warning => process.emitWarning(warning));

export async function wasmBytes(format: 'xlsx' | 'pptx'): Promise<Uint8Array> {
  const entry = format === 'xlsx' ? import.meta.resolve('@betteroffice/xlsx') : import.meta.resolve('@betteroffice/pptx');
  const directory = entry.endsWith('/src/index.ts') ? './wasm/generated' : './generated';
  return readFile(new URL(`${directory}/${format}_wasm_bg.wasm`, entry));
}

export async function loadCanvas() {
  try { return await import('@napi-rs/canvas'); }
  catch { throw new DocumentToolError('RENDER_UNAVAILABLE', 'Previews on this platform need @napi-rs/canvas installed.'); }
}
