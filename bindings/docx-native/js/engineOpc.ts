import { createRequire } from 'node:module';
import type * as Wasm from '../../../packages/docx/src/wasm/generated/opc/ooxml_opc.js';

export type { InitInput, InitOutput, SyncInitInput } from '../../../packages/docx/src/wasm/generated/opc/ooxml_opc.js';

type Engine = Pick<
  typeof Wasm,
  | 'rezip_docx'
  | 'sanitizeOoxml'
  | 'unzip_docx'
>;

const binding = createRequire(import.meta.url)('../index.js') as Engine;

export const rezip_docx = binding.rezip_docx;
export const sanitizeOoxml = binding.sanitizeOoxml;
export const unzip_docx = binding.unzip_docx;

export const engineKind: 'wasm' | 'native' = 'native';

export const initSync = (() => ({ memory: undefined })) as unknown as typeof Wasm.initSync;

async function initNativeEngine(): Promise<{ memory: undefined }> {
  return { memory: undefined };
}

export default initNativeEngine as unknown as typeof Wasm.default;
