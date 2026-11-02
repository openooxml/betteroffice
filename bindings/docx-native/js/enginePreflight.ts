import { createRequire } from 'node:module';
import type * as Wasm from '../../../packages/docx/src/wasm/generated/preflight/docx_preflight.js';

export type { InitInput, InitOutput, SyncInitInput } from '../../../packages/docx/src/wasm/generated/preflight/docx_preflight.js';

type Engine = Pick<
  typeof Wasm,
  | 'font_preflight'
>;

const binding = createRequire(import.meta.url)('../index.js') as Engine;

export const font_preflight = binding.font_preflight;

export const engineKind: 'wasm' | 'native' = 'native';

export const initSync = (() => ({ memory: undefined })) as unknown as typeof Wasm.initSync;

async function initNativeEngine(): Promise<{ memory: undefined }> {
  return { memory: undefined };
}

export default initNativeEngine as unknown as typeof Wasm.default;
