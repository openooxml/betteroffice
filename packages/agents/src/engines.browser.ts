import { DocumentToolError } from './types';
import { createEngineLoader } from './engine-loader';
import { wasm } from './engine-importers';

export const engines = createEngineLoader(undefined, wasm, () => {});
export async function wasmBytes(_format: 'xlsx' | 'pptx'): Promise<undefined> { return undefined; }
export async function loadCanvas(): Promise<typeof import('@napi-rs/canvas')> {
  throw new DocumentToolError('RENDER_UNAVAILABLE', 'Previews on this platform need @napi-rs/canvas installed.');
}
