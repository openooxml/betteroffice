import { createRequire } from 'node:module';
import type * as Wasm from '../../../packages/docx/src/wasm/generated/query/docx_query.js';

export type { InitInput, InitOutput, SyncInitInput } from '../../../packages/docx/src/wasm/generated/query/docx_query.js';

type Engine = Pick<
  typeof Wasm,
  | 'close_display_list'
  | 'hit_test_regions_by_handle'
  | 'hit_test_regions_json'
  | 'open_display_list'
  | 'range_rects_by_handle'
  | 'range_rects_json'
  | 'range_rects_on_pages_by_handle'
  | 'range_rects_on_pages_json'
  | 'range_rects_region_by_handle'
  | 'range_rects_region_json'
  | 'update_display_list'
  | 'vertical_move_by_handle'
  | 'vertical_move_json'
>;

const binding = createRequire(import.meta.url)('../index.js') as Engine;

export const close_display_list = binding.close_display_list;
export const hit_test_regions_by_handle = binding.hit_test_regions_by_handle;
export const hit_test_regions_json = binding.hit_test_regions_json;
export const open_display_list = binding.open_display_list;
export const range_rects_by_handle = binding.range_rects_by_handle;
export const range_rects_json = binding.range_rects_json;
export const range_rects_on_pages_by_handle = binding.range_rects_on_pages_by_handle;
export const range_rects_on_pages_json = binding.range_rects_on_pages_json;
export const range_rects_region_by_handle = binding.range_rects_region_by_handle;
export const range_rects_region_json = binding.range_rects_region_json;
export const update_display_list = binding.update_display_list;
export const vertical_move_by_handle = binding.vertical_move_by_handle;
export const vertical_move_json = binding.vertical_move_json;

export const engineKind: 'wasm' | 'native' = 'native';

export const initSync = (() => ({ memory: undefined })) as unknown as typeof Wasm.initSync;

async function initNativeEngine(): Promise<{ memory: undefined }> {
  return { memory: undefined };
}

export default initNativeEngine as unknown as typeof Wasm.default;
