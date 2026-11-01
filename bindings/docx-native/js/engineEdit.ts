import { createRequire } from 'node:module';
import type * as Wasm from '../../../packages/docx/src/wasm/generated/edit/docx_edit.js';

export type { DecodedFontFace, InitInput, InitOutput, SyncInitInput } from '../../../packages/docx/src/wasm/generated/edit/docx_edit.js';

type Engine = Pick<
  typeof Wasm,
  | 'EditSession'
  | 'RetainedLayoutMeta'
  | 'RunInheritance'
  | 'build_display_list_json'
  | 'clear_measure_fonts'
  | 'close_display_list'
  | 'decodeFontFace'
  | 'decodeTiffPng'
  | 'export_docx_markdown_json'
  | 'export_docx_structured_json'
  | 'find_docx_content_controls_json'
  | 'hit_test_json'
  | 'hit_test_regions_by_handle'
  | 'hit_test_regions_json'
  | 'install_panic_hook'
  | 'layout_document_json'
  | 'list_docx_content_controls_json'
  | 'measure_paragraph_json'
  | 'open_display_list'
  | 'outline_glyph_json'
  | 'paragraphPatchWriterVersion'
  | 'parse_docx_relationships'
  | 'parse_docx_s2'
  | 'parse_docx_s3'
  | 'parse_docx_s4'
  | 'parse_docx_s5'
  | 'parse_docx_s6'
  | 'parse_docx_s7'
  | 'parse_docx_s8'
  | 'parse_docx_s9'
  | 'parse_relationships_xml'
  | 'range_rects_by_handle'
  | 'range_rects_json'
  | 'range_rects_on_pages_by_handle'
  | 'range_rects_on_pages_json'
  | 'range_rects_region_by_handle'
  | 'range_rects_region_json'
  | 'register_measure_font'
  | 'register_substitute_measure_font'
  | 'render_docx_markdown_json'
  | 'render_docx_markdown_with_pages_json'
  | 'reset_wasm_peak_bytes'
  | 'rezip_docx'
  | 'sanitizeOoxml'
  | 'serialize_docx_s10'
  | 'serialize_docx_s11'
  | 'serialize_docx_s12'
  | 'set_wasm_heap_limit'
  | 'unzip_docx'
  | 'update_display_list'
  | 'vertical_move_by_handle'
  | 'vertical_move_json'
  | 'wasm_failed_allocation_bytes'
  | 'wasm_heap_counted'
  | 'wasm_live_bytes'
  | 'wasm_peak_bytes'
  | 'write_docx_s13_wasm'
  | 'write_docx_s13_wasm_with_warnings'
>;

const binding = createRequire(import.meta.url)('../index.js') as Engine;

export const EditSession = binding.EditSession;
export type EditSession = Wasm.EditSession;
export const RetainedLayoutMeta = binding.RetainedLayoutMeta;
export type RetainedLayoutMeta = Wasm.RetainedLayoutMeta;
export const RunInheritance = binding.RunInheritance;
export type RunInheritance = Wasm.RunInheritance;
export const build_display_list_json = binding.build_display_list_json;
export const clear_measure_fonts = binding.clear_measure_fonts;
export const close_display_list = binding.close_display_list;
export const decodeFontFace = binding.decodeFontFace;
export const decodeTiffPng = binding.decodeTiffPng;
export const export_docx_markdown_json = binding.export_docx_markdown_json;
export const export_docx_structured_json = binding.export_docx_structured_json;
export const find_docx_content_controls_json = binding.find_docx_content_controls_json;
export const hit_test_json = binding.hit_test_json;
export const hit_test_regions_by_handle = binding.hit_test_regions_by_handle;
export const hit_test_regions_json = binding.hit_test_regions_json;
export const install_panic_hook = binding.install_panic_hook;
export const layout_document_json = binding.layout_document_json;
export const list_docx_content_controls_json = binding.list_docx_content_controls_json;
export const measure_paragraph_json = binding.measure_paragraph_json;
export const open_display_list = binding.open_display_list;
export const outline_glyph_json = binding.outline_glyph_json;
export const paragraphPatchWriterVersion = binding.paragraphPatchWriterVersion;
export const parse_docx_relationships = binding.parse_docx_relationships;
export const parse_docx_s2 = binding.parse_docx_s2;
export const parse_docx_s3 = binding.parse_docx_s3;
export const parse_docx_s4 = binding.parse_docx_s4;
export const parse_docx_s5 = binding.parse_docx_s5;
export const parse_docx_s6 = binding.parse_docx_s6;
export const parse_docx_s7 = binding.parse_docx_s7;
export const parse_docx_s8 = binding.parse_docx_s8;
export const parse_docx_s9 = binding.parse_docx_s9;
export const parse_relationships_xml = binding.parse_relationships_xml;
export const range_rects_by_handle = binding.range_rects_by_handle;
export const range_rects_json = binding.range_rects_json;
export const range_rects_on_pages_by_handle = binding.range_rects_on_pages_by_handle;
export const range_rects_on_pages_json = binding.range_rects_on_pages_json;
export const range_rects_region_by_handle = binding.range_rects_region_by_handle;
export const range_rects_region_json = binding.range_rects_region_json;
export const register_measure_font = binding.register_measure_font;
export const register_substitute_measure_font = binding.register_substitute_measure_font;
export const render_docx_markdown_json = binding.render_docx_markdown_json;
export const render_docx_markdown_with_pages_json = binding.render_docx_markdown_with_pages_json;
export const reset_wasm_peak_bytes = binding.reset_wasm_peak_bytes;
export const rezip_docx = binding.rezip_docx;
export const sanitizeOoxml = binding.sanitizeOoxml;
export const serialize_docx_s10 = binding.serialize_docx_s10;
export const serialize_docx_s11 = binding.serialize_docx_s11;
export const serialize_docx_s12 = binding.serialize_docx_s12;
export const set_wasm_heap_limit = binding.set_wasm_heap_limit;
export const unzip_docx = binding.unzip_docx;
export const update_display_list = binding.update_display_list;
export const vertical_move_by_handle = binding.vertical_move_by_handle;
export const vertical_move_json = binding.vertical_move_json;
export const wasm_failed_allocation_bytes = binding.wasm_failed_allocation_bytes;
export const wasm_heap_counted = binding.wasm_heap_counted;
export const wasm_live_bytes = binding.wasm_live_bytes;
export const wasm_peak_bytes = binding.wasm_peak_bytes;
export const write_docx_s13_wasm = binding.write_docx_s13_wasm;
export const write_docx_s13_wasm_with_warnings = binding.write_docx_s13_wasm_with_warnings;

export const engineKind: 'wasm' | 'native' = 'native';

export const initSync = (() => ({ memory: undefined })) as unknown as typeof Wasm.initSync;

async function initNativeEngine(): Promise<{ memory: undefined }> {
  return { memory: undefined };
}

export default initNativeEngine as unknown as typeof Wasm.default;
