//! Facade: re-export every wasm-bindgen entrypoint packages/docx consumes.

pub use docx_edit::wasm::EditSession;

pub use docx_layout::{
    build_display_list_json, clear_measure_fonts, close_display_list,
    hit_test_regions_by_handle, hit_test_regions_json, layout_document_json,
    measure_paragraph_json, open_display_list, outline_glyph_json, range_rects_by_handle,
    range_rects_json, range_rects_region_by_handle, range_rects_region_json,
    register_measure_font, register_substitute_measure_font, update_display_list,
    vertical_move_by_handle, vertical_move_json,
};

pub use docx_parse::{
    decode_tiff_png, parse_docx_s9, parse_relationships_xml, serialize_docx_s10,
    serialize_docx_s11, serialize_docx_s12, write_docx_s13_wasm,
};

pub use ooxml_opc::{rezip_docx, unzip_docx};
