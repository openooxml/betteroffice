//! Target-neutral DOCX engine boundary shared by the wasm and native backends.

pub use docx_edit::RunInheritance;
pub use docx_edit::engine::RetainedLayoutMeta;
pub use docx_edit::session::{
    EditSessionCore as Document, EngineError, EngineErrorKind, EngineResult, Host,
    UpdateSubscription,
};

pub fn renderer_version() -> &'static str {
    env!("CARGO_PKG_VERSION")
}

pub fn export_docx_structured_json(bytes: &[u8], options: &str) -> EngineResult<String> {
    docx_edit::session::free::export_docx_structured_json(bytes, options)
        .map_err(EngineError::message)
}

pub fn export_docx_markdown_json(bytes: &[u8], options: &str) -> EngineResult<String> {
    docx_edit::session::free::export_docx_markdown_json(bytes, options)
        .map_err(EngineError::message)
}

pub fn list_docx_content_controls_json(bytes: &[u8], options: &str) -> EngineResult<String> {
    docx_edit::session::free::list_docx_content_controls_json(bytes, options)
        .map_err(EngineError::message)
}

pub fn find_docx_content_controls_json(
    bytes: &[u8],
    query: &str,
    options: &str,
) -> EngineResult<String> {
    docx_edit::session::free::find_docx_content_controls_json(bytes, query, options)
        .map_err(EngineError::message)
}

pub fn render_docx_markdown_json(content: &str, options: &str) -> EngineResult<String> {
    docx_edit::session::free::render_docx_markdown_json(content, options)
        .map_err(EngineError::message)
}

pub fn render_docx_markdown_with_pages_json(content: &str, options: &str) -> EngineResult<String> {
    docx_edit::session::free::render_docx_markdown_with_pages_json(content, options)
        .map_err(EngineError::message)
}

pub fn build_display_list_json(input: &str) -> EngineResult<String> {
    docx_layout::build_display_list_json_str(input).map_err(EngineError::message)
}

pub fn layout_document_json(input: &str) -> EngineResult<String> {
    docx_layout::layout_document_json_str(input).map_err(EngineError::message)
}

pub fn measure_paragraph_json(input: &str) -> EngineResult<String> {
    docx_layout::measure_paragraph_json_str(input).map_err(EngineError::message)
}

pub fn register_measure_font(bytes: &[u8]) -> EngineResult<u32> {
    docx_layout::register_measure_font_module(bytes).map_err(EngineError::message)
}

pub fn register_substitute_measure_font(base: u32, requested_family: &str) -> EngineResult<u32> {
    docx_layout::register_substitute_measure_font_module(base, requested_family)
        .map_err(EngineError::message)
}

pub fn clear_measure_fonts() -> EngineResult<()> {
    docx_layout::clear_measure_fonts_module().map_err(EngineError::message)
}

pub fn outline_glyph_json(font_id: u32, glyph_id: u32) -> EngineResult<String> {
    docx_layout::outline_glyph_json_str(font_id, glyph_id).map_err(EngineError::message)
}

pub fn open_display_list(display_list: &str) -> EngineResult<u32> {
    docx_layout::session::open_display_list(display_list).map_err(EngineError::message)
}

pub fn close_display_list(handle: u32) {
    docx_layout::session::close_display_list(handle)
}

pub fn update_display_list(handle: u32, update: &str) -> EngineResult<()> {
    docx_layout::session::update_display_list(handle, update).map_err(EngineError::message)
}

pub fn hit_test_json(display_list: &str, page_index: u32, x: f64, y: f64) -> EngineResult<String> {
    docx_layout::hit::hit_test_json(display_list, page_index as usize, x, y)
        .map_err(EngineError::message)
}

pub fn hit_test_regions_json(
    display_list: &str,
    page_index: u32,
    x: f64,
    y: f64,
) -> EngineResult<String> {
    docx_layout::hit::hit_test_regions_json(display_list, page_index as usize, x, y)
        .map_err(EngineError::message)
}

pub fn hit_test_regions_by_handle(
    handle: u32,
    page_index: u32,
    x: f64,
    y: f64,
) -> EngineResult<String> {
    docx_layout::session::hit_test_regions_by_handle(handle, page_index as usize, x, y)
        .map_err(EngineError::message)
}

pub fn range_rects_json(display_list: &str, from: f64, to: f64) -> EngineResult<String> {
    docx_layout::hit::range_rects_json(display_list, from as i64, to as i64)
        .map_err(EngineError::message)
}

pub fn range_rects_by_handle(handle: u32, from: f64, to: f64) -> EngineResult<String> {
    docx_layout::session::range_rects_by_handle(handle, from as i64, to as i64)
        .map_err(EngineError::message)
}

pub fn range_rects_on_pages_json(
    display_list: &str,
    from: f64,
    to: f64,
    first_page: f64,
    last_page: f64,
) -> EngineResult<String> {
    docx_layout::range_rects_on_pages_json_str(display_list, from, to, first_page, last_page)
        .map_err(EngineError::message)
}

pub fn range_rects_on_pages_by_handle(
    handle: u32,
    from: f64,
    to: f64,
    first_page: f64,
    last_page: f64,
) -> EngineResult<String> {
    docx_layout::range_rects_on_pages_by_handle_str(handle, from, to, first_page, last_page)
        .map_err(EngineError::message)
}

pub fn range_rects_region_json(
    display_list: &str,
    region: &str,
    part_id: &str,
    from: f64,
    to: f64,
) -> EngineResult<String> {
    docx_layout::hit::range_rects_region_json(display_list, region, part_id, from as i64, to as i64)
        .map_err(EngineError::message)
}

pub fn range_rects_region_by_handle(
    handle: u32,
    region: &str,
    part_id: &str,
    from: f64,
    to: f64,
) -> EngineResult<String> {
    docx_layout::session::range_rects_region_by_handle(
        handle,
        region,
        part_id,
        from as i64,
        to as i64,
    )
    .map_err(EngineError::message)
}

pub fn vertical_move_json(
    display_list: &str,
    position: f64,
    direction: &str,
    goal_x: f64,
) -> EngineResult<String> {
    docx_layout::hit::vertical_move_json(display_list, position as i64, direction, goal_x)
        .map_err(EngineError::message)
}

pub fn vertical_move_by_handle(
    handle: u32,
    position: f64,
    direction: &str,
    goal_x: f64,
) -> EngineResult<String> {
    docx_layout::session::vertical_move_by_handle(handle, position as i64, direction, goal_x)
        .map_err(EngineError::message)
}

pub fn parse_docx_relationships(data: &[u8]) -> EngineResult<String> {
    docx_parse::parse_docx_relationships_json(data).map_err(EngineError::message)
}

pub fn parse_docx_s2(data: &[u8]) -> EngineResult<String> {
    docx_parse::parse_docx_s2_json(data).map_err(EngineError::message)
}

pub fn parse_docx_s3(data: &[u8]) -> EngineResult<String> {
    docx_parse::parse_docx_s3_json(data).map_err(EngineError::message)
}

pub fn parse_docx_s4(data: &[u8]) -> EngineResult<String> {
    docx_parse::parse_docx_s4_json(data).map_err(EngineError::message)
}

pub fn parse_docx_s5(data: &[u8]) -> EngineResult<String> {
    docx_parse::parse_docx_s5_json(data).map_err(EngineError::message)
}

pub fn parse_docx_s6(data: &[u8]) -> EngineResult<String> {
    docx_parse::parse_docx_s6_json(data).map_err(EngineError::message)
}

pub fn parse_docx_s7(data: &[u8]) -> EngineResult<String> {
    docx_parse::parse_docx_s7_json(data).map_err(EngineError::message)
}

pub fn parse_docx_s8(data: &[u8]) -> EngineResult<String> {
    docx_parse::parse_docx_s8_json(data).map_err(EngineError::message)
}

pub fn parse_docx_s9(data: &[u8], options_json: &str) -> EngineResult<String> {
    docx_parse::parse_docx_s9_json(data, options_json).map_err(EngineError::message)
}

pub fn parse_relationships_xml(xml: &[u8], part_path: &str) -> EngineResult<String> {
    docx_parse::parse_relationships_xml_json(xml, part_path).map_err(EngineError::message)
}

pub fn serialize_docx_s10(request_json: &str) -> EngineResult<String> {
    docx_parse::serialize_docx_s10_json(request_json).map_err(EngineError::message)
}

pub fn serialize_docx_s11(request_json: &str) -> EngineResult<String> {
    docx_parse::serialize_docx_s11_json(request_json).map_err(EngineError::message)
}

pub fn serialize_docx_s12(request_json: &str) -> EngineResult<String> {
    docx_parse::serialize_docx_s12_json(request_json).map_err(EngineError::message)
}

pub fn paragraph_patch_writer_version() -> u32 {
    docx_parse::paragraph_patch_writer_version_native()
}

pub fn write_docx_s13_wasm(request_json: &str, original_docx: &[u8]) -> EngineResult<Vec<u8>> {
    docx_parse::write_docx_s13_json(request_json, original_docx).map_err(EngineError::message)
}

pub fn write_docx_s13_wasm_with_warnings(
    request_json: &str,
    original_docx: &[u8],
) -> EngineResult<Vec<u8>> {
    docx_parse::write_docx_s13_with_warnings_json(request_json, original_docx)
        .map_err(EngineError::message)
}

pub fn decode_tiff_png(data: &[u8]) -> EngineResult<Vec<u8>> {
    docx_parse::decode_tiff_png_bytes(data).map_err(EngineError::message)
}

pub fn unzip_docx(data: &[u8]) -> EngineResult<Vec<(String, Vec<u8>)>> {
    ooxml_opc::unzip_parts(data).map_err(EngineError::message)
}

pub fn rezip_docx(entries: Vec<(String, Vec<u8>)>) -> EngineResult<Vec<u8>> {
    ooxml_opc::rezip_parts(&entries).map_err(EngineError::message)
}

pub fn sanitize_ooxml(data: &[u8], expected_format: &str) -> EngineResult<Vec<u8>> {
    ooxml_opc::sanitize_package_for_format(data, expected_format).map_err(EngineError::message)
}
