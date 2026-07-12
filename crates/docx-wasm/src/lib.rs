//! DOCX wasm boundary.

mod core;

use js_sys::{Object, Reflect, Uint8Array};
use wasm_bindgen::prelude::*;

fn js_error(error: String) -> JsValue {
    JsValue::from_str(&error)
}

#[wasm_bindgen]
pub fn unzip_docx(data: &[u8]) -> Result<JsValue, JsValue> {
    let parts = ooxml_opc::unzip_parts(data).map_err(js_error)?;
    let out = Object::new();
    for (name, bytes) in parts {
        Reflect::set(
            &out,
            &JsValue::from_str(&name),
            &Uint8Array::from(bytes.as_slice()),
        )?;
    }
    Ok(out.into())
}

#[wasm_bindgen]
pub fn rezip_docx(entries: JsValue) -> Result<Vec<u8>, JsValue> {
    let object: Object = entries
        .dyn_into()
        .map_err(|_| JsValue::from_str("rezip_docx: expected an object"))?;
    let mut parts = Vec::new();
    for key in Object::keys(&object).iter() {
        let name = key
            .as_string()
            .ok_or_else(|| JsValue::from_str("rezip_docx: non-string key"))?;
        let value = Reflect::get(&object, &key)?;
        parts.push((name, Uint8Array::new(&value).to_vec()));
    }
    ooxml_opc::rezip_parts(&parts).map_err(js_error)
}

#[wasm_bindgen]
pub fn layout_document_json(input: &str) -> Result<String, JsValue> {
    docx_layout::layout_to_json(input).map_err(js_error)
}

#[wasm_bindgen]
pub fn build_display_list_json(input: &str) -> Result<String, JsValue> {
    core::build_display_list_json(input).map_err(js_error)
}

#[wasm_bindgen]
pub fn hit_test_json(
    display_list: &str,
    page_index: u32,
    x: f64,
    y: f64,
) -> Result<String, JsValue> {
    docx_layout::hit::hit_test_json(display_list, page_index as usize, x, y).map_err(js_error)
}

#[wasm_bindgen]
pub fn range_rects_json(display_list: &str, from: f64, to: f64) -> Result<String, JsValue> {
    docx_layout::hit::range_rects_json(display_list, from as i64, to as i64).map_err(js_error)
}

#[wasm_bindgen]
pub fn range_rects_region_json(
    display_list: &str,
    region: &str,
    r_id: &str,
    from: f64,
    to: f64,
) -> Result<String, JsValue> {
    docx_layout::hit::range_rects_region_json(display_list, region, r_id, from as i64, to as i64)
        .map_err(js_error)
}

#[wasm_bindgen]
pub fn hit_test_regions_json(
    display_list: &str,
    page_index: u32,
    x: f64,
    y: f64,
) -> Result<String, JsValue> {
    docx_layout::hit::hit_test_regions_json(display_list, page_index as usize, x, y)
        .map_err(js_error)
}

#[wasm_bindgen]
pub fn open_display_list(display_list: &str) -> Result<u32, JsValue> {
    docx_layout::session::open_display_list(display_list).map_err(js_error)
}

#[wasm_bindgen]
pub fn close_display_list(handle: u32) {
    docx_layout::session::close_display_list(handle);
}

#[wasm_bindgen]
pub fn hit_test_regions_by_handle(
    handle: u32,
    page_index: u32,
    x: f64,
    y: f64,
) -> Result<String, JsValue> {
    docx_layout::session::hit_test_regions_by_handle(handle, page_index as usize, x, y)
        .map_err(js_error)
}

#[wasm_bindgen]
pub fn range_rects_by_handle(handle: u32, from: f64, to: f64) -> Result<String, JsValue> {
    docx_layout::session::range_rects_by_handle(handle, from as i64, to as i64).map_err(js_error)
}

#[wasm_bindgen]
pub fn range_rects_region_by_handle(
    handle: u32,
    region: &str,
    r_id: &str,
    from: f64,
    to: f64,
) -> Result<String, JsValue> {
    docx_layout::session::range_rects_region_by_handle(handle, region, r_id, from as i64, to as i64)
        .map_err(js_error)
}

#[wasm_bindgen]
pub fn register_measure_font(bytes: &[u8]) -> Result<u32, JsValue> {
    core::register_measure_font(bytes).map_err(js_error)
}

#[wasm_bindgen]
pub fn clear_measure_fonts() {
    core::clear_measure_fonts();
}

#[wasm_bindgen]
pub fn measure_paragraph_json(input: &str) -> Result<String, JsValue> {
    core::measure_paragraph_json(input).map_err(js_error)
}

#[wasm_bindgen]
pub fn outline_glyph_json(font_id: u32, glyph_id: u32) -> Result<String, JsValue> {
    core::outline_glyph_json(font_id, glyph_id).map_err(js_error)
}
