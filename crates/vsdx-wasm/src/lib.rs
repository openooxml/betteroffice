//! VSDX display-list wasm boundary.

use wasm_bindgen::prelude::*;

pub use vsdx_edit::wasm::VsdxDocument;

#[wasm_bindgen]
pub struct VsdxRenderer {
    renderer: vsdx_render::Renderer,
    rendered: Option<vsdx_render::VsdxDisplayList>,
    font_count: u32,
}

#[wasm_bindgen]
impl VsdxRenderer {
    #[wasm_bindgen(constructor)]
    pub fn new() -> VsdxRenderer {
        Self {
            renderer: vsdx_render::Renderer::default(),
            rendered: None,
            font_count: 0,
        }
    }

    #[wasm_bindgen(js_name = registerFont)]
    pub fn register_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
    ) -> Result<u32, JsValue> {
        let handle = self.font_count;
        let next_font_count = self
            .font_count
            .checked_add(1)
            .ok_or_else(|| JsValue::from_str("font handle limit exceeded"))?;
        self.renderer
            .register_font(family, bold, italic, bytes.to_vec())
            .map_err(js_error)?;
        self.font_count = next_font_count;
        Ok(handle)
    }

    #[wasm_bindgen(js_name = layoutPageJson)]
    pub fn layout_page_json(
        &mut self,
        document: &VsdxDocument,
        page_index: u32,
    ) -> Result<String, JsValue> {
        let package = document.session().package().map_err(js_error)?;
        let page_part = package
            .page_part_paths
            .get(page_index as usize)
            .ok_or_else(|| JsValue::from_str("page index is outside the document"))?;
        let rendered = self
            .renderer
            .layout_page(&package, page_part)
            .map_err(js_error)?;
        let json = serde_json::to_string(&rendered).map_err(js_error)?;
        self.rendered = Some(rendered);
        Ok(json)
    }

    #[wasm_bindgen(js_name = hitTestJson)]
    pub fn hit_test_json(&self, x: f32, y: f32) -> Result<String, JsValue> {
        let result = self
            .rendered
            .as_ref()
            .and_then(|rendered| vsdx_render::hit_test(rendered, x, y));
        let result = match result {
            Some(vsdx_render::HitTestResult::Shape { shape_id }) => {
                serde_json::json!({ "kind": "shape", "shapeId": shape_id })
            }
            Some(vsdx_render::HitTestResult::Text { shape_id, position }) => {
                serde_json::json!({ "kind": "text", "shapeId": shape_id, "position": position })
            }
            None => serde_json::Value::Null,
        };
        serde_json::to_string(&result).map_err(js_error)
    }
}

impl Default for VsdxRenderer {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen(js_name = parseVsdxJson)]
pub fn parse_vsdx_json(data: &[u8]) -> Result<String, JsValue> {
    let package = vsdx_parse::parse_vsdx(data).map_err(js_error)?;
    serde_json::to_string(&package).map_err(js_error)
}

#[wasm_bindgen(js_name = rendererVersion)]
pub fn renderer_version() -> String {
    env!("CARGO_PKG_VERSION").to_owned()
}

fn js_error(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}

#[cfg(test)]
mod tests {
    use super::{VsdxDocument, VsdxRenderer, parse_vsdx_json};
    use vsdx_edit::EditCtx;
    use yrs::{Array, Map, MapPrelim, Out, ReadTxn, Transact};

    fn add_formula(document: &VsdxDocument, name: &str, formula: &str) {
        let mut txn = document.session().yrs_doc().transact_mut();
        let sheets = txn.get_map("vsdx:sheets").unwrap();
        let shape = match sheets.get(&txn, "page:1:shape:1") {
            Some(Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        let cells = match shape.get(&txn, "cells") {
            Some(Out::YMap(cells)) => cells,
            _ => unreachable!(),
        };
        let cell = cells.insert(&mut txn, name, MapPrelim::default());
        cell.insert(&mut txn, "name", name);
        cell.insert(&mut txn, "formula", formula);
    }

    fn nested_document() -> VsdxDocument {
        VsdxDocument::open_collaborative(
            include_bytes!("../../vsdx-parse/tests/fixtures/nested-groups.vsdx"),
            1.0,
        )
        .unwrap()
    }

    fn shape_cell(name: &str, formula: &str) -> serde_json::Value {
        serde_json::json!({ "locator": { "cellName": name }, "formula": formula })
    }

    fn added_shape_json(source_id: u32) -> String {
        let mut cells = vec![
            shape_cell("Width", "1"),
            shape_cell("Height", "1"),
            shape_cell("PinX", "12"),
            shape_cell("PinY", "2"),
            shape_cell("LocPinX", "0"),
            shape_cell("LocPinY", "0"),
        ];
        for (index, x, y) in [(0, "0", "0"), (1, "1", "0"), (2, "1", "1"), (3, "0", "1")] {
            cells.push(serde_json::json!({
                "locator": { "section": "Geometry", "rowIndex": index, "cellName": "X" },
                "formula": x
            }));
            cells.push(serde_json::json!({
                "locator": { "section": "Geometry", "rowIndex": index, "cellName": "Y" },
                "formula": y
            }));
        }
        cells.push(serde_json::json!({
            "locator": { "section": "Geometry", "rowIndex": 4, "cellName": "NoShow" },
            "formula": "0"
        }));
        serde_json::json!({ "pageId": "page:1", "draft": { "sourceId": source_id, "name": "Added", "cells": cells } }).to_string()
    }

    fn primitive_ids(value: &serde_json::Value) -> Vec<String> {
        value["primitives"]
            .as_array()
            .unwrap()
            .iter()
            .filter_map(|primitive| primitive.get("id").and_then(serde_json::Value::as_str))
            .map(str::to_owned)
            .collect()
    }

    #[test]
    fn parse_vsdx_json_round_trips_a_fixture() {
        let json = parse_vsdx_json(include_bytes!(
            "../../vsdx-parse/tests/fixtures/foundation.vsdx"
        ))
        .unwrap();
        assert!(json.contains("pagePartPaths"));
    }

    #[test]
    fn layout_page_json_uses_the_current_collaborative_state() {
        let document = VsdxDocument::open_collaborative(
            include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx"),
            1.0,
        )
        .unwrap();
        add_formula(&document, "PinX", "1");
        add_formula(&document, "PinY", "1");
        add_formula(&document, "Width", "1");
        add_formula(&document, "Height", "1");
        let mut renderer = VsdxRenderer::new();
        let before = renderer.layout_page_json(&document, 0).unwrap();
        document
            .session()
            .set_cell_formula(
                &EditCtx::local("test"),
                "page:1",
                "page:1:shape:1",
                "Width",
                "10",
            )
            .unwrap();
        let after = renderer.layout_page_json(&document, 0).unwrap();
        assert_ne!(before, after);
    }

    #[test]
    fn layout_page_json_materializes_added_deleted_and_reordered_shapes() {
        let document = nested_document();
        let mut renderer = VsdxRenderer::new();
        let before: serde_json::Value =
            serde_json::from_str(&renderer.layout_page_json(&document, 0).unwrap()).unwrap();
        let receipt: serde_json::Value =
            serde_json::from_str(&document.add_shape_json(&added_shape_json(1)).unwrap()).unwrap();
        let added_id = receipt["shapeId"].as_str().unwrap();
        let after_add: serde_json::Value =
            serde_json::from_str(&renderer.layout_page_json(&document, 0).unwrap()).unwrap();
        assert!(primitive_ids(&after_add).len() > primitive_ids(&before).len());
        assert!(
            primitive_ids(&after_add)
                .iter()
                .any(|id| id.ends_with(":1"))
        );
        let before_reorder = primitive_ids(&after_add);
        document
            .reorder_shape_json(r#"{"pageId":"page:1","shapeId":"page:1:shape:0","toIndex":0}"#)
            .unwrap();
        let after_reorder: serde_json::Value =
            serde_json::from_str(&renderer.layout_page_json(&document, 0).unwrap()).unwrap();
        assert_ne!(before_reorder, primitive_ids(&after_reorder));
        let mut txn = document.session().yrs_doc().transact_mut();
        let pages = txn.get_map("vsdx:pages").unwrap();
        let page = match pages.get(&txn, "page:1") {
            Some(Out::YMap(page)) => page,
            _ => unreachable!(),
        };
        let order = match page.get(&txn, "shapes") {
            Some(Out::YArray(order)) => order,
            _ => unreachable!(),
        };
        let index = (0..order.len(&txn))
            .find(|index| matches!(order.get(&txn, *index), Some(Out::Any(yrs::Any::String(value))) if value.as_ref() == added_id))
            .unwrap();
        order.remove_range(&mut txn, index, 1);
        drop(txn);
        let after_delete: serde_json::Value =
            serde_json::from_str(&renderer.layout_page_json(&document, 0).unwrap()).unwrap();
        assert_eq!(
            primitive_ids(&after_delete).len(),
            primitive_ids(&after_reorder).len() - 1
        );
    }
}
