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

    /// Lays every document master out once per materialized package.
    #[wasm_bindgen(js_name = masterPreviewsJson)]
    pub fn master_previews_json(&self, document: &VsdxDocument) -> Result<String, JsValue> {
        let package = document.session().package().map_err(js_error)?;
        let previews = package
            .master_sheets
            .keys()
            .map(|id| {
                serde_json::json!({
                    "id": id,
                    "name": package.master_names.get(id),
                    "display": self.renderer.layout_master(&package, *id).ok(),
                })
            })
            .collect::<Vec<_>>();
        serde_json::to_string(&previews).map_err(js_error)
    }

    #[wasm_bindgen(js_name = pageLayersJson)]
    pub fn page_layers_json(
        &self,
        document: &VsdxDocument,
        page_index: u32,
    ) -> Result<String, JsValue> {
        let package = document.session().package().map_err(js_error)?;
        let page_part = package
            .page_part_paths
            .get(page_index as usize)
            .ok_or_else(|| JsValue::from_str("page index is outside the document"))?;
        let layers = self.renderer.effective_page_layers(&package, page_part);
        serde_json::to_string(&layers).map_err(js_error)
    }

    #[wasm_bindgen(js_name = validateJson)]
    pub fn validate_json(&self, document: &VsdxDocument) -> Result<String, JsValue> {
        let package = document.session().package().map_err(js_error)?;
        let report = vsdx_validate::validate_package(&package);
        serde_json::to_string(&report.issues).map_err(js_error)
    }

    #[wasm_bindgen(js_name = validatePageJson)]
    pub fn validate_page_json(
        &self,
        document: &VsdxDocument,
        page_index: u32,
    ) -> Result<String, JsValue> {
        let package = document.session().package().map_err(js_error)?;
        let page_part = package
            .page_part_paths
            .get(page_index as usize)
            .ok_or_else(|| JsValue::from_str("page index is outside the document"))?;
        let issues = vsdx_validate::validate_page(&package, page_part);
        serde_json::to_string(&issues).map_err(js_error)
    }

    #[wasm_bindgen(js_name = setLayerVisible)]
    pub fn set_layer_visible(&mut self, page_part: &str, index: u32, visible: bool) {
        self.renderer.set_layer_override(page_part, index, visible);
    }

    #[wasm_bindgen(js_name = clearLayerVisibility)]
    pub fn clear_layer_visibility(&mut self) {
        self.renderer.clear_layer_overrides();
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

    #[wasm_bindgen(js_name = exportPdf)]
    pub fn export_pdf(&self, document: &VsdxDocument) -> Result<Vec<u8>, JsValue> {
        let package = document.session().package().map_err(js_error)?;
        self.renderer.export_pdf(&package).map_err(js_error)
    }

    #[wasm_bindgen(js_name = exportSvgJson)]
    pub fn export_svg_json(&self, document: &VsdxDocument) -> Result<String, JsValue> {
        let package = document.session().package().map_err(js_error)?;
        let pages = self.renderer.export_svg(&package).map_err(js_error)?;
        serde_json::to_string(&pages).map_err(js_error)
    }

    #[cfg(feature = "raster")]
    #[wasm_bindgen(js_name = exportPng)]
    pub fn export_png(
        &self,
        document: &VsdxDocument,
        page_index: u32,
        scale: f32,
    ) -> Result<Vec<u8>, JsValue> {
        let package = document.session().package().map_err(js_error)?;
        vsdx_raster::render_page(&self.renderer, &package, page_index as usize, scale)
            .map(|page| page.bytes)
            .map_err(js_error)
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

    fn added_shape_json() -> String {
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
        serde_json::json!({ "pageId": "page:1", "draft": { "name": "Added", "cells": cells } })
            .to_string()
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
        let receipt: serde_json::Value =
            serde_json::from_str(&document.add_shape_json(&added_shape_json()).unwrap()).unwrap();
        let shape_id = receipt["shapeId"].as_str().unwrap();
        let mut renderer = VsdxRenderer::new();
        let before = renderer.layout_page_json(&document, 0).unwrap();
        document
            .session()
            .set_cell_formula(&EditCtx::local("test"), "page:1", shape_id, "Width", "10")
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
            serde_json::from_str(&document.add_shape_json(&added_shape_json()).unwrap()).unwrap();
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
        let delete = serde_json::json!({ "pageId": "page:1", "shapeId": added_id }).to_string();
        document.delete_shape_json(&delete).unwrap();
        let after_delete: serde_json::Value =
            serde_json::from_str(&renderer.layout_page_json(&document, 0).unwrap()).unwrap();
        assert_eq!(
            primitive_ids(&after_delete).len(),
            primitive_ids(&after_reorder).len() - 1
        );
    }

    #[test]
    fn add_connector_json_glues_two_shapes_and_survives_save() {
        let document = VsdxDocument::open_collaborative(
            include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx"),
            1.0,
        )
        .unwrap();
        let rect = |pin_x: &str| {
            serde_json::json!({
                "pageId": "page:1",
                "draft": {
                    "name": "Rect",
                    "cells": [
                        { "locator": { "cellName": "Width" }, "formula": "1" },
                        { "locator": { "cellName": "Height" }, "formula": "1" },
                        { "locator": { "cellName": "PinX" }, "formula": pin_x },
                        { "locator": { "cellName": "PinY" }, "formula": "1" },
                        { "locator": { "cellName": "LocPinX" }, "formula": "0" },
                        { "locator": { "cellName": "LocPinY" }, "formula": "0" },
                    ],
                }
            })
            .to_string()
        };
        let from: serde_json::Value =
            serde_json::from_str(&document.add_shape_json(&rect("1")).unwrap()).unwrap();
        let to: serde_json::Value =
            serde_json::from_str(&document.add_shape_json(&rect("5")).unwrap()).unwrap();
        let connector = serde_json::json!({
            "pageId": "page:1",
            "draft": {
                "name": "Connector",
                "cells": [
                    { "locator": { "cellName": "OneD" }, "formula": "1" },
                    { "locator": { "cellName": "BeginX" }, "formula": "1" },
                    { "locator": { "cellName": "BeginY" }, "formula": "2" },
                    { "locator": { "cellName": "EndX" }, "formula": "4" },
                    { "locator": { "cellName": "EndY" }, "formula": "2" },
                ],
            },
            "from": { "shapeId": from["shapeId"] },
            "to": { "shapeId": to["shapeId"], "toCell": "PinX" },
        })
        .to_string();
        let receipt: serde_json::Value =
            serde_json::from_str(&document.add_connector_json(&connector).unwrap()).unwrap();
        assert!(receipt["shapeId"].as_str().unwrap().contains(":added:"));
        let mut renderer = VsdxRenderer::new();
        let live = renderer.layout_page_json(&document, 0).unwrap();
        let display: serde_json::Value = serde_json::from_str(&live).unwrap();
        let connector = display["primitives"]
            .as_array()
            .unwrap()
            .iter()
            .find(|primitive| primitive["id"] == "visio/pages/page1.xml:4")
            .unwrap();
        assert_eq!(
            connector["path"],
            serde_json::json!([
                { "type": "move", "x": 1.0, "y": 1.0 },
                { "type": "line", "x": 5.0, "y": 1.0 },
            ])
        );
        let saved = document.save().unwrap();
        let reparsed = vsdx_parse::parse_vsdx(&saved).unwrap();
        let part = reparsed.page_part_paths[0].clone();
        let from_cells = reparsed.page_contents[&part]
            .connects()
            .filter_map(|connect| connect.from_cell.clone())
            .collect::<Vec<_>>();
        assert!(from_cells.iter().any(|cell| cell == "BeginX"));
        assert!(from_cells.iter().any(|cell| cell == "EndX"));
        let reopened = VsdxDocument::open_collaborative(&saved, 2.0).unwrap();
        let mut reopened_renderer = VsdxRenderer::new();
        assert_eq!(
            live,
            reopened_renderer.layout_page_json(&reopened, 0).unwrap()
        );
        assert!(
            document
                .session()
                .add_connector(
                    &vsdx_edit::EditCtx::local("test"),
                    "page:1",
                    &vsdx_edit::ShapeDraft {
                        name: None,
                        master: None,
                        cells: Vec::new(),
                    },
                    &vsdx_edit::ConnectorGlue {
                        shape_id: from["shapeId"].as_str().unwrap().to_owned(),
                        to_cell: None,
                    },
                    &vsdx_edit::ConnectorGlue {
                        shape_id: to["shapeId"].as_str().unwrap().to_owned(),
                        to_cell: None,
                    },
                )
                .is_err()
        );
    }

    #[test]
    fn set_connector_route_json_reroutes_the_painted_path() {
        let document = VsdxDocument::open_collaborative(
            include_bytes!("../../vsdx-parse/tests/fixtures/connector-route-style.vsdx"),
            1.0,
        )
        .unwrap();
        let mut renderer = VsdxRenderer::new();
        let before: serde_json::Value =
            serde_json::from_str(&renderer.layout_page_json(&document, 0).unwrap()).unwrap();
        let receipt: serde_json::Value = serde_json::from_str(
            &document
                .set_connector_route_json(
                    &serde_json::json!({
                        "pageId": "page:1",
                        "shapeId": "page:1:shape:1",
                        "points": [
                            { "x": 1.0, "y": 1.0 },
                            { "x": 1.0, "y": 3.0 },
                            { "x": 4.0, "y": 3.0 },
                        ],
                    })
                    .to_string(),
                )
                .unwrap(),
        )
        .unwrap();
        assert_eq!(receipt["points"], 3);
        let after: serde_json::Value =
            serde_json::from_str(&renderer.layout_page_json(&document, 0).unwrap()).unwrap();
        let path = after["primitives"]
            .as_array()
            .unwrap()
            .iter()
            .find(|primitive| primitive["id"] == "visio/pages/page1.xml:1")
            .unwrap()["path"]
            .clone();
        assert_eq!(
            path,
            serde_json::json!([
                { "type": "move", "x": 1.0, "y": 1.0 },
                { "type": "line", "x": 1.0, "y": 3.0 },
                { "type": "line", "x": 4.0, "y": 3.0 },
                { "type": "line", "x": 4.0, "y": 3.0 },
            ])
        );
        assert_ne!(before, after);
        let reopened = VsdxDocument::open_collaborative(&document.save().unwrap(), 2.0).unwrap();
        let mut reopened_renderer = VsdxRenderer::new();
        assert_eq!(
            after,
            serde_json::from_str::<serde_json::Value>(
                &reopened_renderer.layout_page_json(&reopened, 0).unwrap()
            )
            .unwrap()
        );
    }

    #[test]
    fn layout_of_reordered_added_shapes_matches_the_saved_document() {
        let document = VsdxDocument::open_collaborative(
            include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx"),
            1.0,
        )
        .unwrap();
        document.add_shape_json(&added_shape_json()).unwrap();
        let second: serde_json::Value =
            serde_json::from_str(&document.add_shape_json(&added_shape_json()).unwrap()).unwrap();
        document
            .reorder_shape_json(
                &serde_json::json!({
                    "pageId": "page:1",
                    "shapeId": second["shapeId"],
                    "toIndex": 0,
                })
                .to_string(),
            )
            .unwrap();
        let mut live_renderer = VsdxRenderer::new();
        let live = live_renderer.layout_page_json(&document, 0).unwrap();
        let reopened = VsdxDocument::open_collaborative(&document.save().unwrap(), 2.0).unwrap();
        let mut reopened_renderer = VsdxRenderer::new();
        assert_eq!(
            live,
            reopened_renderer.layout_page_json(&reopened, 0).unwrap()
        );
    }

    #[test]
    fn page_layers_json_lists_effective_visibility() {
        let document = VsdxDocument::open_collaborative(
            include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx"),
            1.0,
        )
        .unwrap();
        let mut renderer = VsdxRenderer::new();
        let layers: serde_json::Value =
            serde_json::from_str(&renderer.page_layers_json(&document, 0).unwrap()).unwrap();
        assert_eq!(layers, serde_json::Value::Array(Vec::new()));
        renderer.set_layer_visible("visio/pages/page1.xml", 0, false);
        renderer.layout_page_json(&document, 0).unwrap();
        renderer.clear_layer_visibility();
    }

    #[test]
    fn master_previews_list_every_document_master_once() {
        let document = VsdxDocument::open_collaborative(
            include_bytes!("../../vsdx-parse/tests/fixtures/document-stencil.vsdx"),
            1.0,
        )
        .unwrap();
        let renderer = VsdxRenderer::new();
        let previews: serde_json::Value =
            serde_json::from_str(&renderer.master_previews_json(&document).unwrap()).unwrap();
        let previews = previews.as_array().unwrap();
        assert_eq!(previews.len(), 2);
        assert_eq!(previews[0]["id"], serde_json::json!(1));
        assert_eq!(previews[0]["name"], serde_json::json!("Stencil-Rect"));
        assert_eq!(
            previews[0]["display"]["contractVersion"],
            serde_json::json!(vsdx_render::CONTRACT_VERSION)
        );
        assert_eq!(
            previews[0]["display"]["primitives"]
                .as_array()
                .unwrap()
                .len(),
            2
        );
        assert_eq!(previews[1]["id"], serde_json::json!(2));
        assert_eq!(previews[1]["name"], serde_json::json!("Stencil-Tri"));
        assert_eq!(
            previews[1]["display"]["primitives"]
                .as_array()
                .unwrap()
                .len(),
            2
        );
    }

    #[test]
    fn export_svg_json_renders_one_vector_page_per_diagram_page() {
        let document = VsdxDocument::open_collaborative(
            include_bytes!("../../vsdx-parse/tests/fixtures/text-accounting.vsdx"),
            1.0,
        )
        .unwrap();
        let renderer = VsdxRenderer::new();
        let pages: Vec<String> =
            serde_json::from_str(&renderer.export_svg_json(&document).unwrap()).unwrap();
        assert_eq!(pages.len(), 1);
        assert!(pages[0].starts_with("<svg xmlns=\"http://www.w3.org/2000/svg\""));
        assert!(pages[0].contains("<text"));
    }

    #[cfg(feature = "raster")]
    #[test]
    fn export_png_renders_scaled_raster_pages() {
        let document = VsdxDocument::open_collaborative(
            include_bytes!("../../vsdx-parse/tests/fixtures/text-accounting.vsdx"),
            1.0,
        )
        .unwrap();
        let renderer = VsdxRenderer::new();
        let first = renderer.export_png(&document, 0, 1.0).unwrap();
        assert_eq!(&first[0..8], &[137, 80, 78, 71, 13, 10, 26, 10]);
        let second = renderer.export_png(&document, 0, 2.0).unwrap();
        assert!(second.len() > first.len());
    }
}
