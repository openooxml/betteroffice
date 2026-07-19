//! PPTX display-list wasm boundary.

use wasm_bindgen::prelude::*;

pub use pptx_edit::wasm::PptxDocument;

#[wasm_bindgen]
pub struct PptxRenderer {
    renderer: pptx_render::SlideRenderer,
    rendered: Option<pptx_render::RenderedSlide>,
}

#[wasm_bindgen]
impl PptxRenderer {
    #[wasm_bindgen(constructor)]
    pub fn new() -> PptxRenderer {
        Self {
            renderer: pptx_render::SlideRenderer::new(),
            rendered: None,
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
        self.renderer
            .register_font(family, bold, italic, bytes)
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = layoutSlideJson)]
    pub fn layout_slide_json(
        &mut self,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        let session = document.session();
        let deck = session.snapshot().map_err(js_error)?;
        let rendered = self
            .renderer
            .layout_slide(session.package(), &deck, slide_index as usize)
            .map_err(js_error)?;
        let json = serde_json::to_string(&rendered.display_list).map_err(js_error)?;
        self.rendered = Some(rendered);
        Ok(json)
    }

    #[wasm_bindgen(js_name = hitTestJson)]
    pub fn hit_test_json(&self, x: f32, y: f32) -> Result<String, JsValue> {
        let result = self
            .rendered
            .as_ref()
            .and_then(|rendered| rendered.hit_test(x, y));
        serde_json::to_string(&result).map_err(js_error)
    }
}

impl Default for PptxRenderer {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen(js_name = parsePptxJson)]
pub fn parse_pptx_json(data: &[u8]) -> Result<String, JsValue> {
    let package = pptx_parse::parse_pptx(data).map_err(js_error)?;
    serde_json::to_string(&package).map_err(js_error)
}

#[wasm_bindgen(js_name = compileSlideJson)]
pub fn compile_slide_json(slide_json: &str) -> Result<String, JsValue> {
    pptx_render::compile_json(slide_json).map_err(|error| JsValue::from_str(&error))
}

#[wasm_bindgen(js_name = rendererVersion)]
pub fn renderer_version() -> String {
    env!("CARGO_PKG_VERSION").to_owned()
}

fn js_error(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}
