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
        let scope = slide_scope(session, slide_index)?;
        let rendered = self
            .renderer
            .layout_scoped_slide(session.package(), &scope)
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

    #[wasm_bindgen(js_name = layoutProposalSlideJson)]
    pub fn layout_proposal_slide_json(
        &self,
        document: &PptxDocument,
        id: &str,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        let preview = document
            .session()
            .proposal_preview_session(id)
            .map_err(js_error)?;
        let scope = slide_scope(&preview, slide_index)?;
        let rendered = self
            .renderer
            .layout_scoped_slide(preview.package(), &scope)
            .map_err(js_error)?;
        serde_json::to_string(&rendered.display_list).map_err(js_error)
    }

    #[wasm_bindgen(js_name = layoutProposalDiffSlideJson)]
    pub fn layout_proposal_diff_slide_json(
        &self,
        document: &PptxDocument,
        id: &str,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        let session = document.session();
        let preview = session.preview_proposal_diff(id).map_err(js_error)?;
        let rendered = self
            .renderer
            .layout_slide(session.package(), &preview.snapshot, slide_index as usize)
            .map_err(js_error)?;
        serde_json::to_string(&serde_json::json!({
            "proposal": preview.proposal,
            "snapshot": preview.snapshot,
            "textChanges": preview.text_changes,
            "frame": rendered.display_list,
        }))
        .map_err(js_error)
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

#[wasm_bindgen(js_name = decodeTiffPng)]
pub fn decode_tiff_png(data: &[u8]) -> Result<Vec<u8>, JsValue> {
    ooxml_drawingml::media::decode_tiff_png(data).map_err(js_error)
}

fn slide_scope(
    session: &pptx_edit::DeckSession,
    slide_index: u32,
) -> Result<pptx_edit::SlideScope, JsValue> {
    session
        .slide_scope(slide_index as usize)
        .map_err(|error| match error {
            pptx_edit::EditError::OutOfBounds { .. } => js_error(
                pptx_render::RenderError::SlideNotFound(slide_index as usize),
            ),
            error => js_error(error),
        })
}

fn js_error(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}
