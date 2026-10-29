//! PPTX display-list wasm boundary.

use wasm_bindgen::prelude::*;

pub use pptx_edit::wasm::PptxDocument;

#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(js_namespace = performance, js_name = now)]
    fn performance_now() -> f64;
}

#[wasm_bindgen]
pub struct PptxRenderer {
    inner: pptx_engine::Renderer,
}

#[wasm_bindgen]
impl PptxRenderer {
    #[wasm_bindgen(constructor)]
    pub fn new() -> PptxRenderer {
        Self {
            inner: pptx_engine::Renderer::new(),
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
        self.inner
            .register_font(family, bold, italic, bytes)
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = registerFallbackFont)]
    pub fn register_fallback_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
    ) -> Result<u32, JsValue> {
        self.inner
            .register_fallback_font(family, bold, italic, bytes)
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = saveBytes)]
    pub fn save_bytes(&self, document: &PptxDocument) -> Result<Vec<u8>, JsValue> {
        self.inner.save_bytes(document.session()).map_err(js_error)
    }

    #[wasm_bindgen(js_name = layoutSlideJson)]
    pub fn layout_slide_json(
        &mut self,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        self.inner
            .layout_slide_json(document.session(), slide_index)
            .map_err(js_error)
    }

    /// requested font faces before layout, for one slide or the whole deck.
    #[wasm_bindgen(js_name = requestedFontFacesJson)]
    pub fn requested_font_faces_json(
        &self,
        document: &PptxDocument,
        slide_id: Option<String>,
    ) -> Result<String, JsValue> {
        self.inner
            .requested_font_faces_json(document.session(), slide_id.as_deref())
            .map_err(js_error)
    }

    /// `layoutSlideJson` with its stages timed, returned as
    /// `{"layout": ..., "profile": {"scopeMs", "layoutMs", "serializeMs"}}`.
    #[wasm_bindgen(js_name = layoutSlideProfiledJson)]
    pub fn layout_slide_profiled_json(
        &mut self,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        self.inner
            .layout_slide_profiled_json(document.session(), slide_index, &mut performance_now)
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = hitTestJson)]
    pub fn hit_test_json(&self, x: f32, y: f32) -> Result<String, JsValue> {
        self.inner.hit_test_json(x, y).map_err(js_error)
    }

    #[wasm_bindgen(js_name = setActiveSlide)]
    pub fn set_active_slide(&mut self, document: &PptxDocument, id: &str, key: &str) -> bool {
        self.inner.set_active_slide(document.session(), id, key)
    }

    #[wasm_bindgen(js_name = snapshotWithLayoutKeysJson)]
    pub fn snapshot_with_layout_keys_json(
        &mut self,
        document: &PptxDocument,
    ) -> Result<String, JsValue> {
        self.inner
            .snapshot_with_layout_keys_json(document.session())
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = slideLayoutKey)]
    pub fn slide_layout_key(
        &mut self,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        self.inner
            .slide_layout_key(document.session(), slide_index)
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = hitTestSlideJson)]
    pub fn hit_test_slide_json(
        &mut self,
        document: &PptxDocument,
        id: &str,
        x: f32,
        y: f32,
    ) -> Result<String, JsValue> {
        self.inner
            .hit_test_slide_json(document.session(), id, x, y)
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = layoutProposalSlideJson)]
    pub fn layout_proposal_slide_json(
        &self,
        document: &PptxDocument,
        id: &str,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        self.inner
            .layout_proposal_slide_json(document.session(), id, slide_index)
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = layoutProposalDiffSlideJson)]
    pub fn layout_proposal_diff_slide_json(
        &self,
        document: &PptxDocument,
        id: &str,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        self.inner
            .layout_proposal_diff_slide_json(document.session(), id, slide_index)
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
    pptx_engine::parse_pptx_json(data).map_err(js_error)
}

#[wasm_bindgen(js_name = parsePptxJsonWithoutMedia)]
pub fn parse_pptx_json_without_media(data: &[u8]) -> Result<String, JsValue> {
    pptx_engine::parse_pptx_json_without_media(data).map_err(js_error)
}

/// Structured export of PPTX bytes as a snapshot, with the options of
/// `PptxDocument.exportStructuredJson`: `{"ok":true,"content"}` or `{"ok":false,"failure"}`.
/// Bytes that are not a readable PPTX throw.
#[wasm_bindgen(js_name = exportPptxStructuredJson)]
pub fn export_pptx_structured_json(data: &[u8], options: &str) -> Result<String, JsValue> {
    pptx_engine::export_pptx_structured_json(data, options).map_err(js_error)
}

/// `exportPptxStructuredJson` rendered as Markdown.
#[wasm_bindgen(js_name = exportPptxMarkdownJson)]
pub fn export_pptx_markdown_json(data: &[u8], options: &str) -> Result<String, JsValue> {
    pptx_engine::export_pptx_markdown_json(data, options).map_err(js_error)
}

/// Renders schema-version-1 structured content as Markdown; `options` is `{"maxBytes"?}`.
#[wasm_bindgen(js_name = renderPptxMarkdownJson)]
pub fn render_pptx_markdown_json(content: &str, options: &str) -> Result<String, JsValue> {
    pptx_engine::render_pptx_markdown_json(content, options).map_err(js_error)
}

#[wasm_bindgen(js_name = compileSlideJson)]
pub fn compile_slide_json(slide_json: &str) -> Result<String, JsValue> {
    pptx_engine::compile_slide_json(slide_json).map_err(js_error)
}

#[wasm_bindgen(js_name = rendererVersion)]
pub fn renderer_version() -> String {
    pptx_engine::renderer_version().to_owned()
}

#[wasm_bindgen(js_name = decodeTiffPng)]
pub fn decode_tiff_png(data: &[u8]) -> Result<Vec<u8>, JsValue> {
    pptx_engine::decode_tiff_png(data).map_err(js_error)
}

fn js_error(error: pptx_engine::EngineError) -> JsValue {
    JsValue::from_str(&error.message)
}
