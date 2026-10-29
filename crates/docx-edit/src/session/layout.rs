use super::{EditSessionCore, EngineError, EngineResult};

impl EditSessionCore {
    /// Measurement input JSON in, `ParagraphExtent` JSON out. Also records the
    /// paragraph's immutable width/font envelope under its stable block id, so
    /// a later resident edit re-measures only the changed block. Errors with
    /// the engine's message for input it cannot measure.
    pub fn measure_paragraph_json(&self, input: &str) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        self.engine
            .measure_paragraph_json(input)
            .map_err(EngineError::message)
    }

    /// `{ measured, options }` JSON in, `Layout` JSON out. Both the measured
    /// input and the resulting layout are retained for the resident edit path.
    /// Errors on unparseable input or on layout failure.
    pub fn layout_document_json(&self, input: &str) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        self.engine
            .layout_document_json(input)
            .map_err(EngineError::message)
    }

    /// Region-layout input JSON in, the font families and sizes that input
    /// needs as JSON out, so the host can register fonts before laying out.
    pub fn layout_font_requirements_json(&self, input: &str) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        self.engine
            .layout_font_requirements_json(input)
            .map_err(EngineError::message)
    }

    /// Region-layout input JSON in, a paginated envelope with section and page
    /// regions already composed as JSON out — ready for the display builder
    /// with no host-side layout mutation. Retains the pass for resident edits.
    pub fn layout_document_with_regions_json(&self, input: &str) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        self.engine
            .layout_document_with_regions_json(input)
            .map_err(EngineError::message)
    }

    /// Same full region pass as [`Self::layout_document_with_regions_json`],
    /// but the reply carries only `{ layout, headersFooters?, notesConverged }`
    /// — the measured arena stays retained wasm-side and is fetched on demand
    /// via [`Self::retained_kernel_inputs_json`].
    pub fn layout_document_with_regions_retained_json(&self, input: &str) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        self.engine
            .layout_document_with_regions_retained_json(input)
            .map_err(EngineError::message)
    }

    pub fn layout_document_with_regions_retained_meta(
        &self,
        input: &str,
    ) -> EngineResult<crate::engine::RetainedLayoutMeta> {
        let _fonts = self.fonts.enter();
        self.engine
            .layout_document_with_regions_retained_meta(input)
            .map_err(EngineError::message)
    }

    pub fn retained_layout_json(&self) -> EngineResult<String> {
        self.engine
            .retained_layout_json()
            .map_err(EngineError::message)
    }

    /// [`Self::layout_document_with_regions_retained_json`] without the reply,
    /// for a caller that reads only the retained state.
    pub fn layout_document_with_regions_retained(&self, input: &str) -> EngineResult<()> {
        let _fonts = self.fonts.enter();
        self.engine
            .layout_document_with_regions_retained(input)
            .map_err(EngineError::message)
    }

    /// [`Self::layout_document_with_regions_retained`] after one edit: the
    /// resident path when that edit was the [`Self::edit_text_resident`] taken
    /// last; see `EngineSession::layout_document_with_regions_after_edit`.
    pub fn layout_document_with_regions_after_edit(&self, input: &str) -> EngineResult<()> {
        let _fonts = self.fonts.enter();
        let edit = self
            .resident_text_edit
            .take()
            .filter(|(_, epoch)| *epoch == self.engine.doc_epoch());
        let story = edit.as_ref().map(|(story, _)| story.as_str());
        self.engine
            .layout_document_with_regions_after_edit(input, story)
            .map_err(EngineError::message)
    }

    /// The retained region layout's `headersFooters` JSON, when it has any.
    pub fn retained_headers_footers_json(&self) -> EngineResult<Option<String>> {
        self.engine
            .retained_headers_footers_json()
            .map_err(EngineError::message)
    }

    /// The retained region layout of the first `pages` pages only; see
    /// `EngineSession::layout_document_with_regions_prefix_retained_json`.
    pub fn layout_document_with_regions_prefix_retained_json(
        &self,
        input: &str,
        pages: u32,
    ) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        self.engine
            .layout_document_with_regions_prefix_retained_json(input, pages as usize)
            .map_err(EngineError::message)
    }

    /// The retained region layout of exactly the first `pages` pages when no
    /// later content can change them, and of the whole document otherwise; see
    /// `EngineSession::layout_document_with_regions_exact_prefix_retained`.
    /// Returns `{ provisional, pages, documentWidth? }` as JSON.
    pub fn layout_document_with_regions_exact_prefix_retained(
        &self,
        input: &str,
        pages: u32,
    ) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        let layout = self
            .engine
            .layout_document_with_regions_exact_prefix_retained(input, pages as usize)
            .map_err(EngineError::message)?;
        serde_json::to_string(&layout).map_err(EngineError::message)
    }

    /// Begins `layout_document_with_regions_retained_json` as a pass measured a
    /// step at a time; see `EngineSession::begin_region_layout`. Returns the
    /// progress JSON, with `layoutJson` once the pass is complete.
    pub fn begin_region_layout(&self, input: &str) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        let progress = self
            .engine
            .begin_region_layout(input)
            .map_err(EngineError::message)?;
        serde_json::to_string(&progress).map_err(EngineError::message)
    }

    /// Measures up to `blocks` more body blocks of the begun pass; see
    /// `EngineSession::resume_region_layout`.
    pub fn resume_region_layout(&self, blocks: u32) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        let progress = self
            .engine
            .resume_region_layout(blocks as usize)
            .map_err(EngineError::message)?;
        serde_json::to_string(&progress).map_err(EngineError::message)
    }

    /// Retained `{ measured, options }` for the main-thread display-list
    /// fallback after a retained-only region layout.
    pub fn retained_kernel_inputs_json(&self) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        self.engine
            .retained_kernel_inputs_json()
            .map_err(EngineError::message)
    }
}
