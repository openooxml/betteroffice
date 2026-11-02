use super::{EditSessionCore, EngineError, EngineResult};

pub(super) fn validate_frame_epoch(epoch: f64) -> EngineResult<u64> {
    const MAX_SAFE_INTEGER: f64 = 9_007_199_254_740_991.0;
    if !(epoch.is_finite() && epoch >= 0.0 && epoch.fract() == 0.0 && epoch <= MAX_SAFE_INTEGER) {
        return Err(EngineError::message(
            "expected_frame_epoch must be a non-negative safe integer",
        ));
    }
    Ok(epoch as u64)
}

impl EditSessionCore {
    /// Makes the next display frame a full one whatever epoch its caller
    /// passes, for a host that switches to this engine from another.
    pub fn reset_frame_base(&self) {
        let _fonts = self.fonts.enter();
        self.engine.reset_frame_base();
    }

    /// `{ measured, options, layout }` JSON in, `DisplayList` JSON out, built
    /// against the same resident font store this session measures with.
    pub fn build_display_list_json(&self, input: &str) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        self.engine
            .build_display_list_json(input)
            .map_err(EngineError::message)
    }

    /// Display-only input JSON in, one binary `FrameDelta` v1 out (exposed as
    /// a transferable `Uint8Array`). `expected_frame_epoch` is the epoch of the
    /// frame the caller currently holds; pass `0` for the first frame. A
    /// mismatch makes the engine emit a full frame instead of a delta, and the
    /// returned frame's epoch is always greater than `expected_frame_epoch`. Errors
    /// unless the epoch is a non-negative safe integer, and on build failure.
    pub fn build_display_list_frame(
        &self,
        input: &str,
        expected_frame_epoch: f64,
    ) -> EngineResult<Vec<u8>> {
        let _fonts = self.fonts.enter();
        let epoch = validate_frame_epoch(expected_frame_epoch)?;
        self.engine
            .build_display_list_frame(input, epoch)
            .map_err(EngineError::message)
    }

    /// Limit full display builds to pages `start..end` plus the pages already
    /// built; the others stay unbuilt placeholders carrying their geometry
    /// until [`Self::build_display_pages_frame`] builds them.
    pub fn set_display_window(&self, start: u32, end: u32) {
        let _fonts = self.fonts.enter();
        self.engine
            .set_display_window(Some(start as usize..(end.max(start)) as usize));
    }

    /// Keep every previously built page while windowed builds are on.
    pub fn set_display_retain_built_pages(&self, retain: bool) {
        let _fonts = self.fonts.enter();
        self.engine.set_display_retain_built_pages(retain);
    }

    /// Let an eligible resident text edit re-lower only its paragraph. Off by default.
    pub fn set_local_lowering(&self, enabled: bool) {
        self.engine.set_local_lowering(enabled);
    }

    /// Apply admitted paragraph-local text batches directly. Off by default.
    pub fn set_direct_batches(&self, enabled: bool) {
        self.engine.doc().set_direct_batches(enabled);
    }

    /// The number of text batches applied directly to the live document.
    pub fn direct_batches_applied(&self) -> f64 {
        self.engine.doc().direct_batches_applied() as f64
    }

    /// Limit incremental rebuilds to the display window and caret pages. Off by default.
    pub fn set_windowed_incremental_builds(&self, enabled: bool) {
        let _fonts = self.fonts.enter();
        self.engine.set_windowed_incremental_builds(enabled);
    }

    /// Build the listed pages that are still unbuilt and return a FrameDelta
    /// v1 carrying them; `expected_frame_epoch` works as for
    /// [`Self::build_display_list_frame`].
    pub fn build_display_pages_frame(
        &self,
        pages: Vec<u32>,
        expected_frame_epoch: f64,
    ) -> EngineResult<Vec<u8>> {
        let _fonts = self.fonts.enter();
        let epoch = validate_frame_epoch(expected_frame_epoch)?;
        let pages: Vec<usize> = pages.into_iter().map(|page| page as usize).collect();
        self.engine
            .build_display_pages_frame(&pages, epoch)
            .map_err(EngineError::message)
    }

    /// Release display pages; an empty result means the request was superseded.
    pub fn release_display_pages_frame(
        &self,
        pages: Vec<u32>,
        expected_frame_epoch: f64,
    ) -> EngineResult<Vec<u8>> {
        let _fonts = self.fonts.enter();
        let epoch = validate_frame_epoch(expected_frame_epoch)?;
        let pages: Vec<usize> = pages.into_iter().map(|page| page as usize).collect();
        self.engine
            .release_display_pages_frame(&pages, epoch)
            .map_err(EngineError::message)
    }
}
