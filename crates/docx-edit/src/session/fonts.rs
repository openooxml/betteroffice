use super::{EditSessionCore, EngineError, EngineResult};

#[cfg(test)]
#[path = "../../tests/support/base_latin_coverage.rs"]
mod coverage_fixture;

#[cfg(test)]
#[test]
fn base_latin_misses_follow_transformed_text_and_font_chains() {
    let core = EditSessionCore::new(9670, super::Host::deterministic());
    coverage_fixture::check(
        |bytes| core.register_measure_font(bytes).unwrap(),
        |input| core.measure_paragraph_json(input).is_ok(),
        || core.take_base_latin_miss(),
    );
}

impl EditSessionCore {
    /// Registers raw sfnt bytes in this session's resident measurement store
    /// and returns the font id that measurement and display inputs reference.
    /// Errors on bytes the font parser rejects.
    pub fn register_measure_font<'a>(
        &self,
        bytes: impl Into<std::borrow::Cow<'a, [u8]>>,
    ) -> EngineResult<u32> {
        let _fonts = self.fonts.enter();
        docx_layout::register_measure_font_bytes(bytes).map_err(EngineError::message)
    }

    /// Registers a measurement view of `base` carrying the vertical metrics
    /// and advance pitch Word measures `requested_family` with, for a face
    /// this host had to substitute. Returns `base` for a family whose metrics
    /// are unknown.
    pub fn register_substitute_measure_font(
        &self,
        base: u32,
        requested_family: &str,
    ) -> EngineResult<u32> {
        let _fonts = self.fonts.enter();
        docx_layout::register_substitute_measure_font_id(base, requested_family)
            .map_err(EngineError::message)
    }

    /// Replaces the outlines of measurement font `font_id` with `bytes`, a face
    /// whose every table but `CFF ` is unchanged (a sliced face reassembled
    /// with more chunks). Errors, leaving the font as it was, for other bytes.
    pub fn update_measure_font_outlines(&self, font_id: u32, bytes: &[u8]) -> EngineResult<()> {
        let _fonts = self.fonts.enter();
        docx_layout::update_measure_font_outlines(font_id, bytes).map_err(EngineError::message)
    }

    /// The outline of one glyph as JSON in font units, y-up. Errors when the
    /// glyph cannot be extracted.
    pub fn outline_glyph_json(&self, font_id: u32, glyph_id: u32) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        docx_layout::outline_glyph_json_resident(font_id, glyph_id).map_err(EngineError::message)
    }

    pub fn decode_font(&self, bytes: &[u8]) -> EngineResult<Vec<u8>> {
        crate::decode_resident_font(bytes).map_err(EngineError::message)
    }

    pub fn take_base_latin_miss(&self) -> bool {
        let _fonts = self.fonts.enter();
        docx_layout::take_base_latin_miss()
    }

    /// Drops every registered measurement font (ids restart at zero) and
    /// invalidates the retained paragraph measurement templates, so the next
    /// edit must pass back through the full layout path.
    pub fn clear_measure_fonts(&self) {
        let _fonts = self.fonts.enter();
        docx_layout::clear_measure_fonts();
        self.engine.clear_measurement_templates();
    }

    /// Typing echo model of one font at one shaping size from this session's
    /// resident font store: `{"horizontal","glyphIds","advances","kerning",
    /// "unsafePairs"}`, indexed by the characters of `alphabet`. `kern` is the
    /// run's kerning gate. Errors for an unknown font or a non-positive size.
    pub fn echo_glyph_set_json(
        &self,
        font_id: u32,
        size: f32,
        kern: bool,
        alphabet: &str,
    ) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        docx_layout::echo_glyph_set_json(font_id, size, kern, alphabet)
            .map_err(EngineError::message)
    }
}
