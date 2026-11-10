use wasm_bindgen::prelude::*;

#[wasm_bindgen]
#[derive(Default)]
pub struct MetafileConverter {
    session: ooxml_metafile::SvgSession,
}

#[wasm_bindgen]
impl MetafileConverter {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self::default()
    }

    #[wasm_bindgen(js_name = toSvg)]
    pub fn to_svg(&mut self, bytes: &[u8]) -> Option<Vec<u8>> {
        self.session.convert(bytes)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const TEXT: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/metafile-text.emf");
    const BITMAP: &[u8] =
        include_bytes!("../../pptx-render/tests/fixtures/metafile-bitmap-line.emf");

    #[test]
    fn matches_session_conversion() {
        let mut converter = MetafileConverter::new();
        let mut cache = pptx_edit::MetafileSvgCache::default();
        for (key, bytes) in [("text", TEXT), ("bitmap", BITMAP)] {
            let expected = cache.resolve(key, bytes).unwrap();
            assert_eq!(
                converter.to_svg(bytes).unwrap().as_slice(),
                expected.as_ref()
            );
        }
    }

    #[test]
    fn refuses_non_metafiles_and_oversized_sources() {
        let mut converter = MetafileConverter::new();
        assert!(converter.to_svg(b"not a metafile").is_none());
        let mut bytes = TEXT.to_vec();
        bytes.resize(16 * 1024 * 1024 + 1, 0);
        assert!(converter.to_svg(&bytes).is_none());
    }
}
