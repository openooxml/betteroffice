use wasm_bindgen::prelude::*;

#[wasm_bindgen(js_name = metafileToSvg)]
pub fn metafile_to_svg(bytes: &[u8]) -> Option<Vec<u8>> {
    if bytes.len() > 16 * 1024 * 1024 || !ooxml_metafile::is_metafile(bytes) {
        return None;
    }
    let picture = ooxml_metafile::ReplayBudget::default();
    let mut replay = ooxml_metafile::ReplayBudget {
        work: 4 * picture.work,
        pixels: 4 * picture.pixels,
    };
    let mut output = 64 * 1024 * 1024;
    ooxml_metafile::to_svg_with_limits(bytes, &mut replay, &mut output)
        .ok()
        .map(|svg| svg.markup.into_bytes())
}

#[cfg(test)]
mod tests {
    use super::*;

    const TEXT: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/metafile-text.emf");
    const BITMAP: &[u8] =
        include_bytes!("../../pptx-render/tests/fixtures/metafile-bitmap-line.emf");

    #[test]
    fn matches_session_conversion() {
        for bytes in [TEXT, BITMAP] {
            let expected = pptx_edit::MetafileSvgCache::default()
                .resolve("picture", bytes)
                .unwrap();
            assert_eq!(metafile_to_svg(bytes).unwrap().as_slice(), expected.as_ref());
        }
    }

    #[test]
    fn refuses_non_metafiles_and_oversized_sources() {
        assert!(metafile_to_svg(b"not a metafile").is_none());
        let mut bytes = TEXT.to_vec();
        bytes.resize(16 * 1024 * 1024 + 1, 0);
        assert!(metafile_to_svg(&bytes).is_none());
    }
}
