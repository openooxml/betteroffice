pub const MANIFEST: &str = include_str!(concat!(env!("FONT_PACK_ASSETS"), "/manifest.json"));

pub fn face(file: &str) -> Option<&'static [u8]> {
    betteroffice_fonts_cjk_sc::face(file)
        .or_else(|| betteroffice_fonts_cjk_tc::face(file))
        .or_else(|| betteroffice_fonts_cjk_jp::face(file))
        .or_else(|| betteroffice_fonts_cjk_kr::face(file))
        .or_else(|| betteroffice_fonts_cjk_serif_sc::face(file))
}

#[cfg(test)]
mod tests {
    use super::{MANIFEST, face};

    #[test]
    fn manifest_faces_have_the_encoded_length() {
        let manifest: serde_json::Value = serde_json::from_str(MANIFEST).unwrap();
        assert!(!manifest["faces"].as_array().unwrap().is_empty());
        for entry in manifest["faces"].as_array().unwrap() {
            let file = entry["file"].as_str().unwrap();
            assert_eq!(
                face(file).unwrap().len() as u64,
                entry["encodedLength"].as_u64().unwrap()
            );
        }
        assert!(face("unknown.ttf.zst").is_none());
    }
}
