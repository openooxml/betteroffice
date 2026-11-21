pub const MANIFEST: &str = include_str!(concat!(env!("FONT_PACK_ASSETS"), "/manifest.json"));

static FACE_0: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansArmenian-Bold.ttf.zst"
));
static FACE_1: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansArmenian-Regular.ttf.zst"
));
static FACE_2: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansBengali-Bold.ttf.zst"
));
static FACE_3: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansBengali-Regular.ttf.zst"
));
static FACE_4: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansCherokee-Bold.ttf.zst"
));
static FACE_5: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansCherokee-Regular.ttf.zst"
));
static FACE_6: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansDevanagari-Bold.ttf.zst"
));
static FACE_7: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansDevanagari-Regular.ttf.zst"
));
static FACE_8: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansEthiopic-Bold.ttf.zst"
));
static FACE_9: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansEthiopic-Regular.ttf.zst"
));
static FACE_10: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansGeorgian-Bold.ttf.zst"
));
static FACE_11: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansGeorgian-Regular.ttf.zst"
));
static FACE_12: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansGujarati-Bold.ttf.zst"
));
static FACE_13: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansGujarati-Regular.ttf.zst"
));
static FACE_14: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansGurmukhi-Bold.ttf.zst"
));
static FACE_15: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansGurmukhi-Regular.ttf.zst"
));
static FACE_16: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansKannada-Bold.ttf.zst"
));
static FACE_17: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansKannada-Regular.ttf.zst"
));
static FACE_18: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansKhmer-Bold.ttf.zst"
));
static FACE_19: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansKhmer-Regular.ttf.zst"
));
static FACE_20: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansLao-Bold.ttf.zst"
));
static FACE_21: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansLao-Regular.ttf.zst"
));
static FACE_22: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansLaoLooped-Bold.ttf.zst"
));
static FACE_23: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansLaoLooped-Regular.ttf.zst"
));
static FACE_24: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansMalayalam-Bold.ttf.zst"
));
static FACE_25: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansMalayalam-Regular.ttf.zst"
));
static FACE_26: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansMongolian-Regular.ttf.zst"
));
static FACE_27: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansMyanmar-Bold.ttf.zst"
));
static FACE_28: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansMyanmar-Regular.ttf.zst"
));
static FACE_29: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansOriya-Bold.ttf.zst"
));
static FACE_30: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansOriya-Regular.ttf.zst"
));
static FACE_31: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansSinhala-Bold.ttf.zst"
));
static FACE_32: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansSinhala-Regular.ttf.zst"
));
static FACE_33: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansSyriac-Regular.ttf.zst"
));
static FACE_34: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansTamil-Bold.ttf.zst"
));
static FACE_35: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansTamil-Regular.ttf.zst"
));
static FACE_36: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansTelugu-Bold.ttf.zst"
));
static FACE_37: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansTelugu-Regular.ttf.zst"
));
static FACE_38: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansThaana-Bold.ttf.zst"
));
static FACE_39: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansThaana-Regular.ttf.zst"
));
static FACE_40: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansThai-Bold.ttf.zst"
));
static FACE_41: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansThai-Regular.ttf.zst"
));
static FACE_42: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansTifinagh-Regular.ttf.zst"
));
static FACE_43: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSerifTibetan-Bold.ttf.zst"
));
static FACE_44: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSerifTibetan-Regular.ttf.zst"
));

pub fn face(file: &str) -> Option<&'static [u8]> {
    match file {
        "NotoSansArmenian-Bold.ttf.zst" => Some(FACE_0),
        "NotoSansArmenian-Regular.ttf.zst" => Some(FACE_1),
        "NotoSansBengali-Bold.ttf.zst" => Some(FACE_2),
        "NotoSansBengali-Regular.ttf.zst" => Some(FACE_3),
        "NotoSansCherokee-Bold.ttf.zst" => Some(FACE_4),
        "NotoSansCherokee-Regular.ttf.zst" => Some(FACE_5),
        "NotoSansDevanagari-Bold.ttf.zst" => Some(FACE_6),
        "NotoSansDevanagari-Regular.ttf.zst" => Some(FACE_7),
        "NotoSansEthiopic-Bold.ttf.zst" => Some(FACE_8),
        "NotoSansEthiopic-Regular.ttf.zst" => Some(FACE_9),
        "NotoSansGeorgian-Bold.ttf.zst" => Some(FACE_10),
        "NotoSansGeorgian-Regular.ttf.zst" => Some(FACE_11),
        "NotoSansGujarati-Bold.ttf.zst" => Some(FACE_12),
        "NotoSansGujarati-Regular.ttf.zst" => Some(FACE_13),
        "NotoSansGurmukhi-Bold.ttf.zst" => Some(FACE_14),
        "NotoSansGurmukhi-Regular.ttf.zst" => Some(FACE_15),
        "NotoSansKannada-Bold.ttf.zst" => Some(FACE_16),
        "NotoSansKannada-Regular.ttf.zst" => Some(FACE_17),
        "NotoSansKhmer-Bold.ttf.zst" => Some(FACE_18),
        "NotoSansKhmer-Regular.ttf.zst" => Some(FACE_19),
        "NotoSansLao-Bold.ttf.zst" => Some(FACE_20),
        "NotoSansLao-Regular.ttf.zst" => Some(FACE_21),
        "NotoSansLaoLooped-Bold.ttf.zst" => Some(FACE_22),
        "NotoSansLaoLooped-Regular.ttf.zst" => Some(FACE_23),
        "NotoSansMalayalam-Bold.ttf.zst" => Some(FACE_24),
        "NotoSansMalayalam-Regular.ttf.zst" => Some(FACE_25),
        "NotoSansMongolian-Regular.ttf.zst" => Some(FACE_26),
        "NotoSansMyanmar-Bold.ttf.zst" => Some(FACE_27),
        "NotoSansMyanmar-Regular.ttf.zst" => Some(FACE_28),
        "NotoSansOriya-Bold.ttf.zst" => Some(FACE_29),
        "NotoSansOriya-Regular.ttf.zst" => Some(FACE_30),
        "NotoSansSinhala-Bold.ttf.zst" => Some(FACE_31),
        "NotoSansSinhala-Regular.ttf.zst" => Some(FACE_32),
        "NotoSansSyriac-Regular.ttf.zst" => Some(FACE_33),
        "NotoSansTamil-Bold.ttf.zst" => Some(FACE_34),
        "NotoSansTamil-Regular.ttf.zst" => Some(FACE_35),
        "NotoSansTelugu-Bold.ttf.zst" => Some(FACE_36),
        "NotoSansTelugu-Regular.ttf.zst" => Some(FACE_37),
        "NotoSansThaana-Bold.ttf.zst" => Some(FACE_38),
        "NotoSansThaana-Regular.ttf.zst" => Some(FACE_39),
        "NotoSansThai-Bold.ttf.zst" => Some(FACE_40),
        "NotoSansThai-Regular.ttf.zst" => Some(FACE_41),
        "NotoSansTifinagh-Regular.ttf.zst" => Some(FACE_42),
        "NotoSerifTibetan-Bold.ttf.zst" => Some(FACE_43),
        "NotoSerifTibetan-Regular.ttf.zst" => Some(FACE_44),
        _ => None,
    }
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
