pub const MANIFEST: &str = include_str!(concat!(env!("FONT_PACK_ASSETS"), "/manifest.json"));

static FACE_0: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Caladea-Bold.ttf.zst"));
static FACE_1: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Caladea-BoldItalic.ttf.zst"
));
static FACE_2: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Caladea-Italic.ttf.zst"));
static FACE_3: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Caladea-Regular.ttf.zst"
));
static FACE_4: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Carlito-Bold.ttf.zst"));
static FACE_5: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Carlito-BoldItalic.ttf.zst"
));
static FACE_6: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Carlito-Italic.ttf.zst"));
static FACE_7: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Carlito-Regular.ttf.zst"
));
static FACE_8: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/ComicRelief-Bold.ttf.zst"
));
static FACE_9: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/ComicRelief-Regular.ttf.zst"
));
static FACE_10: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/DMSans-Bold.ttf.zst"));
static FACE_11: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/DMSans-BoldItalic.ttf.zst"
));
static FACE_12: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/DMSans-Italic.ttf.zst"));
static FACE_13: &[u8] =
    include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/DMSans-Regular.ttf.zst"));
static FACE_14: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/DMSerifDisplay-Italic.ttf.zst"
));
static FACE_15: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/DMSerifDisplay-Regular.ttf.zst"
));
static FACE_16: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Gelasio-Bold.ttf.zst"));
static FACE_17: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Gelasio-BoldItalic.ttf.zst"
));
static FACE_18: &[u8] =
    include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Gelasio-Italic.ttf.zst"));
static FACE_19: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Gelasio-Regular.ttf.zst"
));
static FACE_20: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Heebo-Bold.ttf.zst"));
static FACE_21: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Heebo-Regular.ttf.zst"));
static FACE_22: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Inter-Bold.ttf.zst"));
static FACE_23: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Inter-BoldItalic.ttf.zst"
));
static FACE_24: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Inter-Italic.ttf.zst"));
static FACE_25: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Inter-Regular.ttf.zst"));
static FACE_26: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/LiberationMono-Bold.ttf.zst"
));
static FACE_27: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/LiberationMono-BoldItalic.ttf.zst"
));
static FACE_28: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/LiberationMono-Italic.ttf.zst"
));
static FACE_29: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/LiberationMono-Regular.ttf.zst"
));
static FACE_30: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/LiberationSans-Bold.ttf.zst"
));
static FACE_31: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/LiberationSans-BoldItalic.ttf.zst"
));
static FACE_32: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/LiberationSans-Italic.ttf.zst"
));
static FACE_33: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/LiberationSans-Regular.ttf.zst"
));
static FACE_34: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/LiberationSerif-Bold.ttf.zst"
));
static FACE_35: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/LiberationSerif-BoldItalic.ttf.zst"
));
static FACE_36: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/LiberationSerif-Italic.ttf.zst"
));
static FACE_37: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/LiberationSerif-Regular.ttf.zst"
));
static FACE_38: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Montserrat-Bold.ttf.zst"
));
static FACE_39: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Montserrat-BoldItalic.ttf.zst"
));
static FACE_40: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Montserrat-Italic.ttf.zst"
));
static FACE_41: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Montserrat-Regular.ttf.zst"
));
static FACE_42: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoNaskhArabic-Regular.ttf.zst"
));
static FACE_43: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansArabic-Bold.ttf.zst"
));
static FACE_44: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansArabic-Regular.ttf.zst"
));
static FACE_45: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansHebrew-Bold.ttf.zst"
));
static FACE_46: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/NotoSansHebrew-Regular.ttf.zst"
));
static FACE_47: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/OpenSans-Bold.ttf.zst"));
static FACE_48: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/OpenSans-BoldItalic.ttf.zst"
));
static FACE_49: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/OpenSans-Italic.ttf.zst"
));
static FACE_50: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/OpenSans-Regular.ttf.zst"
));
static FACE_51: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Oswald-Bold.ttf.zst"));
static FACE_52: &[u8] =
    include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Oswald-Regular.ttf.zst"));
static FACE_53: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Poppins-Bold.ttf.zst"));
static FACE_54: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Poppins-BoldItalic.ttf.zst"
));
static FACE_55: &[u8] =
    include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Poppins-Italic.ttf.zst"));
static FACE_56: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Poppins-Regular.ttf.zst"
));
static FACE_57: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Roboto-Bold.ttf.zst"));
static FACE_58: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Roboto-BoldItalic.ttf.zst"
));
static FACE_59: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Roboto-Italic.ttf.zst"));
static FACE_60: &[u8] =
    include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Roboto-Regular.ttf.zst"));
static FACE_61: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Selawik-Bold.ttf.zst"));
static FACE_62: &[u8] = include_bytes!(concat!(env!("FONT_PACK_ASSETS"), "/Selawik-Light.ttf.zst"));
static FACE_63: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Selawik-Regular.ttf.zst"
));
static FACE_64: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Selawik-Semibold.ttf.zst"
));
static FACE_65: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/Selawik-Semilight.ttf.zst"
));
static FACE_66: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/SourceSans3-Bold.ttf.zst"
));
static FACE_67: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/SourceSans3-BoldItalic.ttf.zst"
));
static FACE_68: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/SourceSans3-Italic.ttf.zst"
));
static FACE_69: &[u8] = include_bytes!(concat!(
    env!("FONT_PACK_ASSETS"),
    "/SourceSans3-Regular.ttf.zst"
));

pub fn face(file: &str) -> Option<&'static [u8]> {
    match file {
        "Caladea-Bold.ttf.zst" => Some(FACE_0),
        "Caladea-BoldItalic.ttf.zst" => Some(FACE_1),
        "Caladea-Italic.ttf.zst" => Some(FACE_2),
        "Caladea-Regular.ttf.zst" => Some(FACE_3),
        "Carlito-Bold.ttf.zst" => Some(FACE_4),
        "Carlito-BoldItalic.ttf.zst" => Some(FACE_5),
        "Carlito-Italic.ttf.zst" => Some(FACE_6),
        "Carlito-Regular.ttf.zst" => Some(FACE_7),
        "ComicRelief-Bold.ttf.zst" => Some(FACE_8),
        "ComicRelief-Regular.ttf.zst" => Some(FACE_9),
        "DMSans-Bold.ttf.zst" => Some(FACE_10),
        "DMSans-BoldItalic.ttf.zst" => Some(FACE_11),
        "DMSans-Italic.ttf.zst" => Some(FACE_12),
        "DMSans-Regular.ttf.zst" => Some(FACE_13),
        "DMSerifDisplay-Italic.ttf.zst" => Some(FACE_14),
        "DMSerifDisplay-Regular.ttf.zst" => Some(FACE_15),
        "Gelasio-Bold.ttf.zst" => Some(FACE_16),
        "Gelasio-BoldItalic.ttf.zst" => Some(FACE_17),
        "Gelasio-Italic.ttf.zst" => Some(FACE_18),
        "Gelasio-Regular.ttf.zst" => Some(FACE_19),
        "Heebo-Bold.ttf.zst" => Some(FACE_20),
        "Heebo-Regular.ttf.zst" => Some(FACE_21),
        "Inter-Bold.ttf.zst" => Some(FACE_22),
        "Inter-BoldItalic.ttf.zst" => Some(FACE_23),
        "Inter-Italic.ttf.zst" => Some(FACE_24),
        "Inter-Regular.ttf.zst" => Some(FACE_25),
        "LiberationMono-Bold.ttf.zst" => Some(FACE_26),
        "LiberationMono-BoldItalic.ttf.zst" => Some(FACE_27),
        "LiberationMono-Italic.ttf.zst" => Some(FACE_28),
        "LiberationMono-Regular.ttf.zst" => Some(FACE_29),
        "LiberationSans-Bold.ttf.zst" => Some(FACE_30),
        "LiberationSans-BoldItalic.ttf.zst" => Some(FACE_31),
        "LiberationSans-Italic.ttf.zst" => Some(FACE_32),
        "LiberationSans-Regular.ttf.zst" => Some(FACE_33),
        "LiberationSerif-Bold.ttf.zst" => Some(FACE_34),
        "LiberationSerif-BoldItalic.ttf.zst" => Some(FACE_35),
        "LiberationSerif-Italic.ttf.zst" => Some(FACE_36),
        "LiberationSerif-Regular.ttf.zst" => Some(FACE_37),
        "Montserrat-Bold.ttf.zst" => Some(FACE_38),
        "Montserrat-BoldItalic.ttf.zst" => Some(FACE_39),
        "Montserrat-Italic.ttf.zst" => Some(FACE_40),
        "Montserrat-Regular.ttf.zst" => Some(FACE_41),
        "NotoNaskhArabic-Regular.ttf.zst" => Some(FACE_42),
        "NotoSansArabic-Bold.ttf.zst" => Some(FACE_43),
        "NotoSansArabic-Regular.ttf.zst" => Some(FACE_44),
        "NotoSansHebrew-Bold.ttf.zst" => Some(FACE_45),
        "NotoSansHebrew-Regular.ttf.zst" => Some(FACE_46),
        "OpenSans-Bold.ttf.zst" => Some(FACE_47),
        "OpenSans-BoldItalic.ttf.zst" => Some(FACE_48),
        "OpenSans-Italic.ttf.zst" => Some(FACE_49),
        "OpenSans-Regular.ttf.zst" => Some(FACE_50),
        "Oswald-Bold.ttf.zst" => Some(FACE_51),
        "Oswald-Regular.ttf.zst" => Some(FACE_52),
        "Poppins-Bold.ttf.zst" => Some(FACE_53),
        "Poppins-BoldItalic.ttf.zst" => Some(FACE_54),
        "Poppins-Italic.ttf.zst" => Some(FACE_55),
        "Poppins-Regular.ttf.zst" => Some(FACE_56),
        "Roboto-Bold.ttf.zst" => Some(FACE_57),
        "Roboto-BoldItalic.ttf.zst" => Some(FACE_58),
        "Roboto-Italic.ttf.zst" => Some(FACE_59),
        "Roboto-Regular.ttf.zst" => Some(FACE_60),
        "Selawik-Bold.ttf.zst" => Some(FACE_61),
        "Selawik-Light.ttf.zst" => Some(FACE_62),
        "Selawik-Regular.ttf.zst" => Some(FACE_63),
        "Selawik-Semibold.ttf.zst" => Some(FACE_64),
        "Selawik-Semilight.ttf.zst" => Some(FACE_65),
        "SourceSans3-Bold.ttf.zst" => Some(FACE_66),
        "SourceSans3-BoldItalic.ttf.zst" => Some(FACE_67),
        "SourceSans3-Italic.ttf.zst" => Some(FACE_68),
        "SourceSans3-Regular.ttf.zst" => Some(FACE_69),
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
