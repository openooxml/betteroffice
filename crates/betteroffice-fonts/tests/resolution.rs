#![cfg(feature = "manager")]

use betteroffice_fonts::catalog::BundledFontFace;
use betteroffice_fonts::{
    LastResortOffice, resolve_bundled_family_face, resolve_digit_advance,
    resolve_family_script_fallback_face, resolve_glyph_fallback_face, resolve_last_resort_choice,
    resolve_last_resort_face, resolve_light_variant_face, resolve_metric_compat_face,
    resolve_metric_compat_family, resolve_script_fallback_face,
};
use serde_json::{Value, json};

fn file(face: Option<&BundledFontFace>) -> Value {
    face.map_or(Value::Null, |face| json!(face.file))
}

#[test]
fn replays_typescript_resolution_cases() {
    let cases: Vec<Value> =
        serde_json::from_str(include_str!("../data/resolve-cases.json")).unwrap();
    assert!(cases.len() > 500);
    for case in cases {
        let query = &case["query"];
        let bold = query["bold"].as_bool().unwrap();
        let italic = query["italic"].as_bool().unwrap();
        if let (Some(family), Some(script)) = (query["family"].as_str(), query["script"].as_str()) {
            let answer = json!({
                "scriptFallbackFace": file(resolve_family_script_fallback_face(family, script, bold, italic)),
            });
            assert_eq!(answer, case["answer"], "{query}");
            continue;
        }
        if let Some(script) = query["script"].as_str() {
            assert!(char::from_u32(query["codePoint"].as_u64().unwrap() as u32).is_some());
            let answer = json!({
                "scriptFallbackFace": file(resolve_script_fallback_face(script, bold, italic)),
                "glyphFallbackFace": file(resolve_glyph_fallback_face("Segoe UI", bold, italic, LastResortOffice::PowerPoint)),
            });
            assert_eq!(answer, case["answer"], "{query}");
            continue;
        }
        let family = query["family"].as_str().unwrap();
        let mut choices = serde_json::Map::new();
        let mut last_resort = serde_json::Map::new();
        let mut glyph_fallback = serde_json::Map::new();
        for (key, office) in [
            ("powerpoint", LastResortOffice::PowerPoint),
            ("word", LastResortOffice::Word),
        ] {
            let choice = resolve_last_resort_choice(family, bold, italic, office);
            choices.insert(key.into(), json!({ "face": choice.face.file, "onlyIfPrimaryFails": choice.only_if_primary_fails }));
            last_resort.insert(
                key.into(),
                file(Some(resolve_last_resort_face(family, bold, italic, office))),
            );
            glyph_fallback.insert(
                key.into(),
                file(resolve_glyph_fallback_face(family, bold, italic, office)),
            );
        }
        let answer = json!({
            "metricCompatFamily": resolve_metric_compat_family(family),
            "metricCompatFace": file(resolve_metric_compat_face(family, bold, italic)),
            "bundledFamilyFace": file(resolve_bundled_family_face(family, bold, italic)),
            "lightVariantFace": file(resolve_light_variant_face(family, bold, italic)),
            "digitAdvance": resolve_digit_advance(family, bold, italic),
            "lastResortChoice": choices,
            "lastResortFace": last_resort,
            "glyphFallbackFace": glyph_fallback,
        });
        assert_eq!(answer, case["answer"], "{query}");
    }
    assert!(resolve_script_fallback_face("unknown", false, false).is_none());
}

#[test]
fn naskh_styled_families_fall_back_to_naskh_for_arabic_in_every_style() {
    for family in [
        "Simplified Arabic",
        "Traditional Arabic",
        " traditional ARABIC ",
    ] {
        for bold in [false, true] {
            for italic in [false, true] {
                assert_eq!(
                    resolve_family_script_fallback_face(family, "arabic", bold, italic)
                        .map(|face| face.file.as_str()),
                    Some("NotoNaskhArabic-Regular.ttf.zst"),
                    "{family} bold={bold} italic={italic}"
                );
            }
        }
        assert_eq!(
            resolve_family_script_fallback_face(family, "hebrew", false, false)
                .map(|face| face.file.as_str()),
            Some("NotoSansHebrew-Regular.ttf.zst")
        );
    }
    for (family, bold, file) in [
        ("Arial", false, "NotoSansArabic-Regular.ttf.zst"),
        ("Arial", true, "NotoSansArabic-Bold.ttf.zst"),
        ("Times New Roman", false, "NotoSansArabic-Regular.ttf.zst"),
    ] {
        assert_eq!(
            resolve_family_script_fallback_face(family, "arabic", bold, false)
                .map(|face| face.file.as_str()),
            Some(file)
        );
    }
}
