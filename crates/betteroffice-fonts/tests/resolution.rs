#![cfg(feature = "manager")]

use betteroffice_fonts::catalog::{BundledFontFace, catalog};
use betteroffice_fonts::{
    LastResortOffice, resolve_bundled_family_face, resolve_digit_advance,
    resolve_family_script_fallback_face, resolve_glyph_fallback_face, resolve_last_resort_choice,
    resolve_last_resort_face, resolve_light_variant_face, resolve_metric_compat_face,
    resolve_metric_compat_family, resolve_script_fallback_face, resolve_word_family_substitute,
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
            "wordFamilySubstitute": resolve_word_family_substitute(family)
                .map(|entry| json!({ "remote": entry.remote, "generic": entry.generic })),
        });
        assert_eq!(answer, case["answer"], "{query}");
    }
    assert!(resolve_script_fallback_face("unknown", false, false).is_none());
}

#[test]
fn light_variants_follow_the_requested_style_of_bold_capable_families() {
    for base in [
        "Calibri",
        "Cambria",
        "Arial",
        "Times New Roman",
        "Courier New",
        "Georgia",
        "Inter",
        "Roboto",
        "Source Sans Pro",
        "DM Sans",
        "Open Sans",
        "Montserrat",
        "Poppins",
        "Oswald",
        "Heebo",
        "Comic Sans MS",
    ] {
        assert_eq!(
            resolve_metric_compat_face(base, true, false)
                .unwrap()
                .weight,
            700
        );
        for suffix in &catalog().lighter_than_regular {
            let family = format!("{base} {suffix}");
            for (bold, italic) in [(true, false), (true, true), (false, false), (false, true)] {
                let label = format!("{family} bold={bold} italic={italic}");
                let expected = resolve_metric_compat_face(base, bold, italic);
                assert_eq!(
                    resolve_light_variant_face(&family, bold, italic),
                    expected,
                    "{label}"
                );
                for (office, sans) in [
                    (LastResortOffice::Word, "Arial"),
                    (LastResortOffice::PowerPoint, "Calibri"),
                ] {
                    let fallback = || resolve_metric_compat_face(sans, bold, italic).unwrap();
                    let last = expected.unwrap_or_else(fallback);
                    assert_eq!(
                        resolve_last_resort_face(&family, bold, italic, office),
                        last,
                        "{label} {office:?}"
                    );
                }
            }
        }
    }
}

#[test]
fn lighter_members_of_segoe_ui_keep_their_light_faces_when_bold() {
    for base in ["Segoe UI", "Segoe UI Semibold"] {
        for suffix in &catalog().lighter_than_regular {
            for italic in [false, true] {
                assert_eq!(
                    resolve_light_variant_face(&format!("{base} {suffix}"), true, italic),
                    resolve_metric_compat_face(base, false, italic),
                    "{base} {suffix} italic={italic}"
                );
            }
            let family = format!("{base} {suffix}");
            if resolve_metric_compat_family(&family).is_some() {
                continue;
            }
            for bold in [false, true] {
                for (office, sans) in [
                    (LastResortOffice::Word, "Arial"),
                    (LastResortOffice::PowerPoint, "Calibri"),
                ] {
                    assert_eq!(
                        resolve_last_resort_face(&family, bold, false, office).family,
                        "Selawik",
                        "{family} bold={bold} {office:?}"
                    );
                    assert_eq!(
                        resolve_glyph_fallback_face(&family, bold, false, office),
                        resolve_metric_compat_face(sans, false, false),
                        "{family} bold={bold} {office:?}"
                    );
                }
            }
        }
    }
}

#[test]
fn bold_light_variants_keep_existing_faces_when_the_substitute_has_no_bold() {
    for base in [
        "DM Serif Display",
        "Malgun Gothic",
        "Microsoft YaHei",
        "Microsoft JhengHei",
        "SimSun",
        "MS Gothic",
        "MS PGothic",
        "MS UI Gothic",
    ] {
        for suffix in &catalog().lighter_than_regular {
            let family = format!("{base} {suffix}");
            for italic in [false, true] {
                let expected = resolve_metric_compat_face(base, false, italic).unwrap();
                assert_eq!(
                    resolve_light_variant_face(&family, true, italic),
                    Some(expected),
                    "{family} bold=true italic={italic}"
                );
                for office in [LastResortOffice::Word, LastResortOffice::PowerPoint] {
                    assert_eq!(
                        resolve_last_resort_face(&family, true, italic, office),
                        expected,
                        "{family} bold=true italic={italic} {office:?}"
                    );
                }
            }
        }
    }
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
}

#[test]
fn arabic_in_latin_families_falls_back_to_naskh_in_every_style() {
    for bold in [false, true] {
        for italic in [false, true] {
            let default = resolve_script_fallback_face("arabic", bold, italic).unwrap();
            assert_eq!(default.file, "NotoNaskhArabic-Regular.ttf.zst");
            for family in ["Arial", "Times New Roman"] {
                assert_eq!(
                    resolve_family_script_fallback_face(family, "arabic", bold, italic)
                        .map(|face| face.file.as_str()),
                    Some(default.file.as_str()),
                );
            }
        }
        assert_eq!(
            resolve_bundled_family_face("Noto Sans Arabic", bold, false)
                .map(|face| face.family.as_str()),
            Some("Noto Sans Arabic"),
        );
        assert_eq!(
            resolve_family_script_fallback_face(" Noto Sans Arabic ", "arabic", bold, true)
                .map(|face| (face.family.as_str(), face.weight)),
            Some(("Noto Sans Arabic", if bold { 700 } else { 400 })),
        );
    }
}
