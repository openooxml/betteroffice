/// The CJK coverage bucket a BCP 47 language tag selects for Han text: `ja` Japanese,
/// `ko` Korean, `zh` with `Hant` or region TW/HK/MO Traditional, other `zh` Simplified.
pub fn han_script_for_language(tag: &str) -> Option<&'static str> {
    let mut parts = tag.split(['-', '_']).filter(|part| !part.is_empty());
    let language = parts.next()?;
    if language.eq_ignore_ascii_case("ja") {
        return Some("cjk-jp");
    }
    if language.eq_ignore_ascii_case("ko") {
        return Some("cjk-kr");
    }
    if !language.eq_ignore_ascii_case("zh") {
        return None;
    }
    let mut traditional = false;
    for part in parts {
        if part.eq_ignore_ascii_case("hans") {
            return Some("cjk-sc");
        }
        if ["hant", "tw", "hk", "mo"]
            .iter()
            .any(|tag| part.eq_ignore_ascii_case(tag))
        {
            traditional = true;
        }
    }
    Some(if traditional { "cjk-tc" } else { "cjk-sc" })
}

/// Script coverage buckets for paragraph text, preserving context across runs.
pub fn paragraph_scripts_with_han_fallback<'a>(
    texts: impl IntoIterator<Item = &'a str>,
    include_han_fallback: bool,
) -> Vec<String> {
    paragraph_scripts_with_languages(
        texts.into_iter().map(|text| (text, None)),
        include_han_fallback,
    )
}

/// Script coverage buckets for a paragraph's runs, each with the language that decides its
/// Han characters. Kana and Hangul anywhere in the paragraph keep deciding as before; a run
/// whose language maps through [`han_script_for_language`] takes that bucket instead, and
/// `include_han_fallback` adds Simplified only for Han in runs without such a language.
pub fn paragraph_scripts_with_languages<'a>(
    texts: impl IntoIterator<Item = (&'a str, Option<&'a str>)>,
    include_han_fallback: bool,
) -> Vec<String> {
    scripts_of(&mut texts.into_iter(), include_han_fallback)
}

fn is_han(point: u32) -> bool {
    matches!(
        point,
        0x3000..=0x303f
            | 0x3400..=0x4dbf
            | 0x4e00..=0x9fff
            | 0xf900..=0xfaff
            | 0xfe30..=0xfe4f
            | 0xff00..=0xff65
            | 0x20000..=0x3ffff
    )
}

/// Whether `text` has a character whose face [`han_script_for_language`] chooses.
pub fn contains_han(text: &str) -> bool {
    text.chars().any(|character| is_han(character as u32))
}

fn scripts_of<'a>(
    texts: &mut dyn Iterator<Item = (&'a str, Option<&'a str>)>,
    include_han_fallback: bool,
) -> Vec<String> {
    let mut han_unknown = false;
    let mut han_buckets = [false; 4];
    let mut kana = false;
    let mut hangul = false;
    let mut arabic = false;
    let mut hebrew = false;
    let mut intl = Vec::new();
    for (text, language) in texts {
        let mut han = false;
        for character in text.chars() {
            let point = character as u32;
            match point {
                0x0590..=0x05ff | 0xfb1d..=0xfb4f => hebrew = true,
                0x0600..=0x06ff
                | 0x0750..=0x077f
                | 0x0870..=0x08ff
                | 0xfb50..=0xfdff
                | 0xfe70..=0xfeff => arabic = true,
                0x1100..=0x11ff
                | 0x3130..=0x318f
                | 0xa960..=0xa97f
                | 0xac00..=0xd7ff
                | 0xffa0..=0xffdc => hangul = true,
                0x3040..=0x30ff | 0x31f0..=0x31ff | 0xff66..=0xff9f => kana = true,
                point if is_han(point) => han = true,
                _ => {}
            }
            let script = match point {
                0x0e00..=0x0e7f => Some("thai"),
                0x0e80..=0x0eff => Some("lao"),
                0x1780..=0x17ff | 0x19e0..=0x19ff => Some("khmer"),
                0x1000..=0x109f | 0xa9e0..=0xa9ff | 0xaa60..=0xaa7f | 0x116d0..=0x116ff => {
                    Some("myanmar")
                }
                0x0900..=0x097f | 0xa8e0..=0xa8ff | 0x11b00..=0x11b5f => Some("devanagari"),
                0x0980..=0x09ff => Some("bengali"),
                0x0a00..=0x0a7f => Some("gurmukhi"),
                0x0a80..=0x0aff => Some("gujarati"),
                0x0b00..=0x0b7f => Some("oriya"),
                0x0b80..=0x0bff | 0x11fc0..=0x11fff => Some("tamil"),
                0x0c00..=0x0c7f => Some("telugu"),
                0x0c80..=0x0cff => Some("kannada"),
                0x0d00..=0x0d7f => Some("malayalam"),
                0x0d80..=0x0dff | 0x111e0..=0x111ff => Some("sinhala"),
                0x0530..=0x058f | 0xfb13..=0xfb17 => Some("armenian"),
                0x10a0..=0x10ff | 0x1c90..=0x1cbf | 0x2d00..=0x2d2f => Some("georgian"),
                0x1200..=0x137f
                | 0x1380..=0x139f
                | 0x2d80..=0x2ddf
                | 0xab00..=0xab2f
                | 0x1e7e0..=0x1e7ff => Some("ethiopic"),
                0x0f00..=0x0fff => Some("tibetan"),
                0x1800..=0x18af | 0x11660..=0x1167f => Some("mongolian"),
                0x0700..=0x074f | 0x0860..=0x086f => Some("syriac"),
                0x0780..=0x07bf => Some("thaana"),
                0x13a0..=0x13ff | 0xab70..=0xabbf => Some("cherokee"),
                0x2d30..=0x2d7f => Some("tifinagh"),
                _ => None,
            };
            let Some(script) = script else {
                continue;
            };
            if !intl.contains(&script) {
                intl.push(script);
            }
        }
        if han {
            match language.and_then(han_script_for_language) {
                Some("cjk-jp") => han_buckets[0] = true,
                Some("cjk-kr") => han_buckets[1] = true,
                Some("cjk-tc") => han_buckets[2] = true,
                Some(_) => han_buckets[3] = true,
                None => han_unknown = true,
            }
        }
    }
    let mut scripts = Vec::new();
    if kana || han_buckets[0] {
        scripts.push("cjk-jp".to_owned());
    }
    if hangul || han_buckets[1] {
        scripts.push("cjk-kr".to_owned());
    }
    if han_buckets[2] {
        scripts.push("cjk-tc".to_owned());
    }
    if han_buckets[3] || (han_unknown && (include_han_fallback || (!kana && !hangul))) {
        scripts.push("cjk-sc".to_owned());
    }
    if arabic {
        scripts.push("arabic".to_owned());
    }
    if hebrew {
        scripts.push("hebrew".to_owned());
    }
    scripts.extend(intl.into_iter().map(str::to_owned));
    scripts
}

#[cfg(test)]
mod tests {
    use super::{
        han_script_for_language, paragraph_scripts_with_han_fallback,
        paragraph_scripts_with_languages,
    };

    #[test]
    fn cjk_uses_paragraph_context_across_runs() {
        for (texts, expected) in [
            (vec!["Latin"], vec![]),
            (vec!["日本語", "かな"], vec!["cjk-jp"]),
            (vec!["漢字", "한글"], vec!["cjk-kr"]),
            (vec!["漢字"], vec!["cjk-sc"]),
            (vec!["かな", "한글", "漢字"], vec!["cjk-jp", "cjk-kr"]),
            (vec!["ｶﾅ"], vec!["cjk-jp"]),
        ] {
            assert_eq!(paragraph_scripts_with_han_fallback(texts, false), expected);
        }
    }

    #[test]
    fn preview_keeps_han_fallback_after_contextual_scripts() {
        assert_eq!(
            paragraph_scripts_with_han_fallback(["日本語かな한글"], true),
            ["cjk-jp", "cjk-kr", "cjk-sc"]
        );
    }

    #[test]
    fn collects_rtl_and_international_scripts_once() {
        assert_eq!(
            paragraph_scripts_with_han_fallback(["العربية עברית नमस्ते ไทย नमस्ते"], false),
            ["arabic", "hebrew", "devanagari", "thai"]
        );
    }

    #[test]
    fn language_tags_select_han_buckets() {
        for (tag, expected) in [
            ("ja-JP", Some("cjk-jp")),
            ("JA", Some("cjk-jp")),
            ("ko_KR", Some("cjk-kr")),
            ("zh-CN", Some("cjk-sc")),
            ("zh-SG", Some("cjk-sc")),
            ("zh", Some("cjk-sc")),
            ("zh-TW", Some("cjk-tc")),
            ("zh-HK", Some("cjk-tc")),
            ("zh-MO", Some("cjk-tc")),
            ("zh-Hant", Some("cjk-tc")),
            ("zh-Hant-CN", Some("cjk-tc")),
            ("zh-Hans-HK", Some("cjk-sc")),
            ("en-US", None),
            ("", None),
            ("jav", None),
        ] {
            assert_eq!(han_script_for_language(tag), expected, "{tag}");
        }
    }

    #[test]
    fn run_languages_decide_han_without_dropping_faces() {
        for (runs, fallback, expected) in [
            (vec![("日本語", Some("ja-JP"))], false, vec!["cjk-jp"]),
            (
                vec![("日本語", Some("ja-JP")), ("かな", Some("ja-JP"))],
                true,
                vec!["cjk-jp"],
            ),
            (
                vec![("日本語", None), ("かな", None)],
                true,
                vec!["cjk-jp", "cjk-sc"],
            ),
            (
                vec![("かな", Some("ja-JP")), ("繁體", Some("zh-TW"))],
                false,
                vec!["cjk-jp", "cjk-tc"],
            ),
            (vec![("漢字", Some("ko-KR"))], false, vec!["cjk-kr"]),
            (vec![("漢字", Some("en-US"))], false, vec!["cjk-sc"]),
            (
                vec![("漢字", Some("ja-JP")), ("汉字", Some("zh-CN"))],
                false,
                vec!["cjk-jp", "cjk-sc"],
            ),
            (vec![("Latin", Some("ja-JP"))], false, vec![]),
        ] {
            assert_eq!(
                paragraph_scripts_with_languages(runs.clone(), fallback),
                expected,
                "{runs:?}"
            );
        }
    }
}
