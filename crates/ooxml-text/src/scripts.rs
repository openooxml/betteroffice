/// Script coverage buckets for paragraph text, preserving context across runs.
pub fn paragraph_scripts_with_han_fallback<'a>(
    texts: impl IntoIterator<Item = &'a str>,
    include_han_fallback: bool,
) -> Vec<String> {
    let mut han = false;
    let mut kana = false;
    let mut hangul = false;
    let mut arabic = false;
    let mut hebrew = false;
    let mut intl = Vec::new();
    for text in texts {
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
                0x3000..=0x303f
                | 0x3400..=0x4dbf
                | 0x4e00..=0x9fff
                | 0xf900..=0xfaff
                | 0xfe30..=0xfe4f
                | 0xff00..=0xff65
                | 0x20000..=0x3ffff => han = true,
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
    }
    let mut scripts = Vec::new();
    if kana {
        scripts.push("cjk-jp".to_owned());
    }
    if hangul {
        scripts.push("cjk-kr".to_owned());
    }
    if han && (include_han_fallback || (!kana && !hangul)) {
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
    use super::paragraph_scripts_with_han_fallback;

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
}
