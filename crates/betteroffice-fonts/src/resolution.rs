use serde::Deserialize;

use crate::catalog::{BundledFontFace, catalog};

#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum LastResortOffice {
    #[default]
    PowerPoint,
    Word,
}

impl LastResortOffice {
    fn key(self) -> &'static str {
        match self {
            Self::PowerPoint => "powerpoint",
            Self::Word => "word",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct LastResortChoice {
    pub face: &'static BundledFontFace,
    pub only_if_primary_fails: bool,
}

fn whitespace(ch: char) -> bool {
    (ch.is_whitespace() && ch != '\u{85}') || ch == '\u{feff}'
}

fn key(family: &str) -> String {
    family.trim_matches(whitespace).to_lowercase()
}

fn matches_style(face: &BundledFontFace, bold: bool, italic: bool) -> bool {
    face.weight == if bold { 700 } else { 400 }
        && face.style == if italic { "italic" } else { "normal" }
}

pub fn resolve_metric_compat_family(word_family: &str) -> Option<&'static str> {
    let key = key(word_family);
    let key = catalog()
        .word_family_aliases
        .get(&key)
        .map_or(key.as_str(), String::as_str);
    catalog().faces.iter().rev().find_map(|face| {
        face.metric_compat_with
            .as_ref()
            .filter(|name| name.to_lowercase() == key)
            .map(|_| face.family.as_str())
    })
}

pub fn resolve_metric_compat_face(
    word_family: &str,
    bold: bool,
    italic: bool,
) -> Option<&'static BundledFontFace> {
    let family = resolve_metric_compat_family(word_family)?;
    if family == "Selawik" {
        if italic {
            return None;
        }
        let key = key(word_family);
        let target = if bold && key == "segoe ui semibold" {
            "segoe ui"
        } else {
            &key
        };
        let weight = if bold && target == "segoe ui" {
            700
        } else {
            400
        };
        return catalog().faces.iter().find(|face| {
            face.metric_compat_with
                .as_ref()
                .is_some_and(|name| name.to_lowercase() == target)
                && face.weight == weight
        });
    }
    family_face(family, bold, italic).or_else(|| family_face(family, false, false))
}

fn family_face(family: &str, bold: bool, italic: bool) -> Option<&'static BundledFontFace> {
    catalog()
        .faces
        .iter()
        .find(|face| face.family == family && matches_style(face, bold, italic))
}

pub fn resolve_bundled_family_face(
    family: &str,
    bold: bool,
    italic: bool,
) -> Option<&'static BundledFontFace> {
    if resolve_metric_compat_family(family) == Some("Selawik") {
        return if italic {
            None
        } else {
            resolve_metric_compat_face(family, bold, false)
        };
    }
    let key = key(family);
    let bundled = catalog()
        .faces
        .iter()
        .rev()
        .find(|face| face.family.to_lowercase() == key)
        .map(|face| face.family.as_str())
        .or_else(|| {
            catalog()
                .unvendored_noto_serif
                .get(&key)
                .map(String::as_str)
        })
        .or_else(|| resolve_metric_compat_family(family))?;
    family_face(bundled, bold, italic)
}

pub fn resolve_script_fallback_face(
    script: &str,
    bold: bool,
    italic: bool,
) -> Option<&'static BundledFontFace> {
    let files = catalog().script_fallback_order.get(script)?;
    let faces: Vec<_> = files
        .iter()
        .filter_map(|file| catalog().faces.iter().find(|face| &face.file == file))
        .collect();
    faces
        .iter()
        .copied()
        .find(|face| matches_style(face, bold, italic))
        .or_else(|| {
            faces
                .iter()
                .copied()
                .find(|face| matches_style(face, bold, false))
        })
        .or_else(|| {
            faces
                .iter()
                .copied()
                .find(|face| matches_style(face, false, false))
        })
        .or_else(|| faces.first().copied())
}

fn variant(family: &str, italic: bool, heavy: bool) -> Option<&'static BundledFontFace> {
    let family = family.trim_matches(whitespace);
    if family.ends_with('-') {
        return None;
    }
    let mut words: Vec<_> = family
        .split(|ch| whitespace(ch) || ch == '-')
        .filter(|word| !word.is_empty())
        .collect();
    if words.len() < 2 {
        return None;
    }
    let suffix = words.pop()?.to_lowercase();
    let variants = if heavy {
        &catalog().heavier_than_regular
    } else {
        &catalog().lighter_than_regular
    };
    if !variants.contains(&suffix) {
        return None;
    }
    resolve_metric_compat_face(&words.join(" "), heavy, italic)
}

pub fn resolve_light_variant_face(
    family: &str,
    _bold: bool,
    italic: bool,
) -> Option<&'static BundledFontFace> {
    variant(family, italic, false)
}

fn looks_serif(family: &str) -> bool {
    let lower = family.to_lowercase();
    let pattern = &catalog().serif_exclusion_pattern;
    let exclusions = [
        pattern.replace("[ -]?", ""),
        pattern.replace("[ -]?", " "),
        pattern.replace("[ -]?", "-"),
    ];
    !exclusions.iter().any(|pattern| lower.contains(pattern))
        && catalog()
            .serif_markers
            .iter()
            .any(|marker| lower.contains(marker))
}

fn looks_mono(family: &str) -> bool {
    let lower = family.to_lowercase();
    let word = catalog()
        .mono_pattern
        .strip_prefix(r"(^|[\s-])")
        .and_then(|pattern| pattern.strip_suffix(r"([\s-]|$)"))
        .expect("bundled monospace pattern");
    catalog()
        .mono_markers
        .iter()
        .any(|marker| lower.contains(marker))
        || lower
            .split(|ch| whitespace(ch) || ch == '-')
            .any(|candidate| candidate == word)
}

fn resolve_base_face(
    family: &str,
    mut bold: bool,
    italic: bool,
    office: LastResortOffice,
    exclude_glyph_fallback: bool,
) -> &'static BundledFontFace {
    let light = variant(family, italic, false);
    let allowed = |face: &&BundledFontFace| {
        !exclude_glyph_fallback || !catalog().glyph_fallback_families.contains(&face.family)
    };
    if let Some(face) = light.filter(allowed) {
        return face;
    }
    if light.is_some() {
        bold = false;
    }
    if let Some(face) = variant(family, italic, true).filter(allowed) {
        return face;
    }
    let families = &catalog().last_resort_families[office.key()];
    if looks_mono(family) {
        return family_face(&families.mono, bold, italic).expect("bundled monospace face");
    }
    if let Some(matched) = catalog().width_matched_substitutes.get(&key(family))
        && let Some(face) = resolve_metric_compat_face(&matched.family, bold, italic)
    {
        return face;
    }
    let base = if looks_serif(family) {
        &families.serif
    } else {
        &families.sans
    };
    family_face(base, bold, italic).expect("bundled base face")
}

pub fn resolve_last_resort_choice(
    family: &str,
    bold: bool,
    italic: bool,
    office: LastResortOffice,
) -> LastResortChoice {
    let primary = resolve_metric_compat_face(family, bold, italic);
    let base = resolve_base_face(family, bold, italic, office, primary.is_some());
    if primary.is_some_and(|primary| primary.file == base.file) {
        let alternate = if base.family == catalog().last_resort_families["word"].sans {
            &catalog().last_resort_families["powerpoint"].sans
        } else {
            &catalog().last_resort_families["word"].sans
        };
        LastResortChoice {
            face: family_face(alternate, bold, italic).expect("bundled alternate base face"),
            only_if_primary_fails: true,
        }
    } else {
        LastResortChoice {
            face: base,
            only_if_primary_fails: false,
        }
    }
}

pub fn resolve_last_resort_face(
    family: &str,
    bold: bool,
    italic: bool,
    office: LastResortOffice,
) -> &'static BundledFontFace {
    resolve_last_resort_choice(family, bold, italic, office).face
}

pub fn resolve_glyph_fallback_face(
    family: &str,
    bold: bool,
    italic: bool,
    office: LastResortOffice,
) -> Option<&'static BundledFontFace> {
    let primary = resolve_metric_compat_face(family, bold, italic);
    let face = primary.unwrap_or_else(|| resolve_base_face(family, bold, italic, office, false));
    if !catalog().glyph_fallback_families.contains(&face.family) {
        return None;
    }
    Some(if primary.is_some() {
        resolve_last_resort_face(family, bold, italic, office)
    } else {
        resolve_base_face(family, bold, italic, office, true)
    })
}

pub fn resolve_digit_advance(family: &str, bold: bool, italic: bool) -> Option<f64> {
    let style = match (bold, italic) {
        (false, false) => "regular",
        (true, false) => "bold",
        (false, true) => "italic",
        (true, true) => "boldItalic",
    };
    catalog()
        .width_matched_substitutes
        .get(&key(family))?
        .digit_advances
        .get(style)
        .copied()
}
