use std::collections::BTreeMap;
use std::sync::OnceLock;

use serde::Deserialize;

use crate::FontPack;

#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct BundledFontFace {
    pub pack: FontPack,
    pub file: String,
    pub family: String,
    pub weight: u16,
    pub style: String,
    pub metric_compat_with: Option<String>,
    pub script: Option<String>,
    pub byte_length: usize,
    pub version: String,
    pub codec: String,
    pub dict: Option<String>,
    pub encoded_length: usize,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WidthMatchedSubstitute {
    pub family: String,
    pub digit_advances: BTreeMap<String, f64>,
}

#[derive(Debug, Deserialize)]
pub struct LastResortFamilies {
    pub sans: String,
    pub serif: String,
    pub mono: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Catalog {
    pub faces: Vec<BundledFontFace>,
    pub word_family_aliases: BTreeMap<String, String>,
    pub unvendored_noto_serif: BTreeMap<String, String>,
    pub lighter_than_regular: Vec<String>,
    pub heavier_than_regular: Vec<String>,
    pub serif_markers: Vec<String>,
    pub serif_exclusion_pattern: String,
    pub mono_markers: Vec<String>,
    pub mono_pattern: String,
    pub width_matched_substitutes: BTreeMap<String, WidthMatchedSubstitute>,
    pub script_fallback_order: BTreeMap<String, Vec<String>>,
    pub last_resort_families: BTreeMap<String, LastResortFamilies>,
    pub glyph_fallback_families: Vec<String>,
}

pub fn catalog() -> &'static Catalog {
    static CATALOG: OnceLock<Catalog> = OnceLock::new();
    CATALOG.get_or_init(|| {
        serde_json::from_str(include_str!("../data/catalog.json"))
            .expect("valid bundled font catalog")
    })
}
