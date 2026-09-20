use std::collections::BTreeMap;

use ooxml_drawingml::Theme;
use serde::{Deserialize, Serialize};

use crate::patch::ElementSpan;
use crate::{Relationship, Sheet, XmlRecord};

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct VsdxPackage {
    pub document_part_path: String,
    pub pages_part_path: Option<String>,
    pub masters_part_path: Option<String>,
    pub page_part_paths: Vec<String>,
    pub master_part_paths: Vec<String>,
    pub theme_part_paths: Vec<String>,
    #[serde(default)]
    pub themes: BTreeMap<u32, Theme>,
    #[serde(default)]
    pub theme_effects: BTreeMap<u32, ThemeEffects>,
    pub windows_part_path: Option<String>,
    pub relationships: BTreeMap<String, Vec<Relationship>>,
    pub document_sheet: Option<Sheet>,
    pub style_sheets: Vec<Sheet>,
    pub colors: Vec<XmlRecord>,
    pub face_names: Vec<XmlRecord>,
    pub page_sheets: BTreeMap<u32, Sheet>,
    pub master_sheets: BTreeMap<u32, Sheet>,
    pub page_part_ids: BTreeMap<String, u32>,
    #[serde(default)]
    pub page_names: BTreeMap<u32, String>,
    pub master_part_ids: BTreeMap<String, u32>,
    #[serde(default)]
    pub master_names: BTreeMap<u32, String>,
    pub page_contents: BTreeMap<String, Sheet>,
    pub master_contents: BTreeMap<String, Sheet>,
    #[serde(skip)]
    pub(crate) parts: Vec<PackagePart>,
    /// Source container bytes, kept so unchanged members can be re-emitted
    /// with their original compressed payload on save.
    #[serde(skip)]
    pub(crate) source_container: Option<ooxml_opc::SourceContainer>,
}

impl VsdxPackage {
    pub fn part_bytes(&self, path: &str) -> Option<&[u8]> {
        self.parts
            .iter()
            .find(|part| part.path == path)
            .map(|part| part.bytes.as_slice())
    }

    /// Replaces bytes and invalidates lexical provenance.
    pub fn replace_part(&mut self, path: &str, bytes: Vec<u8>) -> bool {
        let Some(part) = self.parts.iter_mut().find(|part| part.path == path) else {
            return false;
        };
        part.bytes = bytes;
        part.spans.clear();
        true
    }

    pub fn add_part(&mut self, path: impl Into<String>, bytes: Vec<u8>) {
        self.parts.push(PackagePart {
            path: path.into(),
            bytes,
            spans: Vec::new(),
        });
    }

    pub fn element_spans(&self, path: &str) -> Option<&[ElementSpan]> {
        self.parts
            .iter()
            .find(|part| part.path == path)
            .map(|part| part.spans.as_slice())
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct PackagePart {
    pub path: String,
    pub bytes: Vec<u8>,
    pub spans: Vec<ElementSpan>,
}

/// Theme shadow and variant data behind `V='Themed'` effect cells.
#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeEffects {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scheme_id: Option<u32>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub effect_styles: Vec<ThemeEffectStyle>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub variation_schemes: Vec<ThemeVariationScheme>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub variation_colors: Vec<Vec<Option<String>>>,
}

/// One `a:effectStyle` entry; index-selected by QuickStyleEffectsMatrix.
#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeEffectStyle {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outer_shadow: Option<ThemeOuterShadow>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub has_bevel: bool,
}

/// An `a:outerShdw` in EMUs and 60000ths of a degree.
#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeOuterShadow {
    #[serde(default)]
    pub blur_emu: i64,
    #[serde(default)]
    pub dist_emu: i64,
    #[serde(default)]
    pub direction_60k: i64,
    #[serde(default)]
    pub color: ThemeEffectColor,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alpha_1000pct: Option<i64>,
}

/// An `a:outerShdw` colour reference.
#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ThemeEffectColor {
    Srgb(String),
    Scheme(String),
    #[default]
    Placeholder,
}

/// One `vt:variationStyleScheme`: effect style per variant position.
#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeVariationScheme {
    #[serde(default)]
    pub embellishment: u32,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub effect_indexes: Vec<u32>,
}

fn is_false(value: &bool) -> bool {
    !*value
}
