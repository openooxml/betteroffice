use std::collections::BTreeMap;

use ooxml_drawingml::Theme;
use serde::{Deserialize, Serialize};

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
    pub windows_part_path: Option<String>,
    pub relationships: BTreeMap<String, Vec<Relationship>>,
    pub document_sheet: Option<Sheet>,
    pub style_sheets: Vec<Sheet>,
    pub colors: Vec<XmlRecord>,
    pub face_names: Vec<XmlRecord>,
    pub page_sheets: BTreeMap<u32, Sheet>,
    pub master_sheets: BTreeMap<u32, Sheet>,
    /// Catalogued Page ID for each page-content part.
    pub page_part_ids: BTreeMap<String, u32>,
    /// Catalogued Master ID for each master-content part.
    pub master_part_ids: BTreeMap<String, u32>,
    pub page_contents: BTreeMap<String, Sheet>,
    pub master_contents: BTreeMap<String, Sheet>,
    #[serde(skip)]
    pub(crate) parts: Vec<PackagePart>,
}

impl VsdxPackage {
    pub fn part_bytes(&self, path: &str) -> Option<&[u8]> {
        self.parts
            .iter()
            .find(|part| part.path == path)
            .map(|part| part.bytes.as_slice())
    }

    pub fn replace_part(&mut self, path: &str, bytes: Vec<u8>) -> bool {
        let Some(part) = self.parts.iter_mut().find(|part| part.path == path) else {
            return false;
        };
        part.bytes = bytes;
        true
    }

    pub fn add_part(&mut self, path: impl Into<String>, bytes: Vec<u8>) {
        self.parts.push(PackagePart {
            path: path.into(),
            bytes,
        });
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct PackagePart {
    pub path: String,
    pub bytes: Vec<u8>,
}
