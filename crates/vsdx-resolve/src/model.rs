use std::collections::BTreeMap;

use ooxml_drawingml::GeometryPathCommand;
use serde::{Deserialize, Serialize};
use thiserror::Error;
use vsdx_parse::Cell;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum Provenance {
    Local,
    MasterShape,
    Master,
    StyleLine,
    StyleFill,
    StyleText,
    Page,
    Document,
    Default,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct ResolvedCell {
    pub cell: Cell,
    pub provenance: Provenance,
}

pub type ResolvedValue = ResolvedCell;

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum Lookup {
    Found(ResolvedValue),
    Deleted,
    Absent,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct ResolvedRow {
    pub key: String,
    pub deleted: bool,
    pub row_type: Option<String>,
    pub cells: BTreeMap<String, Lookup>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct ResolvedSection {
    pub name: String,
    pub deleted: bool,
    pub rows: BTreeMap<String, ResolvedRow>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct ResolvedShape {
    pub deleted: bool,
    pub cells: BTreeMap<String, Lookup>,
    pub sections: BTreeMap<String, ResolvedSection>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum ResolvedTextToken {
    Literal(String),
    CharacterRun {
        index: u32,
        properties: BTreeMap<String, Lookup>,
    },
    ParagraphRun {
        index: u32,
        properties: BTreeMap<String, Lookup>,
    },
    Tab {
        index: u32,
        properties: BTreeMap<String, Lookup>,
    },
    Field {
        index: u32,
        properties: BTreeMap<String, Lookup>,
    },
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub enum GeometryIssue {
    UnsupportedRowType(String),
    UnevaluatedCell { row_type: String, cell: String },
    MissingCell { row_type: String, cell: String },
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct RealizedGeometry {
    pub commands: Vec<GeometryPathCommand>,
    pub issues: Vec<GeometryIssue>,
}

#[derive(Debug, Error, PartialEq, Eq)]
pub enum ResolveError {
    #[error("page content not found: {0}")]
    MissingPage(String),
    #[error("shape not found: {0}")]
    MissingShape(u32),
    #[error("inheritance cycle: {0}")]
    Cycle(String),
}
