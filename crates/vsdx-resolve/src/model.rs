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
    pub row_order: Vec<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct ResolvedShape {
    pub deleted: bool,
    pub cells: BTreeMap<String, Lookup>,
    pub sections: BTreeMap<String, ResolvedSection>,
}

impl ResolvedShape {
    pub fn cell(&self, name: &str) -> Option<&Lookup> {
        self.cells.get(name).or_else(|| {
            let (section, reference) = name.split_once('.')?;
            let section = self
                .sections
                .get(section)
                .or_else(|| {
                    self.sections
                        .get(section.trim_end_matches(|c: char| c.is_ascii_digit()))
                })
                .or_else(|| {
                    section
                        .strip_suffix('s')
                        .and_then(|name| self.sections.get(name))
                })?;
            let (row, cell) = reference.split_once('.').map_or_else(
                || match section.rows.get(&format!("N:{reference}")) {
                    Some(row) => (
                        Some(row),
                        if section.name == "User" { "Value" } else { "X" },
                    ),
                    None if !reference.chars().any(|c| c.is_ascii_digit()) => {
                        (section.rows.get("IX:0"), reference)
                    }
                    None => {
                        let split = reference
                            .char_indices()
                            .rev()
                            .take_while(|(_, c)| c.is_ascii_digit())
                            .last()
                            .map_or(reference.len(), |(index, _)| index);
                        let (cell, index) = reference.split_at(split);
                        let row = index
                            .parse::<u32>()
                            .ok()
                            .and_then(|index| {
                                if section.name == "Scratch" {
                                    index.checked_sub(1)
                                } else {
                                    Some(index)
                                }
                            })
                            .and_then(|index| section.rows.get(&format!("IX:{index}")));
                        (row, cell)
                    }
                },
                |(row, cell)| (section.rows.get(&format!("N:{row}")), cell),
            );
            row.and_then(|row| row.cells.get(cell))
        })
    }

    /// Returns the shape's explicit theme selection, when it has one.
    pub fn theme_index(&self) -> Option<u32> {
        self.index_cell("ThemeIndex")
    }

    /// Returns the shape's explicit colour-scheme selection, when it has one.
    pub fn color_scheme_index(&self) -> Option<u32> {
        self.index_cell("ColorSchemeIndex")
    }

    fn index_cell(&self, name: &str) -> Option<u32> {
        match self.cells.get(name) {
            Some(Lookup::Found(value)) => value.cell.value.as_deref()?.parse().ok(),
            _ => None,
        }
    }
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
    #[error("master not found: {0}")]
    MissingMaster(u32),
    #[error("inheritance cycle: {0}")]
    Cycle(String),
}
