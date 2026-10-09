use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};
use vsdx_parse::{Cell, Row, VsdxPackage};

use crate::{Lookup, ResolvedShape};

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct PageLayer {
    pub index: u32,
    pub name: String,
    pub visible: bool,
    pub print: bool,
    pub lock: bool,
    pub active: bool,
    pub color: String,
    pub status: String,
}

/// Page-sheet `Layer` rows in row-index order; empty when the page defines none.
pub fn page_layers(package: &VsdxPackage, page_part: &str) -> Vec<PageLayer> {
    let Some(page_id) = package.page_part_ids.get(page_part) else {
        return Vec::new();
    };
    let Some(sheet) = package.page_sheets.get(page_id) else {
        return Vec::new();
    };
    let Some(section) = sheet
        .sections()
        .find(|section| section.name == "Layer" && !section.del)
    else {
        return Vec::new();
    };
    let mut layers = Vec::new();
    for (ordinal, row) in section.rows().enumerate() {
        if row.del {
            continue;
        }
        let cells = row_cells(row);
        layers.push(PageLayer {
            index: row.index.unwrap_or(ordinal as u32),
            name: cell_text(&cells, "Name"),
            visible: cell_flag(&cells, "Visible", true),
            print: cell_flag(&cells, "Print", true),
            lock: cell_flag(&cells, "Lock", false),
            active: cell_flag(&cells, "Active", false),
            color: cell_text(&cells, "Color"),
            status: cell_text(&cells, "Status"),
        });
    }
    layers.sort_by_key(|layer| layer.index);
    layers
}

/// Semicolon-separated `LayerMember` indices; empty when the shape is on no layer.
pub fn shape_layer_indices(resolved: &ResolvedShape) -> Vec<u32> {
    let Some(Lookup::Found(value)) = resolved.cells.get("LayerMember") else {
        return Vec::new();
    };
    let raw = value
        .cell
        .value
        .as_deref()
        .or(value.cell.formula.as_deref())
        .unwrap_or("");
    let mut indices = raw
        .split(';')
        .filter_map(|part| part.trim().parse::<u32>().ok())
        .collect::<Vec<_>>();
    indices.sort_unstable();
    indices.dedup();
    indices
}

/// True when the shape is on at least one layer and every one of them is invisible.
pub fn shape_hidden_by_layers(resolved: &ResolvedShape, layers: &[PageLayer]) -> bool {
    let indices = shape_layer_indices(resolved);
    !indices.is_empty()
        && indices.iter().all(|index| {
            layers
                .iter()
                .find(|layer| layer.index == *index)
                .is_some_and(|layer| !layer.visible)
        })
}

fn row_cells(row: &Row) -> BTreeMap<&str, &Cell> {
    row.cells().map(|cell| (cell.name.as_str(), cell)).collect()
}

fn cell_text(cells: &BTreeMap<&str, &Cell>, name: &str) -> String {
    cells
        .get(name)
        .filter(|cell| !cell.del)
        .and_then(|cell| cell.value.as_deref().or(cell.formula.as_deref()))
        .unwrap_or_default()
        .to_owned()
}

fn cell_flag(cells: &BTreeMap<&str, &Cell>, name: &str, default: bool) -> bool {
    let Some(cell) = cells.get(name).filter(|cell| !cell.del) else {
        return default;
    };
    if let Some(value) = cell.value.as_deref() {
        return parse_flag(value, default);
    }
    if let Some(formula) = cell.formula.as_deref()
        && let Some(number) = evaluate_flag(formula)
    {
        return number != 0.0;
    }
    default
}

fn parse_flag(raw: &str, default: bool) -> bool {
    let trimmed = raw.trim();
    if trimmed.eq_ignore_ascii_case("true") {
        return true;
    }
    if trimmed.eq_ignore_ascii_case("false") {
        return false;
    }
    trimmed
        .parse::<f64>()
        .map(|value| value != 0.0)
        .unwrap_or(default)
}

fn evaluate_flag(formula: &str) -> Option<f64> {
    vsdx_formula::evaluate_number(
        formula,
        vsdx_formula::Limits {
            max_depth: 256,
            max_nodes: 1024,
            max_tokens: 2048,
        },
        &mut |_| None,
    )
}
