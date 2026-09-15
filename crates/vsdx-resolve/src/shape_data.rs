use serde::{Deserialize, Serialize};

use crate::{Lookup, ResolvedShape};

pub const PROPERTY_SECTION: &str = "Property";

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum ShapeDataType {
    String,
    FixedList,
    Number,
    Boolean,
    VariableList,
    Date,
    Duration,
    Currency,
}

impl ShapeDataType {
    pub fn from_type_value(value: &str) -> Option<Self> {
        match value.trim() {
            "0" => Some(Self::String),
            "1" => Some(Self::FixedList),
            "2" => Some(Self::Number),
            "3" => Some(Self::Boolean),
            "4" => Some(Self::VariableList),
            "5" => Some(Self::Date),
            "6" => Some(Self::Duration),
            "7" => Some(Self::Currency),
            _ => None,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct ShapeDataValue {
    pub formula: Option<String>,
    pub value: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct ShapeDataProperty {
    pub row: String,
    pub label: String,
    pub prompt: Option<String>,
    pub data_type: ShapeDataType,
    pub format: Option<String>,
    pub value: ShapeDataValue,
    pub sort_key: Option<String>,
    pub invisible: bool,
    pub ask: bool,
}

/// Typed Shape Data rows in `SortKey` order, document order when it is absent.
/// Invisible rows are included with their flag set: hidden values still feed
/// `Field` formulas, so filtering is the caller's decision.
pub fn shape_data(shape: &ResolvedShape) -> Vec<ShapeDataProperty> {
    let mut sections: Vec<_> = shape
        .sections
        .values()
        .filter(|section| section.name == PROPERTY_SECTION && !section.deleted)
        .collect();
    sections.sort_by(|left, right| {
        (left.index.unwrap_or(0), left.name.clone())
            .cmp(&(right.index.unwrap_or(0), right.name.clone()))
    });
    let mut rows = Vec::new();
    for section in sections {
        for key in &section.row_order {
            let Some(row) = section.rows.get(key) else {
                continue;
            };
            if row.deleted {
                continue;
            }
            let name = key
                .strip_prefix("N:")
                .unwrap_or(key.as_str())
                .split('\u{1f}')
                .next()
                .unwrap_or(key.as_str());
            let invisible = cell_truthy(row, "Invisible");
            let label = cell_text(row, "Label").unwrap_or_else(|| name.to_owned());
            rows.push(ShapeDataProperty {
                row: name.to_owned(),
                label,
                prompt: cell_text(row, "Prompt"),
                data_type: cell_text(row, "Type")
                    .as_deref()
                    .and_then(ShapeDataType::from_type_value)
                    .unwrap_or(ShapeDataType::String),
                format: cell_text(row, "Format"),
                value: ShapeDataValue {
                    formula: cell_formula(row, "Value"),
                    value: cell_value(row, "Value"),
                },
                sort_key: cell_text(row, "SortKey"),
                invisible,
                ask: cell_truthy(row, "Ask") || cell_truthy(row, "Verify"),
            });
        }
    }
    rows.sort_by(|left, right| sort_key(left).cmp(sort_key(right)));
    rows
}

fn sort_key(property: &ShapeDataProperty) -> &str {
    property.sort_key.as_deref().unwrap_or("")
}

fn lookup<'a>(row: &'a crate::ResolvedRow, name: &str) -> Option<&'a crate::ResolvedCell> {
    match row.cells.get(name) {
        Some(Lookup::Found(cell)) => Some(cell),
        _ => None,
    }
}

fn cell_formula(row: &crate::ResolvedRow, name: &str) -> Option<String> {
    lookup(row, name)?.cell.formula.clone()
}

fn cell_value(row: &crate::ResolvedRow, name: &str) -> Option<String> {
    lookup(row, name)?.cell.value.clone()
}

fn cell_text(row: &crate::ResolvedRow, name: &str) -> Option<String> {
    let cell = lookup(row, name)?;
    cell.cell
        .value
        .clone()
        .filter(|value| !value.is_empty())
        .or_else(|| cell.cell.formula.as_deref().and_then(unquote_literal))
}

fn cell_truthy(row: &crate::ResolvedRow, name: &str) -> bool {
    let Some(cell) = lookup(row, name) else {
        return false;
    };
    if let Some(formula) = cell.cell.formula.as_deref() {
        let normalized = formula.trim().trim_start_matches('=').trim();
        if normalized.eq_ignore_ascii_case("true") {
            return true;
        }
        if normalized.eq_ignore_ascii_case("false") {
            return false;
        }
        if let Some(number) = vsdx_formula::evaluate_number(
            normalized,
            vsdx_formula::Limits {
                max_depth: 256,
                max_nodes: 8192,
                max_tokens: 16384,
            },
            &mut |_| None,
        ) {
            return number != 0.0;
        }
    }
    cell.cell.value.as_deref().is_some_and(truthy)
}

fn truthy(value: &str) -> bool {
    let value = value.trim();
    value.eq_ignore_ascii_case("true") || value.parse::<f64>().is_ok_and(|number| number != 0.0)
}

fn unquote_literal(formula: &str) -> Option<String> {
    let literal = formula.trim().trim_start_matches('=').trim();
    literal
        .strip_prefix('"')
        .and_then(|inner| inner.strip_suffix('"'))
        .map(|inner| inner.replace("\"\"", "\""))
}

#[cfg(test)]
mod tests {
    use vsdx_parse::{
        Cell, Row, RowChild, Section, SectionChild, Shape, ShapeChild, ShapesChild, Sheet,
        SheetChild,
    };

    use crate::{Lookup, Provenance, ResolvedCell, ResolvedRow, ResolvedShape, Resolver};

    use super::*;

    fn cell(name: &str, value: &str) -> Cell {
        Cell {
            name: name.into(),
            formula: None,
            value: Some(value.into()),
            unit: None,
            del: false,
            other_attrs: Vec::new(),
        }
    }

    fn formula_cell(name: &str, formula: &str, value: Option<&str>) -> Cell {
        Cell {
            name: name.into(),
            formula: Some(formula.into()),
            value: value.map(str::to_owned),
            unit: None,
            del: false,
            other_attrs: Vec::new(),
        }
    }

    fn property_row(name: &str, cells: Vec<Cell>) -> Row {
        Row {
            index: None,
            name: Some(name.into()),
            local_name: None,
            row_type: None,
            del: false,
            children: cells.into_iter().map(RowChild::Cell).collect(),
            other_attrs: Vec::new(),
        }
    }

    fn resolved_property(rows: Vec<(String, Vec<Cell>)>) -> ResolvedShape {
        let mut shape = ResolvedShape::default();
        let mut section = crate::ResolvedSection {
            name: PROPERTY_SECTION.into(),
            ..Default::default()
        };
        for (name, cells) in rows {
            let key = format!("N:{name}");
            section.row_order.push(key.clone());
            let mut resolved_row = ResolvedRow {
                key: key.clone(),
                ..Default::default()
            };
            for cell in cells {
                resolved_row.cells.insert(
                    cell.name.clone(),
                    Lookup::Found(ResolvedCell {
                        cell,
                        provenance: Provenance::Local,
                    }),
                );
            }
            section.rows.insert(key, resolved_row);
        }
        shape.sections.insert(PROPERTY_SECTION.into(), section);
        shape
    }

    fn full_row(
        label: &str,
        value: &str,
        data_type: &str,
        sort_key: Option<&str>,
        invisible: &str,
    ) -> Vec<Cell> {
        let mut cells = vec![
            cell("Label", label),
            cell("Value", value),
            cell("Type", data_type),
            cell("Invisible", invisible),
        ];
        if let Some(sort_key) = sort_key {
            cells.push(cell("SortKey", sort_key));
        }
        cells
    }

    #[test]
    fn maps_every_type_value() {
        for (value, expected) in [
            ("0", ShapeDataType::String),
            ("1", ShapeDataType::FixedList),
            ("2", ShapeDataType::Number),
            ("3", ShapeDataType::Boolean),
            ("4", ShapeDataType::VariableList),
            ("5", ShapeDataType::Date),
            ("6", ShapeDataType::Duration),
            ("7", ShapeDataType::Currency),
        ] {
            let shape = resolved_property(vec![(
                "Row".into(),
                full_row("Label", "v", value, None, "0"),
            )]);
            assert_eq!(shape_data(&shape)[0].data_type, expected, "{value}");
        }
        for value in ["", "9", "text"] {
            let shape = resolved_property(vec![(
                "Row".into(),
                full_row("Label", "v", value, None, "0"),
            )]);
            assert_eq!(
                shape_data(&shape)[0].data_type,
                ShapeDataType::String,
                "{value}"
            );
        }
    }

    #[test]
    fn sorts_by_sort_key_and_falls_back_to_document_order() {
        let shape = resolved_property(vec![
            ("Zebra".into(), full_row("Zebra", "1", "0", Some("b"), "0")),
            ("Alpha".into(), full_row("Alpha", "2", "0", Some("a"), "0")),
            (
                "Middle".into(),
                full_row("Middle", "3", "0", Some("b"), "0"),
            ),
        ]);
        let rows: Vec<_> = shape_data(&shape)
            .iter()
            .map(|property| property.row.clone())
            .collect();
        assert_eq!(rows, ["Alpha", "Zebra", "Middle"]);
        let shape = resolved_property(vec![
            ("Zebra".into(), full_row("Zebra", "1", "0", None, "0")),
            ("Alpha".into(), full_row("Alpha", "2", "0", None, "0")),
        ]);
        let rows: Vec<_> = shape_data(&shape)
            .iter()
            .map(|property| property.row.clone())
            .collect();
        assert_eq!(rows, ["Zebra", "Alpha"]);
    }

    #[test]
    fn hides_invisible_rows_and_labels_fall_back_to_row_names() {
        let shape = resolved_property(vec![
            ("Shown".into(), full_row("Shown label", "1", "0", None, "0")),
            (
                "Hidden".into(),
                full_row("Hidden label", "2", "0", None, "1"),
            ),
            ("Unlabelled".into(), vec![cell("Value", "3")]),
        ]);
        let properties = shape_data(&shape);
        assert_eq!(properties.len(), 3);
        assert!(!properties[0].invisible);
        assert!(properties[1].invisible);
        assert_eq!(properties[2].label, "Unlabelled");
        assert_eq!(properties[2].data_type, ShapeDataType::String);
    }

    #[test]
    fn evaluates_literal_invisible_formulas_and_falls_back_to_cached_values() {
        let shape = resolved_property(vec![
            (
                "Literal".into(),
                vec![
                    cell("Label", "Literal"),
                    formula_cell("Invisible", "TRUE", None),
                ],
            ),
            (
                "Numeric".into(),
                vec![
                    cell("Label", "Numeric"),
                    formula_cell("Invisible", "1", None),
                ],
            ),
            (
                "Reference".into(),
                vec![
                    cell("Label", "Reference"),
                    formula_cell("Invisible", "NOT(Prop.Other)", Some("1")),
                ],
            ),
            (
                "False".into(),
                vec![cell("Label", "False"), formula_cell("Invisible", "0", None)],
            ),
        ]);
        let properties = shape_data(&shape);
        assert!(properties[0].invisible);
        assert!(properties[1].invisible);
        assert!(properties[2].invisible);
        assert!(!properties[3].invisible);
    }

    #[test]
    fn resolves_inherited_master_properties() {
        let mut package: vsdx_parse::VsdxPackage = serde_json::from_value(serde_json::json!({
            "documentPartPath": "", "pagesPartPath": null, "mastersPartPath": null,
            "pagePartPaths": [], "masterPartPaths": [], "themePartPaths": [], "windowsPartPath": null,
            "relationships": {}, "documentSheet": null, "styleSheets": [], "colors": [], "faceNames": [],
            "pageSheets": {}, "masterSheets": {}, "pagePartIds": {}, "masterPartIds": {},
            "pageContents": {}, "masterContents": {}
        }))
        .unwrap();
        let master = Shape {
            id: 1,
            name: None,
            name_u: None,
            shape_type: None,
            master: None,
            master_shape: None,
            line_style: None,
            fill_style: None,
            text_style: None,
            del: false,
            other_attrs: Vec::new(),
            children: vec![ShapeChild::Section(Section {
                name: PROPERTY_SECTION.into(),
                index: None,
                del: false,
                other_attrs: Vec::new(),
                children: vec![SectionChild::Row(property_row(
                    "Device",
                    full_row("Device name", "Amplifier", "0", None, "0"),
                ))],
            })],
        };
        package.master_contents.insert(
            "master".into(),
            Sheet {
                id: None,
                other_attrs: Vec::new(),
                children: vec![SheetChild::Shapes(vec![ShapesChild::Shape(master)])],
            },
        );
        package.master_part_ids.insert("master".into(), 7);
        let instance = Shape {
            id: 3,
            name: None,
            name_u: None,
            shape_type: None,
            master: Some(7),
            master_shape: None,
            line_style: None,
            fill_style: None,
            text_style: None,
            del: false,
            other_attrs: Vec::new(),
            children: Vec::new(),
        };
        package.page_contents.insert(
            "page".into(),
            Sheet {
                id: None,
                other_attrs: Vec::new(),
                children: vec![SheetChild::Shapes(vec![ShapesChild::Shape(instance)])],
            },
        );
        let resolved = Resolver::new(&package).resolve_shape("page", 3).unwrap();
        let properties = shape_data(&resolved);
        assert_eq!(properties.len(), 1);
        assert_eq!(properties[0].row, "Device");
        assert_eq!(properties[0].label, "Device name");
        assert_eq!(properties[0].value.value.as_deref(), Some("Amplifier"));
        assert!(properties.iter().all(|property| !property.invisible));
    }

    #[test]
    fn skips_deleted_rows() {
        let mut shape =
            resolved_property(vec![("Gone".into(), full_row("Gone", "1", "0", None, "0"))]);
        let section = shape.sections.get_mut(PROPERTY_SECTION).unwrap();
        section.rows.get_mut("N:Gone").unwrap().deleted = true;
        assert!(shape_data(&shape).is_empty());
    }
}
