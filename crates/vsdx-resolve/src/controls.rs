//! Yellow control-handle positions from the Visio `Control` section.

use serde::{Deserialize, Serialize};

use crate::ResolvedShape;

/// A resolved, draggable control handle in shape-local inches.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ControlHandle {
    pub row: String,
    pub x: f64,
    pub y: f64,
    pub x_con: f64,
    pub y_con: f64,
}

impl ControlHandle {
    /// A handle is hidden when either behaviour cell selects a hidden variant.
    pub fn hidden(&self) -> bool {
        self.x_con >= 5.0 || self.y_con >= 5.0
    }

    /// A handle is axis-locked when its behaviour cell selects a locked variant.
    pub fn locked_x(&self) -> bool {
        is_locked(self.x_con)
    }

    /// A handle is axis-locked when its behaviour cell selects a locked variant.
    pub fn locked_y(&self) -> bool {
        is_locked(self.y_con)
    }
}

fn is_locked(behavior: f64) -> bool {
    behavior.is_finite() && (behavior.round() as i64).rem_euclid(5) == 1
}

/// Resolves control handles, skipping rows without a finite X/Y position.
pub fn control_handles(
    shape: &ResolvedShape,
    value: impl Fn(&str) -> Option<f64> + Copy,
) -> Vec<ControlHandle> {
    let Some(section) = shape.sections.get("Control") else {
        return Vec::new();
    };
    if section.deleted {
        return Vec::new();
    }
    let mut handles = Vec::new();
    for key in &section.row_order {
        let Some(name) = key.strip_prefix("N:") else {
            continue;
        };
        let Some(row) = section.rows.get(key) else {
            continue;
        };
        if row.deleted {
            continue;
        }
        let (Some(x), Some(y)) = (
            value(&format!("Control.{name}.X")),
            value(&format!("Control.{name}.Y")),
        ) else {
            continue;
        };
        if !x.is_finite() || !y.is_finite() {
            continue;
        }
        handles.push(ControlHandle {
            row: name.to_owned(),
            x,
            y,
            x_con: value(&format!("Control.{name}.XCon")).unwrap_or(0.0),
            y_con: value(&format!("Control.{name}.YCon")).unwrap_or(0.0),
        });
    }
    handles
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use vsdx_parse::Cell;

    use super::*;
    use crate::{Lookup, Provenance, ResolvedCell, ResolvedRow, ResolvedSection};

    fn resolved(cells: &[(&str, &str)]) -> ResolvedShape {
        let mut shape = ResolvedShape::default();
        let mut section = ResolvedSection {
            name: "Control".to_owned(),
            ..Default::default()
        };
        for (row, name, formula) in cells.iter().map(|(key, formula)| {
            let (row, name) = key.split_once('.').unwrap();
            (row, name, formula)
        }) {
            let key = format!("N:{row}");
            if !section.rows.contains_key(&key) {
                section.row_order.push(key.clone());
            }
            section
                .rows
                .entry(key.clone())
                .or_insert_with(|| ResolvedRow {
                    key,
                    ..Default::default()
                })
                .cells
                .insert(
                    name.to_owned(),
                    Lookup::Found(ResolvedCell {
                        cell: Cell {
                            name: name.to_owned(),
                            formula: Some(formula.to_string()),
                            value: None,
                            unit: None,
                            del: false,
                            other_attrs: Vec::new(),
                        },
                        provenance: Provenance::Local,
                    }),
                );
        }
        shape.sections.insert("Control".to_owned(), section);
        shape
    }

    fn values(entries: &[(&str, f64)]) -> BTreeMap<String, f64> {
        entries
            .iter()
            .map(|(name, value)| ((*name).to_owned(), *value))
            .collect()
    }

    #[test]
    fn hidden_variants_hide_while_plain_variants_show() {
        for (x_con, y_con, hidden) in [
            (0.0, 0.0, false),
            (1.0, 1.0, false),
            (4.0, 3.0, false),
            (5.0, 0.0, true),
            (0.0, 5.0, true),
            (6.0, 1.0, true),
            (9.0, 9.0, true),
        ] {
            let handle = ControlHandle {
                row: "Row_1".into(),
                x: 0.0,
                y: 0.0,
                x_con,
                y_con,
            };
            assert_eq!(handle.hidden(), hidden, "{x_con}/{y_con}");
        }
    }

    #[test]
    fn locked_variants_pin_only_their_axis() {
        let handle = ControlHandle {
            row: "Row_1".into(),
            x: 0.0,
            y: 0.0,
            x_con: 1.0,
            y_con: 0.0,
        };
        assert!(handle.locked_x());
        assert!(!handle.locked_y());
        let hidden_locked = ControlHandle {
            x_con: 6.0,
            y_con: 6.0,
            ..handle.clone()
        };
        assert!(hidden_locked.locked_x());
        assert!(hidden_locked.locked_y());
        assert!(hidden_locked.hidden());
    }

    #[test]
    fn handles_resolve_named_rows_and_skip_broken_ones() {
        let shape = resolved(&[
            ("Row_1.X", "Width*0"),
            ("Row_1.Y", "Height*0.5"),
            ("Row_1.XCon", "1"),
            ("Row_2.X", "Width*0.5"),
        ]);
        let table = values(&[
            ("Control.Row_1.X", 0.0),
            ("Control.Row_1.Y", 0.5),
            ("Control.Row_1.XCon", 1.0),
            ("Control.Row_2.X", 0.5),
        ]);
        let handles = control_handles(&shape, |name| table.get(name).copied());
        assert_eq!(
            handles,
            [ControlHandle {
                row: "Row_1".into(),
                x: 0.0,
                y: 0.5,
                x_con: 1.0,
                y_con: 0.0,
            }]
        );
    }

    #[test]
    fn plural_control_references_reach_the_singular_section() {
        let shape = resolved(&[("Row_1.X", "Width*0"), ("Row_1.Y", "Height*0.5")]);
        for reference in ["Controls.Row_1", "Controls.Row_1.X", "Controls.Row_1.Y"] {
            assert!(
                matches!(shape.cell(reference), Some(Lookup::Found(_))),
                "{reference}"
            );
        }
    }
}
