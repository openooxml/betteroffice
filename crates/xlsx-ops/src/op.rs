//! wire types for the op log: the serializable `Op` vocabulary, a `CellState`
//! mirror of the model's non-serde `Cell`, provenance tags, and transactions.

use serde::{Deserialize, Serialize};
use xlsx_model::{
    Cell, CellRange, CellRef, CellValue, ChartAnchor, ColId, ColStyle, ConditionalFormatting,
    DefinedName, DifferentialFormat, FreezePane, Hyperlink, RowId, SheetChart, SheetId,
};

use crate::formatting::{CapturedFormat, NumberFormatMutation, StylePatch};

/// serializable mirror of `xlsx_model::Cell`, which deliberately has no serde.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellState {
    #[serde(default)]
    pub value: CellValue,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<u32>,
}

impl From<&Cell> for CellState {
    fn from(c: &Cell) -> Self {
        Self {
            value: c.value.clone(),
            formula: c.formula.clone(),
            style: c.style,
        }
    }
}

impl From<Cell> for CellState {
    fn from(c: Cell) -> Self {
        Self {
            value: c.value,
            formula: c.formula,
            style: c.style,
        }
    }
}

impl From<CellState> for Cell {
    fn from(s: CellState) -> Self {
        Cell {
            value: s.value,
            formula: s.formula,
            style: s.style,
        }
    }
}

/// the invertible edit vocabulary. width/height are `Option` so "reset to
/// default" is expressible and invertible.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum Op {
    SetCell {
        sheet: SheetId,
        at: CellRef,
        cell: CellState,
    },
    InsertRows {
        sheet: SheetId,
        at: RowId,
        count: u32,
    },
    DeleteRows {
        sheet: SheetId,
        at: RowId,
        count: u32,
    },
    InsertCols {
        sheet: SheetId,
        at: ColId,
        count: u32,
    },
    DeleteCols {
        sheet: SheetId,
        at: ColId,
        count: u32,
    },
    SetColWidth {
        sheet: SheetId,
        col: ColId,
        width: Option<f64>,
    },
    SetColVisibility {
        sheet: SheetId,
        col: ColId,
        hidden: bool,
    },
    #[doc(hidden)]
    RestoreColDimension {
        sheet: SheetId,
        col: ColId,
        width: Option<f64>,
        hidden_width: Option<f64>,
    },
    SetRowHeight {
        sheet: SheetId,
        row: RowId,
        height: Option<f64>,
    },
    SetRowVisibility {
        sheet: SheetId,
        row: RowId,
        hidden: bool,
    },
    #[doc(hidden)]
    RestoreRowDimension {
        sheet: SheetId,
        row: RowId,
        height: Option<f64>,
        hidden_height: Option<f64>,
    },
    SetFreezePane {
        sheet: SheetId,
        pane: Option<FreezePane>,
    },
    #[doc(hidden)]
    SetHyperlinks {
        sheet: SheetId,
        hyperlinks: Vec<Hyperlink>,
    },
    SetHyperlink {
        sheet: SheetId,
        hyperlink: Hyperlink,
    },
    RemoveHyperlink {
        sheet: SheetId,
        range: CellRange,
    },
    SetAutoFilter {
        sheet: SheetId,
        #[serde(with = "auto_filter_range")]
        range: Option<CellRange>,
    },
    #[doc(hidden)]
    RestoreAutoFilter {
        sheet: SheetId,
        filter: Option<xlsx_model::AutoFilter>,
    },
    AddConditionalFormat {
        sheet: SheetId,
        range: CellRange,
        kind: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        operator: Option<String>,
        formulas: Vec<String>,
        format: DifferentialFormat,
    },
    RemoveConditionalFormat {
        sheet: SheetId,
        id: String,
    },
    #[doc(hidden)]
    SetConditionalFormats {
        sheet: SheetId,
        formats: Vec<ConditionalFormatting>,
    },
    #[doc(hidden)]
    RestoreColStyles {
        sheet: SheetId,
        styles: Vec<ColStyle>,
    },
    #[doc(hidden)]
    RestoreRowStyles {
        sheet: SheetId,
        #[serde(deserialize_with = "row_keyed")]
        styles: std::collections::BTreeMap<RowId, u32>,
    },
    /// Absolute replacement emitted as the inverse of a chart remap.
    #[doc(hidden)]
    SetCharts {
        sheet: SheetId,
        charts: Vec<SheetChart>,
    },
    /// Repin one chart frame, addressed by its `SheetChart::frame_id`. A frame
    /// id names an ordinal in a drawing, and a drawing another editor has since
    /// added to, reordered or thinned renumbers those ordinals, so the id alone
    /// cannot survive being stored. `part` and `from` are the chart part and
    /// the anchor the frame held when the op was recorded, and a replay that
    /// finds either changed is refused instead of moving whoever now sits
    /// there.
    SetChartAnchor {
        sheet: SheetId,
        frame: String,
        part: String,
        from: ChartAnchor,
        to: ChartAnchor,
    },
    MergeCells {
        sheet: SheetId,
        range: CellRange,
    },
    UnmergeCells {
        sheet: SheetId,
        range: CellRange,
    },
    PatchRangeStyle {
        sheet: SheetId,
        range: CellRange,
        patch: StylePatch,
    },
    SetRangeNumberFormat {
        sheet: SheetId,
        range: CellRange,
        format: NumberFormatMutation,
    },
    ApplyRangeFormat {
        sheet: SheetId,
        range: CellRange,
        format: CapturedFormat,
    },
    AddSheet {
        index: usize,
        name: String,
    },
    RemoveSheet {
        index: usize,
    },
    MoveSheet {
        from: usize,
        to: usize,
    },
    RenameSheet {
        sheet: SheetId,
        name: String,
    },
    SetDefinedName {
        name: String,
        refers_to: String,
        scope: Option<SheetId>,
        hidden: Option<bool>,
    },
    DeleteDefinedName {
        name: String,
        scope: Option<SheetId>,
    },
    /// Absolute replacement emitted as the inverse of structural sheet ops.
    #[doc(hidden)]
    SetDefinedNames {
        defined_names: Vec<DefinedName>,
    },
    #[doc(hidden)]
    RestoreSheet {
        sheet: SheetId,
        name: String,
        formulas: Vec<(SheetId, CellRef, CellState)>,
    },
}

/// who authored a transaction.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum Provenance {
    User,
    Agent { id: String },
    Remote { actor: String },
}

/// an atomic, provenance-tagged batch of ops. `proposed = true` marks a HITL
/// proposal awaiting accept or reject; it is not yet applied.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Transaction {
    pub ops: Vec<Op>,
    pub author: Provenance,
    #[serde(default)]
    pub proposed: bool,
}

impl Transaction {
    pub fn new(ops: Vec<Op>, author: Provenance) -> Self {
        Self {
            ops,
            author,
            proposed: false,
        }
    }

    /// a proposal awaiting review — same payload, not yet committed.
    pub fn proposal(ops: Vec<Op>, author: Provenance) -> Self {
        Self {
            ops,
            author,
            proposed: true,
        }
    }
}

/// Integer map keys arrive as strings once a tagged enum has buffered them.
fn row_keyed<'de, D: serde::Deserializer<'de>>(
    deserializer: D,
) -> Result<std::collections::BTreeMap<RowId, u32>, D::Error> {
    #[derive(PartialEq, Eq, PartialOrd, Ord)]
    struct Row(RowId);
    impl<'de> Deserialize<'de> for Row {
        fn deserialize<D: serde::Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
            struct Visitor;
            impl serde::de::Visitor<'_> for Visitor {
                type Value = Row;
                fn expecting(&self, formatter: &mut std::fmt::Formatter) -> std::fmt::Result {
                    formatter.write_str("a row index")
                }
                fn visit_u64<E: serde::de::Error>(self, value: u64) -> Result<Row, E> {
                    RowId::try_from(value).map(Row).map_err(E::custom)
                }
                fn visit_str<E: serde::de::Error>(self, value: &str) -> Result<Row, E> {
                    value.parse().map(Row).map_err(E::custom)
                }
            }
            deserializer.deserialize_any(Visitor)
        }
    }
    let rows = std::collections::BTreeMap::<Row, u32>::deserialize(deserializer)?;
    Ok(rows
        .into_iter()
        .map(|(row, style)| (row.0, style))
        .collect())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn row_style_restore_round_trips_through_json() {
        let op = Op::RestoreRowStyles {
            sheet: SheetId(0),
            styles: std::collections::BTreeMap::from([(3, 1), (70_000, 2)]),
        };
        let json = serde_json::to_string(&vec![op.clone()]).unwrap();
        assert_eq!(serde_json::from_str::<Vec<Op>>(&json).unwrap(), vec![op]);
    }

    #[test]
    fn cell_state_round_trips_through_model_cell() {
        let cell = Cell {
            value: CellValue::Number { value: 3.5 },
            formula: Some("1+2.5".into()),
            style: Some(7),
        };
        let state = CellState::from(&cell);
        assert_eq!(Cell::from(state), cell);
    }
}

mod auto_filter_range {
    use serde::{Deserialize, Deserializer, Serializer};
    use xlsx_model::CellRange;

    pub fn serialize<S: Serializer>(
        range: &Option<CellRange>,
        serializer: S,
    ) -> Result<S::Ok, S::Error> {
        serializer.serialize_some(&range.map(|range| range.to_a1()))
    }

    pub fn deserialize<'de, D: Deserializer<'de>>(
        deserializer: D,
    ) -> Result<Option<CellRange>, D::Error> {
        Option::<String>::deserialize(deserializer)?
            .map(|range| CellRange::parse_a1(&range).map_err(serde::de::Error::custom))
            .transpose()
    }
}
