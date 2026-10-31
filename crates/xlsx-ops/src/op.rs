//! wire types for the op log: the serializable `Op` vocabulary, a `CellState`
//! mirror of the model's non-serde `Cell`, provenance tags, and transactions.

use serde::de::{Error as _, IntoDeserializer};
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

macro_rules! define_ops {
    (
        $(#[$enum_attr:meta])*
        pub enum Op {
            $(
                $(#[$variant_attr:meta])*
                $variant:ident => $tag:literal {
                    $(
                        $(#[$field_attr:meta])*
                        $field:ident: $field_type:ty,
                    )*
                },
            )*
        }
    ) => {
        $(#[$enum_attr])*
        #[derive(Debug, Clone, PartialEq, Serialize)]
        #[serde(tag = "type", rename_all = "camelCase")]
        pub enum Op {
            $(
                $(#[$variant_attr])*
                $variant {
                    $($(#[$field_attr])* $field: $field_type,)*
                },
            )*
        }

        mod wire {
            use super::*;

            #[derive(Deserialize)]
            #[serde(rename_all = "camelCase")]
            pub(super) enum Op {
                $(
                    $(#[$variant_attr])*
                    $variant {
                        $($(#[$field_attr])* $field: $field_type,)*
                    },
                )*
            }
        }

        use wire::Op as Wire;

        impl From<Wire> for Op {
            fn from(wire: Wire) -> Self {
                match wire {
                    $(Wire::$variant { $($field,)* } => Self::$variant { $($field,)* },)*
                }
            }
        }

        #[derive(Deserialize)]
        #[serde(variant_identifier, rename_all = "camelCase")]
        enum Tag {
            $($variant,)*
        }

        impl Tag {
            fn name(&self) -> &'static str {
                match self {
                    $(Self::$variant => $tag,)*
                }
            }

            fn fields(&self) -> &'static [&'static str] {
                match self {
                    $(Self::$variant => &[$(stringify!($field)),*],)*
                }
            }
        }

        #[cfg(any(test, feature = "fuzzing"))]
        mod oracle {
            use super::*;

            $(#[$enum_attr])*
            #[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
            #[serde(tag = "type", rename_all = "camelCase")]
            pub enum Op {
                $(
                    $(#[$variant_attr])*
                    $variant {
                        $($(#[$field_attr])* $field: $field_type,)*
                    },
                )*
            }
        }

        #[cfg(any(test, feature = "fuzzing"))]
        pub use oracle::Op as OracleOp;

        #[cfg(any(test, feature = "fuzzing"))]
        impl From<OracleOp> for Op {
            fn from(oracle: OracleOp) -> Self {
                match oracle {
                    $(OracleOp::$variant { $($field,)* } => Self::$variant { $($field,)* },)*
                }
            }
        }
    };
}

define_ops! {
    /// the invertible edit vocabulary. width/height are `Option` so "reset to
    /// default" is expressible and invertible.
    pub enum Op {
        SetCell => "setCell" {
            sheet: SheetId,
            at: CellRef,
            cell: CellState,
        },
        InsertRows => "insertRows" {
            sheet: SheetId,
            at: RowId,
            count: u32,
        },
        DeleteRows => "deleteRows" {
            sheet: SheetId,
            at: RowId,
            count: u32,
        },
        InsertCols => "insertCols" {
            sheet: SheetId,
            at: ColId,
            count: u32,
        },
        DeleteCols => "deleteCols" {
            sheet: SheetId,
            at: ColId,
            count: u32,
        },
        SetColWidth => "setColWidth" {
            sheet: SheetId,
            col: ColId,
            width: Option<f64>,
        },
        SetColVisibility => "setColVisibility" {
            sheet: SheetId,
            col: ColId,
            hidden: bool,
        },
        #[doc(hidden)]
        RestoreColDimension => "restoreColDimension" {
            sheet: SheetId,
            col: ColId,
            width: Option<f64>,
            hidden_width: Option<f64>,
        },
        SetRowHeight => "setRowHeight" {
            sheet: SheetId,
            row: RowId,
            height: Option<f64>,
        },
        SetRowVisibility => "setRowVisibility" {
            sheet: SheetId,
            row: RowId,
            hidden: bool,
        },
        #[doc(hidden)]
        RestoreRowDimension => "restoreRowDimension" {
            sheet: SheetId,
            row: RowId,
            height: Option<f64>,
            hidden_height: Option<f64>,
        },
        SetFreezePane => "setFreezePane" {
            sheet: SheetId,
            pane: Option<FreezePane>,
        },
        #[doc(hidden)]
        SetHyperlinks => "setHyperlinks" {
            sheet: SheetId,
            hyperlinks: Vec<Hyperlink>,
        },
        SetHyperlink => "setHyperlink" {
            sheet: SheetId,
            hyperlink: Hyperlink,
        },
        RemoveHyperlink => "removeHyperlink" {
            sheet: SheetId,
            range: CellRange,
        },
        SetAutoFilter => "setAutoFilter" {
            sheet: SheetId,
            #[serde(with = "auto_filter_range")]
            range: Option<CellRange>,
        },
        #[doc(hidden)]
        RestoreAutoFilter => "restoreAutoFilter" {
            sheet: SheetId,
            filter: Option<xlsx_model::AutoFilter>,
        },
        AddConditionalFormat => "addConditionalFormat" {
            sheet: SheetId,
            range: CellRange,
            kind: String,
            #[serde(default, skip_serializing_if = "Option::is_none")]
            operator: Option<String>,
            formulas: Vec<String>,
            format: DifferentialFormat,
        },
        RemoveConditionalFormat => "removeConditionalFormat" {
            sheet: SheetId,
            id: String,
        },
        #[doc(hidden)]
        SetConditionalFormats => "setConditionalFormats" {
            sheet: SheetId,
            formats: Vec<ConditionalFormatting>,
        },
        #[doc(hidden)]
        RestoreColStyles => "restoreColStyles" {
            sheet: SheetId,
            styles: Vec<ColStyle>,
        },
        #[doc(hidden)]
        RestoreRowStyles => "restoreRowStyles" {
            sheet: SheetId,
            #[serde(deserialize_with = "row_keyed")]
            styles: std::collections::BTreeMap<RowId, u32>,
        },
        /// Absolute replacement emitted as the inverse of a chart remap.
        #[doc(hidden)]
        SetCharts => "setCharts" {
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
        SetChartAnchor => "setChartAnchor" {
            sheet: SheetId,
            frame: String,
            part: String,
            from: ChartAnchor,
            to: ChartAnchor,
        },
        MergeCells => "mergeCells" {
            sheet: SheetId,
            range: CellRange,
        },
        UnmergeCells => "unmergeCells" {
            sheet: SheetId,
            range: CellRange,
        },
        PatchRangeStyle => "patchRangeStyle" {
            sheet: SheetId,
            range: CellRange,
            patch: StylePatch,
        },
        SetRangeNumberFormat => "setRangeNumberFormat" {
            sheet: SheetId,
            range: CellRange,
            format: NumberFormatMutation,
        },
        ApplyRangeFormat => "applyRangeFormat" {
            sheet: SheetId,
            range: CellRange,
            format: CapturedFormat,
        },
        AddSheet => "addSheet" {
            index: usize,
            name: String,
        },
        RemoveSheet => "removeSheet" {
            index: usize,
        },
        MoveSheet => "moveSheet" {
            from: usize,
            to: usize,
        },
        RenameSheet => "renameSheet" {
            sheet: SheetId,
            name: String,
        },
        SetDefinedName => "setDefinedName" {
            name: String,
            refers_to: String,
            scope: Option<SheetId>,
            hidden: Option<bool>,
        },
        DeleteDefinedName => "deleteDefinedName" {
            name: String,
            scope: Option<SheetId>,
        },
        /// Absolute replacement emitted as the inverse of structural sheet ops.
        #[doc(hidden)]
        SetDefinedNames => "setDefinedNames" {
            defined_names: Vec<DefinedName>,
        },
        #[doc(hidden)]
        RestoreSheet => "restoreSheet" {
            sheet: SheetId,
            name: String,
            formulas: Vec<(SheetId, CellRef, CellState)>,
        },
    }
}

struct Remaining<M>(M);

impl<'de, M: serde::de::MapAccess<'de>> serde::de::MapAccess<'de> for Remaining<M> {
    type Error = M::Error;

    fn next_key_seed<K: serde::de::DeserializeSeed<'de>>(
        &mut self,
        seed: K,
    ) -> Result<Option<K::Value>, Self::Error> {
        let Some(key) = self.0.next_key::<std::borrow::Cow<'de, str>>()? else {
            return Ok(None);
        };
        if key == "type" {
            return Err(M::Error::duplicate_field("type"));
        }
        seed.deserialize(key.into_deserializer()).map(Some)
    }

    fn next_value_seed<V: serde::de::DeserializeSeed<'de>>(
        &mut self,
        seed: V,
    ) -> Result<V::Value, Self::Error> {
        self.0.next_value_seed(seed)
    }
}

struct TaggedMap<M> {
    tag: &'static str,
    map: M,
}

impl<'de, M: serde::de::MapAccess<'de>> serde::de::EnumAccess<'de> for TaggedMap<M> {
    type Error = M::Error;
    type Variant = Self;

    fn variant_seed<V: serde::de::DeserializeSeed<'de>>(
        self,
        seed: V,
    ) -> Result<(V::Value, Self::Variant), Self::Error> {
        seed.deserialize(self.tag.into_deserializer())
            .map(|tag| (tag, self))
    }
}

impl<'de, M: serde::de::MapAccess<'de>> serde::de::VariantAccess<'de> for TaggedMap<M> {
    type Error = M::Error;

    fn unit_variant(self) -> Result<(), Self::Error> {
        Err(M::Error::custom("expected a struct variant"))
    }

    fn newtype_variant_seed<T: serde::de::DeserializeSeed<'de>>(
        self,
        _seed: T,
    ) -> Result<T::Value, Self::Error> {
        Err(M::Error::custom("expected a struct variant"))
    }

    fn tuple_variant<V: serde::de::Visitor<'de>>(
        self,
        _len: usize,
        _visitor: V,
    ) -> Result<V::Value, Self::Error> {
        Err(M::Error::custom("expected a struct variant"))
    }

    fn struct_variant<V: serde::de::Visitor<'de>>(
        self,
        _fields: &'static [&'static str],
        visitor: V,
    ) -> Result<V::Value, Self::Error> {
        visitor.visit_map(Remaining(self.map))
    }
}

struct TaggedDeserializer<M>(TaggedMap<M>);

impl<'de, M: serde::de::MapAccess<'de>> serde::Deserializer<'de> for TaggedDeserializer<M> {
    type Error = M::Error;

    fn deserialize_any<V: serde::de::Visitor<'de>>(
        self,
        visitor: V,
    ) -> Result<V::Value, Self::Error> {
        visitor.visit_enum(self.0)
    }

    serde::forward_to_deserialize_any! {
        bool i8 i16 i32 i64 i128 u8 u16 u32 u64 u128 f32 f64 char str string bytes
        byte_buf option unit unit_struct newtype_struct seq tuple tuple_struct map
        struct enum identifier ignored_any
    }
}

enum Captured {
    Null,
    Bool(bool),
    U64(u64),
    I64(i64),
    F64(f64),
    String(String),
    Seq(Vec<Captured>),
    Map(Vec<(String, Captured)>),
}

impl<'de> Deserialize<'de> for Captured {
    fn deserialize<D: serde::Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        struct Visitor;

        impl<'de> serde::de::Visitor<'de> for Visitor {
            type Value = Captured;

            fn expecting(&self, formatter: &mut std::fmt::Formatter) -> std::fmt::Result {
                formatter.write_str("a JSON value")
            }

            fn visit_unit<E: serde::de::Error>(self) -> Result<Self::Value, E> {
                Ok(Captured::Null)
            }

            fn visit_none<E: serde::de::Error>(self) -> Result<Self::Value, E> {
                Ok(Captured::Null)
            }

            fn visit_some<D: serde::Deserializer<'de>>(
                self,
                deserializer: D,
            ) -> Result<Self::Value, D::Error> {
                Captured::deserialize(deserializer)
            }

            fn visit_bool<E: serde::de::Error>(self, value: bool) -> Result<Self::Value, E> {
                Ok(Captured::Bool(value))
            }

            fn visit_u64<E: serde::de::Error>(self, value: u64) -> Result<Self::Value, E> {
                Ok(Captured::U64(value))
            }

            fn visit_i64<E: serde::de::Error>(self, value: i64) -> Result<Self::Value, E> {
                Ok(Captured::I64(value))
            }

            fn visit_f64<E: serde::de::Error>(self, value: f64) -> Result<Self::Value, E> {
                Ok(Captured::F64(value))
            }

            fn visit_str<E: serde::de::Error>(self, value: &str) -> Result<Self::Value, E> {
                Ok(Captured::String(value.to_owned()))
            }

            fn visit_string<E: serde::de::Error>(self, value: String) -> Result<Self::Value, E> {
                Ok(Captured::String(value))
            }

            fn visit_seq<S: serde::de::SeqAccess<'de>>(
                self,
                mut seq: S,
            ) -> Result<Self::Value, S::Error> {
                let mut values = Vec::new();
                while let Some(value) = seq.next_element::<Captured>()? {
                    values.push(value);
                }
                Ok(Captured::Seq(values))
            }

            fn visit_map<M: serde::de::MapAccess<'de>>(
                self,
                mut map: M,
            ) -> Result<Self::Value, M::Error> {
                let mut values = Vec::new();
                while let Some(entry) = map.next_entry::<String, Captured>()? {
                    values.push(entry);
                }
                Ok(Captured::Map(values))
            }
        }

        deserializer.deserialize_any(Visitor)
    }
}

impl Serialize for Captured {
    fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        match self {
            Self::Null => serializer.serialize_unit(),
            Self::Bool(value) => serializer.serialize_bool(*value),
            Self::U64(value) => serializer.serialize_u64(*value),
            Self::I64(value) => serializer.serialize_i64(*value),
            Self::F64(value) => serializer.serialize_f64(*value),
            Self::String(value) => serializer.serialize_str(value),
            Self::Seq(values) => values.serialize(serializer),
            Self::Map(values) => {
                use serde::ser::SerializeMap;

                let mut map = serializer.serialize_map(Some(values.len()))?;
                for (key, value) in values {
                    map.serialize_entry(key, value)?;
                }
                map.end()
            }
        }
    }
}

enum Body<'de> {
    Map(Vec<(std::borrow::Cow<'de, str>, Captured)>),
    Seq(Vec<Captured>),
}

struct OpVisitor;

impl<'de> serde::de::Visitor<'de> for OpVisitor {
    type Value = Op;

    fn expecting(&self, formatter: &mut std::fmt::Formatter) -> std::fmt::Result {
        formatter.write_str("internally tagged enum Op")
    }

    fn visit_map<M: serde::de::MapAccess<'de>>(self, mut map: M) -> Result<Self::Value, M::Error> {
        let first = map
            .next_key::<std::borrow::Cow<'de, str>>()?
            .ok_or_else(|| M::Error::missing_field("type"))?;
        if first == "type" {
            let tag = map.next_value::<Tag>()?;
            return Wire::deserialize(TaggedDeserializer(TaggedMap {
                tag: tag.name(),
                map,
            }))
            .map(Op::from);
        }

        let mut tag = None;
        let value = map.next_value::<Captured>()?;
        let mut fields = vec![(first, value)];
        while let Some(key) = map.next_key::<std::borrow::Cow<'de, str>>()? {
            if key == "type" {
                if tag.is_some() {
                    return Err(M::Error::duplicate_field("type"));
                }
                tag = Some(map.next_value::<Tag>()?);
            } else {
                let value = map.next_value::<Captured>()?;
                fields.push((key, value));
            }
        }
        let tag = tag.ok_or_else(|| M::Error::missing_field("type"))?;
        decode_body(tag, Body::Map(fields))
    }

    fn visit_seq<S: serde::de::SeqAccess<'de>>(self, mut seq: S) -> Result<Self::Value, S::Error> {
        let tag = seq
            .next_element::<Tag>()?
            .ok_or_else(|| S::Error::missing_field("type"))?;
        let mut fields = Vec::new();
        while let Some(value) = seq.next_element::<Captured>()? {
            fields.push(value);
        }
        decode_body(tag, Body::Seq(fields))
    }
}

impl<'de> Deserialize<'de> for Op {
    fn deserialize<D: serde::Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        deserializer.deserialize_any(OpVisitor)
    }
}

fn decode_body<E: serde::de::Error>(tag: Tag, body: Body<'_>) -> Result<Op, E> {
    let mut text = format!("{{\"type\":\"{}\"", tag.name());
    match body {
        Body::Map(fields) => {
            for (key, value) in fields {
                text.push(',');
                text.push_str(&serde_json::to_string(&key).map_err(E::custom)?);
                text.push(':');
                text.push_str(&serde_json::to_string(&value).map_err(E::custom)?);
            }
        }
        Body::Seq(fields) => {
            let names = tag.fields();
            if fields.len() != names.len() {
                return Err(E::custom(format!(
                    "invalid length {}, expected {} fields in sequence",
                    fields.len(),
                    names.len()
                )));
            }
            for (key, value) in names.iter().zip(fields) {
                text.push(',');
                text.push_str(&serde_json::to_string(key).map_err(E::custom)?);
                text.push(':');
                text.push_str(&serde_json::to_string(&value).map_err(E::custom)?);
            }
        }
    }
    text.push('}');
    let mut deserializer = serde_json::Deserializer::from_str(&text);
    serde::Deserializer::deserialize_any(&mut deserializer, OpVisitor)
        .map_err(|error| E::custom(wire_error_message(&error)))
}

fn wire_error_message(error: &serde_json::Error) -> String {
    let message = error.to_string();
    let suffix = format!(" at line {} column {}", error.line(), error.column());
    message.strip_suffix(&suffix).unwrap_or(&message).to_owned()
}

#[cfg(test)]
#[path = "op_decode_tests.rs"]
mod op_decode_tests;

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
