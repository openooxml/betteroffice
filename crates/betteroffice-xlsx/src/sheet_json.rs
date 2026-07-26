//! bounded readers for the per-sheet JSON a peer can put in the shared
//! document. every cap is charged while the payload streams, so an oversized
//! graph is refused instead of being materialized first and rejected after.

use std::fmt;

use serde::Deserialize;
use serde::de::{
    DeserializeSeed, Deserializer, Error as DeError, IgnoredAny, MapAccess, SeqAccess, Visitor,
};
use xlsx_model::{
    AnchorCell, AnchorEditAs, AnchorExtent, AnchorPos, CellRange, ChartAnchor, ChartRef,
    ChartRefKind, Hyperlink, MAX_COLS, MAX_ROWS, SheetChart,
};

pub(crate) const MAX_CHARTS_PER_SHEET: usize = 4_096;
/// upper bound on the anchor index a chart may claim, matching the parser's
/// per-drawing anchor cap.
pub(crate) const MAX_CHART_ANCHORS_PER_DRAWING: usize = 4_096;
pub(crate) const MAX_CHART_REFS_PER_CHART: usize = 16_384;
pub(crate) const MAX_CHART_FIELD_BYTES: usize = 32_767;
/// upper bound on one sheet's encoded chart state: [`MAX_CHARTS_PER_SHEET`] at
/// a realistic 2 KiB per chart. a charted workbook needs three orders of
/// magnitude less.
pub(crate) const MAX_CHART_STATE_BYTES: usize = 8 * 1024 * 1024;
pub(crate) const MAX_HYPERLINKS_PER_SHEET: usize = 65_536;
pub(crate) const MAX_HYPERLINK_FIELD_BYTES: usize = 32_767;
/// upper bound on one sheet's encoded hyperlink state, sized the same way:
/// [`MAX_HYPERLINKS_PER_SHEET`] at a realistic 256 bytes per link.
pub(crate) const MAX_HYPERLINK_STATE_BYTES: usize = 16 * 1024 * 1024;

const TOO_MANY_CHARTS: &str = "sheet has too many charts";
const CHART_UNNAMED: &str = "sheet chart does not name its part and drawing";
const CHART_TOO_LARGE: &str = "sheet chart exceeds its size limit";
const CHART_REF_TOO_LONG: &str = "sheet chart reference exceeds its length limit";
const TOO_MANY_HYPERLINKS: &str = "sheet has too many hyperlinks";
const HYPERLINK_OUT_OF_BOUNDS: &str = "sheet hyperlink range is out of bounds";
const HYPERLINK_NO_DESTINATION: &str = "sheet hyperlink has no destination";
const HYPERLINK_FIELD_TOO_LONG: &str = "sheet hyperlink field exceeds its length limit";

/// read a sheet's charts, refusing the payload the moment it outgrows a cap.
pub(crate) fn decode_charts(json: &str) -> Result<Vec<SheetChart>, String> {
    if json.len() > MAX_CHART_STATE_BYTES {
        return Err("sheet chart state exceeds its size limit".to_string());
    }
    let mut deserializer = serde_json::Deserializer::from_str(json);
    ChartsSeed
        .deserialize(&mut deserializer)
        .and_then(|charts| deserializer.end().map(|()| charts))
        .map_err(|error| format!("sheet charts are invalid: {error}"))
}

/// read a sheet's hyperlinks under the same discipline as [`decode_charts`].
pub(crate) fn decode_hyperlinks(json: &str) -> Result<Vec<Hyperlink>, String> {
    if json.len() > MAX_HYPERLINK_STATE_BYTES {
        return Err("sheet hyperlink state exceeds its size limit".to_string());
    }
    let mut deserializer = serde_json::Deserializer::from_str(json);
    HyperlinksSeed
        .deserialize(&mut deserializer)
        .and_then(|hyperlinks| deserializer.end().map(|()| hyperlinks))
        .map_err(|error| format!("sheet hyperlinks are invalid: {error}"))
}

struct BoundedString {
    limit: usize,
    message: &'static str,
}

impl<'de> DeserializeSeed<'de> for BoundedString {
    type Value = String;

    fn deserialize<D: Deserializer<'de>>(self, deserializer: D) -> Result<Self::Value, D::Error> {
        deserializer.deserialize_str(self)
    }
}

impl Visitor<'_> for BoundedString {
    type Value = String;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("a string")
    }

    fn visit_str<E: DeError>(self, value: &str) -> Result<Self::Value, E> {
        if value.len() > self.limit {
            return Err(E::custom(self.message));
        }
        Ok(value.to_owned())
    }

    fn visit_string<E: DeError>(self, value: String) -> Result<Self::Value, E> {
        if value.len() > self.limit {
            return Err(E::custom(self.message));
        }
        Ok(value)
    }
}

struct BoundedOptionString {
    limit: usize,
    message: &'static str,
}

impl<'de> DeserializeSeed<'de> for BoundedOptionString {
    type Value = Option<String>;

    fn deserialize<D: Deserializer<'de>>(self, deserializer: D) -> Result<Self::Value, D::Error> {
        deserializer.deserialize_option(self)
    }
}

impl<'de> Visitor<'de> for BoundedOptionString {
    type Value = Option<String>;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("an optional string")
    }

    fn visit_none<E: DeError>(self) -> Result<Self::Value, E> {
        Ok(None)
    }

    fn visit_unit<E: DeError>(self) -> Result<Self::Value, E> {
        Ok(None)
    }

    fn visit_some<D: Deserializer<'de>>(self, deserializer: D) -> Result<Self::Value, D::Error> {
        BoundedString {
            limit: self.limit,
            message: self.message,
        }
        .deserialize(deserializer)
        .map(Some)
    }
}

macro_rules! keyword_seed {
    ($seed:ident, $value:ty, $expecting:literal, $unknown:literal, $($name:literal => $variant:expr),+ $(,)?) => {
        struct $seed;

        impl<'de> DeserializeSeed<'de> for $seed {
            type Value = $value;

            fn deserialize<D: Deserializer<'de>>(
                self,
                deserializer: D,
            ) -> Result<Self::Value, D::Error> {
                deserializer.deserialize_str(self)
            }
        }

        impl Visitor<'_> for $seed {
            type Value = $value;

            fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
                formatter.write_str($expecting)
            }

            fn visit_str<E: DeError>(self, value: &str) -> Result<Self::Value, E> {
                match value {
                    $($name => Ok($variant),)+
                    _ => Err(E::custom($unknown)),
                }
            }
        }
    };
}

keyword_seed!(
    ChartRefKindSeed,
    ChartRefKind,
    "a chart reference kind",
    "unknown chart reference kind",
    "seriesName" => ChartRefKind::SeriesName,
    "categories" => ChartRefKind::Categories,
    "values" => ChartRefKind::Values,
    "bubbleSize" => ChartRefKind::BubbleSize,
    "title" => ChartRefKind::Title,
    "dataLabels" => ChartRefKind::DataLabels,
    "other" => ChartRefKind::Other,
);

keyword_seed!(
    AnchorEditAsSeed,
    AnchorEditAs,
    "a chart anchor edit mode",
    "unknown chart anchor edit mode",
    "twoCell" => AnchorEditAs::TwoCell,
    "oneCell" => AnchorEditAs::OneCell,
    "absolute" => AnchorEditAs::Absolute,
);

enum AnchorKind {
    TwoCell,
    OneCell,
    Absolute,
}

keyword_seed!(
    AnchorKindSeed,
    AnchorKind,
    "a chart anchor kind",
    "unknown chart anchor kind",
    "twoCell" => AnchorKind::TwoCell,
    "oneCell" => AnchorKind::OneCell,
    "absolute" => AnchorKind::Absolute,
);

macro_rules! field_key {
    ($name:ident, $visitor:ident, $expecting:literal, $($key:literal => $variant:ident),+ $(,)?) => {
        enum $name {
            $($variant,)+
            Other,
        }

        impl<'de> Deserialize<'de> for $name {
            fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
                deserializer.deserialize_identifier($visitor)
            }
        }

        struct $visitor;

        impl Visitor<'_> for $visitor {
            type Value = $name;

            fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
                formatter.write_str($expecting)
            }

            fn visit_str<E: DeError>(self, value: &str) -> Result<Self::Value, E> {
                Ok(match value {
                    $($key => $name::$variant,)+
                    _ => $name::Other,
                })
            }
        }
    };
}

field_key!(
    ChartField,
    ChartFieldVisitor,
    "a sheet chart field",
    "part" => Part,
    "drawing" => Drawing,
    "anchorIndex" => AnchorIndex,
    "anchor" => Anchor,
    "refs" => Refs,
);

field_key!(
    AnchorField,
    AnchorFieldVisitor,
    "a chart anchor field",
    "kind" => Kind,
    "from" => From,
    "to" => To,
    "edit_as" => EditAs,
    "extent" => Extent,
    "pos" => Pos,
);

field_key!(
    ChartRefField,
    ChartRefFieldVisitor,
    "a chart reference field",
    "kind" => Kind,
    "formula" => Formula,
);

field_key!(
    HyperlinkField,
    HyperlinkFieldVisitor,
    "a sheet hyperlink field",
    "range" => Range,
    "external_target" => ExternalTarget,
    "location" => Location,
    "tooltip" => Tooltip,
    "display" => Display,
);

struct ChartsSeed;

impl<'de> DeserializeSeed<'de> for ChartsSeed {
    type Value = Vec<SheetChart>;

    fn deserialize<D: Deserializer<'de>>(self, deserializer: D) -> Result<Self::Value, D::Error> {
        deserializer.deserialize_seq(self)
    }
}

impl<'de> Visitor<'de> for ChartsSeed {
    type Value = Vec<SheetChart>;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("an array of sheet charts")
    }

    fn visit_seq<A: SeqAccess<'de>>(self, mut seq: A) -> Result<Self::Value, A::Error> {
        let mut charts = Vec::new();
        while let Some(chart) = seq.next_element_seed(ChartSeed { seen: charts.len() })? {
            charts.push(chart);
        }
        Ok(charts)
    }
}

struct ChartSeed {
    seen: usize,
}

impl<'de> DeserializeSeed<'de> for ChartSeed {
    type Value = SheetChart;

    fn deserialize<D: Deserializer<'de>>(self, deserializer: D) -> Result<Self::Value, D::Error> {
        if self.seen >= MAX_CHARTS_PER_SHEET {
            return Err(D::Error::custom(TOO_MANY_CHARTS));
        }
        deserializer.deserialize_map(self)
    }
}

impl<'de> Visitor<'de> for ChartSeed {
    type Value = SheetChart;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("a sheet chart")
    }

    fn visit_map<A: MapAccess<'de>>(self, mut map: A) -> Result<Self::Value, A::Error> {
        let mut part = None;
        let mut drawing = None;
        let mut anchor_index = None;
        let mut anchor = None;
        let mut refs = None;
        while let Some(field) = map.next_key::<ChartField>()? {
            match field {
                ChartField::Part => {
                    if part.is_some() {
                        return Err(A::Error::duplicate_field("part"));
                    }
                    part = Some(map.next_value_seed(BoundedString {
                        limit: MAX_CHART_FIELD_BYTES,
                        message: CHART_TOO_LARGE,
                    })?);
                }
                ChartField::Drawing => {
                    if drawing.is_some() {
                        return Err(A::Error::duplicate_field("drawing"));
                    }
                    drawing = Some(map.next_value_seed(BoundedString {
                        limit: MAX_CHART_FIELD_BYTES,
                        message: CHART_TOO_LARGE,
                    })?);
                }
                ChartField::AnchorIndex => {
                    if anchor_index.is_some() {
                        return Err(A::Error::duplicate_field("anchorIndex"));
                    }
                    anchor_index = Some(map.next_value::<usize>()?);
                }
                ChartField::Anchor => {
                    if anchor.is_some() {
                        return Err(A::Error::duplicate_field("anchor"));
                    }
                    anchor = Some(map.next_value_seed(AnchorSeed)?);
                }
                ChartField::Refs => {
                    if refs.is_some() {
                        return Err(A::Error::duplicate_field("refs"));
                    }
                    refs = Some(map.next_value_seed(ChartRefsSeed)?);
                }
                ChartField::Other => {
                    map.next_value::<IgnoredAny>()?;
                }
            }
        }
        let part: String = part.ok_or_else(|| A::Error::missing_field("part"))?;
        let drawing: String = drawing.ok_or_else(|| A::Error::missing_field("drawing"))?;
        if part.is_empty() || drawing.is_empty() {
            return Err(A::Error::custom(CHART_UNNAMED));
        }
        Ok(SheetChart {
            part,
            drawing,
            anchor_index: anchor_index.ok_or_else(|| A::Error::missing_field("anchorIndex"))?,
            anchor: anchor.ok_or_else(|| A::Error::missing_field("anchor"))?,
            refs: refs.ok_or_else(|| A::Error::missing_field("refs"))?,
        })
    }
}

struct AnchorSeed;

impl<'de> DeserializeSeed<'de> for AnchorSeed {
    type Value = ChartAnchor;

    fn deserialize<D: Deserializer<'de>>(self, deserializer: D) -> Result<Self::Value, D::Error> {
        deserializer.deserialize_map(self)
    }
}

impl<'de> Visitor<'de> for AnchorSeed {
    type Value = ChartAnchor;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("a chart anchor")
    }

    fn visit_map<A: MapAccess<'de>>(self, mut map: A) -> Result<Self::Value, A::Error> {
        let mut kind = None;
        let mut from = None;
        let mut to = None;
        let mut edit_as = None;
        let mut extent = None;
        let mut pos = None;
        while let Some(field) = map.next_key::<AnchorField>()? {
            match field {
                AnchorField::Kind => {
                    if kind.is_some() {
                        return Err(A::Error::duplicate_field("kind"));
                    }
                    kind = Some(map.next_value_seed(AnchorKindSeed)?);
                }
                AnchorField::From => {
                    if from.is_some() {
                        return Err(A::Error::duplicate_field("from"));
                    }
                    from = Some(map.next_value::<AnchorCell>()?);
                }
                AnchorField::To => {
                    if to.is_some() {
                        return Err(A::Error::duplicate_field("to"));
                    }
                    to = Some(map.next_value::<AnchorCell>()?);
                }
                AnchorField::EditAs => {
                    if edit_as.is_some() {
                        return Err(A::Error::duplicate_field("edit_as"));
                    }
                    edit_as = Some(map.next_value_seed(AnchorEditAsSeed)?);
                }
                AnchorField::Extent => {
                    if extent.is_some() {
                        return Err(A::Error::duplicate_field("extent"));
                    }
                    extent = Some(map.next_value::<AnchorExtent>()?);
                }
                AnchorField::Pos => {
                    if pos.is_some() {
                        return Err(A::Error::duplicate_field("pos"));
                    }
                    pos = Some(map.next_value::<AnchorPos>()?);
                }
                AnchorField::Other => {
                    map.next_value::<IgnoredAny>()?;
                }
            }
        }
        match kind.ok_or_else(|| A::Error::missing_field("kind"))? {
            AnchorKind::TwoCell => Ok(ChartAnchor::TwoCell {
                from: from.ok_or_else(|| A::Error::missing_field("from"))?,
                to: to.ok_or_else(|| A::Error::missing_field("to"))?,
                edit_as: edit_as.ok_or_else(|| A::Error::missing_field("edit_as"))?,
            }),
            AnchorKind::OneCell => Ok(ChartAnchor::OneCell {
                from: from.ok_or_else(|| A::Error::missing_field("from"))?,
                extent: extent.ok_or_else(|| A::Error::missing_field("extent"))?,
            }),
            AnchorKind::Absolute => Ok(ChartAnchor::Absolute {
                pos: pos.ok_or_else(|| A::Error::missing_field("pos"))?,
                extent: extent.ok_or_else(|| A::Error::missing_field("extent"))?,
            }),
        }
    }
}

struct ChartRefsSeed;

impl<'de> DeserializeSeed<'de> for ChartRefsSeed {
    type Value = Vec<ChartRef>;

    fn deserialize<D: Deserializer<'de>>(self, deserializer: D) -> Result<Self::Value, D::Error> {
        deserializer.deserialize_seq(self)
    }
}

impl<'de> Visitor<'de> for ChartRefsSeed {
    type Value = Vec<ChartRef>;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("an array of chart references")
    }

    fn visit_seq<A: SeqAccess<'de>>(self, mut seq: A) -> Result<Self::Value, A::Error> {
        let mut refs = Vec::new();
        while let Some(reference) = seq.next_element_seed(ChartRefSeed { seen: refs.len() })? {
            refs.push(reference);
        }
        Ok(refs)
    }
}

struct ChartRefSeed {
    seen: usize,
}

impl<'de> DeserializeSeed<'de> for ChartRefSeed {
    type Value = ChartRef;

    fn deserialize<D: Deserializer<'de>>(self, deserializer: D) -> Result<Self::Value, D::Error> {
        if self.seen >= MAX_CHART_REFS_PER_CHART {
            return Err(D::Error::custom(CHART_TOO_LARGE));
        }
        deserializer.deserialize_map(self)
    }
}

impl<'de> Visitor<'de> for ChartRefSeed {
    type Value = ChartRef;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("a chart reference")
    }

    fn visit_map<A: MapAccess<'de>>(self, mut map: A) -> Result<Self::Value, A::Error> {
        let mut kind = None;
        let mut formula = None;
        while let Some(field) = map.next_key::<ChartRefField>()? {
            match field {
                ChartRefField::Kind => {
                    if kind.is_some() {
                        return Err(A::Error::duplicate_field("kind"));
                    }
                    kind = Some(map.next_value_seed(ChartRefKindSeed)?);
                }
                ChartRefField::Formula => {
                    if formula.is_some() {
                        return Err(A::Error::duplicate_field("formula"));
                    }
                    formula = Some(map.next_value_seed(BoundedString {
                        limit: MAX_CHART_FIELD_BYTES,
                        message: CHART_REF_TOO_LONG,
                    })?);
                }
                ChartRefField::Other => {
                    map.next_value::<IgnoredAny>()?;
                }
            }
        }
        Ok(ChartRef {
            kind: kind.ok_or_else(|| A::Error::missing_field("kind"))?,
            formula: formula.ok_or_else(|| A::Error::missing_field("formula"))?,
        })
    }
}

struct HyperlinksSeed;

impl<'de> DeserializeSeed<'de> for HyperlinksSeed {
    type Value = Vec<Hyperlink>;

    fn deserialize<D: Deserializer<'de>>(self, deserializer: D) -> Result<Self::Value, D::Error> {
        deserializer.deserialize_seq(self)
    }
}

impl<'de> Visitor<'de> for HyperlinksSeed {
    type Value = Vec<Hyperlink>;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("an array of sheet hyperlinks")
    }

    fn visit_seq<A: SeqAccess<'de>>(self, mut seq: A) -> Result<Self::Value, A::Error> {
        let mut hyperlinks = Vec::new();
        while let Some(hyperlink) = seq.next_element_seed(HyperlinkSeed {
            seen: hyperlinks.len(),
        })? {
            hyperlinks.push(hyperlink);
        }
        Ok(hyperlinks)
    }
}

struct HyperlinkSeed {
    seen: usize,
}

impl<'de> DeserializeSeed<'de> for HyperlinkSeed {
    type Value = Hyperlink;

    fn deserialize<D: Deserializer<'de>>(self, deserializer: D) -> Result<Self::Value, D::Error> {
        if self.seen >= MAX_HYPERLINKS_PER_SHEET {
            return Err(D::Error::custom(TOO_MANY_HYPERLINKS));
        }
        deserializer.deserialize_map(self)
    }
}

impl<'de> Visitor<'de> for HyperlinkSeed {
    type Value = Hyperlink;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("a sheet hyperlink")
    }

    fn visit_map<A: MapAccess<'de>>(self, mut map: A) -> Result<Self::Value, A::Error> {
        let mut range = None;
        let mut external_target = None;
        let mut location = None;
        let mut tooltip = None;
        let mut display = None;
        while let Some(field) = map.next_key::<HyperlinkField>()? {
            let (slot, name) = match field {
                HyperlinkField::Range => {
                    if range.is_some() {
                        return Err(A::Error::duplicate_field("range"));
                    }
                    range = Some(map.next_value::<CellRange>()?);
                    continue;
                }
                HyperlinkField::ExternalTarget => (&mut external_target, "external_target"),
                HyperlinkField::Location => (&mut location, "location"),
                HyperlinkField::Tooltip => (&mut tooltip, "tooltip"),
                HyperlinkField::Display => (&mut display, "display"),
                HyperlinkField::Other => {
                    map.next_value::<IgnoredAny>()?;
                    continue;
                }
            };
            if slot.is_some() {
                return Err(A::Error::duplicate_field(name));
            }
            *slot = Some(map.next_value_seed(BoundedOptionString {
                limit: MAX_HYPERLINK_FIELD_BYTES,
                message: HYPERLINK_FIELD_TOO_LONG,
            })?);
        }
        let hyperlink = Hyperlink {
            range: range.ok_or_else(|| A::Error::missing_field("range"))?,
            external_target: external_target.flatten(),
            location: location.flatten(),
            tooltip: tooltip.flatten(),
            display: display.flatten(),
        };
        let range = hyperlink.range;
        if range.start.row > range.end.row
            || range.start.col > range.end.col
            || range.end.row >= MAX_ROWS
            || range.end.col >= MAX_COLS
        {
            return Err(A::Error::custom(HYPERLINK_OUT_OF_BOUNDS));
        }
        if hyperlink
            .external_target
            .as_deref()
            .is_none_or(str::is_empty)
            && hyperlink.location.as_deref().is_none_or(str::is_empty)
        {
            return Err(A::Error::custom(HYPERLINK_NO_DESTINATION));
        }
        Ok(hyperlink)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::cell::Cell;
    use xlsx_model::CellRef;

    thread_local! {
        static LIVE: Cell<usize> = const { Cell::new(0) };
        static PEAK: Cell<usize> = const { Cell::new(0) };
    }

    /// `System`, plus a per-thread high-water mark so a test can assert what a
    /// payload actually costs before it is refused.
    struct PeakAlloc;

    unsafe impl GlobalAlloc for PeakAlloc {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            let _ = LIVE.try_with(|live| {
                let value = live.get().saturating_add(layout.size());
                live.set(value);
                let _ = PEAK.try_with(|peak| peak.set(peak.get().max(value)));
            });
            unsafe { System.alloc(layout) }
        }

        unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
            let _ = LIVE.try_with(|live| live.set(live.get().saturating_sub(layout.size())));
            unsafe { System.dealloc(ptr, layout) }
        }
    }

    #[global_allocator]
    static ALLOCATOR: PeakAlloc = PeakAlloc;

    fn peak_bytes<T>(body: impl FnOnce() -> T) -> usize {
        LIVE.with(|live| live.set(0));
        PEAK.with(|peak| peak.set(0));
        let value = body();
        let peak = PEAK.with(Cell::get);
        drop(value);
        peak
    }

    fn hostile_charts(count: usize) -> String {
        const ELEMENT: &str = r#"{"part":"a","drawing":"b","anchorIndex":0,"anchor":{"kind":"absolute","pos":{"x":0,"y":0},"extent":{"cx":0,"cy":0}},"refs":[]}"#;
        let mut json = String::with_capacity(count * (ELEMENT.len() + 1) + 2);
        json.push('[');
        for index in 0..count {
            if index > 0 {
                json.push(',');
            }
            json.push_str(ELEMENT);
        }
        json.push(']');
        json
    }

    fn chart(refs: usize) -> SheetChart {
        SheetChart {
            part: "xl/charts/chart1.xml".to_string(),
            drawing: "xl/drawings/drawing1.xml".to_string(),
            anchor_index: 0,
            anchor: ChartAnchor::TwoCell {
                from: AnchorCell {
                    col: 1,
                    col_off: 2,
                    row: 3,
                    row_off: 4,
                },
                to: AnchorCell {
                    col: 9,
                    col_off: 8,
                    row: 7,
                    row_off: 6,
                },
                edit_as: AnchorEditAs::OneCell,
            },
            refs: (0..refs)
                .map(|index| ChartRef {
                    kind: ChartRefKind::Values,
                    formula: format!("Sheet1!$B${index}"),
                })
                .collect(),
        }
    }

    fn hyperlink() -> Hyperlink {
        Hyperlink {
            range: CellRange::new(CellRef::new(0, 0), CellRef::new(3, 3)),
            external_target: Some("https://example.com".to_string()),
            location: None,
            tooltip: Some("go".to_string()),
            display: None,
        }
    }

    #[test]
    fn every_anchor_and_reference_kind_round_trips_through_the_bounded_reader() {
        let anchors = [
            ChartAnchor::TwoCell {
                from: AnchorCell::default(),
                to: AnchorCell::default(),
                edit_as: AnchorEditAs::TwoCell,
            },
            ChartAnchor::TwoCell {
                from: AnchorCell::default(),
                to: AnchorCell::default(),
                edit_as: AnchorEditAs::OneCell,
            },
            ChartAnchor::TwoCell {
                from: AnchorCell::default(),
                to: AnchorCell::default(),
                edit_as: AnchorEditAs::Absolute,
            },
            ChartAnchor::OneCell {
                from: AnchorCell::default(),
                extent: AnchorExtent { cx: 5, cy: 6 },
            },
            ChartAnchor::Absolute {
                pos: AnchorPos { x: 7, y: 8 },
                extent: AnchorExtent { cx: 9, cy: 10 },
            },
        ];
        let kinds = [
            ChartRefKind::SeriesName,
            ChartRefKind::Categories,
            ChartRefKind::Values,
            ChartRefKind::BubbleSize,
            ChartRefKind::Title,
            ChartRefKind::DataLabels,
            ChartRefKind::Other,
        ];
        let charts = anchors
            .into_iter()
            .enumerate()
            .map(|(index, anchor)| SheetChart {
                anchor,
                anchor_index: index,
                refs: kinds
                    .iter()
                    .map(|&kind| ChartRef {
                        kind,
                        formula: "Sheet1!$A$1".to_string(),
                    })
                    .collect(),
                ..chart(0)
            })
            .collect::<Vec<_>>();
        let json = serde_json::to_string(&charts).unwrap();
        assert_eq!(decode_charts(&json).unwrap(), charts);
        assert_eq!(decode_charts("[]").unwrap(), Vec::<SheetChart>::new());
    }

    #[test]
    fn hyperlinks_round_trip_through_the_bounded_reader() {
        let hyperlinks = vec![
            hyperlink(),
            Hyperlink {
                external_target: None,
                location: Some("Sheet2!A1".to_string()),
                display: Some("jump".to_string()),
                ..hyperlink()
            },
        ];
        let json = serde_json::to_string(&hyperlinks).unwrap();
        assert_eq!(decode_hyperlinks(&json).unwrap(), hyperlinks);
    }

    #[test]
    fn oversized_collections_are_refused_before_the_graph_is_built() {
        let charts = vec![chart(0); MAX_CHARTS_PER_SHEET + 1];
        let json = serde_json::to_string(&charts).unwrap();
        assert!(json.len() < MAX_CHART_STATE_BYTES);
        assert!(decode_charts(&json).unwrap_err().contains(TOO_MANY_CHARTS));

        let refs = vec![chart(MAX_CHART_REFS_PER_CHART + 1)];
        let json = serde_json::to_string(&refs).unwrap();
        assert!(json.len() < MAX_CHART_STATE_BYTES);
        assert!(decode_charts(&json).unwrap_err().contains(CHART_TOO_LARGE));

        let hyperlinks = vec![hyperlink(); MAX_HYPERLINKS_PER_SHEET + 1];
        let json = serde_json::to_string(&hyperlinks).unwrap();
        assert!(json.len() < MAX_HYPERLINK_STATE_BYTES);
        assert!(
            decode_hyperlinks(&json)
                .unwrap_err()
                .contains(TOO_MANY_HYPERLINKS)
        );
    }

    #[test]
    fn oversized_state_is_refused_without_parsing() {
        let json = " ".repeat(MAX_CHART_STATE_BYTES + 1);
        assert_eq!(
            decode_charts(&json).unwrap_err(),
            "sheet chart state exceeds its size limit"
        );
        let json = " ".repeat(MAX_HYPERLINK_STATE_BYTES + 1);
        assert_eq!(
            decode_hyperlinks(&json).unwrap_err(),
            "sheet hyperlink state exceeds its size limit"
        );
    }

    #[test]
    fn oversized_fields_are_refused() {
        let long = "a".repeat(MAX_CHART_FIELD_BYTES + 1);
        let charts = vec![SheetChart {
            part: long.clone(),
            ..chart(0)
        }];
        let json = serde_json::to_string(&charts).unwrap();
        assert!(decode_charts(&json).unwrap_err().contains(CHART_TOO_LARGE));

        let charts = vec![SheetChart {
            refs: vec![ChartRef {
                kind: ChartRefKind::Values,
                formula: long.clone(),
            }],
            ..chart(0)
        }];
        let json = serde_json::to_string(&charts).unwrap();
        assert!(
            decode_charts(&json)
                .unwrap_err()
                .contains(CHART_REF_TOO_LONG)
        );

        let hyperlinks = vec![Hyperlink {
            tooltip: Some("a".repeat(MAX_HYPERLINK_FIELD_BYTES + 1)),
            ..hyperlink()
        }];
        let json = serde_json::to_string(&hyperlinks).unwrap();
        assert!(
            decode_hyperlinks(&json)
                .unwrap_err()
                .contains(HYPERLINK_FIELD_TOO_LONG)
        );
    }

    #[test]
    fn malformed_and_semantically_invalid_payloads_are_refused() {
        let charts = vec![SheetChart {
            part: String::new(),
            ..chart(0)
        }];
        assert!(
            decode_charts(&serde_json::to_string(&charts).unwrap())
                .unwrap_err()
                .contains(CHART_UNNAMED)
        );
        assert!(decode_charts("[{}]").is_err());
        assert!(decode_charts("[] []").is_err());
        assert!(decode_charts("{}").is_err());
        assert!(
            decode_charts(
                r#"[{"part":"p","drawing":"d","anchorIndex":0,"anchor":{"kind":"nope"},"refs":[]}]"#
            )
            .is_err()
        );

        let hyperlinks = vec![Hyperlink {
            external_target: None,
            location: None,
            ..hyperlink()
        }];
        assert!(
            decode_hyperlinks(&serde_json::to_string(&hyperlinks).unwrap())
                .unwrap_err()
                .contains(HYPERLINK_NO_DESTINATION)
        );
        let hyperlinks = vec![Hyperlink {
            range: CellRange {
                start: CellRef::new(0, 0),
                end: CellRef::new(MAX_ROWS, 0),
            },
            ..hyperlink()
        }];
        assert!(
            decode_hyperlinks(&serde_json::to_string(&hyperlinks).unwrap())
                .unwrap_err()
                .contains(HYPERLINK_OUT_OF_BOUNDS)
        );
    }

    #[test]
    fn a_hostile_payload_costs_the_cap_not_the_payload() {
        let count = (MAX_CHART_STATE_BYTES - 2) / 132;
        let json = hostile_charts(count);
        assert!(json.len() <= MAX_CHART_STATE_BYTES);
        assert!(count > MAX_CHARTS_PER_SHEET * 10);

        let bounded = peak_bytes(|| {
            let error = decode_charts(&json).unwrap_err();
            assert!(error.contains(TOO_MANY_CHARTS), "{error}");
        });
        let unbounded = peak_bytes(|| serde_json::from_str::<Vec<SheetChart>>(&json).unwrap());

        assert!(bounded < 2 * 1024 * 1024, "bounded peak {bounded} bytes");
        assert!(
            bounded * 4 < unbounded,
            "bounded {bounded} bytes, unbounded {unbounded} bytes"
        );
    }

    #[test]
    fn unknown_fields_are_ignored_without_buffering() {
        let json = r#"[{"part":"p","drawing":"d","anchorIndex":0,"anchor":{"kind":"absolute","pos":{"x":0,"y":0},"extent":{"cx":0,"cy":0},"junk":[1,2,3]},"refs":[{"kind":"other","formula":"f","junk":{"a":1}}],"junk":"x"}]"#;
        let charts = decode_charts(json).unwrap();
        assert_eq!(charts.len(), 1);
        assert_eq!(charts[0].refs[0].formula, "f");

        let junk = "0,".repeat(512 * 1024);
        let json = format!(
            r#"[{{"part":"p","drawing":"d","anchorIndex":0,"anchor":{{"kind":"absolute","pos":{{"x":0,"y":0}},"extent":{{"cx":0,"cy":0}},"junk":[{junk}0]}},"refs":[]}}]"#
        );
        let bounded = peak_bytes(|| decode_charts(&json).unwrap());
        let unbounded = peak_bytes(|| serde_json::from_str::<Vec<SheetChart>>(&json).unwrap());
        assert!(bounded < 64 * 1024, "bounded peak {bounded} bytes");
        assert!(
            bounded * 16 < unbounded,
            "bounded {bounded} bytes, unbounded {unbounded} bytes"
        );
    }
}
