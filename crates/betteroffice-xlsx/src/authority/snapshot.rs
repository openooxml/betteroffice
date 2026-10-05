use std::collections::LinkedList;

use xlsx_model::styles::{
    Alignment, Border, BorderEdge, BorderStyle, Color, Fill, Font, HAlign, Theme, VAlign, Xf,
};
use xlsx_model::{AnchorCell, ChartRefKind};

use crate::snapshot::growth::Growth;
use crate::snapshot::header::{SnapshotHeader, SnapshotMode};
use crate::snapshot::wire::{ChunkKind, Reader, Writer, frame};
use crate::snapshot::yrs_split::{
    CausalState, SplitError, UpdateCursor, split_fallback_v1_bounded, split_update_v1_bounded,
};
use crate::snapshot::{SnapshotBudget, SnapshotError, SnapshotProgress, SnapshotResult};

use super::*;

trait Codec: Sized {
    fn write(&self, w: &mut Writer);
    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self>;
}

macro_rules! primitive_codec {
    ($ty:ty, $method:ident) => {
        impl Codec for $ty {
            fn write(&self, w: &mut Writer) {
                w.$method(*self);
            }

            fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
                r.$method()
            }
        }
    };
}

primitive_codec!(bool, bool);
primitive_codec!(u8, u8);
primitive_codec!(u32, var_u32);
primitive_codec!(u64, var_u64);
primitive_codec!(usize, var_usize);
primitive_codec!(i64, var_i64);
primitive_codec!(f64, f64);

impl Codec for u16 {
    fn write(&self, w: &mut Writer) {
        w.var_u32(u32::from(*self));
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        u16::try_from(r.var_u32()?)
            .map_err(|_| SnapshotError::new("snapshot integer overflows u16"))
    }
}

impl Codec for String {
    fn write(&self, w: &mut Writer) {
        w.str(self);
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        Ok(r.str()?.to_owned())
    }
}

impl<T: Codec> Codec for Option<T> {
    fn write(&self, w: &mut Writer) {
        w.option(self.as_ref(), |w, value| value.write(w));
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        r.option(T::read)
    }
}

impl<T: Codec> Codec for Vec<T> {
    fn write(&self, w: &mut Writer) {
        w.var_usize(self.len());
        for value in self {
            value.write(w);
        }
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        let count = r.var_usize()?;
        let mut values = Vec::new();
        for _ in 0..count {
            values.push(T::read(r)?);
        }
        Ok(values)
    }
}

impl<K: Codec + Ord, V: Codec> Codec for BTreeMap<K, V> {
    fn write(&self, w: &mut Writer) {
        w.var_usize(self.len());
        for (key, value) in self {
            key.write(w);
            value.write(w);
        }
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        let count = r.var_usize()?;
        let mut values = Self::new();
        for _ in 0..count {
            let key = K::read(r)?;
            if values
                .last_key_value()
                .is_some_and(|(last, _)| last >= &key)
            {
                return Err(SnapshotError::new("snapshot map keys are not increasing"));
            }
            values.insert(key, V::read(r)?);
        }
        Ok(values)
    }
}

impl<A: Codec, B: Codec> Codec for (A, B) {
    fn write(&self, w: &mut Writer) {
        let (first, second) = self;
        first.write(w);
        second.write(w);
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        Ok((A::read(r)?, B::read(r)?))
    }
}

macro_rules! struct_codec {
    ($ty:ident { $($field:ident),+ $(,)? }) => {
        impl Codec for $ty {
            fn write(&self, w: &mut Writer) {
                let $ty { $($field),+ } = self;
                $($field.write(w);)+
            }

            fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
                Ok(Self { $($field: Codec::read(r)?),+ })
            }
        }
    };
}

macro_rules! enum_codec {
    ($ty:ident { $($variant:ident = $tag:literal),+ $(,)? }) => {
        impl Codec for $ty {
            fn write(&self, w: &mut Writer) {
                w.u8(match self { $(Self::$variant => $tag),+ });
            }

            fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
                match r.u8()? {
                    $($tag => Ok(Self::$variant),)+
                    _ => Err(SnapshotError::new(concat!("invalid snapshot ", stringify!($ty)))),
                }
            }
        }
    };
}

enum_codec!(DateSystem { V1900 = 0, V1904 = 1 });
enum_codec!(BorderStyle {
    Thin = 0, Medium = 1, Thick = 2, Dashed = 3, Dotted = 4, Double = 5, Hair = 6,
});
enum_codec!(HAlign {
    General = 0, Left = 1, Center = 2, Right = 3, Fill = 4, Justify = 5,
    CenterContinuous = 6, Distributed = 7,
});
enum_codec!(VAlign { Top = 0, Center = 1, Bottom = 2, Justify = 3, Distributed = 4 });
enum_codec!(AnchorEditAs { TwoCell = 0, OneCell = 1, Absolute = 2 });
enum_codec!(ChartRefKind {
    SeriesName = 0, Categories = 1, Values = 2, BubbleSize = 3, Title = 4,
    DataLabels = 5, Other = 6,
});

impl Codec for SheetId {
    fn write(&self, w: &mut Writer) {
        let SheetId(value) = self;
        value.write(w);
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        Ok(Self(r.var_u32()?))
    }
}

struct_codec!(CellRef {
    row,
    col,
    abs_row,
    abs_col
});
struct_codec!(CellRange { start, end });
struct_codec!(FreezePane {
    rows,
    cols,
    top_left
});
struct_codec!(DefinedName {
    name,
    formula,
    local_sheet,
    hidden
});
struct_codec!(Table {
    name,
    sheet,
    range,
    header_rows,
    totals_rows,
    columns
});
struct_codec!(Hyperlink {
    range,
    external_target,
    location,
    tooltip,
    display
});
struct_codec!(SheetFormat {
    default_row_height_pt,
    custom_height,
    zero_height
});
struct_codec!(ColStyle { first, last, xf });
struct_codec!(HiddenDimensions {
    col_widths,
    row_heights
});
struct_codec!(Font {
    name,
    size_pt,
    bold,
    italic,
    underline,
    strike,
    color
});
struct_codec!(BorderEdge { style, color });
struct_codec!(Border {
    left,
    right,
    top,
    bottom
});
struct_codec!(Alignment {
    h,
    v,
    wrap_text,
    shrink_to_fit
});
struct_codec!(Xf {
    font,
    fill,
    border,
    num_fmt_id,
    alignment
});
struct_codec!(AnchorCell {
    col,
    col_off,
    row,
    row_off
});
struct_codec!(AnchorExtent { cx, cy });
struct_codec!(AnchorPos { x, y });
struct_codec!(ChartRef { kind, formula });
struct_codec!(SheetChart {
    part,
    drawing,
    anchor_index,
    anchor,
    refs
});

impl Codec for Color {
    fn write(&self, w: &mut Writer) {
        match self {
            Self::Rgb(value) => {
                w.u8(0);
                value.write(w);
            }
            Self::Theme { idx, tint } => {
                w.u8(1);
                idx.write(w);
                tint.write(w);
            }
            Self::Indexed(value) => {
                w.u8(2);
                value.write(w);
            }
            Self::Auto => w.u8(3),
        }
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        Ok(match r.u8()? {
            0 => Self::Rgb(Codec::read(r)?),
            1 => Self::Theme {
                idx: Codec::read(r)?,
                tint: Codec::read(r)?,
            },
            2 => Self::Indexed(Codec::read(r)?),
            3 => Self::Auto,
            _ => return Err(SnapshotError::new("invalid snapshot color")),
        })
    }
}

impl Codec for Fill {
    fn write(&self, w: &mut Writer) {
        match self {
            Self::None => w.u8(0),
            Self::Solid(color) => {
                w.u8(1);
                color.write(w);
            }
        }
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        match r.u8()? {
            0 => Ok(Self::None),
            1 => Ok(Self::Solid(Codec::read(r)?)),
            _ => Err(SnapshotError::new("invalid snapshot fill")),
        }
    }
}

impl Codec for Theme {
    fn write(&self, w: &mut Writer) {
        let Self { colors } = self;
        for color in colors {
            color.write(w);
        }
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        let mut colors = std::array::from_fn(|_| String::new());
        for color in &mut colors {
            *color = Codec::read(r)?;
        }
        Ok(Self { colors })
    }
}

impl Codec for ChartAnchor {
    fn write(&self, w: &mut Writer) {
        match self {
            Self::TwoCell { from, to, edit_as } => {
                w.u8(0);
                from.write(w);
                to.write(w);
                edit_as.write(w);
            }
            Self::OneCell { from, extent } => {
                w.u8(1);
                from.write(w);
                extent.write(w);
            }
            Self::Absolute { pos, extent } => {
                w.u8(2);
                pos.write(w);
                extent.write(w);
            }
        }
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        Ok(match r.u8()? {
            0 => Self::TwoCell {
                from: Codec::read(r)?,
                to: Codec::read(r)?,
                edit_as: Codec::read(r)?,
            },
            1 => Self::OneCell {
                from: Codec::read(r)?,
                extent: Codec::read(r)?,
            },
            2 => Self::Absolute {
                pos: Codec::read(r)?,
                extent: Codec::read(r)?,
            },
            _ => return Err(SnapshotError::new("invalid snapshot chart anchor")),
        })
    }
}

const BASE_SECTIONS: usize = 16;

fn base_counts(base: &WorkbookBase) -> [usize; BASE_SECTIONS] {
    let WorkbookBase {
        bootstrap_client_id: _,
        date_system: _,
        defined_names,
        fingerprint: _,
        fingerprints,
        freeze_panes,
        formats,
        col_styles,
        hyperlinks,
        charts,
        hidden_dimensions,
        shared_strings,
        styles,
        tables,
    } = base;
    [
        defined_names.len(),
        fingerprints.len(),
        freeze_panes.len(),
        formats.len(),
        col_styles.len(),
        hyperlinks.len(),
        charts.len(),
        hidden_dimensions.len(),
        shared_strings.len(),
        styles.fonts.len(),
        styles.fills.len(),
        styles.borders.len(),
        styles.cell_xfs.len(),
        styles.num_fmts.len(),
        styles.indexed_colors.len(),
        tables.len(),
    ]
}

fn write_manifest(base: &WorkbookBase, w: &mut Writer) {
    let WorkbookBase {
        bootstrap_client_id,
        date_system,
        defined_names: _,
        fingerprint,
        fingerprints: _,
        freeze_panes: _,
        formats: _,
        col_styles: _,
        hyperlinks: _,
        charts: _,
        hidden_dimensions: _,
        shared_strings: _,
        styles,
        tables: _,
    } = base;
    bootstrap_client_id.write(w);
    date_system.write(w);
    fingerprint.write(w);
    styles.theme.write(w);
    for count in base_counts(base) {
        count.write(w);
    }
}

#[derive(Clone, Default)]
struct BaseCursor {
    section: usize,
    index: usize,
    fingerprint_key: Option<i64>,
}

impl BaseCursor {
    fn normalize(&mut self, counts: &[usize; BASE_SECTIONS]) {
        while self.section > 0
            && self.section <= BASE_SECTIONS
            && self.index == counts[self.section - 1]
        {
            self.section += 1;
            self.index = 0;
        }
    }

    fn next(&mut self, base: &WorkbookBase) -> Option<Vec<u8>> {
        self.normalize(&base_counts(base));
        if self.section > BASE_SECTIONS {
            return None;
        }
        let WorkbookBase {
            bootstrap_client_id: _,
            date_system: _,
            defined_names,
            fingerprint: _,
            fingerprints,
            freeze_panes,
            formats,
            col_styles,
            hyperlinks,
            charts,
            hidden_dimensions,
            shared_strings,
            styles,
            tables,
        } = base;
        let mut w = Writer::new();
        w.u8(self.section as u8);
        let index = self.index;
        match self.section {
            0 => write_manifest(base, &mut w),
            1 => defined_names[index].write(&mut w),
            2 => {
                let (key, values) = match self.fingerprint_key {
                    Some(key) => fingerprints
                        .range((std::ops::Bound::Excluded(key), std::ops::Bound::Unbounded))
                        .next(),
                    None => fingerprints.first_key_value(),
                }
                .expect("fingerprint count matches cursor");
                key.write(&mut w);
                values.write(&mut w);
                self.fingerprint_key = Some(*key);
            }
            3 => freeze_panes[index].write(&mut w),
            4 => formats[index].write(&mut w),
            5 => col_styles[index].write(&mut w),
            6 => hyperlinks[index].write(&mut w),
            7 => charts[index].write(&mut w),
            8 => hidden_dimensions[index].write(&mut w),
            9 => shared_strings[index].write(&mut w),
            10 => styles.fonts[index].write(&mut w),
            11 => styles.fills[index].write(&mut w),
            12 => styles.borders[index].write(&mut w),
            13 => styles.cell_xfs[index].write(&mut w),
            14 => styles.num_fmts[index].write(&mut w),
            15 => styles.indexed_colors[index].write(&mut w),
            16 => tables[index].write(&mut w),
            _ => unreachable!(),
        }
        if self.section == 0 {
            self.section = 1;
        } else {
            self.index += 1;
        }
        Some(w.into_bytes())
    }
}

struct BaseEncoder {
    base: Arc<WorkbookBase>,
    cursor: BaseCursor,
    pending: Option<Vec<u8>>,
    offset: usize,
}

impl BaseEncoder {
    fn new(base: Arc<WorkbookBase>) -> Self {
        Self {
            base,
            cursor: BaseCursor::default(),
            pending: None,
            offset: 0,
        }
    }

    fn next(&mut self, budget: SnapshotBudget) -> Option<Vec<u8>> {
        let mut w = Writer::new();
        let mut records = 0;
        while records < budget.max_records() && w.len() < budget.max_bytes() {
            if self.pending.is_none() {
                let Some(record) = self.cursor.next(&self.base) else {
                    break;
                };
                let mut framed_record = Writer::new();
                framed_record.bytes(&record);
                self.pending = Some(framed_record.into_bytes());
            }
            let record = self.pending.as_ref().expect("encoded base record");
            let count = (record.len() - self.offset).min(budget.max_bytes() - w.len());
            w.raw(&record[self.offset..self.offset + count]);
            self.offset += count;
            if self.offset == record.len() {
                self.pending = None;
                self.offset = 0;
                records += 1;
            }
        }
        (!w.is_empty()).then(|| w.into_bytes())
    }
}

#[derive(Default)]
struct BaseStreamDecoder {
    length: Option<usize>,
    varint: u64,
    shift: u32,
    record: Vec<u8>,
}

impl BaseStreamDecoder {
    fn push(
        &mut self,
        mut payload: &[u8],
        records: &mut LinkedList<Vec<u8>>,
    ) -> SnapshotResult<()> {
        while !payload.is_empty() {
            if let Some(length) = self.length {
                let count = (length - self.record.len()).min(payload.len());
                self.record.extend_from_slice(&payload[..count]);
                payload = &payload[count..];
                if self.record.len() == length {
                    records.push_back(std::mem::take(&mut self.record));
                    self.length = None;
                }
            } else {
                let byte = payload[0];
                payload = &payload[1..];
                let bits = u64::from(byte & 0x7f);
                if self.shift == 63 && bits > 1 {
                    return Err(SnapshotError::new("authority base length overflows u64"));
                }
                self.varint |= bits << self.shift;
                if byte & 0x80 == 0 {
                    let length = usize::try_from(self.varint)
                        .map_err(|_| SnapshotError::new("authority base length overflows usize"))?;
                    if length == 0 {
                        return Err(SnapshotError::new("authority base record is empty"));
                    }
                    self.length = Some(length);
                    self.varint = 0;
                    self.shift = 0;
                } else {
                    self.shift += 7;
                    if self.shift > 63 {
                        return Err(SnapshotError::new("authority base length overflows u64"));
                    }
                }
            }
        }
        Ok(())
    }

    fn is_complete(&self) -> bool {
        self.length.is_none() && self.shift == 0 && self.record.is_empty()
    }
}

pub(crate) struct AuthoritySnapshotEncoder {
    base: BaseEncoder,
    yrs: std::vec::IntoIter<Vec<u8>>,
    base_ordinal: u64,
    yrs_ordinal: u64,
    yrs_count: u64,
    split_fallback: Option<String>,
}

impl AuthoritySnapshotEncoder {
    pub(crate) fn new(a: &WorkbookAuthority, budget: SnapshotBudget) -> SnapshotResult<Self> {
        let WorkbookAuthority {
            doc,
            projection_valid: _,
            snapshot_revision: _,
            base,
            history: _,
            next_sheet_id: _,
            undo_stack: _,
            redo_stack: _,
        } = a;
        let update = doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let (parts, split_fallback) =
            match split_update_v1_bounded(&update, budget.max_records(), budget.max_bytes()) {
                Ok(parts) => (parts, None),
                Err(reason) => (
                    split_fallback_v1_bounded(&update, budget.max_records(), budget.max_bytes())
                        .map_err(|failure| {
                            SnapshotError::new(format!(
                                "invalid Yrs snapshot: {}",
                                failure.reason(),
                            ))
                        })?,
                    Some(reason.reason().to_owned()),
                ),
            };
        let yrs_count = parts.len() as u64;
        Ok(Self {
            base: BaseEncoder::new(Arc::clone(base)),
            yrs: parts.into_iter(),
            base_ordinal: 0,
            yrs_ordinal: 0,
            yrs_count,
            split_fallback,
        })
    }

    pub(crate) fn next(&mut self, budget: SnapshotBudget) -> SnapshotResult<Option<Vec<u8>>> {
        if let Some(payload) = self.base.next(budget) {
            let chunk = frame(ChunkKind::AuthorityBase, self.base_ordinal, &payload);
            self.base_ordinal += 1;
            return Ok(Some(chunk));
        }
        if let Some(payload) = self.yrs.next() {
            let chunk = frame(ChunkKind::Yrs, self.yrs_ordinal, &payload);
            self.yrs_ordinal += 1;
            return Ok(Some(chunk));
        }
        Ok(None)
    }

    pub(crate) fn split_fallback(&self) -> Option<String> {
        self.split_fallback.clone()
    }

    pub(crate) fn chunk_counts(&self, budget: SnapshotBudget) -> (u64, u64) {
        let mut base = BaseEncoder::new(Arc::clone(&self.base.base));
        let mut count = 0;
        while base.next(budget).is_some() {
            count += 1;
        }
        (count, self.yrs_count)
    }
}

impl WorkbookAuthority {
    #[cfg(test)]
    pub(crate) fn snapshot_deletion_history_for_test(&self, count: u32) {
        let map = self.doc.get_or_insert_map("snapshot-deletions");
        {
            let mut txn = self.doc.transact_mut();
            for index in 0..count {
                map.insert(&mut txn, index.to_string(), index);
            }
        }
        let mut txn = self.doc.transact_mut();
        for index in 0..count {
            map.remove(&mut txn, &index.to_string());
        }
    }

    pub(crate) fn snapshot_identity(&self) -> (u64, String, u64) {
        let Self {
            doc,
            projection_valid: _,
            snapshot_revision: _,
            base: _,
            history: _,
            next_sheet_id,
            undo_stack: _,
            redo_stack: _,
        } = self;
        (
            doc.client_id().get(),
            doc.guid().to_string(),
            *next_sheet_id,
        )
    }
}

#[derive(Default)]
struct BaseBuilder {
    base: Option<WorkbookBase>,
    counts: [usize; BASE_SECTIONS],
    cursor: BaseCursor,
    growth: Growth<WorkbookBase>,
}

impl BaseBuilder {
    fn push(&mut self, payload: &[u8]) -> SnapshotResult<()> {
        let mut r = Reader::new(payload);
        self.cursor.normalize(&self.counts);
        let section = usize::from(r.u8()?);
        if section != self.cursor.section || section > BASE_SECTIONS {
            return Err(SnapshotError::new("unexpected authority base record"));
        }
        if section == 0 {
            let bootstrap_client_id = Codec::read(&mut r)?;
            let date_system = Codec::read(&mut r)?;
            let fingerprint = Codec::read(&mut r)?;
            let theme = Codec::read(&mut r)?;
            let mut styles = Stylesheet::default();
            for count in &mut self.counts {
                *count = r.var_usize()?;
            }
            styles.theme = theme;
            self.base = Some(WorkbookBase {
                bootstrap_client_id,
                date_system,
                defined_names: Vec::new(),
                fingerprint,
                fingerprints: BTreeMap::new(),
                freeze_panes: Vec::new(),
                formats: Vec::new(),
                col_styles: Vec::new(),
                hyperlinks: Vec::new(),
                charts: Vec::new(),
                hidden_dimensions: Vec::new(),
                shared_strings: Vec::new(),
                styles,
                tables: Vec::new(),
            });
            self.cursor.section = 1;
        } else {
            let base = self
                .base
                .as_mut()
                .ok_or_else(|| SnapshotError::new("authority base manifest is missing"))?;
            match section {
                1 => base.defined_names.push(Codec::read(&mut r)?),
                2 => {
                    let key = Codec::read(&mut r)?;
                    if self
                        .cursor
                        .fingerprint_key
                        .is_some_and(|previous| previous >= key)
                    {
                        return Err(SnapshotError::new(
                            "snapshot fingerprints are not increasing",
                        ));
                    }
                    base.fingerprints.insert(key, Codec::read(&mut r)?);
                    self.cursor.fingerprint_key = Some(key);
                }
                3 => base.freeze_panes.push(Codec::read(&mut r)?),
                4 => base.formats.push(Codec::read(&mut r)?),
                5 => base.col_styles.push(Codec::read(&mut r)?),
                6 => base.hyperlinks.push(Codec::read(&mut r)?),
                7 => base.charts.push(Codec::read(&mut r)?),
                8 => base.hidden_dimensions.push(Codec::read(&mut r)?),
                9 => base.shared_strings.push(Codec::read(&mut r)?),
                10 => base.styles.fonts.push(Codec::read(&mut r)?),
                11 => base.styles.fills.push(Codec::read(&mut r)?),
                12 => base.styles.borders.push(Codec::read(&mut r)?),
                13 => base.styles.cell_xfs.push(Codec::read(&mut r)?),
                14 => base.styles.num_fmts.push(Codec::read(&mut r)?),
                15 => base.styles.indexed_colors.push(Codec::read(&mut r)?),
                16 => base.tables.push(Codec::read(&mut r)?),
                _ => return Err(SnapshotError::new("invalid authority base section")),
            }
            self.cursor.index += 1;
        }
        r.finish()
    }

    fn advance_capacity(&mut self, record: &[u8], budget: SnapshotBudget) -> SnapshotResult<bool> {
        self.cursor.normalize(&self.counts);
        let section = usize::from(Reader::new(record).u8()?);
        if section != self.cursor.section || section > BASE_SECTIONS {
            return Err(SnapshotError::new("unexpected authority base record"));
        }
        if section == 0 {
            return Ok(true);
        }
        let base = self
            .base
            .as_mut()
            .ok_or_else(|| SnapshotError::new("authority base manifest is missing"))?;
        let g = &mut self.growth;
        match section {
            1 => g.ensure(base, |b| Ok(&mut b.defined_names), budget),
            3 => g.ensure(base, |b| Ok(&mut b.freeze_panes), budget),
            4 => g.ensure(base, |b| Ok(&mut b.formats), budget),
            5 => g.ensure(base, |b| Ok(&mut b.col_styles), budget),
            6 => g.ensure(base, |b| Ok(&mut b.hyperlinks), budget),
            7 => g.ensure(base, |b| Ok(&mut b.charts), budget),
            8 => g.ensure(base, |b| Ok(&mut b.hidden_dimensions), budget),
            9 => g.ensure(base, |b| Ok(&mut b.shared_strings), budget),
            10 => g.ensure(base, |b| Ok(&mut b.styles.fonts), budget),
            11 => g.ensure(base, |b| Ok(&mut b.styles.fills), budget),
            12 => g.ensure(base, |b| Ok(&mut b.styles.borders), budget),
            13 => g.ensure(base, |b| Ok(&mut b.styles.cell_xfs), budget),
            14 => g.ensure(base, |b| Ok(&mut b.styles.num_fmts), budget),
            15 => g.ensure(base, |b| Ok(&mut b.styles.indexed_colors), budget),
            16 => g.ensure(base, |b| Ok(&mut b.tables), budget),
            _ => Ok(true),
        }
    }

    fn is_complete(&self) -> bool {
        let mut cursor = self.cursor.clone();
        cursor.normalize(&self.counts);
        self.base.is_some() && cursor.section > BASE_SECTIONS && !self.growth.is_pending()
    }

    fn finish(self) -> SnapshotResult<WorkbookBase> {
        if !self.is_complete() {
            return Err(SnapshotError::new("authority base is incomplete"));
        }
        self.base
            .ok_or_else(|| SnapshotError::new("authority base is missing"))
    }
}

pub(crate) struct AuthorityHydrator {
    doc: Doc,
    base: BaseBuilder,
    base_stream: BaseStreamDecoder,
    pending_base: LinkedList<Vec<u8>>,
    expected_base_chunks: u64,
    expected_yrs_chunks: u64,
    base_chunks: u64,
    yrs_chunks: u64,
    state_vector: Vec<u8>,
    next_sheet_id: u64,
    failed: bool,
    yrs_cursor: Option<UpdateCursor>,
    causal: CausalState,
    final_vector_offset: usize,
    final_vector_clients: usize,
    vector_validated: bool,
}

fn hydrate_snapshot_doc(doc: &Doc, payload: &[u8]) -> Result<(), String> {
    let update = decode_local_update_v1(payload)?;
    let mut transaction = doc.transact_mut_with(HYDRATE_ORIGIN);
    let events = transaction
        .events()
        .is_some()
        .then(|| std::mem::take(transaction.events_mut()));
    let result = transaction
        .apply_update(update)
        .map_err(|failure| failure.to_string());
    transaction.commit();
    if let Some(events) = events {
        *transaction.events_mut() = events;
    }
    result
}

impl AuthorityHydrator {
    pub(crate) fn new(header: &SnapshotHeader) -> SnapshotResult<Self> {
        match header.mode {
            SnapshotMode::Standalone => {}
        }
        if header.client_id > MAX_SAFE_CLIENT_ID {
            return Err(SnapshotError::new("snapshot client id is out of range"));
        }
        decode_state_vector_v1(&header.state_vector).map_err(SnapshotError::new)?;
        let expected_base_chunks = header.chunk_count(ChunkKind::AuthorityBase);
        let expected_yrs_chunks = header.chunk_count(ChunkKind::Yrs);
        if expected_base_chunks == 0 || expected_yrs_chunks == 0 {
            return Err(SnapshotError::new("snapshot authority chunks are missing"));
        }
        let mut vector = Reader::new(&header.state_vector);
        let final_vector_clients = vector.var_usize()?;
        let final_vector_offset = header.state_vector.len() - vector.rest().len();
        Ok(Self {
            doc: Doc::with_options(Options::with_guid_and_client_id(
                header.guid.as_str().into(),
                ClientID::new(header.client_id),
            )),
            base: BaseBuilder::default(),
            base_stream: BaseStreamDecoder::default(),
            pending_base: LinkedList::new(),
            expected_base_chunks,
            expected_yrs_chunks,
            base_chunks: 0,
            yrs_chunks: 0,
            state_vector: header.state_vector.clone(),
            next_sheet_id: header.next_sheet_id,
            failed: false,
            yrs_cursor: None,
            causal: CausalState::default(),
            final_vector_offset,
            final_vector_clients,
            vector_validated: false,
        })
    }

    pub(crate) fn push_base(&mut self, payload: &[u8]) -> SnapshotResult<()> {
        let result = (|| {
            if self.failed || self.base_chunks >= self.expected_base_chunks {
                return Err(SnapshotError::new("unexpected authority base chunk"));
            }
            if payload.is_empty() {
                return Err(SnapshotError::new("authority base chunk is empty"));
            }
            self.base_stream.push(payload, &mut self.pending_base)?;
            self.base_chunks += 1;
            Ok(())
        })();
        self.failed |= result.is_err();
        result
    }

    #[cfg(test)]
    pub(crate) fn push_yrs(&mut self, payload: &[u8]) -> SnapshotResult<()> {
        let result = (|| {
            if self.failed || self.yrs_chunks >= self.expected_yrs_chunks {
                return Err(SnapshotError::new("unexpected Yrs snapshot chunk"));
            }
            hydrate_snapshot_doc(&self.doc, payload).map_err(SnapshotError::new)?;
            self.yrs_chunks += 1;
            Ok(())
        })();
        self.failed |= result.is_err();
        result
    }

    pub(crate) fn advance_yrs(
        &mut self,
        payload: &[u8],
        budget: SnapshotBudget,
    ) -> SnapshotResult<SnapshotProgress> {
        if self.failed || self.yrs_chunks >= self.expected_yrs_chunks {
            return Err(SnapshotError::new("unexpected Yrs snapshot chunk"));
        }
        let cursor = self.yrs_cursor.get_or_insert_with(UpdateCursor::default);
        if let Some(part) = cursor
            .next(payload, budget.max_records(), budget.max_bytes())
            .map_err(|failure| {
                if failure == SplitError::OversizedStruct {
                    SnapshotError::new("Yrs snapshot record exceeds advance byte budget")
                } else {
                    SnapshotError::new(format!("invalid Yrs snapshot: {}", failure.reason()))
                }
            })?
        {
            if has_pending(&self.doc) {
                self.failed = true;
                return Err(SnapshotError::new(
                    "authority snapshot has pending Yrs state",
                ));
            }
            if let Err(failure) = self.causal.admit(&part.bytes) {
                self.failed = true;
                return Err(SnapshotError::new(format!(
                    "invalid Yrs snapshot: {}",
                    failure.reason(),
                )));
            }
            hydrate_snapshot_doc(&self.doc, &part.bytes).map_err(SnapshotError::new)?;
            if has_pending(&self.doc) {
                self.failed = true;
                return Err(SnapshotError::new(
                    "authority snapshot has pending Yrs state",
                ));
            }
            #[cfg(test)]
            {
                crate::snapshot::step::record(part.records, part.bytes.len());
                if part.bytes[0] == 0 {
                    crate::snapshot::step::delete(part.records);
                }
            }
            if !cursor
                .is_complete(payload)
                .map_err(|_| SnapshotError::new("invalid Yrs snapshot tail"))?
            {
                return Ok(SnapshotProgress::pending());
            }
            self.yrs_cursor = None;
            self.yrs_chunks += 1;
            Ok(SnapshotProgress::ready())
        } else {
            self.yrs_cursor = None;
            self.yrs_chunks += 1;
            Ok(SnapshotProgress::ready())
        }
    }

    #[cfg(test)]
    pub(crate) fn advance(&mut self, budget: SnapshotBudget) -> SnapshotResult<SnapshotProgress> {
        self.advance_base(SnapshotBudget::new(
            budget.max_records(),
            budget.max_bytes().max(std::mem::size_of::<WorkbookBase>()),
        )?)
    }

    fn advance_base(&mut self, budget: SnapshotBudget) -> SnapshotResult<SnapshotProgress> {
        if self.failed {
            return Err(SnapshotError::new("authority snapshot hydration failed"));
        }
        if let Some(record) = self.pending_base.front() {
            if !self.base.advance_capacity(record, budget)? {
                return Ok(SnapshotProgress::pending());
            }
            let record = self
                .pending_base
                .pop_front()
                .ok_or_else(|| SnapshotError::new("authority base record is missing"))?;
            #[cfg(test)]
            crate::snapshot::step::record(1, record.len());
            if let Err(error) = self.base.push(&record) {
                self.failed = true;
                return Err(error);
            }
        }
        if self.pending_base.is_empty() {
            if self.base_chunks == self.expected_base_chunks && !self.base.is_complete() {
                self.failed = true;
                return Err(SnapshotError::new("authority base is incomplete"));
            }
            Ok(SnapshotProgress::ready())
        } else {
            Ok(SnapshotProgress::pending())
        }
    }

    pub(crate) fn has_pending_base(&self) -> bool {
        !self.pending_base.is_empty()
    }

    pub(crate) fn advance_bounded(
        &mut self,
        budget: SnapshotBudget,
    ) -> SnapshotResult<SnapshotProgress> {
        if self
            .pending_base
            .front()
            .is_some_and(|record| record.len() > budget.max_bytes())
        {
            return Err(SnapshotError::new(
                "authority base record exceeds advance byte budget",
            ));
        }
        self.advance_base(budget)
    }

    pub(crate) fn finish_drained(self) -> SnapshotResult<WorkbookAuthority> {
        if !self.causal.is_empty() || !self.vector_validated {
            return Err(SnapshotError::new(
                "authority causal metadata is not drained",
            ));
        }
        self.finish_parts()
    }

    #[cfg(test)]
    pub(crate) fn finish(self) -> SnapshotResult<WorkbookAuthority> {
        self.finish_parts()
    }

    fn finish_parts(self) -> SnapshotResult<WorkbookAuthority> {
        if self.failed
            || self.base_chunks != self.expected_base_chunks
            || self.yrs_chunks != self.expected_yrs_chunks
        {
            return Err(SnapshotError::new(
                "authority snapshot chunks are incomplete",
            ));
        }
        if !self.pending_base.is_empty() || self.yrs_cursor.is_some() {
            return Err(SnapshotError::new(
                "authority snapshot has pending base records",
            ));
        }
        if !self.base_stream.is_complete() {
            return Err(SnapshotError::new(
                "authority snapshot base record is truncated",
            ));
        }
        if has_pending(&self.doc) {
            return Err(SnapshotError::new(
                "authority snapshot has pending Yrs state",
            ));
        }
        let base = self.base.finish()?;
        let authority = WorkbookAuthority::hydrated(self.doc, Arc::new(base), self.next_sheet_id);
        if !self.vector_validated && authority.encode_state_vector_v1() != self.state_vector {
            return Err(SnapshotError::new(
                "authority snapshot state vector differs",
            ));
        }
        Ok(authority)
    }

    pub(crate) fn advance_finalization(
        &mut self,
        budget: SnapshotBudget,
    ) -> SnapshotResult<SnapshotProgress> {
        let mut records = 0;
        let mut bytes = 0;
        while records < budget.max_records() && !self.causal.is_empty() {
            let clock = self.causal.final_clock();
            let (_, count, cost) = self.causal.drain(1, budget.max_bytes() - bytes);
            if count == 0 {
                break;
            }
            if let Some((client, clock)) = clock {
                let mut vector = Reader::new(&self.state_vector[self.final_vector_offset..]);
                if self.final_vector_clients == 0
                    || vector.var_u64()? != client
                    || vector.var_u32()? != clock
                {
                    return Err(SnapshotError::new(
                        "authority snapshot state vector differs",
                    ));
                }
                self.final_vector_clients -= 1;
                self.final_vector_offset = self.state_vector.len() - vector.rest().len();
            }
            records += count;
            bytes += cost;
        }
        #[cfg(test)]
        {
            crate::snapshot::step::record(records, bytes);
            crate::snapshot::step::drain(records);
        }
        if self.causal.is_empty() {
            if self.final_vector_clients != 0 || self.final_vector_offset != self.state_vector.len()
            {
                return Err(SnapshotError::new(
                    "authority snapshot state vector differs",
                ));
            }
            self.vector_validated = true;
        }
        if self.vector_validated && records == 0 {
            Ok(SnapshotProgress::ready())
        } else {
            Ok(SnapshotProgress::pending())
        }
    }
}

#[cfg(test)]
mod tests {
    use std::sync::atomic::{AtomicUsize, Ordering};

    use yrs::updates::encoder::Encode;

    use crate::CalculationResult;
    use crate::snapshot::wire::unframe;

    use super::*;

    fn model() -> WorkbookModel {
        let mut sheet = Sheet::new("Data");
        for row in 0..40 {
            sheet.set_cell(
                CellRef::new(row, row % 4),
                Cell {
                    value: CellValue::Number {
                        value: f64::from(row),
                    },
                    formula: Some(format!("{row}+0")),
                    style: None,
                },
            );
        }
        WorkbookModel {
            sheets: vec![sheet, Sheet::new("Second")],
            ..WorkbookModel::default()
        }
    }

    fn source() -> WorkbookAuthority {
        WorkbookAuthority::from_model(&model()).unwrap()
    }

    fn header(
        a: &WorkbookAuthority,
        encoder: &AuthoritySnapshotEncoder,
        budget: SnapshotBudget,
    ) -> SnapshotHeader {
        let (client_id, guid, next_sheet_id) = a.snapshot_identity();
        let (base_count, yrs_count) = encoder.chunk_counts(budget);
        let mut chunk_counts = [0; 9];
        chunk_counts[ChunkKind::Header as usize - 1] = 1;
        chunk_counts[ChunkKind::AuthorityBase as usize - 1] = base_count;
        chunk_counts[ChunkKind::Yrs as usize - 1] = yrs_count;
        chunk_counts[ChunkKind::End as usize - 1] = 1;
        SnapshotHeader {
            snapshot_id: 1,
            mode: SnapshotMode::Standalone,
            edited_since_open: false,
            recalculated_since_open: false,
            moved_references_since_open: false,
            active_sheet: SheetId(0),
            rand_seed: None,
            model_epoch: 0,
            version_nonce: "version".into(),
            committed_changes: 0,
            last_calculation: CalculationResult::default(),
            calculation_context: None,
            client_id,
            guid,
            next_sheet_id,
            state_vector: a.encode_state_vector_v1(),
            chunk_counts,
        }
    }

    fn push_all(
        encoder: &mut AuthoritySnapshotEncoder,
        hydrator: &mut AuthorityHydrator,
        budget: SnapshotBudget,
    ) -> (u64, u64) {
        let mut base_ordinal = 0;
        let mut yrs_ordinal = 0;
        while let Some(chunk) = encoder.next(budget).unwrap() {
            let (kind, ordinal, payload) = unframe(&chunk).unwrap();
            match kind {
                ChunkKind::AuthorityBase => {
                    assert_eq!(ordinal, base_ordinal);
                    assert!(payload.len() <= budget.max_bytes());
                    hydrator.push_base(payload).unwrap();
                    base_ordinal += 1;
                }
                ChunkKind::Yrs => {
                    assert_eq!(ordinal, yrs_ordinal);
                    hydrator.push_yrs(payload).unwrap();
                    yrs_ordinal += 1;
                }
                _ => panic!("unexpected authority chunk"),
            }
            while !hydrator.advance(budget).unwrap().is_ready() {}
        }
        (base_ordinal, yrs_ordinal)
    }

    fn base_records(base: &WorkbookBase) -> Vec<Vec<u8>> {
        let mut cursor = BaseCursor::default();
        let mut records = Vec::new();
        while let Some(record) = cursor.next(base) {
            records.push(record);
        }
        records
    }

    #[test]
    fn huge_base_manifest_does_not_reserve_declared_strings() {
        let a = source();
        let budget = SnapshotBudget::new(1, 16 * 1024).unwrap();
        let encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        let mut header = header(&a, &encoder, budget);
        header.chunk_counts[ChunkKind::AuthorityBase as usize - 1] = 2;
        let mut manifest = Writer::new();
        manifest.u8(0);
        a.base.bootstrap_client_id.write(&mut manifest);
        a.base.date_system.write(&mut manifest);
        a.base.fingerprint.write(&mut manifest);
        a.base.styles.theme.write(&mut manifest);
        for section in 0..BASE_SECTIONS {
            manifest.var_usize(if section == 8 { 50_000_000 } else { 0 });
        }
        let mut payload = Writer::new();
        payload.bytes(&manifest.into_bytes());
        let payload = payload.into_bytes();
        let mut hydrator = AuthorityHydrator::new(&header).unwrap();
        hydrator.push_base(&payload).unwrap();
        crate::snapshot::step::reset();
        assert!(hydrator.advance_bounded(budget).unwrap().is_ready());
        assert_eq!(
            hydrator
                .base
                .base
                .as_ref()
                .unwrap()
                .shared_strings
                .capacity(),
            0
        );
        let work = crate::snapshot::step::current();
        assert_eq!(work.records, 1);
        assert!(work.bytes <= budget.max_bytes());
        let failure: SnapshotError = hydrator.finish().err().unwrap();
        assert_eq!(
            failure.to_string(),
            "authority snapshot chunks are incomplete"
        );

        header.chunk_counts[ChunkKind::AuthorityBase as usize - 1] = 1;
        let mut hydrator = AuthorityHydrator::new(&header).unwrap();
        hydrator.push_base(&payload).unwrap();
        let failure: SnapshotError = hydrator.advance_bounded(budget).unwrap_err();
        assert_eq!(failure.to_string(), "authority base is incomplete");
        assert_eq!(
            hydrator
                .base
                .base
                .as_ref()
                .unwrap()
                .shared_strings
                .capacity(),
            0
        );
    }

    #[test]
    fn nested_base_count_is_refused_without_declared_allocation() {
        let mut builder = BaseBuilder::default();
        let a = source();
        let mut base = (*a.base).clone();
        base.col_styles = vec![Vec::new()];
        let mut hostile = Writer::new();
        hostile.u8(0);
        base.bootstrap_client_id.write(&mut hostile);
        base.date_system.write(&mut hostile);
        base.fingerprint.write(&mut hostile);
        base.styles.theme.write(&mut hostile);
        for (section, count) in base_counts(&base).into_iter().enumerate() {
            hostile.var_usize(if section == 4 { 50_000_000 } else { count });
        }
        builder.push(&hostile.into_bytes()).unwrap();
        assert_eq!(builder.base.as_ref().unwrap().col_styles.capacity(), 0);
        assert!(builder.base.as_ref().unwrap().col_styles.is_empty());
        assert!(!builder.is_complete());
        builder = BaseBuilder::default();
        let mut manifest = Writer::new();
        manifest.u8(0);
        write_manifest(&base, &mut manifest);
        builder.push(&manifest.into_bytes()).unwrap();
        let mut cursor = BaseCursor {
            section: 1,
            ..BaseCursor::default()
        };
        let mut refused = false;
        while let Some(record) = cursor.next(&base) {
            if record[0] == 5 {
                let mut record = Writer::new();
                record.u8(5);
                record.var_usize(50_000_000);
                let record = record.into_bytes();
                let budget = SnapshotBudget::new(1, 16 * 1024).unwrap();
                assert!(builder.advance_capacity(&record, budget).unwrap());
                let failure: SnapshotError = builder.push(&record).unwrap_err();
                assert_eq!(failure.to_string(), "snapshot payload is truncated");
                assert!(builder.base.as_ref().unwrap().col_styles.capacity() <= 1);
                assert!(builder.base.as_ref().unwrap().col_styles.is_empty());
                refused = true;
                break;
            }
            builder.push(&record).unwrap();
        }
        assert!(refused);
    }

    #[test]
    fn base_storage_growth_moves_only_budgeted_admitted_records() {
        let mut a = source();
        Arc::make_mut(&mut a.base).shared_strings =
            (0..300).map(|index| index.to_string()).collect();
        let budget = SnapshotBudget::new(1, 16 * 1024).unwrap();
        let mut encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        let header = header(&a, &encoder, budget);
        let mut hydrator = AuthorityHydrator::new(&header).unwrap();
        let mut steps = 0;
        let mut base_chunks = 0;
        while let Some(chunk) = encoder.next(budget).unwrap() {
            let (kind, _, payload) = unframe(&chunk).unwrap();
            if kind == ChunkKind::AuthorityBase {
                base_chunks += 1;
                hydrator.push_base(payload).unwrap();
                loop {
                    crate::snapshot::step::reset();
                    let progress = hydrator.advance_bounded(budget).unwrap();
                    let work = crate::snapshot::step::current();
                    assert!(work.records <= budget.max_records());
                    assert!(work.bytes <= budget.max_bytes());
                    steps += 1;
                    if progress.is_ready() {
                        break;
                    }
                }
            } else {
                while !hydrator.advance_yrs(payload, budget).unwrap().is_ready() {}
            }
        }
        assert!(steps > base_chunks);
        let restored = hydrator.finish().unwrap();
        assert_eq!(base_records(&a.base), base_records(&restored.base));
        assert_eq!(
            restored.encode_state_as_update_v1(),
            a.encode_state_as_update_v1()
        );
    }

    #[test]
    fn advance_refuses_out_of_order_structs_and_deletes_before_integration() {
        let foreign = Doc::with_client_id(99);
        let map = foreign.get_or_insert_map("out-of-order");
        map.insert(&mut foreign.transact_mut(), "first", 1);
        let before = foreign.transact().state_vector();
        map.insert(&mut foreign.transact_mut(), "second", 2);
        let structs = foreign.transact().encode_diff_v1(&before);
        let before = foreign.transact().state_vector();
        map.remove(&mut foreign.transact_mut(), "first");
        let deletes = foreign.transact().encode_diff_v1(&before);
        let a = source();
        let budget = SnapshotBudget::new(1, 16 * 1024).unwrap();
        let encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        let header = header(&a, &encoder, budget);
        for payload in [structs, deletes] {
            let mut hydrator = AuthorityHydrator::new(&header).unwrap();
            crate::snapshot::step::reset();
            let failure: SnapshotError = hydrator.advance_yrs(&payload, budget).unwrap_err();
            assert_eq!(
                failure.to_string(),
                "invalid Yrs snapshot: missing_dependency"
            );
            assert!(hydrator.doc.transact().state_vector().is_empty());
            assert!(!has_pending(&hydrator.doc));
            assert_eq!(crate::snapshot::step::current().records, 0);
        }
    }

    #[test]
    fn advance_refuses_recursive_type_deletion_before_integration() {
        let foreign = Doc::with_client_id(99);
        let root = foreign.get_or_insert_map("recursive-delete");
        let nested = root.insert(&mut foreign.transact_mut(), "nested", MapPrelim::default());
        {
            let mut txn = foreign.transact_mut();
            for index in 0..1_000 {
                nested.insert(&mut txn, index.to_string(), index);
            }
        }
        let update = foreign
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let a = source();
        let budget = SnapshotBudget::new(1, 16 * 1024).unwrap();
        let encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        let mut header = header(&a, &encoder, budget);
        header.chunk_counts[ChunkKind::Yrs as usize - 1] = 2;
        let mut hydrator = AuthorityHydrator::new(&header).unwrap();
        while !hydrator.advance_yrs(&update, budget).unwrap().is_ready() {}
        let before = hydrator
            .doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let mut payload = Writer::new();
        payload.raw(&[0, 1]);
        payload.var_u64(99);
        payload.var_u32(1);
        payload.var_u32(0);
        payload.var_u32(1);
        crate::snapshot::step::reset();
        let failure: SnapshotError = hydrator
            .advance_yrs(&payload.into_bytes(), budget)
            .unwrap_err();
        assert_eq!(
            failure.to_string(),
            "invalid Yrs snapshot: retained_deletion"
        );
        assert_eq!(crate::snapshot::step::current().records, 0);
        assert_eq!(
            hydrator
                .doc
                .transact()
                .encode_state_as_update_v1(&StateVector::default()),
            before,
        );
        assert!(!has_pending(&hydrator.doc));
    }

    #[test]
    fn advance_refuses_delete_range_without_admitted_clocks() {
        let a = source();
        let budget = SnapshotBudget::new(1, 16 * 1024).unwrap();
        let encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        let header = header(&a, &encoder, budget);
        let mut hydrator = AuthorityHydrator::new(&header).unwrap();
        let before = hydrator
            .doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let mut payload = Writer::new();
        payload.raw(&[0, 1]);
        payload.var_u64(99);
        payload.var_u32(1);
        payload.var_u32(0);
        payload.var_u32(100_000);
        crate::snapshot::step::reset();
        let failure: SnapshotError = hydrator
            .advance_yrs(&payload.into_bytes(), budget)
            .unwrap_err();
        assert_eq!(
            failure.to_string(),
            "invalid Yrs snapshot: missing_dependency"
        );
        assert_eq!(crate::snapshot::step::current().records, 0);
        assert_eq!(
            hydrator
                .doc
                .transact()
                .encode_state_as_update_v1(&StateVector::default()),
            before,
        );
        assert!(!has_pending(&hydrator.doc));
    }

    #[test]
    fn base_roundtrip_preserves_accepted_fingerprints_and_pools() {
        let mut a = source();
        let base = Arc::make_mut(&mut a.base);
        base.bootstrap_client_id = 17;
        base.date_system = DateSystem::V1904;
        base.fingerprint = "kept-verbatim".into();
        base.fingerprints = BTreeMap::from([
            (
                i64::MIN,
                vec!["later".into(), "first".into(), "later".into()],
            ),
            (3, Vec::new()),
            (i64::MAX, vec![String::new()]),
        ]);
        base.defined_names = vec![
            DefinedName {
                name: "Second".into(),
                formula: String::new(),
                local_sheet: Some(SheetId(0)),
                hidden: true,
            },
            DefinedName {
                name: String::new(),
                formula: "Data!$A$1".into(),
                local_sheet: None,
                hidden: false,
            },
        ];
        base.freeze_panes = vec![
            None,
            Some(FreezePane::new(0, 0, CellRef::parse_a1("$B$2").unwrap())),
        ];
        base.formats = vec![
            SheetFormat {
                default_row_height_pt: Some(f64::from_bits(0x7ff8_0000_0000_0042)),
                custom_height: true,
                zero_height: false,
            },
            SheetFormat {
                default_row_height_pt: Some(-0.0),
                custom_height: false,
                zero_height: true,
            },
            SheetFormat::default(),
        ];
        base.col_styles = vec![
            vec![
                ColStyle {
                    first: 4,
                    last: 6,
                    xf: 1,
                },
                ColStyle {
                    first: 0,
                    last: 7,
                    xf: 0,
                },
            ],
            Vec::new(),
        ];
        let range = CellRange {
            start: CellRef::parse_a1("$B$2").unwrap(),
            end: CellRef::parse_a1("C$4").unwrap(),
        };
        base.hyperlinks = vec![
            vec![Hyperlink {
                range,
                external_target: Some(String::new()),
                location: None,
                tooltip: Some("tip".into()),
                display: Some(String::new()),
            }],
            Vec::new(),
        ];
        base.charts = vec![
            [
                ChartAnchor::TwoCell {
                    from: AnchorCell {
                        col: 2,
                        col_off: i64::MIN,
                        row: 3,
                        row_off: i64::MAX,
                    },
                    to: AnchorCell::default(),
                    edit_as: AnchorEditAs::Absolute,
                },
                ChartAnchor::OneCell {
                    from: AnchorCell::default(),
                    extent: AnchorExtent { cx: -3, cy: 7 },
                },
                ChartAnchor::Absolute {
                    pos: AnchorPos { x: i64::MIN, y: 1 },
                    extent: AnchorExtent {
                        cx: 0,
                        cy: i64::MAX,
                    },
                },
            ]
            .into_iter()
            .map(|anchor| SheetChart {
                part: "xl/charts/chart1.xml".into(),
                drawing: String::new(),
                anchor_index: usize::MAX,
                anchor,
                refs: vec![
                    ChartRef {
                        kind: ChartRefKind::Title,
                        formula: String::new(),
                    },
                    ChartRef {
                        kind: ChartRefKind::Values,
                        formula: "Data!$B$2".into(),
                    },
                ],
            })
            .collect(),
            Vec::new(),
        ];
        base.hidden_dimensions = vec![HiddenDimensions {
            col_widths: BTreeMap::from([(0, -0.0), (4, 0.0)]),
            row_heights: BTreeMap::from([(2, f64::from_bits(0x7ff8_0000_0000_0001))]),
        }];
        base.shared_strings = vec!["second".into(), String::new(), "second".into()];
        base.styles.fonts = vec![
            Font {
                name: Some(String::new()),
                size_pt: Some(-0.0),
                bold: true,
                italic: false,
                underline: true,
                strike: false,
                color: Some(Color::Theme {
                    idx: 5,
                    tint: f64::from_bits(0x7ff8_0000_0000_0033),
                }),
            },
            Font::default(),
        ];
        base.styles.fills = vec![
            Fill::Solid(Color::Rgb(String::new())),
            Fill::None,
            Fill::Solid(Color::Indexed(0)),
            Fill::Solid(Color::Auto),
        ];
        base.styles.borders = vec![Border {
            left: Some(BorderEdge {
                style: BorderStyle::Double,
                color: None,
            }),
            right: None,
            top: Some(BorderEdge {
                style: BorderStyle::Hair,
                color: Some(Color::Auto),
            }),
            bottom: None,
        }];
        base.styles.cell_xfs = vec![
            Xf {
                font: Some(0),
                fill: None,
                border: Some(0),
                num_fmt_id: Some(0),
                alignment: Some(Alignment {
                    h: Some(HAlign::Distributed),
                    v: Some(VAlign::Top),
                    wrap_text: true,
                    shrink_to_fit: false,
                }),
            },
            Xf::default(),
        ];
        base.styles.num_fmts = vec![
            (u16::MAX, "0.0".into()),
            (0, String::new()),
            (0, "x".into()),
        ];
        base.styles.theme.colors = std::array::from_fn(|index| format!("color-{index}"));
        base.styles.indexed_colors = vec!["#112233".into(), String::new(), "#112233".into()];
        base.tables = vec![Table {
            name: String::new(),
            sheet: SheetId(0),
            range,
            header_rows: 0,
            totals_rows: u32::MAX,
            columns: vec!["second".into(), String::new(), "second".into()],
        }];
        let budget = SnapshotBudget::new(1, 128).unwrap();
        let mut encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        let header = header(&a, &encoder, budget);
        let mut hydrator = AuthorityHydrator::new(&header).unwrap();
        push_all(&mut encoder, &mut hydrator, budget);
        let restored = hydrator.finish().unwrap();
        assert_eq!(base_records(&a.base), base_records(&restored.base));
        assert_eq!(restored.base.bootstrap_client_id, 17);
        assert_eq!(restored.base.fingerprint, "kept-verbatim");
        assert_eq!(restored.base.fingerprints, a.base.fingerprints);
        assert_eq!(restored.base.shared_strings, a.base.shared_strings);
        assert_eq!(restored.base.styles.cell_xfs, a.base.styles.cell_xfs);
        assert_eq!(restored.base.styles.num_fmts, a.base.styles.num_fmts);
        assert_eq!(
            restored.base.styles.indexed_colors,
            a.base.styles.indexed_colors
        );
        assert_eq!(restored.base.tables, a.base.tables);
        assert_eq!(
            restored.base.styles.fonts[0].size_pt.unwrap().to_bits(),
            (-0.0_f64).to_bits(),
        );
    }

    #[test]
    fn hydrate_preserves_ids_clocks_guid_and_content() {
        let mut a = source();
        a.next_sheet_id = 123;
        let expected_update = a.encode_state_as_update_v1();
        let expected_vector = a.encode_state_vector_v1();
        for budget in [
            SnapshotBudget::new(100, usize::MAX).unwrap(),
            SnapshotBudget::new(2, 512).unwrap(),
        ] {
            let mut encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
            assert_eq!(encoder.split_fallback(), None);
            let header = header(&a, &encoder, budget);
            let mut hydrator = AuthorityHydrator::new(&header).unwrap();
            assert_eq!(hydrator.doc.client_id().get(), a.client_id());
            assert_eq!(hydrator.doc.guid(), a.doc.guid());
            assert!(hydrator.doc.transact().state_vector().is_empty());
            assert!(hydrator.advance(budget).unwrap().is_ready());
            let (base_count, yrs_count) = push_all(&mut encoder, &mut hydrator, budget);
            assert_eq!(base_count, header.chunk_count(ChunkKind::AuthorityBase));
            assert_eq!(yrs_count, header.chunk_count(ChunkKind::Yrs));
            if budget.max_bytes() == 512 {
                assert!(yrs_count > 1);
            }
            assert!(hydrator.advance(budget).unwrap().is_ready());
            let restored = hydrator.finish().unwrap();
            assert_eq!(restored.snapshot_identity(), a.snapshot_identity());
            assert_eq!(restored.encode_state_vector_v1(), expected_vector);
            assert_eq!(restored.encode_state_as_update_v1(), expected_update);
            assert_eq!(restored.materialize().unwrap(), a.materialize().unwrap());
            assert!(!restored.has_pending_updates());
        }
    }

    #[test]
    fn hydrate_has_no_pending_history_or_events() {
        let a = source();
        let budget = SnapshotBudget::new(3, 512).unwrap();
        let mut encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        let header = header(&a, &encoder, budget);
        let mut hydrator = AuthorityHydrator::new(&header).unwrap();
        let events = Arc::new(AtomicUsize::new(0));
        let observed = Arc::clone(&events);
        let subscription = hydrator
            .doc
            .observe_update_v1(move |_, _| {
                observed.fetch_add(1, Ordering::SeqCst);
            })
            .unwrap();
        push_all(&mut encoder, &mut hydrator, budget);
        assert!(hydrator.advance(budget).unwrap().is_ready());
        let restored = hydrator.finish().unwrap();
        assert!(!restored.has_pending_updates());
        assert!(restored.history.undo.is_empty());
        assert!(restored.history.redo.is_empty());
        assert!(restored.undo_stack.is_empty());
        assert!(restored.redo_stack.is_empty());
        assert!(!restored.can_undo());
        assert!(!restored.can_redo());
        assert_eq!(events.load(Ordering::SeqCst), 0);
        let map = restored.doc.get_or_insert_map("after-hydration");
        map.insert(&mut restored.doc.transact_mut(), "value", 1);
        assert_eq!(events.load(Ordering::SeqCst), 1);
        drop(subscription);
    }

    #[test]
    fn multi_key_map_fallback_hydrates_through_production_advance() {
        let a = source();
        let map = a.doc.get_or_insert_map("production-fallback");
        let value = Any::Map(Arc::new(HashMap::from([
            ("first".into(), Any::Number(1.25)),
            ("second".into(), Any::from("kept")),
        ])));
        map.insert(&mut a.doc.transact_mut(), "value", value.clone());
        let budget = SnapshotBudget::new(2, 16 * 1024).unwrap();
        let mut encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        assert_eq!(encoder.split_fallback().as_deref(), Some("multi_key_map"));
        let header = header(&a, &encoder, budget);
        assert_eq!(header.chunk_count(ChunkKind::Yrs), 1);
        let mut hydrator = AuthorityHydrator::new(&header).unwrap();
        while let Some(chunk) = encoder.next(budget).unwrap() {
            let (kind, _, payload) = unframe(&chunk).unwrap();
            match kind {
                ChunkKind::AuthorityBase => {
                    hydrator.push_base(payload).unwrap();
                    while !hydrator.advance_bounded(budget).unwrap().is_ready() {}
                }
                ChunkKind::Yrs => loop {
                    crate::snapshot::step::reset();
                    let progress = hydrator.advance_yrs(payload, budget).unwrap();
                    let work = crate::snapshot::step::current();
                    assert!(work.records <= budget.max_records(), "{work:?}");
                    assert!(work.bytes <= budget.max_bytes(), "{work:?}");
                    if progress.is_ready() {
                        break;
                    }
                },
                _ => panic!("unexpected authority chunk"),
            }
        }
        let restored = hydrator.finish().unwrap();
        assert_eq!(restored.snapshot_identity(), a.snapshot_identity());
        assert_eq!(
            restored.encode_state_vector_v1(),
            a.encode_state_vector_v1()
        );
        assert_eq!(restored.materialize().unwrap(), a.materialize().unwrap());
        let txn = restored.doc.transact();
        assert_eq!(
            txn.get_map("production-fallback")
                .unwrap()
                .get(&txn, "value"),
            Some(Out::Any(value)),
        );
        assert!(!restored.has_pending_updates());
    }

    #[test]
    fn refused_split_falls_back_whole_and_reports() {
        let a = source();
        let map = a.doc.get_or_insert_map("snapshot-fallback");
        let value = Any::Map(Arc::new(HashMap::from([
            ("second".into(), Any::Number(1.25)),
            ("first".into(), Any::from("kept")),
        ])));
        map.insert(&mut a.doc.transact_mut(), "value", value.clone());
        let budget = SnapshotBudget::new(2, 128).unwrap();
        let mut encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        assert_eq!(encoder.split_fallback().as_deref(), Some("multi_key_map"));
        let header = header(&a, &encoder, budget);
        assert_eq!(header.chunk_count(ChunkKind::Yrs), 1);
        let mut hydrator = AuthorityHydrator::new(&header).unwrap();
        let advance_budget = SnapshotBudget::new(budget.max_records(), 16 * 1024).unwrap();
        let mut ordinals = [0, 0];
        while let Some(chunk) = encoder.next(budget).unwrap() {
            let (kind, ordinal, payload) = unframe(&chunk).unwrap();
            match kind {
                ChunkKind::AuthorityBase => {
                    assert_eq!(ordinal, ordinals[0]);
                    ordinals[0] += 1;
                    assert!(payload.len() <= budget.max_bytes());
                    hydrator.push_base(payload).unwrap();
                    while !hydrator.advance_bounded(advance_budget).unwrap().is_ready() {}
                }
                ChunkKind::Yrs => {
                    assert_eq!(ordinal, ordinals[1]);
                    ordinals[1] += 1;
                    loop {
                        crate::snapshot::step::reset();
                        let progress = hydrator.advance_yrs(payload, advance_budget).unwrap();
                        let work = crate::snapshot::step::current();
                        assert!(work.records <= advance_budget.max_records());
                        assert!(work.bytes <= advance_budget.max_bytes());
                        if progress.is_ready() {
                            break;
                        }
                    }
                }
                _ => panic!("unexpected authority chunk"),
            }
        }
        let restored = hydrator.finish().unwrap();
        assert_eq!(restored.snapshot_identity(), a.snapshot_identity());
        assert_eq!(
            restored.encode_state_vector_v1(),
            a.encode_state_vector_v1()
        );
        let txn = restored.doc.transact();
        assert_eq!(
            txn.get_map("snapshot-fallback").unwrap().get(&txn, "value"),
            Some(Out::Any(value)),
        );
        assert!(!restored.has_pending_updates());
    }

    #[test]
    fn oversized_struct_falls_back_whole_and_reports() {
        let a = source();
        let budget = SnapshotBudget::new(1, 1).unwrap();
        let mut encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        assert_eq!(
            encoder.split_fallback().as_deref(),
            Some("oversized_struct")
        );
        let header = header(&a, &encoder, budget);
        assert_eq!(header.chunk_count(ChunkKind::Yrs), 1);
        let mut hydrator = AuthorityHydrator::new(&header).unwrap();
        push_all(&mut encoder, &mut hydrator, budget);
        let restored = hydrator.finish().unwrap();
        assert_eq!(
            restored.encode_state_as_update_v1(),
            a.encode_state_as_update_v1()
        );
    }

    #[test]
    fn oversized_inline_text_struct_is_isolated_in_bounded_hydration() {
        let mut model = WorkbookModel::default();
        let mut sheet = Sheet::new("Data");
        for row in 0..1_000 {
            sheet.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: if row == 500 {
                        CellValue::Text {
                            value: "x".repeat(20 * 1024),
                        }
                    } else {
                        CellValue::Number {
                            value: f64::from(row),
                        }
                    },
                    formula: None,
                    style: None,
                },
            );
        }
        model.sheets = vec![sheet];
        let authority = WorkbookAuthority::from_model(&model).unwrap();
        let budget = SnapshotBudget::new(7, 1_024).unwrap();
        let mut encoder = AuthoritySnapshotEncoder::new(&authority, budget).unwrap();
        assert_eq!(encoder.split_fallback(), None);
        let header = header(&authority, &encoder, budget);
        assert!(header.chunk_count(ChunkKind::Yrs) > 2);
        let mut hydrator = AuthorityHydrator::new(&header).unwrap();
        let mut oversized_chunks = 0;
        let mut oversized_steps = 0;
        let mut bounded_steps = 0;
        while let Some(chunk) = encoder.next(budget).unwrap() {
            let (kind, _, payload) = unframe(&chunk).unwrap();
            match kind {
                ChunkKind::AuthorityBase => {
                    hydrator.push_base(payload).unwrap();
                    while !hydrator.advance(budget).unwrap().is_ready() {}
                }
                ChunkKind::Yrs => {
                    let oversized = payload.len() > budget.max_bytes();
                    oversized_chunks += usize::from(oversized);
                    if oversized {
                        assert!(
                            payload
                                .windows(20 * 1024)
                                .any(|bytes| { bytes.iter().all(|byte| *byte == b'x') })
                        );
                    }
                    loop {
                        crate::snapshot::step::reset();
                        let progress = match hydrator.advance_yrs(payload, budget) {
                            Ok(progress) => {
                                let work = crate::snapshot::step::current();
                                assert!(work.records <= budget.max_records(), "{work:?}");
                                assert!(work.bytes <= budget.max_bytes(), "{work:?}");
                                bounded_steps += usize::from(work.records != 0);
                                progress
                            }
                            Err(failure) => {
                                assert!(oversized);
                                assert_eq!(
                                    failure.to_string(),
                                    "Yrs snapshot record exceeds advance byte budget"
                                );
                                assert_eq!(crate::snapshot::step::current().records, 0);
                                let larger = SnapshotBudget::new(1, payload.len()).unwrap();
                                let progress = hydrator.advance_yrs(payload, larger).unwrap();
                                let work = crate::snapshot::step::current();
                                assert_eq!(work.records, 1);
                                assert!(work.bytes > budget.max_bytes());
                                assert!(work.bytes <= larger.max_bytes());
                                oversized_steps += 1;
                                progress
                            }
                        };
                        if progress.is_ready() {
                            break;
                        }
                    }
                }
                _ => panic!("unexpected authority chunk"),
            }
        }
        assert_eq!(oversized_chunks, 1);
        assert_eq!(oversized_steps, 1);
        assert!(bounded_steps > 2);
        let restored = hydrator.finish().unwrap();
        assert_eq!(
            restored.encode_state_vector_v1(),
            authority.encode_state_vector_v1()
        );
        assert_eq!(
            restored.encode_state_as_update_v1(),
            authority.encode_state_as_update_v1()
        );
        assert_eq!(restored.materialize().unwrap(), model);
        assert!(!restored.has_pending_updates());
    }

    #[test]
    fn hydration_rejects_missing_chunks_and_mismatched_state_vectors() {
        let a = source();
        let budget = SnapshotBudget::new(1, 512).unwrap();
        let encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        let expected = header(&a, &encoder, budget);
        assert!(AuthorityHydrator::new(&expected).unwrap().finish().is_err());

        let mut encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        let mut wrong_vector = expected;
        wrong_vector.state_vector = StateVector::default().encode_v1();
        let mut hydrator = AuthorityHydrator::new(&wrong_vector).unwrap();
        push_all(&mut encoder, &mut hydrator, budget);
        assert!(hydrator.finish().is_err());
    }

    #[test]
    fn base_hydration_advances_in_record_budget() {
        let a = source();
        let budget = SnapshotBudget::new(100, usize::MAX).unwrap();
        let mut encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
        let header = header(&a, &encoder, budget);
        let mut hydrator = AuthorityHydrator::new(&header).unwrap();
        while let Some(chunk) = encoder.next(budget).unwrap() {
            let (kind, _, payload) = unframe(&chunk).unwrap();
            match kind {
                ChunkKind::AuthorityBase => hydrator.push_base(payload).unwrap(),
                ChunkKind::Yrs => hydrator.push_yrs(payload).unwrap(),
                _ => panic!("unexpected authority chunk"),
            }
        }
        assert!(hydrator.base.base.is_none());
        let step = SnapshotBudget::new(1, usize::MAX).unwrap();
        let queued = hydrator.pending_base.len();
        assert!(!hydrator.advance(step).unwrap().is_ready());
        assert_eq!(hydrator.pending_base.len(), queued - 1);
        assert!(hydrator.base.base.is_some());
        while !hydrator.advance(step).unwrap().is_ready() {}
        let restored = hydrator.finish().unwrap();
        assert_eq!(
            restored.encode_state_as_update_v1(),
            a.encode_state_as_update_v1()
        );
    }

    #[test]
    fn hydration_rejects_pending_structs_and_delete_sets() {
        let foreign = Doc::with_client_id(99);
        let root = foreign.get_or_insert_map("foreign");
        let map = root.insert(&mut foreign.transact_mut(), "nested", MapPrelim::default());
        map.insert(&mut foreign.transact_mut(), "first", "kept");
        let before_second = foreign.transact().state_vector();
        map.insert(&mut foreign.transact_mut(), "second", "pending");
        let structs = foreign.transact().encode_diff_v1(&before_second);
        let before_delete = foreign.transact().state_vector();
        map.remove(&mut foreign.transact_mut(), "first");
        let deletes = foreign.transact().encode_diff_v1(&before_delete);
        for (update, expect_structs) in [(structs, true), (deletes, false)] {
            let a = source();
            let budget = SnapshotBudget::new(100, usize::MAX).unwrap();
            let mut encoder = AuthoritySnapshotEncoder::new(&a, budget).unwrap();
            let mut header = header(&a, &encoder, budget);
            assert_eq!(header.chunk_count(ChunkKind::Yrs), 1);
            header.state_vector = StateVector::default().encode_v1();
            let mut hydrator = AuthorityHydrator::new(&header).unwrap();
            while let Some(chunk) = encoder.next(budget).unwrap() {
                let (kind, _, payload) = unframe(&chunk).unwrap();
                if kind == ChunkKind::AuthorityBase {
                    hydrator.push_base(payload).unwrap();
                }
            }
            while !hydrator.advance(budget).unwrap().is_ready() {}
            hydrator.push_yrs(&update).unwrap();
            {
                let txn = hydrator.doc.transact();
                if expect_structs {
                    assert!(txn.store().pending_update().is_some());
                    assert!(txn.store().pending_ds().is_none());
                } else {
                    assert!(txn.store().pending_ds().is_some());
                    assert!(txn.store().pending_update().is_none());
                }
                assert!(txn.state_vector().is_empty());
            }
            assert!(hydrator.advance(budget).unwrap().is_ready());
            assert_eq!(
                hydrator.finish().err().unwrap().to_string(),
                "authority snapshot has pending Yrs state"
            );
        }
    }
}
