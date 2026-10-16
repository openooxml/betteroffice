use std::mem::size_of;

use xlsx_model::CellRef;

use crate::package_facts::{
    PackageFacts, ReferenceAreaFacts, ReferenceFacts, SheetFacts, any_uncached_formula,
    chart_part_indices, sheet_facts,
};
use crate::{ParseError, PreservedPackage, SheetVisibility, SourceSheetKind};

#[doc(hidden)]
pub const SNAPSHOT_RECORD_MAX_BYTES: usize = 64 * 1024 * 1024;

#[derive(Default)]
#[doc(hidden)]
pub struct PackageFactsEncoder {
    chart_parts: Option<Vec<usize>>,
    record_index: usize,
    pending: Vec<u8>,
    offset: usize,
    preflight_done: bool,
}

impl PackageFactsEncoder {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn next(
        &mut self,
        package: &PreservedPackage,
        max_bytes: usize,
    ) -> Result<Option<Vec<u8>>, ParseError> {
        if max_bytes == 0 {
            return Err(malformed("facts byte budget must be positive"));
        }
        let charts = self
            .chart_parts
            .get_or_insert_with(|| chart_part_indices(package));
        if !self.preflight_done {
            for index in 0..package.source_sheet_count() {
                let mut writer = Writer::default();
                writer.uint(1);
                encode_sheet(&mut writer, &sheet_facts(package, index));
                check_record_length(writer.0.len())?;
            }
            for reference in package.unpatchable_references() {
                let mut writer = Writer::default();
                writer.uint(2);
                encode_reference(&mut writer, &reference.package_facts());
                check_record_length(writer.0.len())?;
            }
            for &index in charts.iter() {
                let (path, bytes) = &package.parts[index];
                check_record_length(chart_record_length(path.len(), bytes.len()))?;
            }
            self.preflight_done = true;
        }
        if self.offset == self.pending.len() {
            let sheets = package.source_sheet_count();
            let references = package.unpatchable_references();
            let mut writer = Writer::default();
            match self.record_index {
                0 => {
                    writer.uint(0);
                    writer.uint(1);
                    writer.bool(any_uncached_formula(package));
                    writer.usize(sheets);
                    writer.usize(references.len());
                    writer.usize(charts.len());
                }
                index if index <= sheets => {
                    writer.uint(1);
                    encode_sheet(&mut writer, &sheet_facts(package, index - 1));
                }
                index if index <= sheets + references.len() => {
                    writer.uint(2);
                    encode_reference(&mut writer, &references[index - sheets - 1].package_facts());
                }
                index if index <= sheets + references.len() + charts.len() => {
                    writer.uint(3);
                    let chart = charts[index - sheets - references.len() - 1];
                    let (path, bytes) = &package.parts[chart];
                    writer.string(path);
                    writer.bytes(bytes);
                }
                _ => return Ok(None),
            }
            self.set_record(writer)?;
        }
        Ok(Some(self.take_pending(max_bytes)))
    }

    #[doc(hidden)]
    pub fn next_from_facts(
        &mut self,
        facts: &PackageFacts,
        max_bytes: usize,
    ) -> Result<Option<Vec<u8>>, ParseError> {
        if max_bytes == 0 {
            return Err(malformed("facts byte budget must be positive"));
        }
        if !self.preflight_done {
            for sheet in &facts.sheets {
                let mut writer = Writer::default();
                writer.uint(1);
                encode_sheet(&mut writer, sheet);
                check_record_length(writer.0.len())?;
            }
            for reference in &facts.references {
                let mut writer = Writer::default();
                writer.uint(2);
                encode_reference(&mut writer, reference);
                check_record_length(writer.0.len())?;
            }
            for (path, bytes) in &facts.charts {
                check_record_length(chart_record_length(path.len(), bytes.len()))?;
            }
            self.preflight_done = true;
        }
        if self.offset == self.pending.len() {
            let sheets = facts.sheets.len();
            let references = facts.references.len();
            let mut writer = Writer::default();
            match self.record_index {
                0 => {
                    writer.uint(0);
                    writer.uint(1);
                    writer.bool(facts.any_uncached_formula);
                    writer.usize(sheets);
                    writer.usize(references);
                    writer.usize(facts.charts.len());
                }
                index if index <= sheets => {
                    writer.uint(1);
                    encode_sheet(&mut writer, &facts.sheets[index - 1]);
                }
                index if index <= sheets + references => {
                    writer.uint(2);
                    encode_reference(&mut writer, &facts.references[index - sheets - 1]);
                }
                index if index <= sheets + references + facts.charts.len() => {
                    writer.uint(3);
                    let (path, bytes) = &facts.charts[index - sheets - references - 1];
                    writer.string(path);
                    writer.bytes(bytes);
                }
                _ => return Ok(None),
            }
            self.set_record(writer)?;
        }
        Ok(Some(self.take_pending(max_bytes)))
    }

    fn set_record(&mut self, writer: Writer) -> Result<(), ParseError> {
        check_record_length(writer.0.len())?;
        let mut framed = Writer::default();
        framed.bytes(&writer.0);
        self.pending = framed.0;
        self.offset = 0;
        self.record_index += 1;
        Ok(())
    }

    fn take_pending(&mut self, max_bytes: usize) -> Vec<u8> {
        let end = self.offset + max_bytes.min(self.pending.len() - self.offset);
        let payload = self.pending[self.offset..end].to_vec();
        self.offset = end;
        payload
    }
}

fn chart_record_length(path: usize, bytes: usize) -> usize {
    let var_length = |mut value: usize| {
        let mut length = 1;
        while value >= 128 {
            value >>= 7;
            length += 1;
        }
        length
    };
    1usize
        .saturating_add(var_length(path))
        .saturating_add(path)
        .saturating_add(var_length(bytes))
        .saturating_add(bytes)
}

fn check_record_length(length: usize) -> Result<(), ParseError> {
    if length > SNAPSHOT_RECORD_MAX_BYTES {
        Err(malformed("facts record exceeds byte limit"))
    } else {
        Ok(())
    }
}

#[doc(hidden)]
pub struct PackageFactsBuilder {
    facts: PackageFacts,
    expected: Option<(usize, usize, usize)>,
    length: u64,
    shift: u32,
    record_length: Option<usize>,
    record: Vec<u8>,
    failed: bool,
    sheets_growth: VectorGrowth<SheetFacts>,
    references_growth: VectorGrowth<ReferenceFacts>,
    charts_growth: VectorGrowth<(String, Vec<u8>)>,
}

impl Default for PackageFactsBuilder {
    fn default() -> Self {
        Self {
            facts: PackageFacts {
                sheets: Vec::new(),
                references: Vec::new(),
                charts: Vec::new(),
                any_uncached_formula: false,
            },
            expected: None,
            length: 0,
            shift: 0,
            record_length: None,
            record: Vec::new(),
            failed: false,
            sheets_growth: VectorGrowth::default(),
            references_growth: VectorGrowth::default(),
            charts_growth: VectorGrowth::default(),
        }
    }
}

struct VectorGrowth<T> {
    pending: Option<(std::vec::IntoIter<T>, Vec<T>)>,
}

impl<T> Default for VectorGrowth<T> {
    fn default() -> Self {
        Self { pending: None }
    }
}

impl<T> VectorGrowth<T> {
    fn ensure(
        &mut self,
        target: &mut Vec<T>,
        max_records: usize,
        max_bytes: usize,
    ) -> Result<Option<(usize, usize)>, ParseError> {
        if size_of::<T>() > max_bytes || max_records == 0 {
            return Err(malformed("facts storage exceeds advance budget"));
        }
        if self.pending.is_none() {
            if target.len() < target.capacity() {
                return Ok(None);
            }
            let mut storage = Vec::new();
            storage
                .try_reserve_exact(target.len().saturating_mul(2).max(1))
                .map_err(|_| malformed("cannot allocate facts storage"))?;
            if target.is_empty() {
                *target = storage;
                return Ok(None);
            }
            self.pending = Some((std::mem::take(target).into_iter(), storage));
        }
        let (entries, storage) = self.pending.as_mut().unwrap();
        let count = entries
            .len()
            .min(max_records)
            .min(max_bytes / size_of::<T>().max(1));
        storage.extend(entries.by_ref().take(count));
        if entries.len() == 0 {
            *target = self.pending.take().unwrap().1;
        }
        Ok(Some((count, count * size_of::<T>())))
    }
}

impl PackageFactsBuilder {
    pub fn new() -> Self {
        Self::default()
    }

    #[doc(hidden)]
    pub fn advance_capacity(
        &mut self,
        max_records: usize,
        max_bytes: usize,
    ) -> Result<(bool, usize, usize), ParseError> {
        if let Some((records, bytes)) =
            self.sheets_growth
                .ensure(&mut self.facts.sheets, max_records, max_bytes)?
        {
            return Ok((false, records, bytes));
        }
        if let Some((records, bytes)) =
            self.references_growth
                .ensure(&mut self.facts.references, max_records, max_bytes)?
        {
            return Ok((false, records, bytes));
        }
        if let Some((records, bytes)) =
            self.charts_growth
                .ensure(&mut self.facts.charts, max_records, max_bytes)?
        {
            return Ok((false, records, bytes));
        }
        Ok((true, 0, 0))
    }

    pub fn push(&mut self, payload: &[u8]) -> Result<(), ParseError> {
        if self.failed {
            return Err(malformed("facts builder has failed"));
        }
        let result = self.push_inner(payload);
        self.failed = result.is_err();
        result
    }

    fn push_inner(&mut self, mut payload: &[u8]) -> Result<(), ParseError> {
        while !payload.is_empty() {
            if let Some(length) = self.record_length {
                let count = payload.len().min(length - self.record.len());
                self.record
                    .try_reserve(count)
                    .map_err(|_| malformed("cannot allocate facts record"))?;
                self.record.extend_from_slice(&payload[..count]);
                payload = &payload[count..];
                if self.record.len() == length {
                    self.accept_record()?;
                    self.record.clear();
                    self.record_length = None;
                }
            } else {
                let byte = payload[0];
                payload = &payload[1..];
                if self.shift == 63 && byte > 1 {
                    return Err(malformed("facts length overflows"));
                }
                self.length |= u64::from(byte & 0x7f) << self.shift;
                if byte & 0x80 == 0 {
                    let length = usize::try_from(self.length)
                        .map_err(|_| malformed("facts length overflows"))?;
                    if length == 0 || length > SNAPSHOT_RECORD_MAX_BYTES {
                        return Err(malformed("invalid facts record length"));
                    }
                    self.record_length = Some(length);
                    self.length = 0;
                    self.shift = 0;
                } else {
                    self.shift += 7;
                }
            }
        }
        Ok(())
    }

    fn accept_record(&mut self) -> Result<(), ParseError> {
        let mut reader = Reader::new(&self.record);
        let tag = reader.uint()?;
        match self.expected {
            None => {
                if tag != 0 || reader.uint()? != 1 {
                    return Err(malformed("invalid facts header"));
                }
                self.facts.any_uncached_formula = reader.bool()?;
                self.expected = Some((reader.usize()?, reader.usize()?, reader.usize()?));
            }
            Some((sheets, _, _)) if self.facts.sheets.len() < sheets => {
                if tag != 1 {
                    return Err(malformed("expected facts sheet"));
                }
                self.facts.sheets.push(decode_sheet(&mut reader)?);
            }
            Some((_, references, _)) if self.facts.references.len() < references => {
                if tag != 2 {
                    return Err(malformed("expected facts reference"));
                }
                self.facts.references.push(decode_reference(&mut reader)?);
            }
            Some((_, _, charts)) if self.facts.charts.len() < charts => {
                if tag != 3 {
                    return Err(malformed("expected facts chart"));
                }
                self.facts
                    .charts
                    .push((reader.string()?, reader.bytes()?.to_vec()));
            }
            Some(_) => return Err(malformed("unexpected facts record")),
        }
        reader.finish()
    }

    pub fn finish(self) -> Result<PackageFacts, ParseError> {
        let PackageFacts {
            sheets,
            references,
            charts,
            any_uncached_formula,
        } = self.facts;
        if self.failed
            || self.record_length.is_some()
            || self.shift != 0
            || self.sheets_growth.pending.is_some()
            || self.references_growth.pending.is_some()
            || self.charts_growth.pending.is_some()
            || self.expected != Some((sheets.len(), references.len(), charts.len()))
        {
            return Err(malformed("incomplete facts"));
        }
        Ok(PackageFacts {
            sheets,
            references,
            charts,
            any_uncached_formula,
        })
    }
}

fn encode_sheet(writer: &mut Writer, sheet: &SheetFacts) {
    let SheetFacts {
        path,
        kind,
        visibility,
        worksheet,
        protected,
    } = sheet;
    writer.string(path);
    writer.uint(match kind {
        SourceSheetKind::Worksheet => 0,
        SourceSheetKind::Chartsheet => 1,
        SourceSheetKind::Dialogsheet => 2,
        SourceSheetKind::Macrosheet => 3,
        SourceSheetKind::Other => 4,
    });
    writer.uint(match visibility {
        SheetVisibility::Visible => 0,
        SheetVisibility::Hidden => 1,
        SheetVisibility::VeryHidden => 2,
        SheetVisibility::Unknown => 3,
    });
    writer.bool(*worksheet);
    writer.bool(*protected);
}

fn decode_sheet(reader: &mut Reader<'_>) -> Result<SheetFacts, ParseError> {
    Ok(SheetFacts {
        path: reader.string()?,
        kind: match reader.uint()? {
            0 => SourceSheetKind::Worksheet,
            1 => SourceSheetKind::Chartsheet,
            2 => SourceSheetKind::Dialogsheet,
            3 => SourceSheetKind::Macrosheet,
            4 => SourceSheetKind::Other,
            _ => return Err(malformed("invalid facts sheet kind")),
        },
        visibility: match reader.uint()? {
            0 => SheetVisibility::Visible,
            1 => SheetVisibility::Hidden,
            2 => SheetVisibility::VeryHidden,
            3 => SheetVisibility::Unknown,
            _ => return Err(malformed("invalid facts visibility")),
        },
        worksheet: reader.bool()?,
        protected: reader.bool()?,
    })
}

fn encode_reference(writer: &mut Writer, reference: &ReferenceFacts) {
    let ReferenceFacts { part, areas } = reference;
    writer.string(part);
    writer.bool(areas.is_some());
    if let Some(areas) = areas {
        writer.usize(areas.len());
        for area in areas {
            let ReferenceAreaFacts { sheet, end } = area;
            let CellRef {
                row,
                col,
                abs_row,
                abs_col,
            } = end;
            writer.string(sheet);
            writer.uint(u64::from(*row));
            writer.uint(u64::from(*col));
            writer.bool(*abs_row);
            writer.bool(*abs_col);
        }
    }
}

fn decode_reference(reader: &mut Reader<'_>) -> Result<ReferenceFacts, ParseError> {
    let part = reader.string()?;
    let areas = if reader.bool()? {
        let count = reader.usize()?;
        if count > reader.remaining.len() / 5 {
            return Err(malformed("invalid facts area count"));
        }
        let mut areas = Vec::new();
        for _ in 0..count {
            areas.push(ReferenceAreaFacts {
                sheet: reader.string()?,
                end: CellRef {
                    row: reader.u32()?,
                    col: reader.u32()?,
                    abs_row: reader.bool()?,
                    abs_col: reader.bool()?,
                },
            });
        }
        Some(areas)
    } else {
        None
    };
    Ok(ReferenceFacts { part, areas })
}

#[derive(Default)]
struct Writer(Vec<u8>);

impl Writer {
    fn uint(&mut self, mut value: u64) {
        while value >= 0x80 {
            self.0.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        self.0.push(value as u8);
    }

    fn usize(&mut self, value: usize) {
        self.uint(value as u64);
    }

    fn bool(&mut self, value: bool) {
        self.uint(u64::from(value));
    }

    fn bytes(&mut self, value: &[u8]) {
        self.usize(value.len());
        self.0.extend_from_slice(value);
    }

    fn string(&mut self, value: &str) {
        self.bytes(value.as_bytes());
    }
}

struct Reader<'a> {
    remaining: &'a [u8],
}

impl<'a> Reader<'a> {
    fn new(remaining: &'a [u8]) -> Self {
        Self { remaining }
    }

    fn uint(&mut self) -> Result<u64, ParseError> {
        let mut value = 0;
        for shift in (0..=63).step_by(7) {
            let (&byte, rest) = self
                .remaining
                .split_first()
                .ok_or_else(|| malformed("short facts integer"))?;
            self.remaining = rest;
            if shift == 63 && byte > 1 {
                return Err(malformed("facts integer overflows"));
            }
            value |= u64::from(byte & 0x7f) << shift;
            if byte & 0x80 == 0 {
                return Ok(value);
            }
        }
        Err(malformed("facts integer overflows"))
    }

    fn usize(&mut self) -> Result<usize, ParseError> {
        usize::try_from(self.uint()?).map_err(|_| malformed("facts integer overflows"))
    }

    fn u32(&mut self) -> Result<u32, ParseError> {
        u32::try_from(self.uint()?).map_err(|_| malformed("facts coordinate overflows"))
    }

    fn bool(&mut self) -> Result<bool, ParseError> {
        match self.uint()? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(malformed("invalid facts boolean")),
        }
    }

    fn bytes(&mut self) -> Result<&'a [u8], ParseError> {
        let length = self.usize()?;
        if length > self.remaining.len() {
            return Err(malformed("short facts bytes"));
        }
        let (bytes, rest) = self.remaining.split_at(length);
        self.remaining = rest;
        Ok(bytes)
    }

    fn string(&mut self) -> Result<String, ParseError> {
        String::from_utf8(self.bytes()?.to_vec()).map_err(|_| malformed("invalid facts UTF-8"))
    }

    fn finish(self) -> Result<(), ParseError> {
        if self.remaining.is_empty() {
            Ok(())
        } else {
            Err(malformed("trailing facts bytes"))
        }
    }
}

fn malformed(message: &str) -> ParseError {
    ParseError::Malformed(message.to_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn encode_facts(facts: &PackageFacts) -> Vec<u8> {
        let PackageFacts {
            sheets,
            references,
            charts,
            any_uncached_formula,
        } = facts;
        let mut framed = Writer::default();
        let mut header = Writer::default();
        header.uint(0);
        header.uint(1);
        header.bool(*any_uncached_formula);
        header.usize(sheets.len());
        header.usize(references.len());
        header.usize(charts.len());
        framed.bytes(&header.0);
        for sheet in sheets {
            let mut record = Writer::default();
            record.uint(1);
            encode_sheet(&mut record, sheet);
            framed.bytes(&record.0);
        }
        for reference in references {
            let mut record = Writer::default();
            record.uint(2);
            encode_reference(&mut record, reference);
            framed.bytes(&record.0);
        }
        for (path, bytes) in charts {
            let mut record = Writer::default();
            record.uint(3);
            record.string(path);
            record.bytes(bytes);
            framed.bytes(&record.0);
        }
        framed.0
    }

    #[test]
    fn reference_facts_growth_migrates_within_each_advance() {
        let facts = PackageFacts {
            sheets: Vec::new(),
            references: (0..4_097)
                .map(|index| ReferenceFacts {
                    part: format!("part{index}"),
                    areas: None,
                })
                .collect(),
            charts: Vec::new(),
            any_uncached_formula: false,
        };
        let mut encoder = PackageFactsEncoder::new();
        let mut builder = PackageFactsBuilder::new();
        let mut migrated = 0;
        while let Some(payload) = encoder.next_from_facts(&facts, 256).unwrap() {
            loop {
                let (ready, records, bytes) = builder.advance_capacity(1, 256).unwrap();
                assert!(records <= 1);
                assert!(bytes <= 256);
                migrated += records;
                if ready {
                    break;
                }
            }
            builder.push(&payload).unwrap();
        }
        assert!(migrated >= 4_096);
        assert_eq!(builder.finish().unwrap(), facts);
    }

    #[test]
    fn facts_roundtrip_preserves_options_and_order() {
        let package = crate::package_facts::tests::package(true, true, true, true);
        let expected = PackageFacts::from_package(&package);
        let canonical = encode_facts(&expected);
        for max_bytes in [1, 2, 7, 128, usize::MAX] {
            let mut encoder = PackageFactsEncoder::new();
            let mut builder = PackageFactsBuilder::new();
            let mut encoded = Vec::new();
            let mut chunks = 0;
            while let Some(payload) = encoder.next(&package, max_bytes).unwrap() {
                assert!(!payload.is_empty() && payload.len() <= max_bytes);
                encoded.extend_from_slice(&payload);
                builder.push(&payload).unwrap();
                chunks += 1;
            }
            assert!(chunks > 1);
            assert_eq!(encoded, canonical);
            let actual = builder.finish().unwrap();
            assert_eq!(actual, expected);
            assert_eq!(encode_facts(&actual), canonical);
            assert!(encoder.next(&package, max_bytes).unwrap().is_none());
        }

        let mut expected = expected;
        expected.references = vec![
            ReferenceFacts {
                part: String::new(),
                areas: Some(Vec::new()),
            },
            ReferenceFacts {
                part: "first".to_owned(),
                areas: Some(vec![
                    ReferenceAreaFacts {
                        sheet: "dAtA".to_owned(),
                        end: CellRef {
                            row: 0,
                            col: 0,
                            abs_row: true,
                            abs_col: false,
                        },
                    },
                    ReferenceAreaFacts {
                        sheet: String::new(),
                        end: CellRef {
                            row: u32::MAX,
                            col: u32::MAX,
                            abs_row: false,
                            abs_col: true,
                        },
                    },
                ]),
            },
            ReferenceFacts {
                part: "blanket".to_owned(),
                areas: None,
            },
            ReferenceFacts {
                part: "last".to_owned(),
                areas: Some(Vec::new()),
            },
        ];
        expected.charts = vec![
            (String::new(), Vec::new()),
            ("same".to_owned(), vec![0, 255, 0]),
            ("same".to_owned(), Vec::new()),
        ];
        expected.sheets.extend([
            SheetFacts {
                path: String::new(),
                kind: SourceSheetKind::Dialogsheet,
                visibility: SheetVisibility::VeryHidden,
                worksheet: false,
                protected: false,
            },
            SheetFacts {
                path: "macro".to_owned(),
                kind: SourceSheetKind::Macrosheet,
                visibility: SheetVisibility::Unknown,
                worksheet: false,
                protected: true,
            },
            SheetFacts {
                path: "other".to_owned(),
                kind: SourceSheetKind::Other,
                visibility: SheetVisibility::Visible,
                worksheet: true,
                protected: false,
            },
        ]);
        let bytes = encode_facts(&expected);
        let mut builder = PackageFactsBuilder::new();
        for byte in &bytes {
            builder.push(std::slice::from_ref(byte)).unwrap();
        }
        let actual = builder.finish().unwrap();
        assert_eq!(actual, expected);
        assert_eq!(encode_facts(&actual), bytes);
        let view = crate::PackageFactsView::from_facts(&actual);
        assert_eq!(view.unpatchable_reference_part(), Some(""));
        assert_eq!(view.reference_naming_sheet("DATA"), Some("first"));
        assert_eq!(view.reference_moved_by_rows("Data", 0), Some("first"));
        assert_eq!(view.reference_moved_by_rows("Data", 1), Some("blanket"));
        assert_eq!(view.reference_moved_by_cols("Data", 0), Some("first"));
        assert_eq!(view.reference_moved_by_cols("Data", 1), Some("blanket"));
        assert_eq!(view.chart_part_bytes("same"), Some([0, 255, 0].as_slice()));
    }

    #[test]
    fn facts_reject_incomplete_invalid_and_trailing_payloads() {
        let package = crate::package_facts::tests::package(false, true, false, true);
        let bytes = encode_facts(&PackageFacts::from_package(&package));
        for end in 0..bytes.len() {
            let mut builder = PackageFactsBuilder::new();
            builder.push(&bytes[..end]).unwrap();
            assert!(builder.finish().is_err());
        }
        let mut builder = PackageFactsBuilder::new();
        builder.push(&bytes).unwrap();
        assert!(builder.push(&[1, 0]).is_err());
        assert!(builder.finish().is_err());
        for bytes in [
            vec![0],
            vec![0xff; 10],
            vec![6, 0, 1, 2, 0, 0, 0],
            vec![7, 0, 1, 0, 0, 0, 0, 0],
        ] {
            let mut builder = PackageFactsBuilder::new();
            assert!(builder.push(&bytes).is_err());
            assert!(builder.finish().is_err());
        }
        assert!(PackageFactsEncoder::new().next(&package, 0).is_err());
    }
}

#[cfg(test)]
mod ceiling_tests {
    use super::*;

    #[test]
    fn snapshot_oversized_facts_record_refuses_capture_before_chunks() {
        let facts = PackageFacts {
            sheets: Vec::new(),
            references: Vec::new(),
            any_uncached_formula: false,
            charts: vec![(
                "xl/charts/chart1.xml".to_owned(),
                vec![0; SNAPSHOT_RECORD_MAX_BYTES + 1],
            )],
        };
        let mut encoder = PackageFactsEncoder::new();
        assert!(encoder.next_from_facts(&facts, 16_384).is_err());
        assert_eq!(encoder.record_index, 0);
        assert!(encoder.pending.is_empty());
        assert_eq!(encoder.offset, 0);
    }
}
