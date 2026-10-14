use std::collections::VecDeque;
use std::ops::Bound::{Excluded, Unbounded};
use std::sync::Arc;

use xlsx_parse::{SharedStringCells, SheetAxes};

use crate::workbook::PreservedSheetState;

use super::growth::Growth;
use super::wire::{ChunkKind, Reader, Writer, frame, frame_records, packed_record, unframe};
use super::{SnapshotBudget, SnapshotError, SnapshotProgress, SnapshotResult};

const HEADER: u8 = 0;
const ORIGIN: u8 = 1;
const STRINGS: u8 = 2;
const STRING_CELL: u8 = 3;
const AXES: u8 = 4;
const CREATED: u8 = 5;

pub(crate) struct PreservedSnapshotEncoder {
    ordinal: u64,
    started: bool,
    cursor: Cursor,
}

impl PreservedSnapshotEncoder {
    pub(crate) fn new() -> Self {
        Self {
            ordinal: 0,
            started: false,
            cursor: Cursor {
                phase: ORIGIN,
                index: 0,
                strings_started: false,
                after: None,
            },
        }
    }

    pub(crate) fn next(
        &mut self,
        state: &PreservedSheetState,
        budget: SnapshotBudget,
    ) -> SnapshotResult<Option<Vec<u8>>> {
        let mut payload = Writer::new();
        let mut count = 0;
        let mut lengths = Vec::new();
        if !self.started {
            write_header(&mut payload, state)?;
            if frame(ChunkKind::Preserved, self.ordinal, &[]).len() + payload.len()
                > budget.max_bytes()
            {
                return Err(SnapshotError::new(
                    "snapshot preservation header exceeds byte budget",
                ));
            }
            self.started = true;
            count = 1;
            lengths.push(payload.len());
        }
        while count < budget.max_records().min(64) {
            let mut cursor = self.cursor;
            let mut record = Writer::new();
            if !cursor.write_next(&mut record, state)? {
                self.cursor = cursor;
                break;
            }
            lengths.push(record.len());
            let length = frame_records(ChunkKind::Preserved, self.ordinal, &[], &lengths).len()
                + payload.len()
                + record.len();
            if length > budget.max_bytes().min(1024) && count != 0 {
                lengths.pop();
                break;
            }
            if length > budget.max_bytes() {
                if count != 0 {
                    break;
                }
                return Err(SnapshotError::new(
                    "snapshot preservation record exceeds byte budget",
                ));
            }
            payload.raw(&record.into_bytes());
            self.cursor = cursor;
            count += 1;
        }
        if payload.is_empty() {
            return Ok(None);
        }
        let chunk = frame_records(
            ChunkKind::Preserved,
            self.ordinal,
            &payload.into_bytes(),
            &lengths,
        );
        self.ordinal += 1;
        Ok(Some(chunk))
    }
}

fn write_header(w: &mut Writer, state: &PreservedSheetState) -> SnapshotResult<()> {
    let PreservedSheetState {
        origins,
        shared_string_cells,
        axes,
        created,
    } = state;
    let counts = [
        origins.len(),
        shared_string_cells.len(),
        axes.len(),
        created.len(),
    ];
    let mut total = 1usize;
    for count in counts
        .into_iter()
        .chain(shared_string_cells.iter().map(SharedStringCells::len))
    {
        total = total
            .checked_add(count)
            .ok_or_else(|| SnapshotError::new("snapshot preservation count overflows usize"))?;
    }
    w.u8(HEADER);
    w.var_usize(total);
    for count in counts {
        w.var_usize(count);
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct Cursor {
    phase: u8,
    index: usize,
    strings_started: bool,
    after: Option<(u32, u32)>,
}

impl Cursor {
    fn write_next(&mut self, w: &mut Writer, state: &PreservedSheetState) -> SnapshotResult<bool> {
        let PreservedSheetState {
            origins,
            shared_string_cells,
            axes,
            created,
        } = state;
        loop {
            let length = match self.phase {
                ORIGIN => origins.len(),
                STRINGS => shared_string_cells.len(),
                AXES => axes.len(),
                CREATED => created.len(),
                _ => return Ok(false),
            };
            if self.index == length {
                self.phase = if self.phase == STRINGS {
                    AXES
                } else {
                    self.phase + 1
                };
                self.index = 0;
                continue;
            }
            match self.phase {
                ORIGIN => {
                    w.u8(ORIGIN);
                    w.option(origins[self.index], Writer::var_usize);
                }
                STRINGS => {
                    let cells = &shared_string_cells[self.index];
                    if !self.strings_started {
                        w.u8(STRINGS);
                        w.var_usize(cells.len());
                        self.strings_started = true;
                        self.after = None;
                        return Ok(true);
                    }
                    let lower = self.after.map_or(Unbounded, Excluded);
                    if let Some((&(row, col), &index)) = cells.range((lower, Unbounded)).next() {
                        w.u8(STRING_CELL);
                        w.var_u32(row);
                        w.var_u32(col);
                        w.var_usize(index);
                        self.after = Some((row, col));
                        return Ok(true);
                    }
                    self.strings_started = false;
                    self.after = None;
                    self.index += 1;
                    continue;
                }
                AXES => {
                    w.u8(AXES);
                    match &axes[self.index] {
                        None => w.u8(0),
                        Some(value) => {
                            let SheetAxes { rows, cols } = value;
                            if !rows.is_identity() || !cols.is_identity() {
                                return Err(SnapshotError::new(
                                    "snapshot preservation axes are not identity",
                                ));
                            }
                            if value != &SheetAxes::default() {
                                return Err(SnapshotError::new(
                                    "snapshot preservation axes have noninitial limits",
                                ));
                            }
                            w.u8(1);
                        }
                    }
                }
                CREATED => {
                    w.u8(CREATED);
                    w.bool(created[self.index]);
                }
                _ => unreachable!(),
            }
            self.index += 1;
            return Ok(true);
        }
    }
}

pub(crate) struct PreservedSnapshotBuilder {
    state: PreservedSheetState,
    ordinal: u64,
    started: bool,
    failed: bool,
    remaining: usize,
    runs: VecDeque<(u8, usize)>,
    growth: Growth<PreservedSheetState>,
    offset: usize,
}

impl PreservedSnapshotBuilder {
    pub(crate) fn new() -> Self {
        Self {
            state: PreservedSheetState::default(),
            ordinal: 0,
            started: false,
            failed: false,
            remaining: 0,
            runs: VecDeque::new(),
            growth: Growth::default(),
            offset: 0,
        }
    }

    #[cfg(test)]
    pub(crate) fn push(&mut self, payload: &[u8]) -> SnapshotResult<()> {
        self.push_bounded(payload, SnapshotBudget::new(usize::MAX, usize::MAX)?)
    }

    #[cfg(test)]
    pub(crate) fn push_bounded(
        &mut self,
        payload: &[u8],
        budget: SnapshotBudget,
    ) -> SnapshotResult<()> {
        if self.failed {
            return Err(SnapshotError::new(
                "snapshot preservation builder has failed",
            ));
        }
        let result = self.push_inner(payload, budget);
        if result.is_err() {
            self.failed = true;
        }
        result
    }

    pub(crate) fn advance_bounded(
        &mut self,
        chunk: &[u8],
        budget: SnapshotBudget,
    ) -> SnapshotResult<SnapshotProgress> {
        if self.failed {
            return Err(SnapshotError::new(
                "snapshot preservation builder has failed",
            ));
        }
        let result = self.advance_inner(chunk, budget);
        self.failed |= result
            .as_ref()
            .err()
            .is_some_and(|failure| !budget.is_partial() || !failure.is_budget_refusal());
        result
    }

    fn advance_inner(
        &mut self,
        chunk: &[u8],
        budget: SnapshotBudget,
    ) -> SnapshotResult<SnapshotProgress> {
        let (kind, ordinal, payload) = unframe(chunk)?;
        let packed = packed_record(chunk, self.offset)?;
        let record = packed.unwrap_or(&payload[self.offset..]);
        let framing = if self.offset == 0 {
            chunk.len() - payload.len()
        } else {
            0
        };
        let extent = if packed.is_some() {
            record.len() + framing
        } else {
            chunk.len()
        };
        if extent > budget.max_bytes() {
            return Err(SnapshotError::new(
                "snapshot preservation exceeds advance byte budget",
            ));
        }
        if kind != ChunkKind::Preserved || ordinal != self.ordinal || payload.is_empty() {
            return Err(SnapshotError::new(
                "snapshot preservation chunk is missing or reordered",
            ));
        }
        let mut r = Reader::new(record);
        let tag = r.u8()?;
        if self.started {
            if self.runs.front().map(|&(expected, _)| expected) != Some(tag) || self.remaining == 0
            {
                return Err(SnapshotError::new(
                    "snapshot preservation records are reordered",
                ));
            }
            let ready = match tag {
                ORIGIN => self
                    .growth
                    .ensure(&mut self.state, |s| Ok(&mut s.origins), budget)?,
                STRINGS => self.growth.ensure(
                    &mut self.state,
                    |s| {
                        Arc::get_mut(&mut s.shared_string_cells).ok_or_else(|| {
                            SnapshotError::new("snapshot SST cells are already shared")
                        })
                    },
                    budget,
                )?,
                AXES => self
                    .growth
                    .ensure(&mut self.state, |s| Ok(&mut s.axes), budget)?,
                CREATED => self
                    .growth
                    .ensure(&mut self.state, |s| Ok(&mut s.created), budget)?,
                _ => true,
            };
            if !ready {
                return Ok(SnapshotProgress::pending());
            }
            self.read_record(tag, &mut r)?;
        } else {
            if tag != HEADER {
                return Err(SnapshotError::new(
                    "snapshot preservation header is missing",
                ));
            }
            self.read_header(&mut r)?;
        }
        let consumed = record.len() - r.clone().rest().len();
        if packed.is_some() {
            r.finish()?;
        }
        self.offset += consumed;
        super::step::record(1, consumed + framing);
        if self.offset == payload.len() {
            self.offset = 0;
            self.ordinal += 1;
            Ok(SnapshotProgress::ready())
        } else {
            Ok(SnapshotProgress::pending())
        }
    }

    #[cfg(test)]
    fn push_inner(&mut self, chunk: &[u8], budget: SnapshotBudget) -> SnapshotResult<()> {
        if chunk.len() > budget.max_bytes() {
            return Err(SnapshotError::new(
                "snapshot preservation exceeds advance byte budget",
            ));
        }
        let (kind, ordinal, payload) = unframe(chunk)?;
        if kind != ChunkKind::Preserved || ordinal != self.ordinal || payload.is_empty() {
            return Err(SnapshotError::new(
                "snapshot preservation chunk is missing or reordered",
            ));
        }
        let mut r = Reader::new(payload);
        crate::snapshot::step::record(0, chunk.len());
        let mut records = 0;
        while !r.is_empty() {
            if records == budget.max_records() {
                return Err(SnapshotError::new(
                    "snapshot preservation exceeds advance record budget",
                ));
            }
            let tag = r.u8()?;
            if !self.started {
                if tag != HEADER {
                    return Err(SnapshotError::new(
                        "snapshot preservation header is missing",
                    ));
                }
                self.read_header(&mut r)?;
            } else {
                self.read_record(tag, &mut r)?;
            }
            records += 1;
            crate::snapshot::step::record(1, 0);
        }
        r.finish()?;
        self.ordinal += 1;
        Ok(())
    }

    fn read_header(&mut self, r: &mut Reader<'_>) -> SnapshotResult<()> {
        self.remaining = r
            .var_usize()?
            .checked_sub(1)
            .ok_or_else(|| SnapshotError::new("snapshot preservation count is zero"))?;
        for tag in [ORIGIN, STRINGS, AXES, CREATED] {
            let count = r.var_usize()?;
            if count != 0 {
                self.runs.push_back((tag, count));
            }
        }
        self.validate_counts()?;
        if self.runs.is_empty() != (self.remaining == 0) {
            return Err(SnapshotError::new(
                "snapshot preservation counts do not match",
            ));
        }
        self.started = true;
        Ok(())
    }

    fn read_record(&mut self, tag: u8, r: &mut Reader<'_>) -> SnapshotResult<()> {
        let Some((expected, remaining)) = self.runs.front_mut() else {
            return Err(SnapshotError::new("extra snapshot preservation record"));
        };
        if tag != *expected || self.remaining == 0 {
            return Err(SnapshotError::new(
                "snapshot preservation records are reordered",
            ));
        }
        *remaining -= 1;
        if *remaining == 0 {
            self.runs.pop_front();
        }
        self.remaining -= 1;
        match tag {
            ORIGIN => self.state.origins.push(r.option(Reader::var_usize)?),
            STRINGS => {
                let count = r.var_usize()?;
                self.shared_string_cells()?.push(SharedStringCells::new());
                if count != 0 {
                    self.runs.push_front((STRING_CELL, count));
                }
            }
            STRING_CELL => {
                let key = (r.var_u32()?, r.var_u32()?);
                let index = r.var_usize()?;
                let cells = self
                    .shared_string_cells()?
                    .last_mut()
                    .ok_or_else(|| SnapshotError::new("snapshot SST cell has no sheet"))?;
                if cells
                    .last_key_value()
                    .is_some_and(|(&previous, _)| key <= previous)
                {
                    return Err(SnapshotError::new("snapshot SST cells are reordered"));
                }
                cells.insert(key, index);
            }
            AXES => self
                .state
                .axes
                .push(r.option(|_| Ok(SheetAxes::default()))?),
            CREATED => self.state.created.push(r.bool()?),
            _ => return Err(SnapshotError::new("invalid snapshot preservation tag")),
        }
        self.validate_counts()?;
        if self.runs.is_empty() != (self.remaining == 0) {
            return Err(SnapshotError::new(
                "snapshot preservation counts do not match",
            ));
        }
        Ok(())
    }

    fn validate_counts(&self) -> SnapshotResult<()> {
        let minimum = self.runs.iter().try_fold(0usize, |total, &(_, count)| {
            total
                .checked_add(count)
                .ok_or_else(|| SnapshotError::new("snapshot preservation count overflows usize"))
        })?;
        if minimum > self.remaining {
            return Err(SnapshotError::new(
                "snapshot preservation counts do not match",
            ));
        }
        Ok(())
    }

    pub(crate) fn validate_complete(&self) -> SnapshotResult<()> {
        if !self.started
            || self.remaining != 0
            || !self.runs.is_empty()
            || self.offset != 0
            || self.growth.is_pending()
        {
            return Err(SnapshotError::new(
                "snapshot preservation state is incomplete",
            ));
        }
        Ok(())
    }

    pub(crate) fn finish(self) -> SnapshotResult<PreservedSheetState> {
        if self.failed
            || !self.started
            || self.remaining != 0
            || !self.runs.is_empty()
            || self.offset != 0
            || self.growth.is_pending()
        {
            return Err(SnapshotError::new(
                "snapshot preservation state is incomplete",
            ));
        }
        Ok(self.state)
    }

    fn shared_string_cells(&mut self) -> SnapshotResult<&mut Vec<SharedStringCells>> {
        Arc::get_mut(&mut self.state.shared_string_cells)
            .ok_or_else(|| SnapshotError::new("snapshot SST cells are already shared"))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn encode(state: &PreservedSheetState, budget: SnapshotBudget) -> Vec<Vec<u8>> {
        let mut encoder = PreservedSnapshotEncoder::new();
        let mut chunks = Vec::new();
        while let Some(chunk) = encoder.next(state, budget).unwrap() {
            chunks.push(chunk);
        }
        chunks
    }

    #[test]
    fn huge_preservation_manifest_does_not_reserve_declared_sheets() {
        let budget = SnapshotBudget::new(1, 16 * 1024).unwrap();
        let mut w = Writer::new();
        w.u8(HEADER);
        w.var_usize(50_000_001);
        for count in [50_000_000, 0, 0, 0] {
            w.var_usize(count);
        }
        let chunk = frame(ChunkKind::Preserved, 0, &w.into_bytes());
        let mut builder = PreservedSnapshotBuilder::new();
        super::super::step::reset();
        assert!(builder.advance_bounded(&chunk, budget).unwrap().is_ready());
        assert_eq!(builder.state.origins.capacity(), 0);
        assert_eq!(builder.state.shared_string_cells.capacity(), 0);
        assert_eq!(builder.state.axes.capacity(), 0);
        assert_eq!(builder.state.created.capacity(), 0);
        let work = super::super::step::current();
        assert_eq!(work.records, 1);
        assert!(work.bytes <= budget.max_bytes());
        let failure: SnapshotError = builder.finish().err().unwrap();
        assert_eq!(
            failure.to_string(),
            "snapshot preservation state is incomplete"
        );
    }

    #[test]
    fn nested_preservation_count_is_refused_in_production_advance() {
        let budget = SnapshotBudget::new(1, 16 * 1024).unwrap();
        let mut w = Writer::new();
        w.u8(HEADER);
        w.var_usize(2);
        for count in [0, 1, 0, 0] {
            w.var_usize(count);
        }
        let mut builder = PreservedSnapshotBuilder::new();
        builder
            .advance_bounded(&frame(ChunkKind::Preserved, 0, &w.into_bytes()), budget)
            .unwrap();
        assert_eq!(builder.state.shared_string_cells.capacity(), 0);
        let mut w = Writer::new();
        w.u8(STRINGS);
        w.var_usize(50_000_000);
        super::super::step::reset();
        let failure: SnapshotError = builder
            .advance_bounded(&frame(ChunkKind::Preserved, 1, &w.into_bytes()), budget)
            .unwrap_err();
        assert_eq!(
            failure.to_string(),
            "snapshot preservation counts do not match"
        );
        assert!(builder.state.shared_string_cells.capacity() <= 1);
        assert!(builder.state.shared_string_cells[0].is_empty());
        assert!(super::super::step::current().records <= budget.max_records());
    }

    #[test]
    fn preservation_storage_growth_moves_only_budgeted_admitted_records() {
        let state = PreservedSheetState {
            origins: vec![None; 300],
            ..PreservedSheetState::default()
        };
        let budget = SnapshotBudget::new(1, 16 * 1024).unwrap();
        let mut builder = PreservedSnapshotBuilder::new();
        let mut steps = 0;
        for chunk in encode(&state, budget) {
            loop {
                super::super::step::reset();
                let progress = builder.advance_bounded(&chunk, budget).unwrap();
                let work = super::super::step::current();
                assert!(work.records <= budget.max_records());
                assert!(work.bytes <= budget.max_bytes());
                steps += 1;
                if progress.is_ready() {
                    break;
                }
            }
        }
        assert!(steps > 301);
        assert_eq!(builder.finish().unwrap().origins, state.origins);
    }

    #[test]
    fn preservation_roundtrip_keeps_duplicate_sst_indices_and_axis_presence() {
        let state = PreservedSheetState {
            origins: vec![None, Some(0), Some(usize::MAX)],
            shared_string_cells: Arc::new(vec![
                [((0, 0), 0), ((0, 1), 7), ((0, 2), 3), ((5, 0), 7)]
                    .into_iter()
                    .collect(),
                SharedStringCells::new(),
            ]),
            axes: vec![None, Some(SheetAxes::default()), Some(SheetAxes::default())],
            created: vec![false, true, false, true],
        };
        for records in 1..=3 {
            let budget = SnapshotBudget::new(records, 80).unwrap();
            let chunks = encode(&state, budget);
            let mut builder = PreservedSnapshotBuilder::new();
            for chunk in &chunks {
                assert!(chunk.len() <= 80);
                builder.push(chunk).unwrap();
            }
            let decoded = builder.finish().unwrap();
            assert_eq!(decoded.origins, state.origins);
            assert_eq!(decoded.shared_string_cells, state.shared_string_cells);
            assert_eq!(decoded.axes, state.axes);
            assert_eq!(decoded.created, state.created);
            assert_eq!(encode(&decoded, budget), chunks);
        }
        let mut builder = PreservedSnapshotBuilder::new();
        for chunk in encode(
            &PreservedSheetState::default(),
            SnapshotBudget::new(1, 80).unwrap(),
        ) {
            builder.push(&chunk).unwrap();
        }
        assert!(builder.finish().unwrap().origins.is_empty());
    }

    #[test]
    fn preservation_rejects_nonidentity_axes_and_incomplete_or_reordered_chunks() {
        let mut axes = SheetAxes::default();
        axes.rows.insert(1, 1);
        let state = PreservedSheetState {
            origins: vec![Some(0)],
            shared_string_cells: Arc::new(vec![SharedStringCells::new()]),
            axes: vec![Some(axes)],
            created: vec![false],
        };
        let budget = SnapshotBudget::new(1, 80).unwrap();
        let mut encoder = PreservedSnapshotEncoder::new();
        loop {
            match encoder.next(&state, budget) {
                Ok(Some(_)) => {}
                Err(_) => break,
                Ok(None) => panic!("nonidentity axes were accepted"),
            }
        }
        let chunks = encode(
            &PreservedSheetState {
                origins: vec![None, Some(0)],
                shared_string_cells: Arc::default(),
                axes: vec![None],
                created: vec![true],
            },
            budget,
        );
        let mut builder = PreservedSnapshotBuilder::new();
        builder.push(&chunks[0]).unwrap();
        assert!(builder.finish().is_err());
        let mut builder = PreservedSnapshotBuilder::new();
        assert!(builder.push(&chunks[1]).is_err());
        assert!(builder.finish().is_err());
        let mut builder = PreservedSnapshotBuilder::new();
        builder.push(&chunks[0]).unwrap();
        assert!(builder.push(&chunks[0]).is_err());
        let mut builder = PreservedSnapshotBuilder::new();
        assert!(builder.push(&frame(ChunkKind::Model, 0, &[0])).is_err());
    }

    #[test]
    fn preservation_byte_budget_errors_leave_the_cursor_retryable() {
        let state = PreservedSheetState {
            origins: vec![Some(usize::MAX)],
            shared_string_cells: Arc::default(),
            axes: Vec::new(),
            created: Vec::new(),
        };
        let mut encoder = PreservedSnapshotEncoder::new();
        assert!(
            encoder
                .next(&state, SnapshotBudget::new(1, 1).unwrap())
                .is_err()
        );
        let header = encoder
            .next(&state, SnapshotBudget::new(1, 80).unwrap())
            .unwrap()
            .unwrap();
        assert!(
            encoder
                .next(&state, SnapshotBudget::new(1, 3).unwrap())
                .is_err()
        );
        let origin = encoder
            .next(&state, SnapshotBudget::new(1, 80).unwrap())
            .unwrap()
            .unwrap();
        assert!(
            encoder
                .next(&state, SnapshotBudget::new(1, 80).unwrap())
                .unwrap()
                .is_none()
        );
        let mut builder = PreservedSnapshotBuilder::new();
        builder.push(&header).unwrap();
        builder.push(&origin).unwrap();
        assert_eq!(builder.finish().unwrap().origins, state.origins);
    }
}
