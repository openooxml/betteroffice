use std::cell::RefCell;
use std::ops::{Deref, DerefMut, Range};
use std::rc::Rc;

use super::*;

#[derive(Default)]
struct Reads {
    ids: BTreeSet<String>,
    hidden_fields: bool,
    expected: Option<Rc<BTreeSet<String>>>,
    unexpected: bool,
}

thread_local! {
    static READS: RefCell<Option<Reads>> = const { RefCell::new(None) };
}

pub(super) struct ReadGuard(Option<Reads>);

impl ReadGuard {
    pub(super) fn new() -> Self {
        Self(READS.with(|reads| reads.replace(Some(Reads::default()))))
    }
}

impl Drop for ReadGuard {
    fn drop(&mut self) {
        READS.with(|reads| reads.replace(self.0.take()));
    }
}

pub(super) fn record_decision(value: &Any) {
    READS.with(|reads| {
        if let Some(reads) = reads.borrow_mut().as_mut()
            && let Some((id, ..)) = crate::queries::revision_parts(value)
        {
            if let Some(expected) = reads.expected.as_ref() {
                reads.unexpected |= !expected.contains(&id);
            } else {
                reads.ids.insert(id);
            }
        }
    });
}

pub(super) fn touch_hidden_fields() {
    READS.with(|reads| {
        if let Some(reads) = reads.borrow_mut().as_mut() {
            reads.hidden_fields = true;
        }
    });
}

#[derive(Default)]
pub(super) struct LoweringOutput {
    pub(super) map: LoweringMap,
    story_base: u32,
    paragraph_base: u32,
}

impl LoweringOutput {
    fn at(position: WalkPosition) -> Self {
        Self {
            story_base: position.stories,
            paragraph_base: position.paragraphs,
            ..Self::default()
        }
    }

    pub(super) fn story_count(&self) -> u32 {
        self.story_base + self.stories.len() as u32
    }

    pub(super) fn paragraph_count(&self) -> u32 {
        self.paragraph_base + self.paragraphs.len() as u32
    }
}

impl Deref for LoweringOutput {
    type Target = LoweringMap;

    fn deref(&self) -> &Self::Target {
        &self.map
    }
}

impl DerefMut for LoweringOutput {
    fn deref_mut(&mut self) -> &mut Self::Target {
        &mut self.map
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub(super) struct WalkPosition {
    pub(super) story_slot: u32,
    pub(super) table_ordinal: u32,
    pub(super) break_ordinal: u32,
    pub(super) story_index: u32,
    pub(super) paragraph_start: u32,
    pub(super) paragraph_pm_start: u64,
    pub(super) paragraph_pm_units: u32,
    pub(super) pm_cursor: u64,
    pub(super) at_block_boundary: bool,
    pub(super) section_margins: SectionMarginsTwips,
    pub(super) stories: u32,
    pub(super) paragraphs: u32,
    pub(super) safe: bool,
}

impl WalkPosition {
    pub(super) fn new(story_slot: u32, pm_base: u64, map: &LoweringOutput) -> Self {
        Self {
            story_slot,
            table_ordinal: 0,
            break_ordinal: 0,
            story_index: 0,
            paragraph_start: 0,
            paragraph_pm_start: pm_base,
            paragraph_pm_units: 0,
            pm_cursor: pm_base,
            at_block_boundary: true,
            section_margins: SectionMarginsTwips::default(),
            stories: map.story_count(),
            paragraphs: map.paragraph_count(),
            safe: true,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
struct BoundaryState {
    position: WalkPosition,
    context: Rc<BoundaryContext>,
}

#[derive(Clone, Debug, PartialEq)]
struct BoundaryContext {
    list_state: ListState,
    opaque_sequences: BTreeSet<String>,
    hidden_field_blocks: BTreeSet<String>,
}

impl Deref for BoundaryState {
    type Target = BoundaryContext;

    fn deref(&self) -> &Self::Target {
        &self.context
    }
}

impl BoundaryState {
    fn capture(
        position: WalkPosition,
        list_state: &ListState,
        opaque_sequences: &BTreeSet<String>,
        hidden_field_blocks: &BTreeSet<String>,
    ) -> Self {
        Self {
            position,
            context: Rc::new(BoundaryContext {
                list_state: list_state.clone(),
                opaque_sequences: opaque_sequences.clone(),
                hidden_field_blocks: hidden_field_blocks.clone(),
            }),
        }
    }

    fn refresh(
        &self,
        position: WalkPosition,
        list_state: &ListState,
        opaque_sequences: &BTreeSet<String>,
        hidden_field_blocks: &BTreeSet<String>,
    ) -> Self {
        if self.list_state == *list_state
            && self.opaque_sequences == *opaque_sequences
            && self.hidden_field_blocks == *hidden_field_blocks
        {
            Self {
                position,
                context: Rc::clone(&self.context),
            }
        } else {
            Self::capture(position, list_state, opaque_sequences, hidden_field_blocks)
        }
    }
}

#[derive(Clone, Debug, Default)]
pub(crate) struct PreviewUnits {
    records: Vec<UnitRecord>,
    comments: Rc<Vec<CommentInterval>>,
    /// The lowering numbered SEQ fields across the whole body.
    sequences: bool,
    untracked_state: bool,
    refresh: Option<Refresh>,
    #[cfg(test)]
    pub(crate) work: RecordingWork,
}

#[cfg(test)]
#[derive(Debug, PartialEq, Eq)]
pub(super) enum AnySnapshot {
    Null,
    Undefined,
    Bool(bool),
    Number(u64),
    BigInt(i64),
    String(String),
    Buffer(Vec<u8>),
    Array(Vec<Self>),
    Map(BTreeMap<String, Self>),
}

#[cfg(test)]
impl From<&Any> for AnySnapshot {
    fn from(value: &Any) -> Self {
        match value {
            Any::Null => Self::Null,
            Any::Undefined => Self::Undefined,
            Any::Bool(value) => Self::Bool(*value),
            Any::Number(value) => Self::Number(value.to_bits()),
            Any::BigInt(value) => Self::BigInt(*value),
            Any::String(value) => Self::String(value.to_string()),
            Any::Buffer(value) => Self::Buffer(value.to_vec()),
            Any::Array(value) => Self::Array(value.iter().map(Self::from).collect()),
            Any::Map(value) => Self::Map(
                value
                    .iter()
                    .map(|(key, value)| (key.clone(), Self::from(value)))
                    .collect(),
            ),
        }
    }
}

#[cfg(test)]
#[test]
fn any_snapshot_preserves_variants_and_number_bits() {
    for (left, right) in [
        (Any::Number(1.0), Any::BigInt(1)),
        (Any::Null, Any::Undefined),
        (Any::Number(f64::NAN), Any::Null),
        (Any::Number(0.0), Any::Number(-0.0)),
    ] {
        let left_snapshot = AnySnapshot::from(&left);
        let right_snapshot = AnySnapshot::from(&right);
        assert_ne!(left_snapshot, right_snapshot);
        assert_eq!(left_snapshot, AnySnapshot::from(&left));
        assert_eq!(right_snapshot, AnySnapshot::from(&right));
    }
}

#[cfg(test)]
impl BoundaryState {
    fn exact_snapshot(&self) -> (String, String) {
        (
            format!("{:?}", self.position),
            format!("{:?}", self.context.as_ref()),
        )
    }
}

#[cfg(test)]
impl PreviewUnits {
    pub(crate) fn snapshot(&self, doc: &EditingDoc) -> impl PartialEq + std::fmt::Debug {
        use yrs::types::ToJson;

        let txn = doc.yrs_doc().transact();
        let records = self
            .records
            .iter()
            .map(|record| {
                let chunks = record
                    .chunks
                    .iter()
                    .map(|chunk| {
                        let insert = AnySnapshot::from(&chunk.insert.to_json(&txn));
                        let attributes = chunk.attributes.as_ref().map(|attrs| {
                            attrs
                                .iter()
                                .map(|(key, value)| (key.to_string(), AnySnapshot::from(value)))
                                .collect::<BTreeMap<_, _>>()
                        });
                        (
                            std::mem::discriminant(&chunk.insert),
                            insert,
                            attributes,
                            format!("{:?}", chunk.ychange),
                        )
                    })
                    .collect::<Vec<_>>();
                (
                    record.ids.as_ref().clone(),
                    record.chunk_range.clone(),
                    chunks,
                    record.blocks.clone(),
                    record.revealable.clone(),
                    record.stories.clone(),
                    record.paragraphs.clone(),
                    record.pm.clone(),
                    record.seed.as_deref().map(BoundaryState::exact_snapshot),
                    record.after.exact_snapshot(),
                    record.capture_raw,
                    record.local_stateful,
                )
            })
            .collect::<Vec<_>>();
        let comments = self
            .comments
            .iter()
            .map(|comment| (comment.start, comment.end, comment.id.to_bits()))
            .collect::<Vec<_>>();
        (records, comments, self.sequences)
    }

    pub(crate) fn raw_ranges(&self) -> Vec<Range<u32>> {
        self.records
            .iter()
            .map(|record| {
                record.seed.as_ref().unwrap().position.story_index
                    ..record.after.position.story_index
            })
            .collect()
    }
}

#[cfg(test)]
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub(crate) struct RecordingWork {
    pub(crate) chunks: usize,
    pub(crate) copied_chunks: usize,
    pub(crate) reused_units: usize,
    pub(crate) refresh_records: usize,
    pub(crate) position_steps: usize,
}

#[derive(Clone, Debug)]
pub(crate) struct TextEdit {
    pub(crate) range: Range<u32>,
    pub(crate) inserted: u32,
    pub(crate) epochs: (u64, u64),
}

#[derive(Debug)]
pub(crate) struct ParagraphEdit {
    pub(crate) raw: Range<u32>,
    pub(crate) pm: Range<u64>,
    pub(crate) delta: i64,
}

pub(crate) fn shift_paragraph_edits(
    units: &mut PreviewUnits,
    edits: &[ParagraphEdit],
) -> Option<()> {
    for edit in edits {
        for record in &mut units.records {
            #[cfg(test)]
            crate::engine::TYPING_EXTRA_WORK.with(|work| {
                let mut counts = work.get();
                counts.preview_position_steps += 1;
                work.set(counts);
            });
            if record.pm.end <= edit.pm.start {
                continue;
            }
            if record.pm.start < edit.pm.end {
                return None;
            }
            let shift_raw = |value: u32| u32::try_from(i64::from(value) + edit.delta).ok();
            let shift_pm = |value: u64| value.checked_add_signed(edit.delta);
            let shift_boundary = |boundary: &mut BoundaryState| -> Option<()> {
                let position = &mut boundary.position;
                if !position.safe || position.story_index < edit.raw.end {
                    return None;
                }
                position.story_index = shift_raw(position.story_index)?;
                position.paragraph_start = shift_raw(position.paragraph_start)?;
                position.paragraph_pm_start = shift_pm(position.paragraph_pm_start)?;
                position.pm_cursor = shift_pm(position.pm_cursor)?;
                Some(())
            };
            shift_boundary(Rc::make_mut(record.seed.as_mut()?))?;
            shift_boundary(Rc::make_mut(&mut record.after))?;
            record.pm = shift_pm(record.pm.start)?..shift_pm(record.pm.end)?;
            record.capture_raw = Some(shift_raw(record.capture_raw?)?);
        }
    }
    Some(())
}

#[derive(Clone, Debug)]
struct Refresh {
    previous: Rc<PreviewUnits>,
    ranges: Vec<Range<u32>>,
    edited: Vec<bool>,
    captures: Vec<u32>,
    cursor: usize,
    valid: bool,
}

pub(crate) fn refresh(
    units: &Rc<PreviewUnits>,
    edit: &TextEdit,
    paragraph_edits: &[ParagraphEdit],
) -> Option<PreviewUnits> {
    let shift = |raw: u32| {
        if raw <= edit.range.start {
            raw
        } else {
            raw.saturating_sub(edit.range.end - edit.range.start) + edit.inserted
        }
    };
    let mut ranges = Vec::new();
    let mut edited = Vec::new();
    let mut captures = Vec::new();
    #[cfg(test)]
    let mut work = RecordingWork::default();
    for record in &units.records {
        #[cfg(test)]
        {
            work.refresh_records += 1;
        }
        let mut start = record.seed.as_ref()?.position.story_index;
        let mut end = record.after.position.story_index;
        let mut capture = record.capture_raw?;
        let mut changed = false;
        for paragraph in paragraph_edits {
            #[cfg(test)]
            {
                work.position_steps += 1;
            }
            changed |= start < paragraph.raw.end && end > paragraph.raw.start;
            let shift_paragraph = |value: u32| {
                if value > paragraph.raw.start {
                    u32::try_from(i64::from(value) + paragraph.delta).ok()
                } else {
                    Some(value)
                }
            };
            start = shift_paragraph(start)?;
            end = shift_paragraph(end)?;
            capture = shift_paragraph(capture)?;
        }
        if !record.after.position.safe {
            return None;
        }
        ranges.push(shift(start)..shift(end));
        edited.push(changed || edit.range.start < end && edit.range.end >= start);
        captures.push(shift(capture));
    }
    Some(PreviewUnits {
        records: Vec::with_capacity(units.records.len()),
        refresh: Some(Refresh {
            previous: Rc::clone(units),
            ranges,
            edited,
            captures,
            cursor: 0,
            valid: true,
        }),
        untracked_state: units.untracked_state,
        #[cfg(test)]
        work,
        ..PreviewUnits::default()
    })
}

#[derive(Clone, Debug)]
struct UnitRecord {
    ids: Rc<BTreeSet<String>>,
    chunk_range: Range<usize>,
    chunks: Rc<Vec<yrs::types::text::Diff<YChange>>>,
    blocks: Range<usize>,
    revealable: Range<usize>,
    stories: Range<usize>,
    paragraphs: Range<usize>,
    pm: Range<u64>,
    seed: Option<Rc<BoundaryState>>,
    after: Rc<BoundaryState>,
    capture_raw: Option<u32>,
    local_stateful: Option<bool>,
}

struct Window {
    position: WalkPosition,
    chunk_start: usize,
    block_start: usize,
    revealable_start: usize,
    span_start: usize,
    paragraph_block_start: usize,
    table_start: usize,
    seed: Option<Rc<BoundaryState>>,
    mutated: bool,
    reused: Option<usize>,
    capture_raw: Option<u32>,
    local_stateful: bool,
}

pub(super) struct UnitRecorder {
    units: PreviewUnits,
    window: Option<Window>,
    certify_local: bool,
}

impl UnitRecorder {
    pub(super) fn new(mut units: PreviewUnits, certify_local: bool) -> Self {
        units.untracked_state |= !certify_local;
        Self {
            units,
            window: None,
            certify_local,
        }
    }

    pub(super) fn read_guard(&self) -> ReadGuard {
        if self.units.refresh.is_some() {
            ReadGuard(READS.with(|reads| reads.replace(None)))
        } else {
            ReadGuard::new()
        }
    }

    pub(super) fn wants_chunk(&mut self, raw: u32) -> bool {
        let Some(refresh) = self.units.refresh.as_mut() else {
            return true;
        };
        let Some(range) = refresh.ranges.get(refresh.cursor) else {
            return false;
        };
        if raw > range.end {
            refresh.valid = false;
            return self.window.is_some();
        }
        raw >= range.start
            && (refresh.edited[refresh.cursor]
                || self.certify_local
                    && refresh.previous.records[refresh.cursor]
                        .local_stateful
                        .is_none()
                || raw == range.start
                || raw == refresh.captures[refresh.cursor]
                || raw == range.end)
    }

    #[allow(clippy::too_many_arguments)]
    pub(super) fn before_chunk<T: ReadTxn>(
        &mut self,
        index: usize,
        diff: &yrs::types::text::Diff<YChange>,
        txn: &T,
        position: WalkPosition,
        list_state: &ListState,
        opaque_sequences: &BTreeSet<String>,
        hidden_field_blocks: &BTreeSet<String>,
        blocks: usize,
        revealable: usize,
        map: &LoweringOutput,
    ) {
        #[cfg(test)]
        {
            self.units.work.chunks += 1;
        }
        let (pilcrow, standalone, is_break) = match &diff.insert {
            Out::YMap(mark) => {
                let kind = shared_map_string(mark, txn, crate::KIND_KEY);
                (
                    kind.as_deref() == Some(crate::PILCROW_KIND),
                    matches!(kind.as_deref(), Some("table" | "blockSdt")),
                    matches!(kind.as_deref(), Some("pageBreak" | "columnBreak")),
                )
            }
            _ => (false, false, false),
        };
        if position.at_block_boundary && !is_break {
            self.end_window(
                index,
                position,
                list_state,
                opaque_sequences,
                hidden_field_blocks,
                blocks,
                revealable,
                map,
            );
        }
        if let Some(refresh) = self.units.refresh.as_ref() {
            let Some(range) = refresh.ranges.get(refresh.cursor) else {
                return;
            };
            if position.story_index < range.start || position.story_index >= range.end {
                return;
            }
        }
        if self.window.is_none() {
            let reused = self.units.refresh.as_mut().and_then(|refresh| {
                refresh.valid &=
                    position.safe && position.story_index == refresh.ranges[refresh.cursor].start;
                (!refresh.edited[refresh.cursor]).then_some(refresh.cursor)
            });
            self.window = Some(Window {
                position,
                chunk_start: index,
                block_start: blocks,
                revealable_start: revealable,
                span_start: map.spans.len(),
                paragraph_block_start: map.paragraph_blocks.len(),
                table_start: map.tables.len(),
                seed: None,
                mutated: is_break,
                reused,
                capture_raw: None,
                local_stateful: false,
            });
            let expected = reused.map(|index| {
                Rc::clone(&self.units.refresh.as_ref().unwrap().previous.records[index].ids)
            });
            READS.with(|reads| {
                *reads.borrow_mut() = Some(Reads {
                    expected,
                    ..Reads::default()
                });
            });
        }
        let window = self.window.as_mut().unwrap();
        if self.certify_local {
            #[cfg(test)]
            crate::engine::TYPING_EXTRA_WORK.with(|work| {
                let mut counts = work.get();
                counts.preview_certifications += 1;
                work.set(counts);
            });
            window.local_stateful |= local::preview_touches_state(diff, txn);
            for key in [INS, DEL] {
                if let Some(value) = attribute(diff.attributes.as_deref(), key) {
                    record_decision(value);
                }
            }
        }
        if !window.mutated && (pilcrow || standalone) {
            let has_reads = READS.with(|reads| {
                reads
                    .borrow()
                    .as_ref()
                    .is_some_and(|reads| !reads.ids.is_empty())
            });
            if window.position.safe && (standalone || has_reads || window.reused.is_some()) {
                let previous = window.reused.and_then(|index| {
                    self.units.refresh.as_ref()?.previous.records[index]
                        .seed
                        .as_ref()
                });
                let seed = if let Some(previous) = previous {
                    previous.refresh(
                        window.position,
                        list_state,
                        opaque_sequences,
                        hidden_field_blocks,
                    )
                } else {
                    BoundaryState::capture(
                        window.position,
                        list_state,
                        opaque_sequences,
                        hidden_field_blocks,
                    )
                };
                window.seed = Some(Rc::new(seed));
                window.capture_raw = Some(position.story_index);
            }
            window.mutated = true;
        }
    }

    #[allow(clippy::too_many_arguments)]
    pub(super) fn end_window(
        &mut self,
        end: usize,
        position: WalkPosition,
        list_state: &ListState,
        opaque_sequences: &BTreeSet<String>,
        hidden_field_blocks: &BTreeSet<String>,
        blocks: usize,
        revealable: usize,
        map: &LoweringOutput,
    ) {
        let Some(mut window) = self.window.take() else {
            return;
        };
        let reads = READS.with(|reads| reads.borrow_mut().take().unwrap());
        if let Some(refresh) = self.units.refresh.as_mut() {
            refresh.valid &= position.safe
                && position.story_index == refresh.ranges[refresh.cursor].end
                && !reads.unexpected;
            refresh.cursor += 1;
        }
        let previous = window
            .reused
            .map(|index| &self.units.refresh.as_ref().unwrap().previous.records[index]);
        if self.certify_local {
            window.local_stateful |= !window.position.safe || !position.safe || reads.hidden_fields;
        }
        if reads.ids.is_empty() && previous.is_none() {
            self.units.untracked_state |= window.local_stateful;
            return;
        }
        if !position.safe || reads.hidden_fields {
            window.seed = None;
        }
        let pm = window.position.pm_cursor..position.pm_cursor;
        let owns_positions = map.spans[window.span_start..]
            .iter()
            .all(|span| pm.contains(&span.pm_start))
            && map.paragraph_blocks[window.paragraph_block_start..]
                .iter()
                .all(|(start, _)| pm.contains(start))
            && map.tables[window.table_start..]
                .iter()
                .all(|(start, ..)| pm.contains(start));
        if !owns_positions {
            window.seed = None;
        }
        self.units.records.push(UnitRecord {
            ids: previous.map_or_else(|| Rc::new(reads.ids), |record| Rc::clone(&record.ids)),
            chunk_range: window.chunk_start..end,
            chunks: previous
                .map_or_else(|| Rc::new(Vec::new()), |record| Rc::clone(&record.chunks)),
            blocks: window.block_start..blocks,
            revealable: window.revealable_start..revealable,
            stories: window.position.stories as usize..position.stories as usize,
            paragraphs: window.position.paragraphs as usize..position.paragraphs as usize,
            pm,
            seed: window.seed,
            capture_raw: window.capture_raw,
            local_stateful: self.certify_local.then(|| {
                window.local_stateful
                    || previous.is_some_and(|record| record.local_stateful == Some(true))
            }),
            after: Rc::new(if let Some(previous) = previous {
                previous
                    .after
                    .refresh(position, list_state, opaque_sequences, hidden_field_blocks)
            } else {
                BoundaryState::capture(position, list_state, opaque_sequences, hidden_field_blocks)
            }),
        });
        #[cfg(test)]
        {
            self.units.work.reused_units += usize::from(previous.is_some());
        }
    }

    pub(super) fn save_chunks(
        &mut self,
        chunks: &[yrs::types::text::Diff<YChange>],
        comments: Rc<Vec<CommentInterval>>,
    ) {
        for record in &mut self.units.records {
            if !record.chunks.is_empty() {
                continue;
            }
            #[cfg(test)]
            {
                self.units.work.copied_chunks += record.chunk_range.len();
            }
            record.chunks = Rc::new(
                chunks[record.chunk_range.clone()]
                    .iter()
                    .map(|chunk| {
                        yrs::types::text::Diff::with_change(
                            chunk.insert.clone(),
                            chunk.attributes.clone(),
                            chunk.ychange.clone(),
                        )
                    })
                    .collect(),
            );
        }
        self.units.comments = comments;
    }

    pub(super) fn finish(mut self, sequences: bool) -> PreviewUnits {
        self.units.sequences = sequences;
        if let Some(refresh) = self.units.refresh.as_mut() {
            refresh.valid &= refresh.cursor == refresh.ranges.len();
        }
        self.units
    }
}

pub(crate) fn lower_recorded(
    doc: &EditingDoc,
    story: &str,
    env: &RenderEnv,
    local: &mut local::LocalLowering,
    record: bool,
) -> Result<
    (
        Vec<LayoutBlock>,
        LoweringMap,
        Vec<LayoutBlock>,
        Option<PreviewUnits>,
    ),
    BridgeError,
> {
    let mut preview = (record && story == "body").then(PreviewUnits::default);
    let mut revealable = Some(Vec::new());
    let (blocks, map) = yrs_doc_to_mapped_layout_blocks_inner(
        doc,
        story,
        env,
        &mut revealable,
        local,
        &mut preview,
    )?;
    local.finish(&blocks, &map);
    Ok((blocks, map, revealable.unwrap_or_default(), preview))
}

pub(crate) fn lower_refreshed(
    doc: &EditingDoc,
    story: &str,
    env: &RenderEnv,
    local: &mut local::LocalLowering,
    units: PreviewUnits,
) -> Result<
    (
        Vec<LayoutBlock>,
        LoweringMap,
        Vec<LayoutBlock>,
        Option<PreviewUnits>,
    ),
    BridgeError,
> {
    let mut preview = Some(units);
    let mut revealable = Some(Vec::new());
    let (blocks, map) = yrs_doc_to_mapped_layout_blocks_inner(
        doc,
        story,
        env,
        &mut revealable,
        local,
        &mut preview,
    )?;
    if preview
        .as_ref()
        .is_some_and(|units| units.refresh.as_ref().is_some_and(|refresh| !refresh.valid))
    {
        *local = if local.legacy {
            local::LocalLowering::fallback(local.enabled)
        } else {
            local::LocalLowering::new(local.enabled)
        };
        return lower_recorded(doc, story, env, local, true);
    }
    if let Some(units) = preview.as_mut() {
        units.refresh = None;
    }
    local.finish(&blocks, &map);
    Ok((blocks, map, revealable.unwrap_or_default(), preview))
}

pub(crate) struct Replay {
    record: usize,
    blocks: Vec<LayoutBlock>,
    map: LoweringMap,
    revealable: Vec<LayoutBlock>,
    ids: BTreeSet<String>,
    local: local::LocalLowering,
}

fn replace_positions<T>(
    entries: &mut Vec<T>,
    replacements: Vec<T>,
    pm: &Range<u64>,
    key: impl Fn(&T) -> u64,
) {
    entries.retain(|entry| !pm.contains(&key(entry)));
    entries.extend(replacements);
    entries.sort_by_key(key);
}

pub(crate) fn targets(
    units: &PreviewUnits,
    changed: &BTreeSet<String>,
    local: &local::LocalLowering,
) -> bool {
    local.enabled && (units.untracked_state || units.records.is_empty())
        || units
            .records
            .iter()
            .any(|record| !record.ids.is_disjoint(changed))
}

pub(crate) fn replayable(units: &PreviewUnits, local: &local::LocalLowering) -> bool {
    !local.enabled || !units.untracked_state && !units.records.is_empty()
}

pub(crate) fn replay(
    doc: &EditingDoc,
    env: &RenderEnv,
    units: &PreviewUnits,
    changed: &BTreeSet<String>,
    current: &[Rc<LayoutBlock>],
    current_map: &LoweringMap,
    local: &local::LocalLowering,
) -> Option<Vec<Replay>> {
    if !replayable(units, local) {
        return None;
    }
    let txn = doc.yrs_doc().transact();
    let story = story_ref(&txn, "body").ok()?;
    let with_media;
    let env = if env.media_tokens && env.media.is_empty() {
        with_media = RenderEnv {
            media: doc.media_sources(),
            ..env.clone()
        };
        &with_media
    } else {
        env
    };
    let mut replays = Vec::new();
    for (index, record) in units.records.iter().enumerate() {
        if record.ids.is_disjoint(changed) {
            continue;
        }
        if local.enabled
            && (record.local_stateful != Some(false)
                || record
                    .chunks
                    .iter()
                    .any(|chunk| local::preview_touches_state(chunk, &txn)))
        {
            return None;
        }
        let seed = record.seed.as_ref()?;
        let mut list_state = seed.list_state.clone();
        let mut opaque_sequences = seed.opaque_sequences.clone();
        let mut output = LoweringOutput::at(seed.position);
        let mut revealed = Some(Vec::new());
        let mut active_stories = BTreeSet::from(["body".to_owned()]);
        let mut replay_local = local::LocalLowering::new(local.enabled && !local.blocked);
        let _reads = ReadGuard::new();
        let (replacement, position, hidden_field_blocks) = walk_story_chunks(
            &txn,
            "body",
            env,
            &mut active_stories,
            &mut list_state,
            CellEdges::default(),
            &mut output,
            &mut opaque_sequences,
            &mut revealed,
            &mut replay_local,
            &story,
            &units.comments,
            &record.chunks,
            seed.position,
            seed.hidden_field_blocks.clone(),
            None,
        )
        .ok()?;
        let after = BoundaryContext {
            list_state,
            opaque_sequences,
            hidden_field_blocks,
        };
        let reads = READS.with(|reads| reads.borrow_mut().take().unwrap());
        if position != record.after.position
            || after != *record.after.context
            || current_map.stories.get(record.stories.clone())? != output.stories.as_slice()
            || reads.hidden_fields
            || !local.replay_preserves_state(&replay_local)
            || output
                .spans
                .iter()
                .any(|span| !record.pm.contains(&span.pm_start))
            || output
                .paragraph_blocks
                .iter()
                .any(|(pm, _)| !record.pm.contains(pm))
            || output.tables.iter().any(|(pm, ..)| !record.pm.contains(pm))
            || units.sequences
                && current.get(record.blocks.clone()).is_none_or(|previous| {
                    previous.iter().any(|block| {
                        docx_layout::sequence_fields::reads_sequence_fields(std::slice::from_ref(
                            block.as_ref(),
                        ))
                    }) || docx_layout::sequence_fields::reads_sequence_fields(&replacement)
                })
        {
            return None;
        }
        replays.push(Replay {
            record: index,
            blocks: replacement,
            map: output.map,
            revealable: revealed.unwrap_or_default(),
            ids: reads.ids,
            local: replay_local,
        });
    }
    Some(replays)
}

pub(crate) fn splice(
    replays: Vec<Replay>,
    blocks: &mut Vec<Rc<LayoutBlock>>,
    map: &mut LoweringMap,
    revealable: &mut Vec<LayoutBlock>,
    units: &mut PreviewUnits,
    local: &mut local::LocalLowering,
) {
    let mut block_shift = 0_isize;
    let mut revealable_shift = 0_isize;
    let mut replays = replays.into_iter().peekable();
    for (index, record) in units.records.iter_mut().enumerate() {
        record.blocks = record.blocks.start.saturating_add_signed(block_shift)
            ..record.blocks.end.saturating_add_signed(block_shift);
        record.revealable = record
            .revealable
            .start
            .saturating_add_signed(revealable_shift)
            ..record
                .revealable
                .end
                .saturating_add_signed(revealable_shift);
        let Some(replay) = replays.next_if(|replay| replay.record == index) else {
            continue;
        };
        block_shift += replay.blocks.len() as isize - record.blocks.len() as isize;
        revealable_shift += replay.revealable.len() as isize - record.revealable.len() as isize;
        let block_end = record.blocks.start + replay.blocks.len();
        let revealable_end = record.revealable.start + replay.revealable.len();
        local.replace_seeds(&record.pm, replay.local);
        blocks.splice(
            record.blocks.clone(),
            replay.blocks.into_iter().map(Rc::new),
        );
        revealable.splice(record.revealable.clone(), replay.revealable);
        map.stories
            .splice(record.stories.clone(), replay.map.stories);
        map.paragraphs
            .splice(record.paragraphs.clone(), replay.map.paragraphs);
        replace_positions(&mut map.spans, replay.map.spans, &record.pm, |span| {
            span.pm_start
        });
        replace_positions(
            &mut map.paragraph_blocks,
            replay.map.paragraph_blocks,
            &record.pm,
            |(pm, _)| *pm,
        );
        replace_positions(
            &mut map.tables,
            replay.map.tables,
            &record.pm,
            |(pm, ..)| *pm,
        );
        record.blocks.end = block_end;
        record.revealable.end = revealable_end;
        record.ids = Rc::new(replay.ids);
    }
    local.refresh_seeds(blocks, map);
}

#[cfg(test)]
#[test]
fn local_certification_refresh_matches_fresh_recording() {
    for previous_enabled in [false, true] {
        for (enabled, legacy) in [(false, false), (true, false), (true, true)] {
            let doc = EditingDoc::new(75272);
            crate::seed_from_docx(&doc, &crate::engine::preview_fixture::nested()).unwrap();
            let env = RenderEnv::default();
            let (blocks, map, _, previous) = lower_recorded(
                &doc,
                "body",
                &env,
                &mut local::LocalLowering::new(previous_enabled),
                true,
            )
            .unwrap();
            let previous = Rc::new(previous.unwrap());
            if !previous_enabled {
                for record in &previous.records {
                    assert_eq!(record.local_stateful, None);
                }
                let changed = BTreeSet::from(["unused".to_owned()]);
                assert!(!targets(
                    &previous,
                    &changed,
                    &local::LocalLowering::new(false)
                ));
                let local = local::LocalLowering::new(true);
                assert!(targets(&previous, &changed, &local));
                let current: Vec<_> = blocks.into_iter().map(Rc::new).collect();
                assert!(replay(&doc, &env, &previous, &changed, &current, &map, &local).is_none());
            }
            let record = previous.records.last().unwrap();
            let raw = record.seed.as_ref().unwrap().position.story_index + 1;
            doc.insert_text(
                &crate::EditCtx::local("", ""),
                crate::Position::new("body", raw),
                "x",
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
            let units = refresh(
                &previous,
                &TextEdit {
                    range: raw..raw,
                    inserted: 1,
                    epochs: (0, 1),
                },
                &[],
            )
            .unwrap();
            let lowering = || {
                if legacy {
                    local::LocalLowering::fallback(enabled)
                } else {
                    local::LocalLowering::new(enabled)
                }
            };
            crate::engine::TYPING_EXTRA_WORK.with(|work| work.set(Default::default()));
            let (_, _, _, refreshed) =
                lower_refreshed(&doc, "body", &env, &mut lowering(), units).unwrap();
            if !enabled || legacy {
                crate::engine::TYPING_EXTRA_WORK.with(|work| {
                    assert_eq!(work.get().preview_certifications, 0);
                });
            }
            let (_, _, _, fresh) =
                lower_recorded(&doc, "body", &env, &mut lowering(), true).unwrap();
            let (refreshed, fresh) = (refreshed.unwrap(), fresh.unwrap());
            assert_eq!(refreshed.snapshot(&doc), fresh.snapshot(&doc));
        }
    }
}

#[cfg(test)]
#[test]
fn paragraph_positions_preserve_intermediate_overflow_and_underflow_refusals() {
    fn boundary(raw: u32, pm: u64) -> Rc<BoundaryState> {
        let mut position = WalkPosition::new(0, pm, &LoweringOutput::default());
        position.story_index = raw;
        position.paragraph_start = raw;
        position.paragraph_pm_start = pm;
        Rc::new(BoundaryState::capture(
            position,
            &ListState::default(),
            &BTreeSet::new(),
            &BTreeSet::new(),
        ))
    }

    for state in 0..5 {
        let mut units = PreviewUnits {
            records: vec![UnitRecord {
                ids: Rc::new(BTreeSet::new()),
                chunk_range: 0..0,
                chunks: Rc::new(Vec::new()),
                blocks: 0..0,
                revealable: 0..0,
                stories: 0..0,
                paragraphs: 0..0,
                pm: 30..35,
                seed: Some(boundary(30, 30)),
                after: boundary(35, 35),
                capture_raw: Some(33),
                local_stateful: Some(false),
            }],
            ..PreviewUnits::default()
        };
        let record = &mut units.records[0];
        let delta = match state {
            0 => {
                Rc::make_mut(record.seed.as_mut().unwrap())
                    .position
                    .paragraph_start = 0;
                -1
            }
            1 => {
                Rc::make_mut(record.seed.as_mut().unwrap())
                    .position
                    .pm_cursor = 0;
                -1
            }
            2 => {
                Rc::make_mut(&mut record.after).position.paragraph_pm_start = u64::MAX;
                1
            }
            3 => {
                record.capture_raw = Some(u32::MAX);
                1
            }
            _ => {
                Rc::make_mut(&mut record.after).position.story_index = u32::MAX;
                1
            }
        };
        let edits = [delta, -delta].map(|delta| ParagraphEdit {
            raw: 0..4,
            pm: 0..4,
            delta,
        });
        if state >= 3 {
            let edit = TextEdit {
                range: 0..0,
                inserted: 0,
                epochs: (0, 1),
            };
            assert!(refresh(&Rc::new(units.clone()), &edit, &edits).is_none());
        }
        assert_eq!(shift_paragraph_edits(&mut units, &edits), None);
    }
}
