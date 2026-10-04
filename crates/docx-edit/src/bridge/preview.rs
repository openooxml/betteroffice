use std::cell::RefCell;
use std::ops::{Deref, DerefMut, Range};
use std::rc::Rc;

use super::*;

#[derive(Default)]
struct Reads {
    ids: BTreeSet<String>,
    hidden_fields: bool,
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
            reads.ids.insert(id);
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
    list_state: ListState,
    opaque_sequences: BTreeSet<String>,
    hidden_field_blocks: BTreeSet<String>,
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
            list_state: list_state.clone(),
            opaque_sequences: opaque_sequences.clone(),
            hidden_field_blocks: hidden_field_blocks.clone(),
        }
    }
}

#[derive(Clone, Debug, Default)]
pub(crate) struct PreviewUnits {
    records: Vec<UnitRecord>,
    comments: Rc<Vec<CommentInterval>>,
    /// The lowering numbered SEQ fields across the whole body.
    sequences: bool,
}

#[derive(Clone, Debug)]
struct UnitRecord {
    ids: BTreeSet<String>,
    chunk_range: Range<usize>,
    chunks: Rc<Vec<yrs::types::text::Diff<YChange>>>,
    blocks: Range<usize>,
    revealable: Range<usize>,
    stories: Range<usize>,
    paragraphs: Range<usize>,
    pm: Range<u64>,
    seed: Option<Rc<BoundaryState>>,
    after: Rc<BoundaryState>,
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
}

pub(super) struct UnitRecorder {
    units: PreviewUnits,
    window: Option<Window>,
}

impl UnitRecorder {
    pub(super) fn new() -> Self {
        Self {
            units: PreviewUnits::default(),
            window: None,
        }
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
        if self.window.is_none() {
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
            });
            READS.with(|reads| *reads.borrow_mut() = Some(Reads::default()));
        }
        let window = self.window.as_mut().unwrap();
        if !window.mutated && (pilcrow || standalone) {
            let has_reads = READS.with(|reads| {
                reads
                    .borrow()
                    .as_ref()
                    .is_some_and(|reads| !reads.ids.is_empty())
            });
            if window.position.safe && (standalone || has_reads) {
                window.seed = Some(Rc::new(BoundaryState::capture(
                    window.position,
                    list_state,
                    opaque_sequences,
                    hidden_field_blocks,
                )));
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
        let reads = READS.with(|reads| reads.borrow_mut().replace(Reads::default()).unwrap());
        if reads.ids.is_empty() {
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
            ids: reads.ids,
            chunk_range: window.chunk_start..end,
            chunks: Rc::new(Vec::new()),
            blocks: window.block_start..blocks,
            revealable: window.revealable_start..revealable,
            stories: window.position.stories as usize..position.stories as usize,
            paragraphs: window.position.paragraphs as usize..position.paragraphs as usize,
            pm,
            seed: window.seed,
            after: Rc::new(BoundaryState::capture(
                position,
                list_state,
                opaque_sequences,
                hidden_field_blocks,
            )),
        });
    }

    pub(super) fn save_chunks(
        &mut self,
        chunks: &[yrs::types::text::Diff<YChange>],
        comments: Rc<Vec<CommentInterval>>,
    ) {
        for record in &mut self.units.records {
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

pub(crate) struct Replay {
    record: usize,
    blocks: Vec<LayoutBlock>,
    map: LoweringMap,
    revealable: Vec<LayoutBlock>,
    ids: BTreeSet<String>,
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

pub(crate) fn targets(units: &PreviewUnits, changed: &BTreeSet<String>) -> bool {
    units
        .records
        .iter()
        .any(|record| !record.ids.is_disjoint(changed))
}

pub(crate) fn replay(
    doc: &EditingDoc,
    env: &RenderEnv,
    units: &PreviewUnits,
    changed: &BTreeSet<String>,
    current: &[LayoutBlock],
) -> Option<Vec<Replay>> {
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
        let seed = record.seed.as_ref()?;
        let mut list_state = seed.list_state.clone();
        let mut opaque_sequences = seed.opaque_sequences.clone();
        let mut output = LoweringOutput::at(seed.position);
        let mut revealed = Some(Vec::new());
        let mut active_stories = BTreeSet::from(["body".to_owned()]);
        let mut local = local::LocalLowering::new(false);
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
            &mut local,
            &story,
            &units.comments,
            &record.chunks,
            seed.position,
            seed.hidden_field_blocks.clone(),
            None,
        )
        .ok()?;
        let after = BoundaryState {
            position,
            list_state,
            opaque_sequences,
            hidden_field_blocks,
        };
        let reads = READS.with(|reads| reads.borrow_mut().take().unwrap());
        if after != *record.after
            || reads.hidden_fields
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
                    docx_layout::sequence_fields::reads_sequence_fields(previous)
                        || docx_layout::sequence_fields::reads_sequence_fields(&replacement)
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
        });
    }
    Some(replays)
}

pub(crate) fn splice(
    replays: Vec<Replay>,
    blocks: &mut Vec<LayoutBlock>,
    map: &mut LoweringMap,
    revealable: &mut Vec<LayoutBlock>,
    units: &mut PreviewUnits,
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
        blocks.splice(record.blocks.clone(), replay.blocks);
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
        record.ids = replay.ids;
    }
}
