//! Endnotes Word sets after the last body text of the document (`docEnd`) or of their section
//! (`sectEnd`), starting on the page that text ends on and continuing onto new pages. They join
//! the body flow as reserved blocks, one per slice of lines the note may not split, so the
//! paginator breaks them across pages as it breaks body text; the note areas are then read back
//! from where those blocks landed.

use std::collections::{HashMap, HashSet};
use std::ops::Range;

use crate::footnotes::{
    NoteContent, NoteKind, NoteRefLocation, NoteSeparatorHeights, OrderedMap, collect_note_refs,
    content_map, map_note_anchors_to_pages, note_item, page_for_map_id,
};
use crate::keep_together::{paragraph_is_unbreakable, paragraph_widow_control};
use crate::measure_blocks::extent_height;
use crate::regions::{DocumentRegions, advances_region_section};
use crate::types::{
    BlockExtent, BlockId, Fragment, ImageExtent, Layout, LayoutBlock, MeasuredBlock,
    NoteAreaContract, NoteLayoutItemContract, Page, ReservationRef, ReservedBlock,
};

/// The reservation owner of the blocks that carry flowed endnotes.
pub const ENDNOTE_FLOW_OWNER: &str = "en-flow";
const BLOCK_PREFIX: &str = "en-flow:";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FlowEnd {
    Document,
    Section,
}

impl FlowEnd {
    fn placement(self) -> &'static str {
        match self {
            Self::Document => "docEnd",
            Self::Section => "sectEnd",
        }
    }
}

/// Where an endnote anchored in `section_index` flows to, or `None` when it stays a page note.
pub fn endnote_flow_end(regions: &DocumentRegions, section_index: usize) -> Option<FlowEnd> {
    match regions
        .note_properties(section_index, NoteKind::Endnote.as_str())
        .position
        .as_deref()
        .unwrap_or("docEnd")
    {
        "docEnd" => Some(FlowEnd::Document),
        "sectEnd" => Some(FlowEnd::Section),
        _ => None,
    }
}

pub fn flows_with_body(kind: NoteKind, regions: &DocumentRegions, section_index: usize) -> bool {
    kind == NoteKind::Endnote && endnote_flow_end(regions, section_index).is_some()
}

pub fn is_flowed(block: &ReservedBlock) -> bool {
    block.reservation.owner == ENDNOTE_FLOW_OWNER
}

fn is_flowed_block(block: &LayoutBlock) -> bool {
    matches!(block, LayoutBlock::Reserved(reserved) if is_flowed(reserved))
}

/// The width a flowed note is set at on `page`: one column of the page's body.
pub fn flow_column_width(page: &Page) -> f64 {
    let margins = page.body_margins.as_ref().unwrap_or(&page.margins);
    let content = page.size.w - margins.left - margins.right;
    let width = match &page.columns {
        Some(columns) if columns.count > 1.0 => {
            (content - (columns.count - 1.0) * columns.gap) / columns.count
        }
        _ => content,
    };
    width.max(1.0)
}

/// The page a flowed note of `section_index` ends on in `pages`.
pub fn flow_target_page<'a>(
    pages: &'a [Page],
    regions: &DocumentRegions,
    section_index: usize,
) -> Option<&'a Page> {
    match endnote_flow_end(regions, section_index)? {
        FlowEnd::Document => pages.last(),
        FlowEnd::Section => pages
            .iter()
            .rev()
            .find(|page| page.region_section_index <= section_index),
    }
}

#[derive(Clone, Debug)]
struct Piece {
    map_id: i64,
    block: usize,
    lines: Range<usize>,
    /// The note content the piece shows; its block may be taller by a separator.
    height: f64,
    /// Index of the piece among its note's pieces, and their count.
    index: usize,
    count: usize,
    /// The first piece after the document's or a section's body text, under the separator.
    opens_run: bool,
    end: FlowEnd,
}

/// The endnotes [`flow_endnotes`] added to a measured flow.
#[derive(Clone, Debug, Default)]
pub struct EndnoteFlow {
    pieces: HashMap<String, Piece>,
    inserted: Vec<usize>,
}

impl EndnoteFlow {
    pub fn is_empty(&self) -> bool {
        self.inserted.is_empty()
    }

    /// Indices of the added blocks in the measured flow, ascending.
    pub fn inserted(&self) -> &[usize] {
        &self.inserted
    }
}

/// The slices of `content` the flow may not split, as `(block, lines, height)`; a block other
/// than a paragraph is one line. Heights add up to what painting the whole note advances. A
/// paragraph kept whole that is taller than `room` splits between lines anyway, as in Word.
fn note_slices(content: &NoteContent, room: f64) -> Vec<(usize, Range<usize>, f64)> {
    let mut slices = Vec::new();
    for (index, (block, measure)) in content.blocks.iter().zip(&content.measures).enumerate() {
        match (block, measure) {
            (LayoutBlock::Paragraph(paragraph), BlockExtent::Paragraph(extent)) => {
                let before = paragraph
                    .attrs
                    .as_ref()
                    .and_then(|attrs| attrs.spacing.as_ref())
                    .and_then(|spacing| spacing.before)
                    .unwrap_or(0.0);
                let count = extent.lines.len();
                if count == 0 {
                    slices.push((index, 0..0, before + extent.total_height));
                    continue;
                }
                let lines: f64 = extent.lines.iter().map(|line| line.line_height).sum();
                let tail = extent.total_height - lines;
                let mut cuts = vec![0];
                if !paragraph_is_unbreakable(paragraph, extent)
                    || before + extent.total_height > room
                {
                    let widow = paragraph_widow_control(paragraph, extent) && count >= 4;
                    let (first, last) = if widow {
                        (2, count - 2)
                    } else {
                        (1, count - 1)
                    };
                    cuts.extend(first..=last);
                    cuts.dedup();
                }
                cuts.push(count);
                cuts.dedup();
                for window in cuts.windows(2) {
                    let range = window[0]..window[1];
                    let mut height: f64 = extent.lines[range.clone()]
                        .iter()
                        .map(|line| line.line_height)
                        .sum();
                    if range.start == 0 {
                        height += before;
                    }
                    if range.end == count {
                        height += tail;
                    }
                    slices.push((index, range, height));
                }
            }
            (_, measure) => {
                let height = extent_height(measure);
                if height > 0.0 {
                    slices.push((index, 0..1, height));
                }
            }
        }
    }
    slices
}

/// Adds the flowing endnotes to `measured` as reserved blocks: the document's after its last
/// body block, a section's before the break that closes it. `pages` is a layout of the same
/// body, which places the notes' references. The first block of each run of notes also holds
/// the endnote separator, so the separator never ends a page alone; a later block that opens
/// a page also holds the continuation separator.
pub fn flow_endnotes(
    measured: &mut Vec<MeasuredBlock>,
    pages: &[Page],
    refs: &[NoteRefLocation],
    contents: &[NoteContent],
    separator_heights: &NoteSeparatorHeights,
    regions: &DocumentRegions,
) -> EndnoteFlow {
    let anchors = map_note_anchors_to_pages(pages, refs);
    let contents = content_map(contents);
    let mut sections = HashMap::new();
    let mut section = 0;
    for entry in measured.iter() {
        for reference in collect_note_refs(std::slice::from_ref(&entry.block)) {
            sections.entry(reference.map_id()).or_insert(section);
        }
        if advances_region_section(&entry.block) {
            section += 1;
        }
    }
    let mut groups: Vec<(Option<usize>, Vec<&NoteContent>)> = Vec::new();
    let mut seen = HashSet::new();
    for reference in refs {
        if reference.note_kind != NoteKind::Endnote {
            continue;
        }
        let map_id = reference.map_id();
        let Some(anchor) = page_for_map_id(pages, &anchors, map_id) else {
            continue;
        };
        let section = sections
            .get(&map_id)
            .copied()
            .unwrap_or(anchor.region_section_index);
        let Some(end) = endnote_flow_end(regions, section) else {
            continue;
        };
        let Some(content) = contents.get(&map_id).copied() else {
            continue;
        };
        if !seen.insert(map_id) {
            continue;
        }
        let key = (end == FlowEnd::Section).then_some(section);
        match groups.iter_mut().find(|(group, _)| *group == key) {
            Some((_, notes)) => notes.push(content),
            None => groups.push((key, vec![content])),
        }
    }

    let section_ends: Vec<usize> = measured
        .iter()
        .enumerate()
        .filter(|(_, entry)| advances_region_section(&entry.block))
        .map(|(index, _)| index)
        .collect();
    let document_end = match section_ends.last() {
        Some(&last) if last + 1 == measured.len() => last,
        _ => measured.len(),
    };

    let mut flow = EndnoteFlow::default();
    let mut insertions: Vec<(usize, Vec<MeasuredBlock>)> = Vec::new();
    for (section, notes) in groups {
        let target = match section {
            Some(section) => pages
                .iter()
                .rev()
                .find(|page| page.region_section_index <= section),
            None => pages.last(),
        };
        let separator = target.map_or_else(
            || separator_heights.height(NoteKind::Endnote, section.unwrap_or(0)),
            |page| separator_heights.height_for_page(NoteKind::Endnote, page),
        );
        let width = target.map_or(0.0, flow_column_width);
        let room = target.map_or(f64::INFINITY, |page| {
            let margins = page.body_margins.as_ref().unwrap_or(&page.margins);
            page.size.h - margins.top - margins.bottom
        });
        let end = if section.is_some() {
            FlowEnd::Section
        } else {
            FlowEnd::Document
        };
        let mut blocks = Vec::new();
        for content in notes {
            let map_id = content.map_id();
            let mut slices = note_slices(content, room - separator);
            if slices.is_empty() && content.height > 0.0 {
                slices.push((0, 0..0, content.height));
            }
            let count = slices.len();
            for (index, (block, lines, height)) in slices.into_iter().enumerate() {
                let opens_run = blocks.is_empty();
                let id = format!("{BLOCK_PREFIX}{map_id}:{block}:{}", lines.start);
                let block_height = if opens_run {
                    height + separator
                } else {
                    height
                };
                flow.pieces.insert(
                    id.clone(),
                    Piece {
                        map_id,
                        block,
                        lines,
                        height,
                        index,
                        count,
                        opens_run,
                        end,
                    },
                );
                blocks.push(MeasuredBlock {
                    block: LayoutBlock::Reserved(ReservedBlock {
                        id: BlockId::Str(id.clone()),
                        reservation: ReservationRef {
                            owner: ENDNOTE_FLOW_OWNER.to_owned(),
                            id,
                        },
                        width: None,
                        height: block_height,
                        page_lead: (!opens_run).then_some(separator),
                    }),
                    measure: BlockExtent::Image(ImageExtent {
                        width,
                        height: block_height,
                    }),
                });
            }
        }
        if blocks.is_empty() {
            continue;
        }
        let at = match section {
            Some(section) => section_ends
                .get(section)
                .copied()
                .unwrap_or(document_end)
                .min(document_end),
            None => document_end,
        };
        insertions.push((at, blocks));
    }
    insertions.sort_by_key(|(at, _)| *at);
    for (at, blocks) in insertions.into_iter().rev() {
        measured.splice(at..at, blocks);
    }
    flow.inserted = measured
        .iter()
        .enumerate()
        .filter(|(_, entry)| is_flowed_block(&entry.block))
        .map(|(index, _)| index)
        .collect();
    flow
}

/// Removes the blocks [`flow_endnotes`] added, in place, from every list aligned with the flow.
pub fn remove_flowed<T>(items: &mut Vec<T>, inserted: &[usize]) {
    for &index in inserted.iter().rev() {
        if index < items.len() {
            items.remove(index);
        }
    }
}

struct Placed {
    x: f64,
    y: f64,
    width: f64,
    height: f64,
    piece: Piece,
}

/// Removes the flowed blocks' fragments from `layout`, returning each with its page and its
/// index among the page's fragments, so a later pass resuming from this layout can put them back.
pub fn take_flowed_fragments(
    layout: &mut Layout,
    flow: &EndnoteFlow,
) -> Vec<(usize, usize, Fragment)> {
    let mut taken = Vec::new();
    if flow.pieces.is_empty() {
        return taken;
    }
    for (page_index, page) in layout.pages.iter_mut().enumerate() {
        let mut index = 0;
        page.fragments.retain(|fragment| {
            let flowed = matches!(
                fragment,
                Fragment::Reserved(reserved)
                    if matches!(&reserved.block_id, BlockId::Str(id) if flow.pieces.contains_key(id))
            );
            if flowed {
                taken.push((page_index, index, fragment.clone()));
            }
            index += 1;
            !flowed
        });
    }
    taken
}

/// Builds each page's endnote area from where the flowed blocks landed on it.
pub fn attach_flowed_endnote_areas(
    layout: &mut Layout,
    flow: &EndnoteFlow,
    contents: &[NoteContent],
    separator_heights: &NoteSeparatorHeights,
) {
    if flow.pieces.is_empty() {
        return;
    }
    let contents = content_map(contents);
    for page in &mut layout.pages {
        let placed: Vec<Placed> = page
            .fragments
            .iter()
            .filter_map(|fragment| {
                let Fragment::Reserved(reserved) = fragment else {
                    return None;
                };
                let BlockId::Str(id) = &reserved.block_id else {
                    return None;
                };
                Some(Placed {
                    x: reserved.x,
                    y: reserved.y,
                    width: reserved.width,
                    height: reserved.height,
                    piece: flow.pieces.get(id)?.clone(),
                })
            })
            .collect();
        let columns = page
            .columns
            .as_ref()
            .is_some_and(|columns| columns.count > 1.0);
        let measured_separator = separator_heights
            .get_for_page(NoteKind::Endnote, page)
            .is_some();
        let mut areas = Vec::new();
        let mut ids = Vec::new();
        let mut start = 0;
        while start < placed.len() {
            let mut end = start + 1;
            while end < placed.len()
                && !placed[end].piece.opens_run
                && placed[end].x == placed[start].x
            {
                end += 1;
            }
            let segment = &placed[start..end];
            start = end;
            if let Some(area) = segment_area(
                segment,
                &contents,
                (columns, measured_separator),
                page.section_id.clone(),
                &mut ids,
            ) {
                areas.push(area);
            }
        }
        if areas.is_empty() {
            continue;
        }
        page.footnote_ids.get_or_insert_with(Vec::new).extend(ids);
        page.note_areas.get_or_insert_with(Vec::new).extend(areas);
    }
}

/// The note area of one run of flowed blocks in one column of a page: a run starts after the
/// body text of the document or a section, so two runs on a page have body text between them.
fn segment_area(
    segment: &[Placed],
    contents: &OrderedMap<i64, &NoteContent>,
    (columns, measured_separator): (bool, bool),
    section_id: Option<String>,
    ids: &mut Vec<f64>,
) -> Option<NoteAreaContract> {
    let first = segment.first()?;
    let y = first.y;
    let bottom = segment
        .iter()
        .map(|placed| placed.y + placed.height)
        .fold(y, f64::max);
    let separator = (first.height - first.piece.height).max(0.0);
    let mut notes: Vec<NoteLayoutItemContract> = Vec::new();
    let mut index = 0;
    while index < segment.len() {
        let map_id = segment[index].piece.map_id;
        let mut last = index;
        while last + 1 < segment.len() && segment[last + 1].piece.map_id == map_id {
            last += 1;
        }
        let (from, to) = (&segment[index].piece, &segment[last].piece);
        if let Some(content) = contents.get(&map_id).copied() {
            let mut item = note_item(content);
            item.height = Some(
                segment[index..=last]
                    .iter()
                    .map(|placed| placed.piece.height)
                    .sum(),
            );
            if from.index > 0 || to.index + 1 < to.count {
                item.block_start = Some(from.block);
                item.line_start = Some(from.lines.start);
                item.block_end = Some(to.block);
                item.line_end = Some(to.lines.end);
            }
            notes.push(item);
            ids.push(map_id as f64);
        }
        index = last + 1;
    }
    if notes.is_empty() {
        return None;
    }
    Some(NoteAreaContract {
        page_index: None,
        section_id,
        kind: Some(NoteKind::Endnote.as_str().to_owned()),
        placement: Some(first.piece.end.placement().to_owned()),
        y: Some(y),
        height: Some(bottom - y),
        columns: Some(1),
        separator: measured_separator.then(|| NoteLayoutItemContract {
            height: Some(separator),
            ..Default::default()
        }),
        notes: Some(notes),
        continuation: (!first.piece.opens_run).then_some(true),
        x: columns.then_some(first.x),
        width: columns.then_some(first.width),
    })
}

#[cfg(test)]
mod tests {
    use serde_json::json;

    use super::*;
    use crate::regions::{NoteProperties, NoteSettings};

    fn note(id: i64, lines: usize) -> NoteContent {
        let block: LayoutBlock = serde_json::from_value(json!({
            "kind": "paragraph",
            "id": format!("en{id}"),
            "runs": [{"kind": "text", "text": "note"}],
            "attrs": {"spacing": {"before": 4}}
        }))
        .unwrap();
        let measure: BlockExtent = serde_json::from_value(json!({
            "kind": "paragraph",
            "totalHeight": lines as f64 * 20.0 + 6.0,
            "lines": vec![json!({
                "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 4,
                "width": 40, "ascent": 15, "descent": 5, "lineHeight": 20
            }); lines]
        }))
        .unwrap();
        NoteContent {
            id,
            note_kind: NoteKind::Endnote,
            display_number: Some(id),
            display_label: None,
            blocks: vec![block],
            measures: vec![measure],
            height: lines as f64 * 20.0 + 6.0,
            anchor: None,
            custom_mark_follows: None,
        }
    }

    #[test]
    fn widow_control_keeps_two_lines_at_both_ends_of_a_split() {
        let slices = note_slices(&note(1, 6), f64::INFINITY);
        let ranges: Vec<_> = slices.iter().map(|(_, lines, _)| lines.clone()).collect();
        assert_eq!(ranges, [0..2, 2..3, 3..4, 4..6]);
        let total: f64 = slices.iter().map(|(_, _, height)| height).sum();
        assert_eq!(total, 4.0 + 6.0 * 20.0 + 6.0);
        assert_eq!(note_slices(&note(2, 3), f64::INFINITY).len(), 1);
        assert_eq!(note_slices(&note(2, 3), 40.0).len(), 3);
    }

    #[test]
    fn page_notes_do_not_flow() {
        let regions = DocumentRegions {
            note_settings: NoteSettings {
                endnote: NoteProperties {
                    position: Some("sectEnd".to_owned()),
                    ..Default::default()
                },
                ..Default::default()
            },
            ..Default::default()
        };
        assert_eq!(endnote_flow_end(&regions, 0), Some(FlowEnd::Section));
        assert_eq!(
            endnote_flow_end(&DocumentRegions::default(), 0),
            Some(FlowEnd::Document)
        );
        assert!(!flows_with_body(NoteKind::Footnote, &regions, 0));
    }
}
