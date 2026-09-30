//! Aggregated read-state queries for host UI surfaces.

use std::collections::HashSet;

use yrs::{Any, Map, Out, ReadTxn, TextRef, Transact};

use crate::op::{OpError, OpResult, para_bounds};
use crate::ops::{Chunk, ChunkKind, capture_pilcrow};
use crate::queries::TextView;
use crate::{
    ChangeInfo, EditingDoc, KIND_KEY, ParagraphId, StoryId, StoryRange, map_string, story_ref,
};

/// Preview cap for [`RevisionInfo::preview`], in Unicode scalar values.
const PREVIEW_MAX_CHARS: usize = 80;

/// One toggle mark aggregated over a range: active on every text unit
/// ([`TriState::On`]), on none ([`TriState::Off`]), or on some ([`TriState::Mixed`]).
/// A range without any text unit reports [`TriState::Off`].
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TriState {
    On,
    Off,
    Mixed,
}

impl TriState {
    fn fold(current: Option<TriState>, active: bool) -> Option<TriState> {
        Some(match (current, active) {
            (None | Some(TriState::On), true) => TriState::On,
            (None | Some(TriState::Off), false) => TriState::Off,
            _ => TriState::Mixed,
        })
    }
}

/// Uniform-or-mixed aggregation of one attribute value across text units.
enum ValueAgg {
    /// No text unit seen yet.
    Empty,
    /// Every text unit so far agrees on this value (`None` = attribute absent).
    Uniform(Option<Any>),
    Mixed,
}

impl ValueAgg {
    fn fold(&mut self, value: Option<&Any>) {
        // Explicit `Any::Null` is the schema's "attribute removed" marker —
        // normalize it to absent so `bold:null` text and untouched text agree.
        let value = value.filter(|value| **value != Any::Null);
        match self {
            Self::Empty => *self = Self::Uniform(value.cloned()),
            Self::Uniform(current) if current.as_ref() == value => {}
            Self::Uniform(_) => *self = Self::Mixed,
            Self::Mixed => {}
        }
    }

    fn uniform(&self) -> Option<&Any> {
        match self {
            Self::Uniform(value) => value.as_ref(),
            _ => None,
        }
    }
}

/// Aggregated selection state over one story range.
#[derive(Clone, Debug, PartialEq)]
pub struct SelectionContextInfo {
    // Tri-state inline marks.
    pub bold: TriState,
    pub italic: TriState,
    pub underline: TriState,
    pub strike: TriState,
    pub superscript: TriState,
    pub subscript: TriState,
    // -- uniform-or-null value marks --
    /// The uniform `fontFamily.ascii`, or `None` when mixed/absent.
    pub font_family: Option<String>,
    /// Uniform font size in half-points, or `None` when mixed or absent.
    pub font_size: Option<f64>,
    /// The uniform text color: the `rgb` hex when set, else the theme color
    /// name; `None` when mixed/absent.
    pub color: Option<String>,
    /// The uniform highlight color name, or `None` when mixed/absent.
    pub highlight: Option<String>,
    // Paragraph state at the range start.
    pub para_id: ParagraphId,
    /// The paragraph's `pStyle`, extracted from `paragraph_properties`.
    pub style_id: Option<String>,
    /// The paragraph's `alignment`, extracted from `paragraph_properties`.
    pub alignment: Option<String>,
    /// The full pilcrow property bag (indent, spacing, `numPr`, and any other
    /// op-set extras) excluding schema identity keys.
    pub paragraph_properties: std::collections::BTreeMap<String, Any>,
    // -- flags --
    /// `start != end`.
    pub has_selection: bool,
    /// The range's ends sit in different paragraphs.
    pub is_multi_paragraph: bool,
    /// The range's story is a table-cell story (referenced from a `table`
    /// embed anywhere in the document).
    pub in_table: bool,
    /// `Some(kind)` when the range covers exactly one non-pilcrow embed unit
    /// (image, drawing, …); the embed's `_kind` discriminator.
    pub embed_kind: Option<String>,
    /// Every aggregated text unit carries a pending `ins` revision.
    pub in_insertion: bool,
    /// Every aggregated text unit carries a pending `del` revision.
    pub in_deletion: bool,
}

/// One tracked-change revision entry from [`EditingDoc::list_revisions`].
///
/// Entries mirror [`EditingDoc::list_changes`] (adjacent same-revision runs
/// coalesced, ordered by position within each story); a revision whose runs
/// are non-adjacent yields one entry per run.
#[derive(Clone, Debug, PartialEq)]
pub struct RevisionInfo {
    pub story: StoryId,
    pub change: ChangeInfo,
    /// The raw text under the change's range (deleted text included), capped
    /// at `PREVIEW_MAX_CHARS` characters. Empty for paragraph-mark revisions.
    pub preview: String,
}

/// Collects every story id referenced as a cell story by a `table` embed
/// (payload `rows[*].cells[*].story`). Nested tables are covered because a
/// nested table's embed lives in a cell story that is itself iterated.
pub(crate) fn table_cell_stories<T: ReadTxn>(doc: &EditingDoc, txn: &T) -> HashSet<String> {
    let mut cells = HashSet::new();
    let Some(stories) = txn.get_map(crate::STORIES) else {
        return cells;
    };
    for (story_id, value) in stories.iter(txn) {
        let Out::YText(story) = value else {
            continue;
        };
        collect_table_cell_stories(doc, txn, story_id, &story, &mut cells);
    }
    cells
}

/// Whether `story` is a table cell story. Cell stories are named
/// `{parent}:t{table}:r{row}c{cell}`, so that parent's tables are read first;
/// only a story they do not list costs the whole-document scan.
pub(crate) fn is_table_cell_story<T: ReadTxn>(doc: &EditingDoc, txn: &T, story: &str) -> bool {
    if let Some(parent) = cell_story_parent(story)
        && let Some(Out::YText(parent_story)) = txn
            .get_map(crate::STORIES)
            .and_then(|stories| stories.get(txn, parent))
    {
        let mut cells = HashSet::new();
        collect_table_cell_stories(doc, txn, parent, &parent_story, &mut cells);
        if cells.contains(story) {
            return true;
        }
    }
    table_cell_stories(doc, txn).contains(story)
}

fn cell_story_parent(story: &str) -> Option<&str> {
    let (rest, cell) = story.rsplit_once(':')?;
    let (parent, table) = rest.rsplit_once(':')?;
    let digits = |value: &str| !value.is_empty() && value.bytes().all(|byte| byte.is_ascii_digit());
    let (row, column) = cell.strip_prefix('r')?.split_once('c')?;
    (digits(table.strip_prefix('t')?) && digits(row) && digits(column)).then_some(parent)
}

fn collect_table_cell_stories<T: ReadTxn>(
    doc: &EditingDoc,
    txn: &T,
    story_id: &str,
    story: &TextRef,
    cells: &mut HashSet<String>,
) {
    for chunk in doc.chunk_snapshot(story_id, story, txn).iter() {
        let ChunkKind::Embed(Some(map)) = &chunk.kind else {
            continue;
        };
        if map_string(map, txn, KIND_KEY).as_deref() != Some("table") {
            continue;
        }
        let Some(Out::Any(Any::Array(rows))) = map.get(txn, "rows") else {
            continue;
        };
        for row in rows.iter() {
            let Any::Map(row) = row else {
                continue;
            };
            let Some(Any::Array(row_cells)) = row.get("cells") else {
                continue;
            };
            for cell in row_cells.iter() {
                let Any::Map(cell) = cell else {
                    continue;
                };
                if let Some(Any::String(story_id)) = cell.get("story") {
                    cells.insert(story_id.to_string());
                }
            }
        }
    }
}

impl EditingDoc {
    /// Aggregates text-unit marks and start-paragraph state over a story range.
    pub fn selection_context(&self, range: &StoryRange) -> OpResult<SelectionContextInfo> {
        if range.end < range.start {
            return Err(OpError::InvalidRange {
                start: range.start,
                end: range.end,
            });
        }
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, &range.story)?;
        let chunks = self.chunk_snapshot(&range.story, &story, &txn);
        let story_len = chunks.last().map_or(0, Chunk::end);
        if range.end > story_len {
            return Err(OpError::OutOfBounds {
                index: range.end,
                len: story_len,
            });
        }

        let bounds = para_bounds(&story, &txn);
        let start_para = bounds
            .iter()
            .find(|para| range.start <= para.pilcrow)
            .ok_or_else(|| OpError::UnknownStory(range.story.clone()))?;
        let end_para = bounds
            .iter()
            .find(|para| range.end <= para.pilcrow)
            .or_else(|| bounds.last())
            .ok_or_else(|| OpError::UnknownStory(range.story.clone()))?;
        let is_multi_paragraph = start_para.para_id != end_para.para_id;

        // Paragraph properties from the start paragraph's pilcrow map.
        let (para_id, para_props) = chunks
            .iter()
            .find_map(|chunk| match &chunk.kind {
                ChunkKind::Pilcrow(map) if chunk.start == start_para.pilcrow => {
                    Some(capture_pilcrow(map, &txn))
                }
                _ => None,
            })
            .ok_or_else(|| OpError::ExpectedPilcrow {
                story: range.story.clone(),
                index: start_para.pilcrow,
            })?;
        let paragraph_properties: std::collections::BTreeMap<String, Any> = para_props
            .into_iter()
            .filter(|(key, _)| !crate::is_identity_key(key))
            .collect();
        let prop_string = |key: &str| match paragraph_properties.get(key) {
            Some(Any::String(value)) => Some(value.to_string()),
            _ => None,
        };

        let is_text_unit = |index: u32| {
            chunks
                .iter()
                .find(|chunk| chunk.start <= index && index < chunk.end())
                .is_some_and(|chunk| matches!(chunk.kind, ChunkKind::Text(_)))
        };

        // The effective mark range: the range itself, or the caret-adjacent
        // text unit (before within the paragraph, else after).
        let (mark_from, mark_to) = if range.start == range.end {
            let at = range.start;
            if at > start_para.start && is_text_unit(at - 1) {
                (at - 1, at)
            } else if at < start_para.pilcrow && is_text_unit(at) {
                (at, at + 1)
            } else {
                (at, at)
            }
        } else {
            (range.start, range.end)
        };

        let mut bold = None;
        let mut italic = None;
        let mut underline = None;
        let mut strike = None;
        let mut superscript = None;
        let mut subscript = None;
        let mut ins = None;
        let mut del = None;
        let mut font_family = ValueAgg::Empty;
        let mut font_size = ValueAgg::Empty;
        let mut color = ValueAgg::Empty;
        let mut highlight = ValueAgg::Empty;
        for chunk in chunks.iter() {
            if chunk.start >= mark_to {
                break;
            }
            if chunk.end() <= mark_from || !matches!(chunk.kind, ChunkKind::Text(_)) {
                continue;
            }
            bold = TriState::fold(bold, chunk.attr_active("bold"));
            italic = TriState::fold(italic, chunk.attr_active("italic"));
            underline = TriState::fold(underline, chunk.attr_active("underline"));
            strike = TriState::fold(strike, chunk.attr_active("strike"));
            superscript = TriState::fold(superscript, chunk.attr_active("superscript"));
            subscript = TriState::fold(subscript, chunk.attr_active("subscript"));
            ins = TriState::fold(ins, chunk.attr_active(crate::INS));
            del = TriState::fold(del, chunk.attr_active(crate::DEL));
            font_family.fold(chunk.attrs.get("fontFamily"));
            font_size.fold(chunk.attrs.get("fontSize"));
            color.fold(chunk.attrs.get("textColor"));
            highlight.fold(chunk.attrs.get("highlight"));
        }

        let map_field = |value: Option<&Any>, key: &str| match value {
            Some(Any::Map(map)) => map.get(key).cloned(),
            _ => None,
        };
        let font_family = match (
            map_field(font_family.uniform(), "ascii"),
            map_field(font_family.uniform(), "hAnsi"),
        ) {
            (Some(Any::String(ascii)), _) => Some(ascii.to_string()),
            (_, Some(Any::String(h_ansi))) => Some(h_ansi.to_string()),
            _ => None,
        };
        let font_size = match (
            map_field(font_size.uniform(), "size"),
            map_field(font_size.uniform(), "sizeCs"),
        ) {
            // Font size remains in half-points.
            (Some(Any::Number(half_points)), _) => Some(half_points),
            (_, Some(Any::Number(half_points_cs))) => Some(half_points_cs),
            _ => None,
        };
        let color = match (
            map_field(color.uniform(), "rgb"),
            map_field(color.uniform(), "themeColor"),
        ) {
            (Some(Any::String(rgb)), _) => Some(rgb.to_string()),
            (_, Some(Any::String(theme))) => Some(theme.to_string()),
            _ => None,
        };
        let highlight = match map_field(highlight.uniform(), "color") {
            Some(Any::String(name)) => Some(name.to_string()),
            _ => None,
        };

        let embed_kind = if range.end == range.start + 1 {
            chunks
                .iter()
                .find(|chunk| chunk.start <= range.start && range.start < chunk.end())
                .and_then(|chunk| match &chunk.kind {
                    ChunkKind::Embed(Some(map)) => {
                        Some(map_string(map, &txn, KIND_KEY).unwrap_or_default())
                    }
                    ChunkKind::Embed(None) => Some(String::new()),
                    _ => None,
                })
        } else {
            None
        };

        Ok(SelectionContextInfo {
            bold: bold.unwrap_or(TriState::Off),
            italic: italic.unwrap_or(TriState::Off),
            underline: underline.unwrap_or(TriState::Off),
            strike: strike.unwrap_or(TriState::Off),
            superscript: superscript.unwrap_or(TriState::Off),
            subscript: subscript.unwrap_or(TriState::Off),
            font_family,
            font_size,
            color,
            highlight,
            para_id,
            style_id: prop_string("pStyle"),
            alignment: prop_string("alignment"),
            paragraph_properties,
            has_selection: range.start != range.end,
            is_multi_paragraph,
            in_table: is_table_cell_story(self, &txn, &range.story),
            embed_kind,
            in_insertion: ins == Some(TriState::On),
            in_deletion: del == Some(TriState::On),
        })
    }

    /// Enumerates every tracked-change revision in the document, across all
    /// stories in sorted story-id order (see [`RevisionInfo`]).
    pub fn list_revisions(&self) -> OpResult<Vec<RevisionInfo>> {
        let story_ids: Vec<String> = {
            let txn = self.yrs_doc().transact();
            let Some(stories) = txn.get_map(crate::STORIES) else {
                return Ok(Vec::new());
            };
            let mut ids: Vec<String> = stories.keys(&txn).map(|key| key.to_string()).collect();
            ids.sort();
            ids
        };
        let mut result = Vec::new();
        for story_id in story_ids {
            let changes = self.story_changes(&story_id)?;
            if changes.is_empty() {
                continue;
            }
            let txn = self.yrs_doc().transact();
            let story = crate::story_ref(&txn, &story_id)?;
            let bounds = crate::op::para_bounds(&story, &txn);
            let views = crate::queries::para_views(
                &txn,
                TextView::Raw,
                &self.chunk_snapshot(&story_id, &story, &txn),
            );
            for (change, _) in changes {
                let preview = if matches!(
                    change.kind,
                    crate::ChangeKind::ParagraphMarkInsertion
                        | crate::ChangeKind::ParagraphMarkDeletion
                        | crate::ChangeKind::ParagraphPropertiesChanged
                        | crate::ChangeKind::TableRowInsertion
                        | crate::ChangeKind::TableRowDeletion
                        | crate::ChangeKind::TableInsertion
                        | crate::ChangeKind::TableDeletion
                ) {
                    String::new()
                } else {
                    let from = crate::op::global_in_bounds(&bounds, &change.range.start)?;
                    let to = crate::op::global_in_bounds(&bounds, &change.range.end)?;
                    if to < from {
                        return Err(OpError::InvalidRange {
                            start: from,
                            end: to,
                        });
                    }
                    let mut full = String::new();
                    for para in &views {
                        para.view_slice_of_raw(from, to, &mut full);
                    }
                    full.chars().take(PREVIEW_MAX_CHARS).collect()
                };
                result.push(RevisionInfo {
                    story: story_id.clone(),
                    change,
                    preview,
                });
            }
        }
        Ok(result)
    }
}

#[cfg(test)]
mod tests {
    use std::collections::{BTreeMap, BTreeSet, HashMap};
    use std::sync::Arc;

    use super::*;
    use crate::{
        ChangeKind, ColorPatch, EditCtx, FontFamilyPatch, FormatPolicy, InlineFormatDelta, Patch,
        Position, RawOp, SimpleFormat,
    };

    const DATE: &str = "2026-07-14T12:00:00Z";

    fn local() -> EditCtx {
        EditCtx::local("Local", DATE)
    }

    fn suggesting(author: &str) -> EditCtx {
        EditCtx::local(author, DATE).suggesting()
    }

    fn seed(text: &str) -> EditingDoc {
        let doc = EditingDoc::new(7);
        doc.create_story("body", text, "Normal", "left").unwrap();
        doc
    }

    fn context(doc: &EditingDoc, start: u32, end: u32) -> SelectionContextInfo {
        doc.selection_context(&StoryRange::new("body", start, end))
            .unwrap()
    }

    #[test]
    fn bold_is_mixed_over_a_spanning_range_and_on_over_the_bold_range() {
        let doc = seed("hello world");
        doc.toggle_format(&local(), StoryRange::new("body", 0, 5), SimpleFormat::Bold)
            .unwrap();

        assert_eq!(context(&doc, 0, 11).bold, TriState::Mixed);
        assert_eq!(context(&doc, 0, 5).bold, TriState::On);
        assert_eq!(context(&doc, 5, 11).bold, TriState::Off);
        // Untouched toggles stay Off on every range.
        assert_eq!(context(&doc, 0, 11).italic, TriState::Off);
    }

    #[test]
    fn caret_reads_the_preceding_unit_and_the_following_at_paragraph_start() {
        let doc = seed("hello world");
        doc.toggle_format(&local(), StoryRange::new("body", 0, 5), SimpleFormat::Bold)
            .unwrap();

        // Caret after the bold run reads the unit before it.
        assert_eq!(context(&doc, 5, 5).bold, TriState::On);
        // Caret at paragraph start reads the unit after it.
        assert_eq!(context(&doc, 0, 0).bold, TriState::On);
        // Caret deep in the plain tail.
        assert_eq!(context(&doc, 8, 8).bold, TriState::Off);
    }

    #[test]
    fn value_marks_report_the_uniform_value_and_null_when_mixed() {
        let doc = seed("hello world");
        let delta = InlineFormatDelta {
            font_family: Patch::Set(FontFamilyPatch {
                ascii: "Georgia".into(),
                h_ansi: None,
            }),
            font_size: Patch::Set(14.0),
            color: Patch::Set(ColorPatch::Rgb("336699".into())),
            ..Default::default()
        };
        doc.format_range(&local(), StoryRange::new("body", 0, 5), &delta)
            .unwrap();

        let styled = context(&doc, 0, 5);
        assert_eq!(styled.font_family.as_deref(), Some("Georgia"));
        assert_eq!(styled.font_size, Some(28.0));
        assert_eq!(styled.color.as_deref(), Some("336699"));

        let spanning = context(&doc, 0, 11);
        assert_eq!(spanning.font_family, None);
        assert_eq!(spanning.font_size, None);
        assert_eq!(spanning.color, None);

        let plain = context(&doc, 6, 11);
        assert_eq!(plain.font_family, None);
        assert_eq!(plain.font_size, None);
    }

    #[test]
    fn script_marks_are_tri_state_and_mutually_exclusive() {
        let doc = seed("hello world");
        doc.toggle_format(
            &local(),
            StoryRange::new("body", 0, 5),
            SimpleFormat::Superscript,
        )
        .unwrap();

        let raised = context(&doc, 0, 5);
        assert_eq!(raised.superscript, TriState::On);
        assert_eq!(raised.subscript, TriState::Off);
        assert_eq!(context(&doc, 0, 11).superscript, TriState::Mixed);
        assert_eq!(context(&doc, 3, 3).superscript, TriState::On);

        doc.toggle_format(
            &local(),
            StoryRange::new("body", 0, 5),
            SimpleFormat::Subscript,
        )
        .unwrap();
        let lowered = context(&doc, 0, 5);
        assert_eq!(lowered.superscript, TriState::Off);
        assert_eq!(lowered.subscript, TriState::On);
    }

    #[test]
    fn highlight_reports_the_uniform_palette_name() {
        let doc = seed("hello world");
        let delta = InlineFormatDelta {
            highlight: Patch::Set("FFFF00".into()),
            ..Default::default()
        };
        doc.format_range(&local(), StoryRange::new("body", 0, 5), &delta)
            .unwrap();

        assert_eq!(context(&doc, 0, 5).highlight.as_deref(), Some("yellow"));
        assert_eq!(context(&doc, 0, 11).highlight, None);
        assert_eq!(context(&doc, 6, 11).highlight, None);
    }

    #[test]
    fn paragraph_state_and_multi_paragraph_flag() {
        let doc = seed("first second");
        let split = doc
            .split_paragraph(&local(), Position::new("body", 5), None)
            .unwrap();
        doc.set_paragraph_attr(&split.first_para_id, "pStyle", Any::from("Heading1"))
            .unwrap();
        doc.set_paragraph_attr(&split.first_para_id, "indentLeft", Any::Number(720.0))
            .unwrap();

        let first = context(&doc, 0, 5);
        assert_eq!(first.para_id, split.first_para_id);
        assert_eq!(first.style_id.as_deref(), Some("Heading1"));
        assert_eq!(first.alignment.as_deref(), Some("left"));
        assert_eq!(
            first.paragraph_properties.get("indentLeft"),
            Some(&Any::Number(720.0))
        );
        assert!(!first.is_multi_paragraph);
        assert!(first.has_selection);

        // [2, 8) crosses the pilcrow at 5 into the second paragraph.
        let spanning = context(&doc, 2, 8);
        assert!(spanning.is_multi_paragraph);
        assert_eq!(spanning.para_id, split.first_para_id);

        let second = context(&doc, 7, 12);
        assert_eq!(second.para_id, split.second_para_id);
        assert_eq!(second.style_id.as_deref(), Some("Normal"));
        assert!(!second.is_multi_paragraph);
    }

    #[test]
    fn tracked_change_flags_and_single_embed() {
        let doc = seed("abc");
        doc.insert_text(
            &suggesting("Reviewer"),
            Position::new("body", 3),
            "NEW",
            FormatPolicy::Plain,
        )
        .unwrap();
        let inserted = context(&doc, 3, 6);
        assert!(inserted.in_insertion);
        assert!(!inserted.in_deletion);
        assert!(!context(&doc, 0, 3).in_insertion);

        doc.apply_raw_ops(
            "body",
            vec![RawOp::InsertEmbed {
                index: 0,
                kind: "image".into(),
                payload: vec![("src".into(), Any::from("media/image1.png"))],
                attrs: yrs::types::Attrs::new(),
            }],
            &local(),
        )
        .unwrap();
        let embed = context(&doc, 0, 1);
        assert_eq!(embed.embed_kind.as_deref(), Some("image"));
        // A wider range is not a single embed.
        assert_eq!(context(&doc, 0, 2).embed_kind, None);
        // Selecting only a pilcrow is not an embed selection.
        let pilcrow_index = doc.story_len("body").unwrap() - 1;
        assert_eq!(
            context(&doc, pilcrow_index, pilcrow_index + 1).embed_kind,
            None
        );
    }

    #[test]
    fn in_table_flags_ranges_inside_a_cell_story() {
        let doc = seed("body text");
        doc.create_story("body:t0:r0c0", "cell", "Normal", "left")
            .unwrap();
        let cell = Any::Map(Arc::new(HashMap::from([(
            "story".into(),
            Any::from("body:t0:r0c0"),
        )])));
        let row = Any::Map(Arc::new(HashMap::from([(
            "cells".into(),
            Any::Array(Arc::from(vec![cell])),
        )])));
        doc.apply_raw_ops(
            "body",
            vec![RawOp::InsertEmbed {
                index: 0,
                kind: "table".into(),
                payload: vec![("rows".into(), Any::Array(Arc::from(vec![row])))],
                attrs: yrs::types::Attrs::new(),
            }],
            &local(),
        )
        .unwrap();

        assert!(
            doc.selection_context(&StoryRange::new("body:t0:r0c0", 0, 4))
                .unwrap()
                .in_table
        );
        assert!(!context(&doc, 1, 4).in_table);
    }

    #[test]
    fn in_table_reads_the_tables_that_list_the_cell() {
        let doc = seed("body text");
        let table = |doc: &EditingDoc, parent: &str, cell_story: &str| {
            let cell = Any::Map(Arc::new(HashMap::from([(
                "story".into(),
                Any::from(cell_story),
            )])));
            let row = Any::Map(Arc::new(HashMap::from([(
                "cells".into(),
                Any::Array(Arc::from(vec![cell])),
            )])));
            doc.apply_raw_ops(
                parent,
                vec![RawOp::InsertEmbed {
                    index: 0,
                    kind: "table".into(),
                    payload: vec![("rows".into(), Any::Array(Arc::from(vec![row])))],
                    attrs: yrs::types::Attrs::new(),
                }],
                &local(),
            )
            .unwrap();
        };
        for story in [
            "body:t0:r0c0",
            "body:t0:r0c0:t0:r0c0",
            "body:t1:r0c0",
            "custom-cell",
        ] {
            doc.create_story(story, "cell", "Normal", "left").unwrap();
        }
        table(&doc, "body", "body:t0:r0c0");
        table(&doc, "body:t0:r0c0", "body:t0:r0c0:t0:r0c0");
        table(&doc, "body:t0:r0c0", "custom-cell");
        let in_table = |story: &str| {
            doc.selection_context(&StoryRange::new(story, 0, 1))
                .unwrap()
                .in_table
        };
        assert!(in_table("body:t0:r0c0"));
        assert!(in_table("body:t0:r0c0:t0:r0c0"), "a nested table's cell");
        assert!(in_table("custom-cell"), "a cell whose name names no parent");
        assert!(!in_table("body:t1:r0c0"), "no table lists it");
        assert!(!in_table("body"));
    }

    #[test]
    fn suggested_insert_and_delete_show_up_in_list_revisions() {
        let doc = seed("alpha beta");
        let insert = doc
            .insert_text(
                &suggesting("Alice"),
                Position::new("body", 10),
                " INSERTED",
                FormatPolicy::Plain,
            )
            .unwrap();
        let delete = doc
            .delete_range(&suggesting("Bob"), StoryRange::new("body", 0, 5))
            .unwrap();

        let revisions = doc.list_revisions().unwrap();
        assert_eq!(revisions.len(), 2);

        let deletion = &revisions[0];
        assert_eq!(deletion.change.kind, ChangeKind::Deletion);
        assert_eq!(deletion.change.author, "Bob");
        assert_eq!(deletion.change.date, DATE);
        assert_eq!(deletion.change.revision_id, delete.revision_ids[0]);
        assert_eq!(deletion.story, "body");
        assert_eq!(deletion.preview, "alpha");

        let insertion = &revisions[1];
        assert_eq!(insertion.change.kind, ChangeKind::Insertion);
        assert_eq!(insertion.change.author, "Alice");
        assert_eq!(insertion.change.revision_id, insert.revision_ids[0]);
        assert_eq!(insertion.preview, " INSERTED");

        // Plain edits never appear.
        let plain = seed("plain");
        plain
            .insert_text(&local(), Position::new("body", 5), "!", FormatPolicy::Plain)
            .unwrap();
        assert!(plain.list_revisions().unwrap().is_empty());
    }

    #[test]
    fn revision_stamps_include_dates_from_adjacent_same_author_insertions() {
        let doc = seed("alpha");
        let first = doc
            .insert_text(
                &suggesting("Alice"),
                Position::new("body", 5),
                " first",
                FormatPolicy::Plain,
            )
            .unwrap();
        let later = "2026-07-14T12:01:00Z";
        let second = doc
            .insert_text(
                &EditCtx::local("Alice", later).suggesting(),
                Position::new("body", 11),
                " second",
                FormatPolicy::Plain,
            )
            .unwrap();
        assert_eq!(second.revision_ids, first.revision_ids);
        let listed = doc.list_revisions().unwrap();
        assert_eq!(listed.len(), 1);
        assert_eq!(listed[0].change.date, DATE);
        let id = first.revision_ids[0].clone();
        assert_eq!(
            doc.revision_stamps(&[id.clone(), "missing".into()])
                .unwrap(),
            BTreeMap::from([(
                id,
                BTreeSet::from([
                    ("Alice".into(), DATE.into()),
                    ("Alice".into(), later.into()),
                ]),
            )])
        );
    }

    #[test]
    fn revision_stamps_collect_text_and_paragraph_stamps_across_stories() {
        let doc = seed("alpha");
        doc.create_story("header", "x", "Normal", "left").unwrap();
        let stamp = |author: &str| {
            Any::Map(Arc::new(HashMap::from([
                ("id".into(), Any::from("shared")),
                ("author".into(), Any::from(author)),
                ("date".into(), Any::from(DATE)),
            ])))
        };
        doc.apply_raw_ops(
            "body",
            vec![RawOp::Format {
                index: 0,
                len: 5,
                attrs: yrs::types::Attrs::from([
                    (crate::INS.into(), stamp("Insert")),
                    (crate::DEL.into(), stamp("Delete")),
                ]),
            }],
            &local(),
        )
        .unwrap();
        doc.apply_raw_ops(
            "header",
            vec![
                RawOp::SetEmbedAttr {
                    index: 1,
                    key: crate::PPR_INS.into(),
                    value: stamp("Mark insert"),
                },
                RawOp::SetEmbedAttr {
                    index: 1,
                    key: crate::PPR_DEL.into(),
                    value: stamp("Mark delete"),
                },
                RawOp::SetEmbedAttr {
                    index: 1,
                    key: crate::PPR_CHANGE.into(),
                    value: Any::Array(Arc::from(vec![
                        stamp("Properties"),
                        stamp("Properties again"),
                    ])),
                },
            ],
            &local(),
        )
        .unwrap();
        let expected: BTreeSet<(String, String)> = [
            "Insert",
            "Delete",
            "Mark insert",
            "Mark delete",
            "Properties",
            "Properties again",
        ]
        .into_iter()
        .map(|author| (author.to_owned(), DATE.to_owned()))
        .collect();
        assert_eq!(
            doc.revision_stamps(&["shared".into()]).unwrap(),
            BTreeMap::from([("shared".into(), expected)])
        );
    }

    #[test]
    fn listed_revisions_match_their_per_change_reads() {
        let doc = seed("alpha beta gamma delta epsilon zeta eta theta iota kappa");
        for at in [45, 36, 27, 18, 11, 5] {
            doc.split_paragraph(&local(), Position::new("body", at), None)
                .unwrap();
        }
        doc.insert_text(
            &suggesting("Alice"),
            Position::new("body", 2),
            "one",
            FormatPolicy::Plain,
        )
        .unwrap();
        doc.delete_range(&suggesting("Bob"), StoryRange::new("body", 14, 18))
            .unwrap();
        doc.split_paragraph(&suggesting("Carol"), Position::new("body", 26), None)
            .unwrap();
        doc.insert_text(
            &suggesting("Dan"),
            Position::new("body", 40),
            "two",
            FormatPolicy::Plain,
        )
        .unwrap();
        let last = doc
            .locate(&doc.loc_at(&Position::new("body", u32::MAX)).unwrap())
            .unwrap();
        doc.delete_range(
            &suggesting("Eve"),
            StoryRange::new("body", last.index - 3, last.index),
        )
        .unwrap();

        let revisions = doc.list_revisions().unwrap();
        assert!(revisions.len() >= 5);
        for revision in &revisions {
            let range = &revision.change.range;
            for loc in [&range.start, &range.end] {
                let position = doc.locate(loc).unwrap();
                assert_eq!(&doc.loc_at(&position).unwrap(), loc);
            }
            let preview: String = if matches!(
                revision.change.kind,
                ChangeKind::ParagraphMarkInsertion | ChangeKind::ParagraphMarkDeletion
            ) {
                String::new()
            } else {
                doc.text_between(range, TextView::Raw)
                    .unwrap()
                    .chars()
                    .take(PREVIEW_MAX_CHARS)
                    .collect()
            };
            assert_eq!(revision.preview, preview, "{revision:?}");
        }

        let txn = doc.yrs_doc().transact();
        let story = crate::story_ref(&txn, "body").unwrap();
        let bounds = para_bounds(&story, &txn);
        let len = bounds.last().unwrap().pilcrow;
        for index in 0..=len + 2 {
            let naive = bounds
                .iter()
                .find(|bounds| index <= bounds.pilcrow)
                .unwrap_or(bounds.last().unwrap());
            let loc = crate::op::loc_in_bounds("body", &bounds, index).unwrap();
            assert_eq!(loc.para, naive.para_id, "index {index}");
            assert_eq!(
                loc.offset,
                index
                    .min(naive.pilcrow)
                    .saturating_sub(naive.start)
                    .min(naive.len())
            );
        }
    }

    #[test]
    fn paragraph_mark_revision_is_listed_with_its_kind() {
        let doc = seed("one two");
        doc.split_paragraph(&suggesting("Alice"), Position::new("body", 3), None)
            .unwrap();
        let revisions = doc.list_revisions().unwrap();
        assert_eq!(revisions.len(), 1);
        assert_eq!(revisions[0].change.kind, ChangeKind::ParagraphMarkInsertion);
        assert_eq!(revisions[0].change.author, "Alice");
        assert_eq!(revisions[0].preview, "");
    }

    #[test]
    fn undo_and_redo_depths_track_the_stacks() {
        let doc = seed("abc");
        let mut undo = doc.undo_manager();
        assert_eq!((undo.undo_depth(), undo.redo_depth()), (0, 0));

        doc.insert_text(&local(), Position::new("body", 3), "!", FormatPolicy::Plain)
            .unwrap();
        assert_eq!(undo.undo_depth(), 1);
        assert_eq!(undo.redo_depth(), 0);

        assert!(undo.undo());
        assert_eq!(undo.undo_depth(), 0);
        assert_eq!(undo.redo_depth(), 1);

        assert!(undo.redo());
        assert_eq!((undo.undo_depth(), undo.redo_depth()), (1, 0));
    }
}
