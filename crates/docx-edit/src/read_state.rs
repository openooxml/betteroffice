//! Aggregated READ-state queries for host UI surfaces (toolbar, a11y live
//! region, tracked-changes sidebar): pure snapshots, no mutation.
//!
//! [`EditingDoc::selection_context`] is the yrs twin of the PM
//! `selectionTracker` plugin's `SelectionContext`: it folds the story's
//! formatted segments over a range into tri-state toggle marks
//! (all-on / all-off / mixed), uniform-or-null value marks, the paragraph
//! properties at the range start, and structural flags (multi-paragraph,
//! table membership, single-embed). [`EditingDoc::list_revisions`] is the
//! doc-wide enumeration behind the sidebar's tracked-changes list — one
//! entry per coalesced `ins`/`del` run or paragraph-mark revision, across
//! every story, with a short raw-text preview.

use std::collections::HashSet;

use yrs::{Any, Map, Out, ReadTxn, Transact};

use crate::op::{LocRange, OpError, OpResult, para_bounds};
use crate::ops::{Chunk, ChunkKind, capture_pilcrow, snapshot};
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

/// The aggregated selection state over one story range — the yrs source for
/// what the PM `selectionTracker` derives from marks today.
#[derive(Clone, Debug, PartialEq)]
pub struct SelectionContextInfo {
    // -- tri-state toggle marks (PM toggleMark presence semantics) --
    pub bold: TriState,
    pub italic: TriState,
    pub underline: TriState,
    pub strike: TriState,
    // -- uniform-or-null value marks --
    /// The uniform `fontFamily.ascii`, or `None` when mixed/absent.
    pub font_family: Option<String>,
    /// The uniform font size in half-points (the PM `TextFormatting.fontSize`
    /// contract and stored `size` attr), or `None` when mixed/absent.
    pub font_size: Option<f64>,
    /// The uniform text color: the `rgb` hex when set, else the theme color
    /// name; `None` when mixed/absent.
    pub color: Option<String>,
    // -- paragraph state at the range start (PM `$from.parent` semantics) --
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
    /// at [`PREVIEW_MAX_CHARS`] characters. Empty for paragraph-mark revisions.
    pub preview: String,
}

/// Collects every story id referenced as a cell story by a `table` embed
/// (payload `rows[*].cells[*].story`). Nested tables are covered because a
/// nested table's embed lives in a cell story that is itself iterated.
fn table_cell_stories<T: ReadTxn>(txn: &T) -> HashSet<String> {
    let mut cells = HashSet::new();
    let Some(stories) = txn.get_map(crate::STORIES) else {
        return cells;
    };
    for (_, value) in stories.iter(txn) {
        let Out::YText(story) = value else {
            continue;
        };
        for chunk in snapshot(&story, txn) {
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
    cells
}

impl EditingDoc {
    /// Aggregates the read state over `[range.start, range.end)` — the yrs
    /// replacement for the PM selection tracker's mark walk.
    ///
    /// A caret (empty range) reads the text unit BEFORE it within the same
    /// paragraph, falling back to the unit after it at a paragraph start —
    /// PM `$from.marks()` semantics. Mark aggregation looks at TEXT units
    /// only (embeds neither veto nor satisfy a toggle, matching PM
    /// `toggleMark`); paragraph state comes from the paragraph containing
    /// the range start.
    pub fn selection_context(&self, range: &StoryRange) -> OpResult<SelectionContextInfo> {
        if range.end < range.start {
            return Err(OpError::InvalidRange {
                start: range.start,
                end: range.end,
            });
        }
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, &range.story)?;
        let chunks = snapshot(&story, &txn);
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
        let paragraph_properties: std::collections::BTreeMap<String, Any> =
            para_props.into_iter().collect();
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
        let mut ins = None;
        let mut del = None;
        let mut font_family = ValueAgg::Empty;
        let mut font_size = ValueAgg::Empty;
        let mut color = ValueAgg::Empty;
        for chunk in &chunks {
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
            ins = TriState::fold(ins, chunk.attr_active(crate::INS));
            del = TriState::fold(del, chunk.attr_active(crate::DEL));
            font_family.fold(chunk.attrs.get("fontFamily"));
            font_size.fold(chunk.attrs.get("fontSize"));
            color.fold(chunk.attrs.get("textColor"));
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
            // Keep the PM selection-state contract: `fontSize` is the raw
            // half-point `w:sz` value. The toolbar converts it to points.
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
            font_family,
            font_size,
            color,
            para_id,
            style_id: prop_string("pStyle"),
            alignment: prop_string("alignment"),
            paragraph_properties,
            has_selection: range.start != range.end,
            is_multi_paragraph,
            in_table: table_cell_stories(&txn).contains(&range.story),
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
            for change in self.list_changes(&story_id)? {
                let range = LocRange {
                    start: change.range.start.clone(),
                    end: change.range.end.clone(),
                };
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
                    let full = self.text_between(&range, TextView::Raw)?;
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
    use std::collections::HashMap;
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
        let mut undo = doc.undo_scope(&["body"]).unwrap();
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
