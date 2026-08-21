use std::collections::{HashMap, HashSet, VecDeque};
use std::fs;
use std::ops::Range;
use std::path::Path;

use anyhow::{Context, Result, anyhow, bail};
use docx_edit::frame_delta::{
    FRAME_DELTA_VERSION, FRAME_FLAG_FULL, FRAME_HEADER_LEN, PAGE_OP_LEN, PAGE_OP_MOVE,
    PAGE_OP_PATCH_POSITIONS, PAGE_OP_REMOVE, PAGE_OP_SHIFT_POSITIONS, PAGE_OP_UPSERT,
};
use docx_edit::{
    EditCtx, EngineSession, FormatPolicy, MergeDirection, ParaAttrDelta, ParaSelector, Patch,
    Position, SegmentContent, SimpleFormat, StoryRange, TriState, UndoSession, story_checksum,
};
use docx_layout::display_list::{DocAttrs, Primitive};
use docx_parse::block::BlockContent;
use docx_parse::document::{DocumentBody, get_paragraph_text};
use docx_parse::inline::{InlineNode, Run, RunContent, RunType};
use docx_parse::paragraph::{Paragraph, ParagraphContent};
use docx_parse::{
    S9PackageWire, S13SaveOptions, S13SaveRequest, SerializerDeterminism, UnderlineValue,
    write_docx_s13,
};
use serde::Deserialize;
use sha2::{Digest, Sha256};

use crate::chrome::{Alignment, EditingState, ToggleState};

const BODY_STORY: &str = "body";
const SAVE_TIME: &str = "1970-01-01T00:00:00.000Z";

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TextLoc {
    pub para_id: String,
    pub offset: u32,
}

#[derive(Clone, Debug)]
struct Selection {
    anchor: TextLoc,
    head: TextLoc,
}

#[derive(Clone, Debug)]
struct ParagraphSpan {
    para_id: String,
    text: String,
    length: u32,
    text_start: u32,
    story_start: u32,
    display_start: i64,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SelectionRect {
    pub page_index: usize,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}

#[derive(Clone, Debug)]
pub struct CaretGeometry {
    pub page_index: usize,
    pub x: f64,
    pub y: f64,
    pub height: f64,
}

#[derive(Clone, Debug)]
pub struct SceneChange {
    pub rebuild_all: bool,
    pub changed_pages: Vec<usize>,
}

#[derive(Clone, Copy, Debug)]
pub enum MoveDirection {
    Left,
    Right,
    Up,
    Down,
    Home,
    End,
}

#[derive(Clone, Copy, Debug)]
pub enum DeleteDirection {
    Backward,
    Forward,
}

pub struct DocxEditor {
    engine: EngineSession,
    undo: UndoSession,
    spans: Vec<ParagraphSpan>,
    selection: Option<Selection>,
    selection_rects: Vec<SelectionRect>,
    vertical_goal_x: Option<f64>,
    frames: FrameTracker,
    save: SaveProjection,
    initial_checksum: u64,
    dirty: bool,
}

impl DocxEditor {
    pub fn new(
        engine: EngineSession,
        source: Vec<u8>,
        package: S9PackageWire,
        initial_frame: &[u8],
    ) -> Result<Self> {
        let spans = paragraph_spans(&engine)?;
        let source_texts = spans
            .iter()
            .map(|span| (span.para_id.clone(), span.text.clone()))
            .collect();
        let source_alignments = engine
            .doc()
            .paragraphs(BODY_STORY)
            .map_err(anyhow::Error::msg)?
            .into_iter()
            .map(|paragraph| {
                (
                    paragraph.para_id,
                    string_property(paragraph.properties.get("alignment")),
                )
            })
            .collect();
        let source_inline = body_inline_projection(engine.doc())?;
        let source_tokens = body_tokens(engine.doc())?;
        let initial_checksum =
            story_checksum(engine.doc(), BODY_STORY).map_err(anyhow::Error::msg)?;
        let undo = UndoSession::new();
        undo.track(engine.doc(), BODY_STORY)
            .map_err(anyhow::Error::msg)?;
        Ok(Self {
            engine,
            undo,
            spans,
            selection: None,
            selection_rects: Vec::new(),
            vertical_goal_x: None,
            frames: FrameTracker::new(initial_frame)?,
            save: SaveProjection::new(
                source,
                package,
                source_texts,
                source_alignments,
                source_inline,
                source_tokens,
            ),
            initial_checksum,
            dirty: false,
        })
    }

    pub fn engine(&self) -> &EngineSession {
        &self.engine
    }

    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    pub fn has_collapsed_selection(&self) -> bool {
        self.selection
            .as_ref()
            .is_some_and(|selection| selection.anchor == selection.head)
    }

    pub fn selection_rects(&self) -> &[SelectionRect] {
        &self.selection_rects
    }

    pub fn editing_state(&self) -> Result<EditingState> {
        let can_undo = self.undo.can_undo();
        let can_redo = self.undo.can_redo();
        let Some(range) = self.selection_range()? else {
            return Ok(EditingState::editable_without_selection(can_undo, can_redo));
        };
        let context = self
            .engine
            .doc()
            .selection_context(&range)
            .map_err(anyhow::Error::msg)?;
        Ok(EditingState {
            bold: toggle_state(context.bold),
            italic: toggle_state(context.italic),
            underline: toggle_state(context.underline),
            alignment: Alignment::from_engine(context.alignment.as_deref()),
            inline_enabled: range.start != range.end,
            alignment_enabled: true,
            can_undo,
            can_redo,
            can_save: true,
        })
    }

    pub fn selection_range(&self) -> Result<Option<StoryRange>> {
        let Some(selection) = &self.selection else {
            return Ok(None);
        };
        let (start, end) = self.ordered_selection(selection)?;
        Ok(Some(StoryRange::new(
            BODY_STORY,
            self.story_index(&start)?,
            self.story_index(&end)?,
        )))
    }

    pub fn hit_test(&self, page_index: usize, x: f64, y: f64) -> Result<Option<TextLoc>> {
        let json = self
            .engine
            .display_hit_test_regions_json(page_index, x, y)
            .map_err(anyhow::Error::msg)?;
        let Some(hit) = serde_json::from_str::<Option<RegionHit>>(&json)? else {
            return Ok(None);
        };
        if hit.region != "body" {
            return Ok(None);
        }
        let Some(position) = hit.pos else {
            return Ok(None);
        };
        let loc = self.loc_from_display_position(position);
        Ok(loc.filter(|loc| {
            self.span(&loc.para_id)
                .is_some_and(|span| loc.offset >= span.text_start)
        }))
    }

    pub fn select_point(&mut self, loc: TextLoc, extend: bool, word: bool) -> Result<()> {
        let loc = self.clamp_loc(loc)?;
        if word {
            let (anchor, head) = self.word_range(&loc)?;
            self.selection = Some(Selection { anchor, head });
        } else if extend {
            if let Some(selection) = &mut self.selection {
                selection.head = loc;
            } else {
                self.selection = Some(Selection {
                    anchor: loc.clone(),
                    head: loc,
                });
            }
        } else {
            self.selection = Some(Selection {
                anchor: loc.clone(),
                head: loc,
            });
        }
        self.vertical_goal_x = None;
        self.refresh_selection_rects()
    }

    pub fn extend_to(&mut self, loc: TextLoc) -> Result<()> {
        self.select_point(loc, true, false)
    }

    pub fn move_selection(&mut self, direction: MoveDirection, extend: bool) -> Result<bool> {
        let Some(current) = self.selection.clone() else {
            return Ok(false);
        };
        if !extend
            && current.anchor != current.head
            && matches!(direction, MoveDirection::Left | MoveDirection::Right)
        {
            let (start, end) = self.ordered_selection(&current)?;
            return self.set_move_selection(
                if matches!(direction, MoveDirection::Left) {
                    start
                } else {
                    end
                },
                false,
            );
        }
        let next = match direction {
            MoveDirection::Up | MoveDirection::Down => {
                let position = self.display_position(&current.head)?;
                let direction = if matches!(direction, MoveDirection::Up) {
                    "up"
                } else {
                    "down"
                };
                let movement = self
                    .engine
                    .display_vertical_move_json(
                        position,
                        direction,
                        self.vertical_goal_x.unwrap_or(f64::NAN),
                    )
                    .map_err(anyhow::Error::msg)?;
                let movement = serde_json::from_str::<Option<VerticalMove>>(&movement)?;
                let Some(movement) = movement else {
                    return Ok(false);
                };
                self.vertical_goal_x = Some(movement.goal_x);
                self.loc_from_display_position(movement.position)
                    .unwrap_or(current.head.clone())
            }
            MoveDirection::Left => {
                self.vertical_goal_x = None;
                self.horizontal_loc(&current.head, false)?
            }
            MoveDirection::Right => {
                self.vertical_goal_x = None;
                self.horizontal_loc(&current.head, true)?
            }
            MoveDirection::Home => {
                self.vertical_goal_x = None;
                let span = self
                    .span(&current.head.para_id)
                    .context("selection paragraph is unavailable")?;
                TextLoc {
                    para_id: span.para_id.clone(),
                    offset: span.text_start,
                }
            }
            MoveDirection::End => {
                self.vertical_goal_x = None;
                let span = self
                    .span(&current.head.para_id)
                    .context("selection paragraph is unavailable")?;
                TextLoc {
                    para_id: span.para_id.clone(),
                    offset: span.length,
                }
            }
        };
        if next == current.head {
            return Ok(false);
        }
        if extend {
            self.selection = Some(Selection {
                anchor: current.anchor,
                head: next,
            });
            self.refresh_selection_rects()?;
            Ok(true)
        } else {
            self.set_move_selection(next, false)
        }
    }

    pub fn insert_text(&mut self, text: &str) -> Result<Option<SceneChange>> {
        if text.is_empty() || text.contains(['\r', '\n']) {
            return Ok(None);
        }
        let Some(selection) = self.selection.clone() else {
            return Ok(None);
        };
        let (start, end) = self.ordered_selection(&selection)?;
        let start_index = self.story_index(&start)?;
        let end_index = self.story_index(&end)?;
        let context = EditCtx::local("", "");
        if start_index == end_index {
            self.engine
                .doc()
                .insert_text(
                    &context,
                    Position::new(BODY_STORY, start_index),
                    text,
                    FormatPolicy::Inherit,
                )
                .map_err(anyhow::Error::msg)?;
        } else {
            self.engine
                .doc()
                .replace_range(
                    &context,
                    StoryRange::new(BODY_STORY, start_index, end_index),
                    text,
                )
                .map_err(anyhow::Error::msg)?;
        }
        let caret_index = start_index
            .checked_add(utf16_len(text))
            .context("caret position overflow")?;
        self.finish_edit_at_index(caret_index).map(Some)
    }

    pub fn delete(&mut self, direction: DeleteDirection) -> Result<Option<SceneChange>> {
        let Some(selection) = self.selection.clone() else {
            return Ok(None);
        };
        if selection.anchor != selection.head {
            let (start, end) = self.ordered_selection(&selection)?;
            let start_index = self.story_index(&start)?;
            let end_index = self.story_index(&end)?;
            self.engine
                .doc()
                .delete_range(
                    &EditCtx::local("", ""),
                    StoryRange::new(BODY_STORY, start_index, end_index),
                )
                .map_err(anyhow::Error::msg)?;
            return self.finish_edit_at_index(start_index).map(Some);
        }
        let caret = selection.head;
        let span_index = self
            .span_index(&caret.para_id)
            .context("selection paragraph is unavailable")?;
        let span = self.spans[span_index].clone();
        let context = EditCtx::local("", "");
        match direction {
            DeleteDirection::Backward if caret.offset > span.text_start => {
                let relative = caret.offset - span.text_start;
                let start = previous_code_point_offset(&span.text, relative);
                let from = span.story_start + span.text_start + start;
                let to = span.story_start + caret.offset;
                self.engine
                    .doc()
                    .delete_range(&context, StoryRange::new(BODY_STORY, from, to))
                    .map_err(anyhow::Error::msg)?;
                self.finish_edit_at_index(from).map(Some)
            }
            DeleteDirection::Forward if caret.offset < span.length => {
                let relative = caret.offset.saturating_sub(span.text_start);
                let end = next_code_point_offset(&span.text, relative);
                let from = span.story_start + caret.offset;
                let to = span.story_start + span.text_start + end;
                self.engine
                    .doc()
                    .delete_range(&context, StoryRange::new(BODY_STORY, from, to))
                    .map_err(anyhow::Error::msg)?;
                self.finish_edit_at_index(from).map(Some)
            }
            DeleteDirection::Backward if span_index > 0 => {
                let previous = self.spans[span_index - 1].clone();
                self.engine
                    .doc()
                    .merge_paragraphs(&context, &span.para_id, MergeDirection::Backward)
                    .map_err(anyhow::Error::msg)?;
                self.finish_edit_at_index(previous.story_start + previous.length)
                    .map(Some)
            }
            DeleteDirection::Forward if span_index + 1 < self.spans.len() => {
                self.engine
                    .doc()
                    .merge_paragraphs(&context, &span.para_id, MergeDirection::Forward)
                    .map_err(anyhow::Error::msg)?;
                self.finish_edit_at_index(span.story_start + span.length)
                    .map(Some)
            }
            _ => Ok(None),
        }
    }

    pub fn enter(&mut self) -> Result<Option<SceneChange>> {
        let Some(selection) = self.selection.clone() else {
            return Ok(None);
        };
        let (start, end) = self.ordered_selection(&selection)?;
        let start_index = self.story_index(&start)?;
        let end_index = self.story_index(&end)?;
        let context = EditCtx::local("", "");
        if start_index != end_index {
            self.engine
                .doc()
                .delete_range(
                    &context,
                    StoryRange::new(BODY_STORY, start_index, end_index),
                )
                .map_err(anyhow::Error::msg)?;
            self.spans = paragraph_spans(&self.engine)?;
        }
        let caret = self
            .loc_from_story_index(start_index)
            .context("collapsed selection did not resolve after deletion")?;
        let receipt = self
            .engine
            .doc()
            .split_paragraph(
                &context,
                Position::new(BODY_STORY, self.story_index(&caret)?),
                None,
            )
            .map_err(anyhow::Error::msg)?;
        self.finish_edit_at_loc(TextLoc {
            para_id: receipt.second_para_id,
            offset: 0,
        })
        .map(Some)
    }

    pub fn toggle_format(&mut self, format: SimpleFormat) -> Result<Option<SceneChange>> {
        let Some(selection) = self.selection.clone() else {
            return Ok(None);
        };
        let Some(range) = self.selection_range()? else {
            return Ok(None);
        };
        if range.start == range.end {
            return Ok(None);
        }
        self.undo.add_undo_barrier();
        self.engine
            .doc()
            .toggle_format(&EditCtx::local("", ""), range, format)
            .map_err(anyhow::Error::msg)?;
        self.undo.add_undo_barrier();
        self.finish_edit_preserving(selection).map(Some)
    }

    pub fn set_alignment(&mut self, alignment: Alignment) -> Result<Option<SceneChange>> {
        let Some(selection) = self.selection.clone() else {
            return Ok(None);
        };
        let Some(range) = self.selection_range()? else {
            return Ok(None);
        };
        let delta = ParaAttrDelta {
            alignment: Patch::Set(alignment.engine_value().to_owned()),
            ..ParaAttrDelta::default()
        };
        self.undo.add_undo_barrier();
        self.engine
            .doc()
            .set_paragraph_attrs(&EditCtx::local("", ""), &ParaSelector::Range(range), &delta)
            .map_err(anyhow::Error::msg)?;
        self.undo.add_undo_barrier();
        self.finish_edit_preserving(selection).map(Some)
    }

    pub fn undo(&mut self) -> Result<Option<SceneChange>> {
        self.apply_history(false)
    }

    pub fn redo(&mut self) -> Result<Option<SceneChange>> {
        self.apply_history(true)
    }

    pub fn caret_geometry(&self) -> Result<Option<CaretGeometry>> {
        let Some(selection) = &self.selection else {
            return Ok(None);
        };
        if selection.anchor != selection.head {
            return Ok(None);
        }
        let span = self
            .span(&selection.head.para_id)
            .context("selection paragraph is unavailable")?;
        let node_offset = selection.head.offset.saturating_sub(span.text_start);
        let snapshot = self
            .engine
            .resident_caret_snapshot(Some((&selection.head.para_id, node_offset)))
            .map_err(anyhow::Error::msg)?;
        Ok(snapshot.caret_rect.map(|rect| CaretGeometry {
            page_index: rect.page_index,
            x: rect.x,
            y: rect.y,
            height: rect.height,
        }))
    }

    pub fn save_to(&mut self, path: &Path) -> Result<()> {
        self.save.sync(self.engine.doc())?;
        let bytes = self.save.serialize()?;
        fs::write(path, bytes).with_context(|| format!("write edited DOCX {}", path.display()))
    }

    fn finish_edit_at_index(&mut self, index: u32) -> Result<SceneChange> {
        self.spans = paragraph_spans(&self.engine)?;
        let loc = self
            .loc_from_story_index(index)
            .context("edited caret no longer resolves")?;
        self.finish_edit_at_loc(loc)
    }

    fn finish_edit_at_loc(&mut self, loc: TextLoc) -> Result<SceneChange> {
        self.spans = paragraph_spans(&self.engine)?;
        let loc = self.clamp_loc(loc)?;
        self.selection = Some(Selection {
            anchor: loc.clone(),
            head: loc,
        });
        self.vertical_goal_x = None;
        let change = self.relayout()?;
        self.refresh_selection_rects()?;
        Ok(change)
    }

    fn finish_edit_preserving(&mut self, selection: Selection) -> Result<SceneChange> {
        self.selection = Some(selection);
        self.vertical_goal_x = None;
        let change = self.relayout()?;
        self.refresh_selection_rects()?;
        Ok(change)
    }

    fn apply_history(&mut self, redo: bool) -> Result<Option<SceneChange>> {
        let selection_indices = self
            .selection
            .as_ref()
            .map(|selection| -> Result<(u32, u32)> {
                Ok((
                    self.story_index(&selection.anchor)?,
                    self.story_index(&selection.head)?,
                ))
            })
            .transpose()?;
        let changed = if redo {
            self.undo.redo()
        } else {
            self.undo.undo()
        };
        if !changed {
            return Ok(None);
        }
        let change = self.relayout()?;
        self.selection = selection_indices.and_then(|(anchor, head)| {
            Some(Selection {
                anchor: self.loc_from_story_index_clamped(anchor)?,
                head: self.loc_from_story_index_clamped(head)?,
            })
        });
        self.vertical_goal_x = None;
        self.refresh_selection_rects()?;
        Ok(Some(change))
    }

    fn relayout(&mut self) -> Result<SceneChange> {
        let frame = self
            .engine
            .apply_and_layout(BODY_STORY, self.frames.frame_epoch)
            .map_err(anyhow::Error::msg)?;
        let change = self.frames.apply(&frame)?;
        self.spans = paragraph_spans(&self.engine)?;
        self.dirty = story_checksum(self.engine.doc(), BODY_STORY).map_err(anyhow::Error::msg)?
            != self.initial_checksum;
        Ok(change)
    }

    fn set_move_selection(&mut self, loc: TextLoc, extend: bool) -> Result<bool> {
        let loc = self.clamp_loc(loc)?;
        if extend {
            if let Some(selection) = &mut self.selection {
                selection.head = loc;
            }
        } else {
            self.selection = Some(Selection {
                anchor: loc.clone(),
                head: loc,
            });
        }
        self.refresh_selection_rects()?;
        Ok(true)
    }

    fn horizontal_loc(&self, loc: &TextLoc, forward: bool) -> Result<TextLoc> {
        let index = self
            .span_index(&loc.para_id)
            .context("selection paragraph is unavailable")?;
        let span = &self.spans[index];
        if forward {
            if loc.offset < span.length {
                let relative = loc.offset.saturating_sub(span.text_start);
                return Ok(TextLoc {
                    para_id: span.para_id.clone(),
                    offset: span.text_start + next_code_point_offset(&span.text, relative),
                });
            }
            if let Some(next) = self.spans.get(index + 1) {
                return Ok(TextLoc {
                    para_id: next.para_id.clone(),
                    offset: next.text_start,
                });
            }
        } else {
            if loc.offset > span.text_start {
                let relative = loc.offset - span.text_start;
                return Ok(TextLoc {
                    para_id: span.para_id.clone(),
                    offset: span.text_start + previous_code_point_offset(&span.text, relative),
                });
            }
            if index > 0 {
                let previous = &self.spans[index - 1];
                return Ok(TextLoc {
                    para_id: previous.para_id.clone(),
                    offset: previous.length,
                });
            }
        }
        Ok(loc.clone())
    }

    fn word_range(&self, loc: &TextLoc) -> Result<(TextLoc, TextLoc)> {
        let span = self
            .span(&loc.para_id)
            .context("selection paragraph is unavailable")?;
        let relative = loc.offset.saturating_sub(span.text_start);
        let (start, end) = word_boundaries(&span.text, relative);
        Ok((
            TextLoc {
                para_id: loc.para_id.clone(),
                offset: span.text_start + start,
            },
            TextLoc {
                para_id: loc.para_id.clone(),
                offset: span.text_start + end,
            },
        ))
    }

    fn refresh_selection_rects(&mut self) -> Result<()> {
        let Some(selection) = &self.selection else {
            self.selection_rects.clear();
            return Ok(());
        };
        if selection.anchor == selection.head {
            self.selection_rects.clear();
            return Ok(());
        }
        let from = self.display_position(&selection.anchor)?;
        let to = self.display_position(&selection.head)?;
        let json = self
            .engine
            .display_range_rects_json(from.min(to), from.max(to))
            .map_err(anyhow::Error::msg)?;
        self.selection_rects = serde_json::from_str(&json)?;
        Ok(())
    }

    fn ordered_selection(&self, selection: &Selection) -> Result<(TextLoc, TextLoc)> {
        if self.story_index(&selection.anchor)? <= self.story_index(&selection.head)? {
            Ok((selection.anchor.clone(), selection.head.clone()))
        } else {
            Ok((selection.head.clone(), selection.anchor.clone()))
        }
    }

    fn clamp_loc(&self, mut loc: TextLoc) -> Result<TextLoc> {
        let span = self
            .span(&loc.para_id)
            .with_context(|| format!("paragraph {:?} is unavailable", loc.para_id))?;
        loc.offset = loc.offset.clamp(span.text_start, span.length);
        Ok(loc)
    }

    fn span(&self, para_id: &str) -> Option<&ParagraphSpan> {
        self.spans.iter().find(|span| span.para_id == para_id)
    }

    fn span_index(&self, para_id: &str) -> Option<usize> {
        self.spans.iter().position(|span| span.para_id == para_id)
    }

    fn story_index(&self, loc: &TextLoc) -> Result<u32> {
        let span = self
            .span(&loc.para_id)
            .with_context(|| format!("paragraph {:?} is unavailable", loc.para_id))?;
        span.story_start
            .checked_add(loc.offset.min(span.length))
            .context("story position overflow")
    }

    fn display_position(&self, loc: &TextLoc) -> Result<i64> {
        let span = self
            .span(&loc.para_id)
            .with_context(|| format!("paragraph {:?} is unavailable", loc.para_id))?;
        Ok(span.display_start
            + 1
            + i64::from(
                loc.offset
                    .clamp(span.text_start, span.length)
                    .saturating_sub(span.text_start),
            ))
    }

    fn loc_from_display_position(&self, position: i64) -> Option<TextLoc> {
        let position = position.max(0);
        let span = self
            .spans
            .iter()
            .rev()
            .find(|span| span.display_start <= position)
            .or_else(|| self.spans.first())?;
        let offset = position
            .saturating_sub(span.display_start + 1)
            .clamp(0, i64::from(span.length.saturating_sub(span.text_start)))
            as u32;
        Some(TextLoc {
            para_id: span.para_id.clone(),
            offset: span.text_start + offset,
        })
    }

    fn loc_from_story_index(&self, index: u32) -> Option<TextLoc> {
        let span = self.spans.iter().find(|span| {
            index >= span.story_start && index <= span.story_start.saturating_add(span.length)
        })?;
        Some(TextLoc {
            para_id: span.para_id.clone(),
            offset: index - span.story_start,
        })
    }

    fn loc_from_story_index_clamped(&self, index: u32) -> Option<TextLoc> {
        self.loc_from_story_index(index).or_else(|| {
            let span = self
                .spans
                .iter()
                .rev()
                .find(|span| span.story_start <= index)
                .or_else(|| self.spans.first())?;
            Some(TextLoc {
                para_id: span.para_id.clone(),
                offset: if index < span.story_start {
                    span.text_start
                } else {
                    span.length
                },
            })
        })
    }
}

fn toggle_state(state: TriState) -> ToggleState {
    match state {
        TriState::On => ToggleState::On,
        TriState::Off => ToggleState::Off,
        TriState::Mixed => ToggleState::Mixed,
    }
}

fn string_property<T: serde::Serialize>(value: Option<&T>) -> Option<String> {
    serde_json::to_value(value?)
        .ok()?
        .as_str()
        .map(str::to_owned)
}

#[derive(Deserialize)]
struct RegionHit {
    region: String,
    pos: Option<i64>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct VerticalMove {
    position: i64,
    goal_x: f64,
}

fn paragraph_spans(engine: &EngineSession) -> Result<Vec<ParagraphSpan>> {
    let display_starts = engine
        .with_display_list(|list| {
            let mut starts = HashMap::new();
            for page in &list.pages {
                for primitive in &page.primitives {
                    let attrs = primitive_attrs(primitive);
                    if let (Some(para_id), Some(start)) = (&attrs.para_id, attrs.fragment_doc_start)
                    {
                        starts
                            .entry(para_id.clone())
                            .and_modify(|current: &mut i64| *current = (*current).min(start))
                            .or_insert(start);
                    }
                }
            }
            starts
        })
        .unwrap_or_default();
    let paragraphs = engine
        .doc()
        .paragraphs(BODY_STORY)
        .map_err(anyhow::Error::msg)?
        .into_iter()
        .map(|paragraph| (paragraph.para_id, paragraph.text))
        .collect::<HashMap<_, _>>();
    let segments = engine
        .doc()
        .story_segments(BODY_STORY)
        .map_err(anyhow::Error::msg)?;
    let mut spans = Vec::new();
    let mut cursor = 0_u32;
    let mut paragraph_start = 0_u32;
    let mut node_start = 0_u32;
    let mut display_start = 0_i64;
    for segment in segments {
        match segment.content {
            SegmentContent::Text(text) => {
                cursor = cursor
                    .checked_add(utf16_len(&text))
                    .context("story length overflow")?;
            }
            SegmentContent::Pilcrow(properties) => {
                let length = cursor
                    .checked_sub(paragraph_start)
                    .context("paragraph span underflow")?;
                let text_start = node_start.saturating_sub(paragraph_start).min(length);
                let text = paragraphs
                    .get(&properties.para_id)
                    .cloned()
                    .unwrap_or_default();
                let para_id = properties.para_id;
                spans.push(ParagraphSpan {
                    display_start: display_starts
                        .get(&para_id)
                        .copied()
                        .unwrap_or(display_start),
                    para_id,
                    text,
                    length,
                    text_start,
                    story_start: paragraph_start,
                });
                cursor = cursor.checked_add(1).context("story length overflow")?;
                paragraph_start = cursor;
                node_start = cursor;
                display_start = display_start
                    .checked_add(i64::from(length) + 2)
                    .context("display position overflow")?;
            }
            SegmentContent::OtherEmbed { kind, .. } => {
                if cursor == node_start && is_block_embed(&kind) {
                    node_start = cursor.checked_add(1).context("story length overflow")?;
                }
                cursor = cursor.checked_add(1).context("story length overflow")?;
            }
        }
    }
    Ok(spans)
}

fn primitive_attrs(primitive: &Primitive) -> &DocAttrs {
    match primitive {
        Primitive::Text(value) => &value.attrs,
        Primitive::GlyphRun(value) => &value.attrs,
        Primitive::Rect(value) => &value.attrs,
        Primitive::Line(value) => &value.attrs,
        Primitive::Image(value) => &value.attrs,
        Primitive::Shape(value) => &value.attrs,
        Primitive::Decoration(value) => &value.attrs,
    }
}

fn is_block_embed(kind: &str) -> bool {
    matches!(kind, "table" | "blockSdt" | "pageBreak" | "columnBreak")
}

fn utf16_len(text: &str) -> u32 {
    text.encode_utf16().count() as u32
}

fn previous_code_point_offset(text: &str, offset: u32) -> u32 {
    let units = text.encode_utf16().collect::<Vec<_>>();
    let offset = usize::try_from(offset)
        .unwrap_or(usize::MAX)
        .min(units.len());
    if offset == 0 {
        return 0;
    }
    if offset > 1
        && (0xdc00..=0xdfff).contains(&units[offset - 1])
        && (0xd800..=0xdbff).contains(&units[offset - 2])
    {
        (offset - 2) as u32
    } else {
        (offset - 1) as u32
    }
}

fn next_code_point_offset(text: &str, offset: u32) -> u32 {
    let units = text.encode_utf16().collect::<Vec<_>>();
    let offset = usize::try_from(offset)
        .unwrap_or(usize::MAX)
        .min(units.len());
    if offset >= units.len() {
        return units.len() as u32;
    }
    if offset + 1 < units.len()
        && (0xd800..=0xdbff).contains(&units[offset])
        && (0xdc00..=0xdfff).contains(&units[offset + 1])
    {
        (offset + 2) as u32
    } else {
        (offset + 1) as u32
    }
}

fn word_boundaries(text: &str, utf16_offset: u32) -> (u32, u32) {
    let chars = text
        .char_indices()
        .map(|(byte, character)| (byte, character, character.len_utf16() as u32))
        .collect::<Vec<_>>();
    if chars.is_empty() {
        return (0, 0);
    }
    let mut positions = Vec::with_capacity(chars.len());
    let mut cursor = 0_u32;
    for (_, character, width) in &chars {
        positions.push((cursor, cursor + *width, *character));
        cursor += *width;
    }
    let offset = utf16_offset.min(cursor.saturating_sub(1));
    let index = positions
        .iter()
        .position(|(start, end, _)| offset >= *start && offset < *end)
        .unwrap_or(positions.len() - 1);
    let class = word_class(positions[index].2);
    let mut start = index;
    let mut end = index + 1;
    while start > 0 && word_class(positions[start - 1].2) == class {
        start -= 1;
    }
    while end < positions.len() && word_class(positions[end].2) == class {
        end += 1;
    }
    if class == WordClass::Other {
        end = start + 1;
    }
    (positions[start].0, positions[end - 1].1)
}

#[derive(Clone, Copy, Eq, PartialEq)]
enum WordClass {
    Word,
    Space,
    Other,
}

fn word_class(character: char) -> WordClass {
    if character.is_whitespace() {
        WordClass::Space
    } else if character.is_alphanumeric() || matches!(character, '_' | '\'' | '\u{2019}' | '-') {
        WordClass::Word
    } else {
        WordClass::Other
    }
}

struct FrameTracker {
    frame_epoch: u64,
    page_ids: Vec<u64>,
}

impl FrameTracker {
    fn new(bytes: &[u8]) -> Result<Self> {
        let frame = ParsedFrame::parse(bytes)?;
        if !frame.full {
            bail!("initial display frame is not full");
        }
        let mut page_ids = vec![0; frame.page_count];
        for op in &frame.ops {
            if op.kind == PAGE_OP_UPSERT && op.page_index < page_ids.len() {
                page_ids[op.page_index] = op.page_id;
            }
        }
        if page_ids.contains(&0) {
            bail!("initial display frame omitted a page");
        }
        Ok(Self {
            frame_epoch: frame.frame_epoch,
            page_ids,
        })
    }

    fn apply(&mut self, bytes: &[u8]) -> Result<SceneChange> {
        let frame = ParsedFrame::parse(bytes)?;
        if !frame.full && frame.base_frame_epoch != self.frame_epoch {
            bail!(
                "display frame base {} does not match {}",
                frame.base_frame_epoch,
                self.frame_epoch
            );
        }
        let reindexed = frame.ops.iter().any(|op| {
            op.kind != PAGE_OP_REMOVE
                && self
                    .page_ids
                    .iter()
                    .position(|page_id| *page_id == op.page_id)
                    .is_some_and(|index| index != op.page_index)
        });
        let structural = frame.full
            || frame.page_count != self.page_ids.len()
            || reindexed
            || frame
                .ops
                .iter()
                .any(|op| matches!(op.kind, PAGE_OP_REMOVE | PAGE_OP_MOVE));
        let mut next_ids = vec![None; frame.page_count];
        let removed = frame
            .ops
            .iter()
            .filter(|op| op.kind == PAGE_OP_REMOVE)
            .map(|op| op.page_id)
            .collect::<HashSet<_>>();
        for op in &frame.ops {
            if op.kind != PAGE_OP_REMOVE && op.page_index < next_ids.len() {
                next_ids[op.page_index] = Some(op.page_id);
            }
        }
        if !frame.full {
            for (index, page_id) in self.page_ids.iter().copied().enumerate() {
                if removed.contains(&page_id)
                    || frame
                        .ops
                        .iter()
                        .any(|op| op.kind != PAGE_OP_REMOVE && op.page_id == page_id)
                {
                    continue;
                }
                if index < next_ids.len() && next_ids[index].is_none() {
                    next_ids[index] = Some(page_id);
                }
            }
            let mut remaining = self
                .page_ids
                .iter()
                .copied()
                .filter(|page_id| {
                    !removed.contains(page_id)
                        && !next_ids.iter().flatten().any(|placed| placed == page_id)
                })
                .collect::<VecDeque<_>>();
            for slot in &mut next_ids {
                if slot.is_none() {
                    *slot = remaining.pop_front();
                }
            }
        }
        let page_ids = next_ids
            .into_iter()
            .collect::<Option<Vec<_>>>()
            .context("display delta did not resolve every page identity")?;
        let mut changed_pages = frame
            .ops
            .iter()
            .filter(|op| op.kind == PAGE_OP_UPSERT)
            .map(|op| op.page_index)
            .collect::<Vec<_>>();
        changed_pages.sort_unstable();
        changed_pages.dedup();
        self.frame_epoch = frame.frame_epoch;
        self.page_ids = page_ids;
        Ok(SceneChange {
            rebuild_all: structural,
            changed_pages,
        })
    }
}

struct ParsedFrame {
    full: bool,
    frame_epoch: u64,
    base_frame_epoch: u64,
    page_count: usize,
    ops: Vec<FrameOp>,
}

impl ParsedFrame {
    fn parse(bytes: &[u8]) -> Result<Self> {
        if bytes.len() < FRAME_HEADER_LEN || bytes.get(..4) != Some(b"FDV1") {
            bail!("invalid FrameDelta header");
        }
        if read_u16(bytes, 4)? != FRAME_DELTA_VERSION {
            bail!("unsupported FrameDelta version");
        }
        let header_len = usize::from(read_u16(bytes, 6)?);
        if header_len != FRAME_HEADER_LEN || read_u32(bytes, 8)? as usize != bytes.len() {
            bail!("inconsistent FrameDelta length");
        }
        let flags = read_u32(bytes, 12)?;
        let page_count = read_u32(bytes, 48)? as usize;
        let op_count = read_u32(bytes, 52)? as usize;
        let ops_offset = read_u32(bytes, 56)? as usize;
        let ops_len = op_count
            .checked_mul(PAGE_OP_LEN)
            .context("FrameDelta operation table overflow")?;
        if ops_offset
            .checked_add(ops_len)
            .is_none_or(|end| end > bytes.len())
        {
            bail!("FrameDelta operation table is truncated");
        }
        let mut ops = Vec::with_capacity(op_count);
        for index in 0..op_count {
            let offset = ops_offset + index * PAGE_OP_LEN;
            let kind = bytes[offset];
            if !matches!(
                kind,
                PAGE_OP_UPSERT
                    | PAGE_OP_REMOVE
                    | PAGE_OP_MOVE
                    | PAGE_OP_PATCH_POSITIONS
                    | PAGE_OP_SHIFT_POSITIONS
            ) {
                bail!("unknown FrameDelta page operation {kind}");
            }
            ops.push(FrameOp {
                kind,
                page_index: read_u32(bytes, offset + 4)? as usize,
                page_id: read_u64(bytes, offset + 8)?,
            });
        }
        Ok(Self {
            full: flags & FRAME_FLAG_FULL != 0,
            frame_epoch: read_u64(bytes, 32)?,
            base_frame_epoch: read_u64(bytes, 40)?,
            page_count,
            ops,
        })
    }
}

struct FrameOp {
    kind: u8,
    page_index: usize,
    page_id: u64,
}

fn read_u16(bytes: &[u8], offset: usize) -> Result<u16> {
    let value = bytes
        .get(offset..offset + 2)
        .context("FrameDelta u16 is truncated")?
        .try_into()
        .map_err(|_| anyhow!("FrameDelta u16 is invalid"))?;
    Ok(u16::from_le_bytes(value))
}

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32> {
    let value = bytes
        .get(offset..offset + 4)
        .context("FrameDelta u32 is truncated")?
        .try_into()
        .map_err(|_| anyhow!("FrameDelta u32 is invalid"))?;
    Ok(u32::from_le_bytes(value))
}

fn read_u64(bytes: &[u8], offset: usize) -> Result<u64> {
    let value = bytes
        .get(offset..offset + 8)
        .context("FrameDelta u64 is truncated")?
        .try_into()
        .map_err(|_| anyhow!("FrameDelta u64 is invalid"))?;
    Ok(u64::from_le_bytes(value))
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct InlineMarks {
    bold: bool,
    italic: bool,
    underline: bool,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct InlineRun {
    start: u32,
    end: u32,
    marks: InlineMarks,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct InlineProjection {
    text: String,
    runs: Vec<InlineRun>,
    has_embed: bool,
}

impl InlineProjection {
    fn marks_at(&self, offset: u32) -> InlineMarks {
        self.runs
            .iter()
            .find(|run| run.start <= offset && offset < run.end)
            .map_or(
                InlineMarks {
                    bold: false,
                    italic: false,
                    underline: false,
                },
                |run| run.marks,
            )
    }

    fn mark_sequence(&self) -> Vec<InlineMarks> {
        self.runs.iter().map(|run| run.marks).collect()
    }
}

struct SaveProjection {
    original: Vec<u8>,
    package: S9PackageWire,
    source_content: Vec<BlockContent>,
    source_texts: HashMap<String, String>,
    source_alignments: HashMap<String, Option<String>>,
    source_inline: HashMap<String, InlineProjection>,
    source_tokens: Vec<BodyToken>,
    seed: String,
    synthetic_ids: HashSet<String>,
    changed_para_ids: Vec<String>,
}

impl SaveProjection {
    fn new(
        original: Vec<u8>,
        mut package: S9PackageWire,
        source_texts: HashMap<String, String>,
        source_alignments: HashMap<String, Option<String>>,
        source_inline: HashMap<String, InlineProjection>,
        source_tokens: Vec<BodyToken>,
    ) -> Self {
        let seed = format!("{:x}", Sha256::digest(&original));
        let mut synthetic_ids = HashSet::new();
        let mut paragraph_index = 0;
        for block in &mut package.document.content {
            if let BlockContent::Paragraph(paragraph) = block {
                if paragraph.para_id.is_none() {
                    let para_id = format!("{BODY_STORY}:p{paragraph_index}");
                    paragraph.para_id = Some(para_id.clone());
                    synthetic_ids.insert(para_id);
                }
                paragraph_index += 1;
            }
        }
        let source_content = package.document.content.clone();
        Self {
            original,
            package,
            source_content,
            source_texts,
            source_alignments,
            source_inline,
            source_tokens,
            seed,
            synthetic_ids,
            changed_para_ids: Vec::new(),
        }
    }

    fn sync(&mut self, document: &docx_edit::EditingDoc) -> Result<()> {
        let paragraphs = document
            .paragraphs(BODY_STORY)
            .map_err(anyhow::Error::msg)?
            .into_iter()
            .map(|paragraph| (paragraph.para_id.clone(), paragraph))
            .collect::<HashMap<_, _>>();
        let inline = body_inline_projection(document)?;
        let tokens = body_tokens(document)?;
        if tokens != self.source_tokens {
            bail!(
                "cannot save DOCX structural edits; paragraph splits, merges, tables, and content controls are not round-tripped"
            );
        }
        let mut content = self.source_content.clone();
        let mut changed_para_ids = Vec::new();
        let mut projected_para_ids = HashSet::new();
        let mut paragraph_number = 0;
        for block in &mut content {
            let BlockContent::Paragraph(paragraph) = block else {
                continue;
            };
            paragraph_number += 1;
            let para_id = paragraph
                .para_id
                .clone()
                .context("source paragraph has no projection identity")?;
            projected_para_ids.insert(para_id.clone());
            let target = paragraphs.get(&para_id).with_context(|| {
                format!("paragraph {paragraph_number} ({para_id}) is unavailable")
            })?;
            let source = self
                .source_texts
                .get(&para_id)
                .cloned()
                .unwrap_or_else(|| get_paragraph_text(paragraph));
            let target_alignment = string_property(target.properties.get("alignment"));
            let source_alignment = self
                .source_alignments
                .get(&para_id)
                .cloned()
                .unwrap_or_default();
            let source_inline = self
                .source_inline
                .get(&para_id)
                .with_context(|| format!("paragraph {para_id} has no inline source projection"))?;
            let target_inline = inline
                .get(&para_id)
                .with_context(|| format!("paragraph {para_id} has no inline target projection"))?;
            let text_changed = source != target.text;
            let alignment_changed = source_alignment != target_alignment;
            let inline_changed = source_inline.runs != target_inline.runs;
            if !text_changed && !alignment_changed && !inline_changed {
                continue;
            }
            if text_changed {
                if source_inline.mark_sequence() != target_inline.mark_sequence() {
                    bail!(
                        "cannot save paragraph {paragraph_number} ({para_id}): text and character formatting changed together"
                    );
                }
                let risks = paragraph_projection_risks(paragraph);
                if !risks.is_empty() {
                    bail!(
                        "cannot save paragraph {paragraph_number} ({para_id}): editing it would lose or flatten {}",
                        risks.join(", ")
                    );
                }
                set_paragraph_text(paragraph, &target.text).map_err(anyhow::Error::msg)?;
            } else if inline_changed {
                let risks = paragraph_projection_risks(paragraph)
                    .into_iter()
                    .filter(|risk| *risk != "differently formatted runs")
                    .collect::<Vec<_>>();
                if source_inline.has_embed || target_inline.has_embed || !risks.is_empty() {
                    let reason = if risks.is_empty() {
                        "non-text inline content".to_owned()
                    } else {
                        risks.join(", ")
                    };
                    bail!(
                        "cannot save paragraph {paragraph_number} ({para_id}): formatting it would lose or flatten {reason}"
                    );
                }
                apply_inline_formatting(paragraph, source_inline, target_inline)?;
            }
            if alignment_changed {
                paragraph.formatting.get_or_insert_default().alignment = target_alignment;
            }
            changed_para_ids.push(para_id);
        }
        for (para_id, target) in &paragraphs {
            let source = self
                .source_texts
                .get(para_id)
                .with_context(|| format!("paragraph {para_id} has no source projection"))?;
            let inline_changed = self.source_inline.get(para_id) != inline.get(para_id);
            let alignment_changed = self
                .source_alignments
                .get(para_id)
                .cloned()
                .unwrap_or_default()
                != string_property(target.properties.get("alignment"));
            if (source != &target.text || inline_changed || alignment_changed)
                && !projected_para_ids.contains(para_id)
            {
                bail!(
                    "cannot save paragraph {para_id}: nested table or content-control paragraphs are not round-tripped"
                );
            }
        }
        self.package.document.content = content;
        self.changed_para_ids = changed_para_ids;
        Ok(())
    }

    fn serialize(&self) -> Result<Vec<u8>> {
        if self.changed_para_ids.is_empty() {
            return Ok(self.original.clone());
        }
        let document = DocumentBody {
            content: self.package.document.content.clone(),
            sections: None,
            final_section_properties: self.package.document.final_section_properties.clone(),
            comments: self.package.document.comments.clone(),
        };
        let request = S13SaveRequest {
            determinism: SerializerDeterminism {
                seed: self.seed.clone(),
                now: SAVE_TIME.to_owned(),
            },
            document,
            header_entries: self.package.header_entries.clone().unwrap_or_default(),
            footer_entries: self.package.footer_entries.clone().unwrap_or_default(),
            footnotes: self.package.footnotes.clone().unwrap_or_default(),
            endnotes: self.package.endnotes.clone().unwrap_or_default(),
            footnote_separators: self.package.footnote_separators.clone().unwrap_or_default(),
            endnote_separators: self.package.endnote_separators.clone().unwrap_or_default(),
            relationship_entries: self.package.relationship_entries.clone(),
            numbering: Some(self.package.numbering.clone()),
            options: S13SaveOptions {
                update_modified_date: false,
                modified_by: None,
            },
            selective: None,
        };
        let projected = write_docx_s13(request, &self.original).map_err(anyhow::Error::from)?;
        patch_document_paragraphs(
            &self.original,
            &projected,
            &self.changed_para_ids,
            &self.synthetic_ids,
        )
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
enum BodyToken {
    Paragraph(String),
    Table,
    BlockSdt,
}

fn body_inline_projection(
    document: &docx_edit::EditingDoc,
) -> Result<HashMap<String, InlineProjection>> {
    let mut projections = HashMap::new();
    let mut text = String::new();
    let mut runs: Vec<InlineRun> = Vec::new();
    let mut cursor = 0_u32;
    let mut has_embed = false;
    for segment in document
        .story_segments(BODY_STORY)
        .map_err(anyhow::Error::msg)?
    {
        match segment.content {
            SegmentContent::Text(value) => {
                let length = utf16_len(&value);
                let marks = InlineMarks {
                    bold: active_property(segment.attributes.get("bold")),
                    italic: active_property(segment.attributes.get("italic")),
                    underline: active_property(segment.attributes.get("underline")),
                };
                if length > 0 {
                    if let Some(previous) = runs.last_mut()
                        && previous.end == cursor
                        && previous.marks == marks
                    {
                        previous.end += length;
                    } else {
                        runs.push(InlineRun {
                            start: cursor,
                            end: cursor + length,
                            marks,
                        });
                    }
                }
                cursor += length;
                text.push_str(&value);
            }
            SegmentContent::Pilcrow(properties) => {
                projections.insert(
                    properties.para_id,
                    InlineProjection {
                        text: std::mem::take(&mut text),
                        runs: std::mem::take(&mut runs),
                        has_embed,
                    },
                );
                cursor = 0;
                has_embed = false;
            }
            SegmentContent::OtherEmbed { .. } => {
                cursor = cursor.checked_add(1).context("inline offset overflow")?;
                has_embed = true;
            }
        }
    }
    Ok(projections)
}

fn active_property<T: serde::Serialize>(value: Option<&T>) -> bool {
    serde_json::to_value(value).is_ok_and(|value| !value.is_null())
}

fn body_tokens(document: &docx_edit::EditingDoc) -> Result<Vec<BodyToken>> {
    let segments = document
        .story_segments(BODY_STORY)
        .map_err(anyhow::Error::msg)?;
    let mut tokens = Vec::new();
    for segment in segments {
        match segment.content {
            SegmentContent::Pilcrow(properties) => {
                tokens.push(BodyToken::Paragraph(properties.para_id));
            }
            SegmentContent::OtherEmbed { kind, .. } if kind == "table" => {
                tokens.push(BodyToken::Table);
            }
            SegmentContent::OtherEmbed { kind, .. } if kind == "blockSdt" => {
                tokens.push(BodyToken::BlockSdt);
            }
            _ => {}
        }
    }
    Ok(tokens)
}

fn paragraph_projection_risks(paragraph: &Paragraph) -> Vec<&'static str> {
    let mut risks = Vec::new();
    if paragraph.property_changes.is_some()
        || paragraph.p_pr_ins.is_some()
        || paragraph.p_pr_del.is_some()
    {
        risks.push("revision marks");
    }
    let mut run_formatting = Vec::new();
    for content in &paragraph.content {
        match content {
            ParagraphContent::Inline(InlineNode::Run(run)) => {
                run_formatting.push(run.formatting.as_ref());
                if run.property_changes.is_some() {
                    risks.push("revision marks");
                }
                for content in &run.content {
                    match content {
                        RunContent::Text { .. } => {}
                        RunContent::FootnoteRef { .. }
                        | RunContent::EndnoteRef { .. }
                        | RunContent::FootnoteRefMark
                        | RunContent::EndnoteRefMark => risks.push("note references"),
                        RunContent::CommentReference { .. } => risks.push("comment references"),
                        RunContent::FieldChar { .. } | RunContent::InstrText { .. } => {
                            risks.push("fields")
                        }
                        RunContent::Drawing { .. }
                        | RunContent::Shape { .. }
                        | RunContent::Chart { .. }
                        | RunContent::OpaqueDrawing { .. } => risks.push("embedded objects"),
                        _ => risks.push("non-text run content"),
                    }
                }
            }
            ParagraphContent::Inline(InlineNode::Hyperlink(_)) => risks.push("hyperlinks"),
            ParagraphContent::Inline(InlineNode::InlineSdt(_)) => risks.push("content controls"),
            ParagraphContent::Inline(InlineNode::BookmarkStart(_))
            | ParagraphContent::Inline(InlineNode::BookmarkEnd(_)) => risks.push("bookmarks"),
            ParagraphContent::Inline(InlineNode::SimpleField(_))
            | ParagraphContent::Inline(InlineNode::ComplexField(_)) => risks.push("fields"),
            ParagraphContent::Inline(InlineNode::Math(_)) => risks.push("math"),
            ParagraphContent::Tracked(_)
            | ParagraphContent::RangeStart(_)
            | ParagraphContent::RangeEnd(_) => risks.push("revision marks"),
            ParagraphContent::CommentRange(_) => risks.push("comment anchors"),
        }
    }
    if let Some(first) = run_formatting.first()
        && run_formatting
            .iter()
            .skip(1)
            .any(|formatting| formatting != first)
    {
        risks.push("differently formatted runs");
    }
    risks.sort_unstable();
    risks.dedup();
    risks
}

fn apply_inline_formatting(
    paragraph: &mut Paragraph,
    source: &InlineProjection,
    target: &InlineProjection,
) -> Result<()> {
    if source.text != target.text {
        bail!("inline formatting projection text changed");
    }
    let chunks =
        simple_run_chunks(paragraph).context("character formatting requires text-only runs")?;
    let paragraph_text = chunks
        .iter()
        .map(|chunk| chunk.text.as_str())
        .collect::<String>();
    if paragraph_text != target.text {
        bail!("character formatting projection does not match paragraph text");
    }
    let mut chunk_ranges = Vec::with_capacity(chunks.len());
    let mut cursor = 0_u32;
    let mut boundaries = vec![0, utf16_len(&target.text)];
    for chunk in chunks {
        let end = cursor + utf16_len(&chunk.text);
        boundaries.extend([cursor, end]);
        chunk_ranges.push((cursor, end, chunk.run));
        cursor = end;
    }
    for run in source.runs.iter().chain(&target.runs) {
        boundaries.extend([run.start, run.end]);
    }
    boundaries.sort_unstable();
    boundaries.dedup();
    let mut content = Vec::new();
    for pair in boundaries.windows(2) {
        let [start, end] = pair else {
            continue;
        };
        if start == end {
            continue;
        }
        let template = chunk_ranges
            .iter()
            .find(|(chunk_start, chunk_end, _)| chunk_start <= start && start < chunk_end)
            .map(|(_, _, run)| run)
            .context("character formatting range has no source run")?;
        let mut run = run_with_text(template, utf16_slice(&target.text, *start, *end));
        let before = source.marks_at(*start);
        let after = target.marks_at(*start);
        if before.bold != after.bold
            || before.italic != after.italic
            || before.underline != after.underline
        {
            let formatting = run.formatting.get_or_insert_default();
            if before.bold != after.bold {
                formatting.bold = Some(after.bold);
            }
            if before.italic != after.italic {
                formatting.italic = Some(after.italic);
            }
            if before.underline != after.underline {
                formatting.underline = Some(UnderlineValue {
                    style: if after.underline { "single" } else { "none" }.to_owned(),
                    color: None,
                });
            }
        }
        content.push(ParagraphContent::Inline(InlineNode::Run(run)));
    }
    paragraph.content = content;
    Ok(())
}

fn set_paragraph_text(paragraph: &mut Paragraph, target: &str) -> Result<(), String> {
    let chunks = simple_run_chunks(paragraph).ok_or_else(|| {
        "paragraph projection unexpectedly encountered complex inline content".to_owned()
    })?;
    let current = chunks
        .iter()
        .map(|chunk| chunk.text.as_str())
        .collect::<String>();
    if current == target {
        return Ok(());
    }
    if target.is_empty() {
        let template = chunks
            .first()
            .map(|chunk| chunk.run.clone())
            .unwrap_or(Run {
                node_type: RunType::Run,
                formatting: None,
                property_changes: None,
                content: Vec::new(),
            });
        paragraph.content = vec![ParagraphContent::Inline(InlineNode::Run(run_with_text(
            &template,
            String::new(),
        )))];
        return Ok(());
    }
    let (prefix, suffix) = common_utf16_edges(&current, target);
    let current_len = utf16_len(&current);
    let target_len = utf16_len(target);
    let replacement = utf16_slice(target, prefix, target_len - suffix);
    let mut output = Vec::new();
    let mut cursor = 0_u32;
    let mut insertion_template = chunks.first().map(|chunk| chunk.run.clone());
    for chunk in &chunks {
        let length = utf16_len(&chunk.text);
        let end = cursor + length;
        if cursor < prefix {
            let keep_end = end.min(prefix);
            let text = utf16_slice(&chunk.text, 0, keep_end - cursor);
            if !text.is_empty() {
                output.push(run_with_text(&chunk.run, text));
                insertion_template = Some(chunk.run.clone());
            }
        } else if insertion_template.is_none() {
            insertion_template = Some(chunk.run.clone());
        }
        cursor = end;
    }
    if !replacement.is_empty() {
        let template = insertion_template
            .or_else(|| chunks.first().map(|chunk| chunk.run.clone()))
            .unwrap_or(Run {
                node_type: RunType::Run,
                formatting: None,
                property_changes: None,
                content: Vec::new(),
            });
        output.push(run_with_text(&template, replacement));
    }
    let suffix_start = current_len - suffix;
    cursor = 0;
    for chunk in &chunks {
        let length = utf16_len(&chunk.text);
        let end = cursor + length;
        if end > suffix_start {
            let start = suffix_start.saturating_sub(cursor);
            let text = utf16_slice(&chunk.text, start, length);
            if !text.is_empty() {
                output.push(run_with_text(&chunk.run, text));
            }
        }
        cursor = end;
    }
    paragraph.content = output
        .into_iter()
        .map(|run| ParagraphContent::Inline(InlineNode::Run(run)))
        .collect();
    Ok(())
}

#[derive(Clone)]
struct RunChunk {
    run: Run,
    text: String,
}

fn simple_run_chunks(paragraph: &Paragraph) -> Option<Vec<RunChunk>> {
    let mut chunks = Vec::new();
    for content in &paragraph.content {
        let ParagraphContent::Inline(InlineNode::Run(run)) = content else {
            return None;
        };
        let mut text = String::new();
        for content in &run.content {
            let RunContent::Text { text: value, .. } = content else {
                return None;
            };
            text.push_str(value);
        }
        chunks.push(RunChunk {
            run: run.clone(),
            text,
        });
    }
    Some(chunks)
}

fn run_with_text(template: &Run, text: String) -> Run {
    Run {
        content: vec![RunContent::Text {
            preserve_space: (text.trim() != text).then_some(true),
            text,
        }],
        ..template.clone()
    }
}

#[derive(Clone)]
struct ParagraphXmlSpan {
    range: Range<usize>,
    para_id: Option<String>,
}

fn patch_document_paragraphs(
    original: &[u8],
    projected: &[u8],
    changed_para_ids: &[String],
    synthetic_ids: &HashSet<String>,
) -> Result<Vec<u8>> {
    let projected_parts = ooxml_opc::unzip_parts(projected).map_err(anyhow::Error::msg)?;
    let projected_document = projected_parts
        .iter()
        .find(|(path, _)| path.eq_ignore_ascii_case("word/document.xml"))
        .map(|(_, bytes)| bytes.as_slice())
        .context("projected DOCX has no word/document.xml")?;
    let projected_spans = paragraph_xml_spans(projected_document)?;
    let mut original_parts = ooxml_opc::unzip_parts(original).map_err(anyhow::Error::msg)?;
    let original_document = original_parts
        .iter_mut()
        .find(|(path, _)| path.eq_ignore_ascii_case("word/document.xml"))
        .map(|(_, bytes)| bytes)
        .context("source DOCX has no word/document.xml")?;
    let original_spans = paragraph_xml_spans(original_document)?;
    if original_spans.len() != projected_spans.len() {
        bail!(
            "cannot save safely: source has {} paragraph spans but projection has {}",
            original_spans.len(),
            projected_spans.len()
        );
    }
    let mut replacements = Vec::with_capacity(changed_para_ids.len());
    for para_id in changed_para_ids {
        let matches = projected_spans
            .iter()
            .enumerate()
            .filter(|(_, span)| span.para_id.as_deref() == Some(para_id))
            .collect::<Vec<_>>();
        let [(index, projected_span)] = matches.as_slice() else {
            bail!("cannot save paragraph {para_id}: projection identity is not unique");
        };
        let mut replacement = projected_document[projected_span.range.clone()].to_vec();
        if synthetic_ids.contains(para_id) {
            let value =
                String::from_utf8(replacement).context("projected paragraph is not UTF-8")?;
            let attribute = format!(" w14:paraId=\"{para_id}\"");
            if !value.contains(&attribute) {
                bail!("cannot remove synthetic paragraph identity {para_id}");
            }
            replacement = value.replacen(&attribute, "", 1).into_bytes();
        }
        replacements.push((original_spans[*index].range.clone(), replacement));
    }
    replacements.sort_unstable_by_key(|(range, _)| std::cmp::Reverse(range.start));
    for (range, replacement) in replacements {
        original_document.splice(range, replacement);
    }
    ooxml_opc::rezip_parts(&original_parts).map_err(anyhow::Error::msg)
}

fn paragraph_xml_spans(xml: &[u8]) -> Result<Vec<ParagraphXmlSpan>> {
    let mut spans = Vec::new();
    let mut open = Vec::new();
    let mut cursor = 0;
    while let Some(relative) = xml[cursor..].iter().position(|byte| *byte == b'<') {
        let start = cursor + relative;
        if xml[start..].starts_with(b"<!--") {
            cursor = find_xml_bytes(xml, start + 4, b"-->")?
                .checked_add(3)
                .context("XML comment position overflow")?;
            continue;
        }
        if xml[start..].starts_with(b"<![CDATA[") {
            cursor = find_xml_bytes(xml, start + 9, b"]]>")?
                .checked_add(3)
                .context("XML CDATA position overflow")?;
            continue;
        }
        let end = xml_tag_end(xml, start)?;
        let tag = &xml[start..=end];
        if is_open_paragraph_tag(tag) {
            let para_id = paragraph_tag_id(tag);
            if is_self_closing_tag(tag) {
                spans.push(ParagraphXmlSpan {
                    range: start..end + 1,
                    para_id,
                });
            } else {
                open.push((start, para_id));
            }
        } else if tag == b"</w:p>" {
            let (paragraph_start, para_id) = open
                .pop()
                .context("word/document.xml has an unmatched paragraph close")?;
            spans.push(ParagraphXmlSpan {
                range: paragraph_start..end + 1,
                para_id,
            });
        }
        cursor = end + 1;
    }
    if !open.is_empty() {
        bail!("word/document.xml has an unclosed paragraph");
    }
    spans.sort_unstable_by_key(|span| span.range.start);
    Ok(spans)
}

fn find_xml_bytes(xml: &[u8], start: usize, needle: &[u8]) -> Result<usize> {
    xml.get(start..)
        .and_then(|tail| {
            tail.windows(needle.len())
                .position(|window| window == needle)
        })
        .map(|relative| start + relative)
        .context("word/document.xml contains an unterminated XML section")
}

fn xml_tag_end(xml: &[u8], start: usize) -> Result<usize> {
    let mut quote = None;
    for (relative, byte) in xml[start + 1..].iter().copied().enumerate() {
        match (quote, byte) {
            (None, b'\'' | b'"') => quote = Some(byte),
            (Some(open), close) if open == close => quote = None,
            (None, b'>') => return Ok(start + relative + 1),
            _ => {}
        }
    }
    bail!("word/document.xml contains an unterminated tag")
}

fn is_open_paragraph_tag(tag: &[u8]) -> bool {
    tag.starts_with(b"<w:p")
        && matches!(
            tag.get(4),
            Some(b'>') | Some(b'/') | Some(b' ' | b'\t' | b'\r' | b'\n')
        )
}

fn is_self_closing_tag(tag: &[u8]) -> bool {
    tag[..tag.len().saturating_sub(1)]
        .iter()
        .rev()
        .find(|byte| !byte.is_ascii_whitespace())
        == Some(&b'/')
}

fn paragraph_tag_id(tag: &[u8]) -> Option<String> {
    let needle = b"w14:paraId";
    let start = tag
        .windows(needle.len())
        .position(|window| window == needle)?;
    let mut cursor = start + needle.len();
    while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
        cursor += 1;
    }
    if tag.get(cursor) != Some(&b'=') {
        return None;
    }
    cursor += 1;
    while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
        cursor += 1;
    }
    let quote = *tag.get(cursor)?;
    if !matches!(quote, b'\'' | b'"') {
        return None;
    }
    cursor += 1;
    let end = tag[cursor..].iter().position(|byte| *byte == quote)? + cursor;
    String::from_utf8(tag[cursor..end].to_vec()).ok()
}

fn common_utf16_edges(left: &str, right: &str) -> (u32, u32) {
    let left_units = left.encode_utf16().collect::<Vec<_>>();
    let right_units = right.encode_utf16().collect::<Vec<_>>();
    let mut prefix = 0;
    while prefix < left_units.len()
        && prefix < right_units.len()
        && left_units[prefix] == right_units[prefix]
    {
        prefix += 1;
    }
    if prefix > 0 && prefix < left_units.len() && (0xdc00..=0xdfff).contains(&left_units[prefix]) {
        prefix -= 1;
    }
    let mut suffix = 0;
    while suffix < left_units.len().saturating_sub(prefix)
        && suffix < right_units.len().saturating_sub(prefix)
        && left_units[left_units.len() - 1 - suffix] == right_units[right_units.len() - 1 - suffix]
    {
        suffix += 1;
    }
    if suffix > 0
        && suffix < left_units.len()
        && (0xd800..=0xdbff).contains(&left_units[left_units.len() - suffix])
    {
        suffix -= 1;
    }
    (prefix as u32, suffix as u32)
}

fn utf16_slice(text: &str, start: u32, end: u32) -> String {
    let units = text.encode_utf16().collect::<Vec<_>>();
    let start = usize::try_from(start)
        .unwrap_or(usize::MAX)
        .min(units.len());
    let end = usize::try_from(end)
        .unwrap_or(usize::MAX)
        .min(units.len())
        .max(start);
    String::from_utf16_lossy(&units[start..end])
}
