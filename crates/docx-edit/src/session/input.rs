use serde::Serialize;

use super::args::{AdjacentStoryUnit, DeleteDirection, adjacent_story_unit, embed_at, index_loc};
use super::frames::validate_frame_epoch;
use super::{EditSessionCore, EngineError, EngineResult};
use crate::{EditCtx, MergeDirection, StoryRange};

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ApplyInputProfile {
    selection_ms: f64,
    edit_ms: f64,
    lower_ms: f64,
    measure_ms: f64,
    paginate_ms: f64,
    display_input_ms: f64,
    display_build_ms: f64,
    display_finalize_ms: f64,
    display_ms: f64,
    encode_ms: f64,
}

impl EditSessionCore {
    pub(crate) fn collapsed_resident_input_selection(&self) -> EngineResult<(String, String, u32)> {
        let selection = self.selection.borrow();
        let selection = selection
            .as_ref()
            .ok_or_else(|| EngineError::message("resident input requires a selection"))?;
        let story = selection.story.clone();
        let (anchor, head) = self
            .engine
            .doc()
            .resolve_cached_sticky_selection(&story, &selection.anchor, &selection.head)
            .map_err(EngineError::message)?;
        if anchor != head {
            return Err(EngineError::message(
                "resident input currently requires a collapsed selection",
            ));
        }
        let loc = index_loc(self.engine.doc(), &story, head)?;
        if !self.engine.can_apply_input(&story, &loc.para_id) {
            return Err(EngineError::message(
                "resident input state is not ready for this paragraph",
            ));
        }
        Ok((story, loc.para_id, head))
    }

    pub(crate) fn delete_resident_units(
        &self,
        direction: &str,
        count: u32,
    ) -> EngineResult<String> {
        self.resident_deleted_units.set(0);
        if count == 0 {
            return Err(EngineError::message(
                "resident delete count must be positive",
            ));
        }
        // Resident layout absorbs one paragraph merge per pass.
        let mut story = None;
        let mut merged = false;
        let mut deleted = 0;
        while deleted < count {
            let step = self
                .collapsed_resident_input_selection()
                .and_then(|selection| {
                    let merges = self.resident_unit_is_pilcrow(direction, &selection)?;
                    if merges && merged {
                        return Ok(None);
                    }
                    Ok(Some((
                        self.delete_resident_input(direction, selection, count == 1)?,
                        merges,
                    )))
                });
            match step {
                Ok(Some((from, merges))) => {
                    story = Some(from);
                    merged |= merges;
                    deleted += 1;
                }
                Err(error) if deleted == 0 => return Err(error),
                Ok(None) | Err(_) => break,
            }
        }
        self.resident_deleted_units.set(deleted);
        Ok(story.expect("the first resident deletion succeeded"))
    }

    fn resident_unit_is_pilcrow(
        &self,
        direction: &str,
        selection: &(String, String, u32),
    ) -> EngineResult<bool> {
        let direction = match direction {
            "backward" => DeleteDirection::Backward,
            "forward" => DeleteDirection::Forward,
            _ => {
                return Err(EngineError::message(
                    "delete direction must be backward or forward",
                ));
            }
        };
        Ok(matches!(
            adjacent_story_unit(self.engine.doc(), &selection.0, selection.2, direction)?,
            Some(AdjacentStoryUnit::Pilcrow)
        ))
    }

    pub(crate) fn delete_resident_input(
        &self,
        direction: &str,
        selection: (String, String, u32),
        local_lowering: bool,
    ) -> EngineResult<String> {
        let (story, para_id, head) = selection;
        let direction = match direction {
            "backward" => DeleteDirection::Backward,
            "forward" => DeleteDirection::Forward,
            _ => {
                return Err(EngineError::message(
                    "delete direction must be backward or forward",
                ));
            }
        };
        let adjacent = adjacent_story_unit(self.engine.doc(), &story, head, direction)?;
        let ctx = EditCtx::local("", "");

        match (direction, adjacent) {
            (_, Some(AdjacentStoryUnit::Table)) => {
                return Err(EngineError::message(
                    "resident input state is not ready for this paragraph",
                ));
            }
            (DeleteDirection::Backward, Some(AdjacentStoryUnit::Content(width))) => {
                self.engine
                    .edit_resident_text(
                        StoryRange::new(&story, head - width, head),
                        None,
                        local_lowering,
                    )
                    .map_err(EngineError::message)?;
            }
            (DeleteDirection::Forward, Some(AdjacentStoryUnit::Content(width))) => {
                self.engine
                    .edit_resident_text(
                        StoryRange::new(&story, head, head + width),
                        None,
                        local_lowering,
                    )
                    .map_err(EngineError::message)?;
            }
            (direction, Some(AdjacentStoryUnit::Pilcrow)) => {
                // Merging forward over a table or other embed would pull it
                // into the paragraph; the host path owns that case.
                if matches!(direction, DeleteDirection::Forward)
                    && embed_at(self.engine.doc(), &story, head + 1)?
                {
                    return Err(EngineError::message(
                        "resident input state is not ready for this paragraph",
                    ));
                }
                let paragraphs = self
                    .engine
                    .doc()
                    .paragraphs(&story)
                    .map_err(EngineError::message)?;
                let paragraph_index = paragraphs
                    .iter()
                    .position(|paragraph| paragraph.para_id == para_id)
                    .ok_or_else(|| {
                        EngineError::message("resident input paragraph no longer resolves")
                    })?;
                let merge = match direction {
                    DeleteDirection::Backward if paragraph_index > 0 => {
                        Some(MergeDirection::Backward)
                    }
                    DeleteDirection::Forward if paragraph_index + 1 < paragraphs.len() => {
                        Some(MergeDirection::Forward)
                    }
                    _ => None,
                };
                let Some(merge) = merge else {
                    return Err(EngineError::message(
                        "resident input has no character in that direction",
                    ));
                };
                self.engine
                    .doc()
                    .merge_paragraphs(&ctx, &para_id, merge)
                    .map_err(EngineError::message)?;
            }
            (_, None) => {
                return Err(EngineError::message(
                    "resident input has no character in that direction",
                ));
            }
        }
        Ok(story)
    }

    /// `{"frameEpoch", "caretRect": {…}|null}` for the session's own collapsed
    /// body selection. `caretRect` is null whenever there is no selection, the
    /// selection is not a collapsed body caret, or the retained layout has no
    /// geometry for it. `frameEpoch` identifies the frame the rect belongs to.
    pub fn resident_caret_snapshot_json(&self) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        let paragraph = {
            let selection = self.selection.borrow();
            let Some(selection) = selection.as_ref() else {
                return serde_json::to_string(
                    &self
                        .engine
                        .resident_caret_snapshot(None)
                        .map_err(EngineError::message)?,
                )
                .map_err(EngineError::message);
            };
            if selection.story != "body" {
                None
            } else {
                let doc = self.engine.doc();
                let anchor = doc
                    .cached_sticky_offset(&selection.story, &selection.anchor)
                    .ok()
                    .flatten();
                let head = if selection.anchor == selection.head {
                    anchor
                } else {
                    doc.cached_sticky_offset(&selection.story, &selection.head)
                        .ok()
                        .flatten()
                };
                match (anchor, head) {
                    (Some(anchor), Some(head)) if anchor == head => {
                        let loc = index_loc(self.engine.doc(), &selection.story, head)?;
                        Some((loc.para_id, loc.node_offset))
                    }
                    _ => None,
                }
            }
        };
        let snapshot = self
            .engine
            .resident_caret_snapshot(
                paragraph
                    .as_ref()
                    .map(|(id, offset)| (id.as_str(), *offset)),
            )
            .map_err(EngineError::message)?;
        serde_json::to_string(&snapshot).map_err(EngineError::message)
    }

    /// Applies one ordinary insertion at this session's collapsed selection
    /// and returns the resulting binary `FrameDelta`. The inserted text
    /// inherits the formatting at the caret; selection, measurement inputs,
    /// pagination checkpoints and display state all stay resident, so nothing
    /// but the frame crosses the boundary.
    ///
    /// Errors when `text` is empty or contains `\r`/`\n`, when
    /// `expected_frame_epoch` is not a non-negative safe integer, when there is
    /// no resident selection or it no longer resolves, when the selection is
    /// not collapsed, or when the resident layout state cannot absorb an edit
    /// in that paragraph. The caller must then run the full layout path.
    pub fn apply_input(&self, text: &str, expected_frame_epoch: f64) -> EngineResult<Vec<u8>> {
        let _fonts = self.fonts.enter();
        if text.is_empty() || text.contains(['\r', '\n']) {
            return Err(EngineError::message(
                "apply_input requires non-empty paragraph-break-free text",
            ));
        }
        let epoch = validate_frame_epoch(expected_frame_epoch)?;

        let selection = self.selection.borrow().clone();
        let selection = selection
            .as_ref()
            .ok_or_else(|| EngineError::message("apply_input requires a resident selection"))?;
        let story = selection.story.clone();
        let (anchor, head) = self
            .engine
            .doc()
            .resolve_cached_sticky_selection(&story, &selection.anchor, &selection.head)
            .map_err(EngineError::message)?;
        if anchor != head {
            return Err(EngineError::message(
                "apply_input currently requires a collapsed selection",
            ));
        }
        let loc = index_loc(self.engine.doc(), &story, head)?;
        if !self.engine.can_apply_input(&story, &loc.para_id) {
            return Err(EngineError::message(
                "resident input state is not ready for this paragraph",
            ));
        }

        let receipt = self
            .engine
            .edit_resident_text(StoryRange::new(&story, head, head), Some(text), true)
            .map_err(EngineError::message)?;
        self.settle_snapped_input(&story, &loc, &receipt)?;
        self.engine
            .apply_and_layout(&story, epoch)
            .map_err(EngineError::message)
    }

    /// Instrumented twin of [`EditSessionCore::apply_input`]: identical arguments,
    /// result and error contract, but it also records stage timings for
    /// [`EditSessionCore::apply_input_profile_json`]. Separate so the ordinary
    /// input path pays no timer calls.
    pub fn apply_input_profiled(
        &self,
        text: &str,
        expected_frame_epoch: f64,
    ) -> EngineResult<Vec<u8>> {
        let _fonts = self.fonts.enter();
        if text.is_empty() || text.contains(['\r', '\n']) {
            return Err(EngineError::message(
                "apply_input requires non-empty paragraph-break-free text",
            ));
        }
        let epoch = validate_frame_epoch(expected_frame_epoch)?;

        let started = self.host.timer_ms();
        let selection = self.selection.borrow().clone();
        let selection = selection
            .as_ref()
            .ok_or_else(|| EngineError::message("apply_input requires a resident selection"))?;
        let story = selection.story.clone();
        let (anchor, head) = self
            .engine
            .doc()
            .resolve_cached_sticky_selection(&story, &selection.anchor, &selection.head)
            .map_err(EngineError::message)?;
        if anchor != head {
            return Err(EngineError::message(
                "apply_input currently requires a collapsed selection",
            ));
        }
        let loc = index_loc(self.engine.doc(), &story, head)?;
        if !self.engine.can_apply_input(&story, &loc.para_id) {
            return Err(EngineError::message(
                "resident input state is not ready for this paragraph",
            ));
        }
        let selection_ms = self.host.timer_ms() - started;

        let started = self.host.timer_ms();
        let receipt = self
            .engine
            .edit_resident_text(StoryRange::new(&story, head, head), Some(text), true)
            .map_err(EngineError::message)?;
        self.settle_snapped_input(&story, &loc, &receipt)?;
        let edit_ms = self.host.timer_ms() - started;
        let (frame, engine_profile) = self
            .engine
            .apply_and_layout_profiled(&story, epoch, &mut || self.host.timer_ms())
            .map_err(EngineError::message)?;
        let profile = ApplyInputProfile {
            selection_ms,
            edit_ms,
            lower_ms: engine_profile.lower_ms,
            measure_ms: engine_profile.measure_ms,
            paginate_ms: engine_profile.paginate_ms,
            display_input_ms: engine_profile.display_input_ms,
            display_build_ms: engine_profile.display_build_ms,
            display_finalize_ms: engine_profile.display_finalize_ms,
            display_ms: engine_profile.display_ms,
            encode_ms: engine_profile.encode_ms,
        };
        *self.last_apply_profile_json.borrow_mut() =
            serde_json::to_string(&profile).map_err(EngineError::message)?;
        Ok(frame)
    }

    /// Deletes up to `count` characters at this session's collapsed selection,
    /// lays out once, and returns the resulting binary `FrameDelta`.
    /// `direction` is `"backward"` or `"forward"`; a surrogate pair is removed
    /// whole. At a paragraph boundary a deletion merges with the neighbouring
    /// paragraph instead. Deleting stops early at the document start or end,
    /// before a second paragraph merge, or at a paragraph the resident state
    /// cannot absorb; [`EditSessionCore::resident_deleted_units`] reports how many
    /// were removed.
    ///
    /// Errors on an unknown `direction`, a zero `count`, when
    /// `expected_frame_epoch` is not a non-negative safe integer, under the
    /// same selection and readiness conditions as [`EditSessionCore::apply_input`],
    /// and when there is no character to delete in that direction.
    pub fn apply_delete(
        &self,
        direction: &str,
        expected_frame_epoch: f64,
        count: u32,
    ) -> EngineResult<Vec<u8>> {
        let _fonts = self.fonts.enter();
        let epoch = validate_frame_epoch(expected_frame_epoch)?;
        let story = self.delete_resident_units(direction, count)?;
        self.engine
            .apply_and_layout(&story, epoch)
            .map_err(EngineError::message)
    }

    /// Characters the last [`EditSessionCore::apply_delete`] removed.
    pub fn resident_deleted_units(&self) -> u32 {
        self.resident_deleted_units.get()
    }

    /// Instrumented twin of [`EditSessionCore::apply_delete`]: identical arguments,
    /// result and error contract, but it also records stage timings for
    /// [`EditSessionCore::apply_input_profile_json`].
    pub fn apply_delete_profiled(
        &self,
        direction: &str,
        expected_frame_epoch: f64,
        count: u32,
    ) -> EngineResult<Vec<u8>> {
        let _fonts = self.fonts.enter();
        let epoch = validate_frame_epoch(expected_frame_epoch)?;

        let started = self.host.timer_ms();
        self.collapsed_resident_input_selection()?;
        let selection_ms = self.host.timer_ms() - started;

        let started = self.host.timer_ms();
        let story = self.delete_resident_units(direction, count)?;
        let edit_ms = self.host.timer_ms() - started;
        let (frame, engine_profile) = self
            .engine
            .apply_and_layout_profiled(&story, epoch, &mut || self.host.timer_ms())
            .map_err(EngineError::message)?;
        let profile = ApplyInputProfile {
            selection_ms,
            edit_ms,
            lower_ms: engine_profile.lower_ms,
            measure_ms: engine_profile.measure_ms,
            paginate_ms: engine_profile.paginate_ms,
            display_input_ms: engine_profile.display_input_ms,
            display_build_ms: engine_profile.display_build_ms,
            display_finalize_ms: engine_profile.display_finalize_ms,
            display_ms: engine_profile.display_ms,
            encode_ms: engine_profile.encode_ms,
        };
        *self.last_apply_profile_json.borrow_mut() =
            serde_json::to_string(&profile).map_err(EngineError::message)?;
        Ok(frame)
    }

    /// Stage timings of the last profiled apply, in milliseconds:
    /// `{"selectionMs","editMs","lowerMs","measureMs","paginateMs",
    /// "displayInputMs","displayBuildMs","displayFinalizeMs","displayMs",
    /// "encodeMs"}`. `{}` before the first profiled call.
    pub fn apply_input_profile_json(&self) -> String {
        self.last_apply_profile_json.borrow().clone()
    }
}
