use serde_json::json;
use yrs::Assoc;

use super::args::{IndexedLoc, attrs_value, index_loc, loc_index, tri_state_value};
use super::{
    EditSessionCore, EngineError, EngineResult, LocalCellPoint, LocalCellSelection, LocalSelection,
};
use crate::presence::{encode_sticky, resolve_sticky_selection};
use crate::{EditTextView, Loc, LocRange, StoryRange, TableRange};

impl EditSessionCore {
    /// Stores this peer's anchor and head as sticky positions, replacing any
    /// previous selection. Both endpoints must lie in `story`. The positions
    /// live outside the yrs document, so they are never serialized as content
    /// or carried in an update. `Assoc::After` makes a collapsed caret advance
    /// with text inserted at it.
    #[allow(clippy::too_many_arguments)]
    pub fn set_selection(
        &self,
        story: &str,
        anchor_para: &str,
        anchor_offset: u32,
        head_para: &str,
        head_offset: u32,
    ) -> EngineResult<()> {
        let anchor_index = loc_index(self.engine.doc(), story, anchor_para, anchor_offset)?;
        let head_index = loc_index(self.engine.doc(), story, head_para, head_offset)?;
        self.select_indexes(story, anchor_index, head_index)
    }

    /// Typing with the caret inside a surrogate pair lands before the pair,
    /// while the sticky head stays inside it: collapse the selection after the
    /// inserted text instead.
    pub(crate) fn settle_snapped_input(
        &self,
        story: &str,
        requested: &IndexedLoc,
        receipt: &crate::Receipt,
    ) -> EngineResult<()> {
        let Some(range) = receipt.range.as_ref() else {
            return Ok(());
        };
        if range.start.para == requested.para_id && range.start.offset == requested.offset {
            return Ok(());
        }
        let end = self
            .engine
            .doc()
            .locate_range(range)
            .map_err(EngineError::message)?
            .end;
        self.select_indexes(story, end, end)
    }

    fn select_indexes(&self, story: &str, anchor_index: u32, head_index: u32) -> EngineResult<()> {
        let doc = self.engine.doc();
        let anchor = doc
            .cached_sticky_index(story, anchor_index, Assoc::After)
            .map_err(EngineError::message)?
            .ok_or_else(|| EngineError::message("selection anchor could not be made sticky"))?;
        let head = if anchor_index == head_index {
            anchor.clone()
        } else {
            doc.cached_sticky_index(story, head_index, Assoc::After)
                .map_err(EngineError::message)?
                .ok_or_else(|| EngineError::message("selection head could not be made sticky"))?
        };
        self.engine.set_resident_caret_head(
            (story == "body" || story.starts_with("body:"))
                .then(|| (story.to_owned(), head.clone())),
        );
        *self.selection.borrow_mut() = Some(LocalSelection {
            story: story.to_owned(),
            anchor,
            head,
        });
        self.undo.select_story(story);
        Ok(())
    }

    /// Encodes one Loc as opaque sticky-position bytes that keep pointing at
    /// the same content as the story is edited. Publish them over an awareness
    /// transport and resolve them with
    /// [`EditSessionCore::resolve_sticky_position`].
    pub fn encode_sticky_position(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
    ) -> EngineResult<Vec<u8>> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let position = self
            .engine
            .doc()
            .cached_sticky_index(story, index, Assoc::After)
            .map_err(EngineError::message)?
            .ok_or_else(|| EngineError::message("position could not be made sticky"))?;
        Ok(encode_sticky(&position))
    }

    /// Resolves sticky bytes from [`EditSessionCore::encode_sticky_position`] back
    /// to `{"story","paraId","offset"}`. Errors when the bytes are malformed
    /// or the position no longer resolves in `story`.
    pub fn resolve_sticky_position(&self, story: &str, position: &[u8]) -> EngineResult<String> {
        let index = self
            .engine
            .doc()
            .resolve_cached_sticky_position(story, position)
            .map_err(EngineError::message)?;
        let loc = index_loc(self.engine.doc(), story, index)?;
        Ok(json!({
            "story": story,
            "paraId": loc.para_id,
            "offset": loc.offset,
        })
        .to_string())
    }

    /// This peer's current selection as `{"anchor":{"story","paraId","offset"},
    /// "head":{…}}`, or `"null"` before [`EditSessionCore::set_selection`] is
    /// called. Errors when an endpoint no longer resolves.
    pub fn selection(&self) -> EngineResult<String> {
        let selection = self.selection.borrow();
        let Some(selection) = selection.as_ref() else {
            return Ok("null".to_owned());
        };
        let (anchor_index, head_index) = self
            .engine
            .doc()
            .resolve_cached_sticky_selection(&selection.story, &selection.anchor, &selection.head)
            .map_err(EngineError::message)?;
        let anchor = index_loc(self.engine.doc(), &selection.story, anchor_index)?;
        let head = index_loc(self.engine.doc(), &selection.story, head_index)?;
        Ok(json!({
            "anchor": {
                "story": selection.story,
                "paraId": anchor.para_id,
                "offset": anchor.offset,
            },
            "head": {
                "story": selection.story,
                "paraId": head.para_id,
                "offset": head.offset,
            }
        })
        .to_string())
    }

    /// This peer's selection in transportable form:
    /// `{"story","anchor":[byte, …],"head":[byte, …]}` with sticky-encoded
    /// endpoints, or `"null"` before a selection is set. A peer resolves it
    /// with [`EditSessionCore::resolve_encoded_selection`].
    pub fn encoded_selection(&self) -> EngineResult<String> {
        let selection = self.selection.borrow();
        let Some(selection) = selection.as_ref() else {
            return Ok("null".to_owned());
        };
        Ok(json!({
            "story": selection.story,
            "anchor": encode_sticky(&selection.anchor),
            "head": encode_sticky(&selection.head),
        })
        .to_string())
    }

    /// Resolves another peer's encoded selection against this replica's
    /// current state: `{"anchor":{"story","paraId","offset"},"head":{…}}`.
    /// Errors on malformed bytes or an endpoint that no longer resolves.
    pub fn resolve_encoded_selection(
        &self,
        story: &str,
        anchor: &[u8],
        head: &[u8],
    ) -> EngineResult<String> {
        let (anchor_index, head_index) =
            resolve_sticky_selection(self.engine.doc(), story, anchor, head)
                .map_err(EngineError::message)?;
        let anchor = index_loc(self.engine.doc(), story, anchor_index)?;
        let head = index_loc(self.engine.doc(), story, head_index)?;
        Ok(json!({
            "anchor": {
                "story": story,
                "paraId": anchor.para_id,
                "offset": anchor.offset,
            },
            "head": {
                "story": story,
                "paraId": head.para_id,
                "offset": head.offset,
            }
        })
        .to_string())
    }

    /// Stores a rectangular anchor-cell to head-cell selection outside the yrs
    /// document. `range_json` is a [`TableRange`]:
    /// `{"anchor":{"story","tableIndex","row","column"},"head":{…}}`. The table
    /// embed is held by a sticky index and each endpoint by its cell story's
    /// stable identity, so the selection survives unrelated edits and follows
    /// inserted or deleted rows and columns. Errors when the two endpoints are
    /// not in the same table, or when either cell does not resolve.
    pub fn set_cell_selection(&self, range_json: &str) -> EngineResult<()> {
        let range: TableRange = serde_json::from_str(range_json).map_err(EngineError::message)?;
        if range.anchor.story != range.head.story
            || range.anchor.table_index != range.head.table_index
        {
            return Err(EngineError::message(
                "a cell selection must stay inside one table",
            ));
        }
        let (anchor, anchor_story) = self
            .engine
            .doc()
            .resolve_cell_identity(&range.anchor)
            .map_err(EngineError::message)?;
        let (head, head_story) = self
            .engine
            .doc()
            .resolve_cell_identity(&range.head)
            .map_err(EngineError::message)?;
        let locator = anchor.table();
        let table_index = self
            .engine
            .doc()
            .table_embed_index(&locator)
            .map_err(EngineError::message)?;
        let table = self
            .engine
            .doc()
            .cached_sticky_index(&locator.story, table_index, Assoc::Before)
            .map_err(EngineError::message)?
            .ok_or_else(|| EngineError::message("table selection could not be made sticky"))?;
        *self.cell_selection.borrow_mut() = Some(LocalCellSelection {
            parent_story: locator.story,
            table,
            anchor: LocalCellPoint {
                cell_story: anchor_story,
                row: anchor.row,
                column: anchor.column,
            },
            head: LocalCellPoint {
                cell_story: head_story,
                row: head.row,
                column: head.column,
            },
        });
        Ok(())
    }

    /// The current local cell selection as a [`TableRange`], or `"null"`
    /// before one is set. An endpoint whose cell was deleted clamps to a
    /// surviving nearby cell. Errors when the table itself no longer resolves.
    pub fn cell_selection(&self) -> EngineResult<String> {
        let selection = self.cell_selection.borrow();
        let Some(selection) = selection.as_ref() else {
            return Ok("null".to_owned());
        };
        let table_index = self
            .engine
            .doc()
            .cached_sticky_offset(&selection.parent_story, &selection.table)
            .map_err(EngineError::message)?
            .ok_or_else(|| EngineError::message("cell selection table no longer resolves"))?;
        let locator = self
            .engine
            .doc()
            .table_locator_at_index(&selection.parent_story, table_index)
            .map_err(EngineError::message)?;
        let anchor = self
            .engine
            .doc()
            .cell_loc_for_story(
                &locator,
                &selection.anchor.cell_story,
                selection.anchor.row,
                selection.anchor.column,
            )
            .map_err(EngineError::message)?;
        let head = self
            .engine
            .doc()
            .cell_loc_for_story(
                &locator,
                &selection.head.cell_story,
                selection.head.row,
                selection.head.column,
            )
            .map_err(EngineError::message)?;
        serde_json::to_string(&TableRange { anchor, head }).map_err(EngineError::message)
    }

    pub fn clear_cell_selection(&self) -> Option<String> {
        let selection = self.cell_selection.borrow_mut().take()?;
        let index = self
            .engine
            .doc()
            .cached_sticky_offset(&selection.parent_story, &selection.table)
            .ok()??;
        let paragraphs = self
            .engine
            .doc()
            .paragraph_index(&selection.parent_story)
            .ok()?;
        let paragraph = paragraphs.para_at(index)?;
        Some(
            json!({ "story": selection.parent_story, "paraId": paragraph.para_id,
                "offset": index.saturating_sub(paragraph.start).min(paragraph.pilcrow - paragraph.start) }).to_string(),
        )
    }

    /// Accepted-view texts around a paragraph-keyed selection:
    /// `{"paraId","selectedText","paragraphText","before","after"}`, with
    /// `\n` between paragraphs and U+FFFC for each inline atom.
    #[allow(clippy::too_many_arguments)]
    pub fn selection_text_json(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
    ) -> EngineResult<String> {
        let range = LocRange::new(
            Loc::new(story, start_para, start_offset),
            Loc::new(story, end_para, end_offset),
        );
        let info = self
            .engine
            .doc()
            .selection_text(&range, EditTextView::Accepted)
            .map_err(EngineError::message)?;
        Ok(json!({
            "paraId": info.para_id,
            "selectedText": info.selected_text,
            "paragraphText": info.paragraph_text,
            "before": info.before,
            "after": info.after,
        })
        .to_string())
    }

    /// Aggregated toolbar and accessibility state over `[start, end)`:
    ///
    /// ```json
    /// {
    ///   "bold": true | false | "mixed", "italic": …, "underline": …, "strike": …,
    ///   "superscript": …, "subscript": …,
    ///   "fontFamily": string|null, "fontSize": number|null, "color": string|null,
    ///   "highlight": string|null,
    ///   "paraId": string, "styleId": string|null, "alignment": string|null,
    ///   "paragraphProperties": {…},
    ///   "hasSelection": bool, "isMultiParagraph": bool, "inTable": bool,
    ///   "isSingleEmbed": bool, "embedKind": string|null, "isImage": bool,
    ///   "inInsertion": bool, "inDeletion": bool
    /// }
    /// ```
    ///
    /// A toggle mark is `"mixed"` when the range disagrees about it; a value
    /// mark is `null` when it is absent OR not uniform, so the two cases are
    /// not distinguished. `isImage` is `embedKind == "image"`, and
    /// `inInsertion`/`inDeletion` report whether the range sits inside a
    /// pending tracked change.
    #[allow(clippy::too_many_arguments)]
    pub fn selection_context(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
    ) -> Result<String, EngineError> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let context = self
            .engine
            .doc()
            .selection_context(&StoryRange::new(story, start, end))
            .map_err(EngineError::message)?;
        let is_single_embed = context.embed_kind.is_some();
        let is_image = context.embed_kind.as_deref() == Some("image");
        Ok(json!({
            "bold": tri_state_value(context.bold),
            "italic": tri_state_value(context.italic),
            "underline": tri_state_value(context.underline),
            "strike": tri_state_value(context.strike),
            "superscript": tri_state_value(context.superscript),
            "subscript": tri_state_value(context.subscript),
            "fontFamily": context.font_family,
            "fontSize": context.font_size,
            "color": context.color,
            "highlight": context.highlight,
            "paraId": context.para_id,
            "styleId": context.style_id,
            "alignment": context.alignment,
            "paragraphProperties": attrs_value(&context.paragraph_properties)?,
            "hasSelection": context.has_selection,
            "isMultiParagraph": context.is_multi_paragraph,
            "inTable": context.in_table,
            "isSingleEmbed": is_single_embed,
            "embedKind": context.embed_kind,
            "isImage": is_image,
            "inInsertion": context.in_insertion,
            "inDeletion": context.in_deletion,
        })
        .to_string())
    }
}
