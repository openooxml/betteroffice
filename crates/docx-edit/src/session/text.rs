use serde_json::{Value, json};
use yrs::Any;

use super::args::{
    apply_mark, edit_ctx, find_para_span, json_to_any, loc_index, parse_inline_format_delta,
    parse_para_attr_delta, text_receipt_json,
};
use super::{EditSessionCore, EngineError, EngineResult};
use crate::{
    EditCtx, FormatPolicy, MergeDirection, ParaAttrDelta, ParaSelector, Position, StoryRange,
};

impl EditSessionCore {
    pub fn insert_text(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        text: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<String> {
        let at = loc_index(self.engine.doc(), story, para_id, offset)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let doc = self.engine.doc();
        let before = doc.committed_epoch();
        let indexed = doc
            .paragraph_index(story)
            .map_err(EngineError::message)?
            .para_at(at)
            .is_some_and(|paragraph| paragraph.node_start <= at && at <= paragraph.pilcrow);
        let mut attrs = None;
        let receipt = self
            .engine
            .doc()
            .insert_text_observed(
                &ctx,
                Position::new(story, at),
                text,
                FormatPolicy::Inherit,
                |effective| attrs = Some(effective.clone()),
            )
            .map_err(EngineError::message)?;
        if indexed {
            if let Some(attrs) = attrs.as_ref() {
                doc.advance_indexes_after_formatted_text_insert(
                    story,
                    before,
                    doc.committed_epoch(),
                    at,
                    text,
                    Some(attrs),
                );
            } else {
                doc.advance_indexes_after_text_insert(
                    story,
                    before,
                    doc.committed_epoch(),
                    at,
                    text,
                );
            }
        }
        Ok(text_receipt_json(receipt))
    }

    pub fn edit_text_resident(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        text: Option<String>,
    ) -> EngineResult<String> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        if text.is_some() && start != end {
            return Err(EngineError::message(
                "resident text insertion needs a collapsed range",
            ));
        }
        if !self
            .engine
            .takes_resident_text_edit(story, start_para, end_para, text.as_deref())
        {
            self.resident_text_edit.replace(None);
            return match text {
                Some(text) => self.insert_text(story, start_para, start_offset, &text, None, None),
                None => self.delete_range(
                    story,
                    start_para,
                    start_offset,
                    end_para,
                    end_offset,
                    None,
                    None,
                ),
            };
        }
        let receipt = self
            .engine
            .edit_resident_text(StoryRange::new(story, start, end), text.as_deref(), true)
            .map_err(EngineError::message)?;
        self.resident_text_edit
            .replace(Some((story.to_owned(), self.engine.doc_epoch())));
        Ok(text_receipt_json(receipt))
    }

    #[allow(clippy::too_many_arguments)]
    pub fn delete_range(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<String> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let doc = self.engine.doc();
        let before = doc.committed_epoch();
        let indexed = !ctx.is_suggesting()
            && doc
                .segment_index(story)
                .map_err(EngineError::message)?
                .is_text_range(start, end)
            && doc
                .paragraph_index(story)
                .map_err(EngineError::message)?
                .para_at(start)
                .is_some_and(|paragraph| paragraph.node_start <= start && end <= paragraph.pilcrow);
        let receipt = self
            .engine
            .doc()
            .delete_range(&ctx, StoryRange::new(story, start, end))
            .map_err(EngineError::message)?;
        if indexed {
            doc.advance_indexes_after_text_delete(story, before, doc.committed_epoch(), start, end);
        }
        Ok(text_receipt_json(receipt))
    }

    #[allow(clippy::too_many_arguments)]
    pub fn replace_range(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        text: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<String> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .replace_range(&ctx, StoryRange::new(story, start, end), text)
            .map_err(EngineError::message)?;
        Ok(text_receipt_json(receipt))
    }

    pub fn insert_line_break(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<String> {
        let at = loc_index(self.engine.doc(), story, para_id, offset)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .insert_hard_break(&ctx, Position::new(story, at))
            .map_err(EngineError::message)?;
        Ok(text_receipt_json(receipt))
    }

    pub fn split_paragraph(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<String> {
        let at = loc_index(self.engine.doc(), story, para_id, offset)?;
        let receipt = if author_name.is_none() && author_date.is_none() {
            let receipt = self
                .engine
                .split_resident_paragraph(Position::new(story, at))
                .map_err(EngineError::message)?;
            self.resident_text_edit
                .replace(Some((story.to_owned(), self.engine.doc_epoch())));
            receipt
        } else {
            let ctx = edit_ctx(author_name, author_date)?;
            self.engine
                .doc()
                .split_paragraph(&ctx, Position::new(story, at), None)
                .map_err(EngineError::message)?
        };
        Ok(json!({
            "firstParaId": receipt.first_para_id,
            "secondParaId": receipt.second_para_id,
            "revisionId": receipt.revision_ids.into_iter().next(),
        })
        .to_string())
    }

    pub fn merge_paragraphs(
        &self,
        story: &str,
        para_id: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<String> {
        // Validate story membership (story-scoped "not found") before merging.
        find_para_span(self.engine.doc(), story, para_id)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .merge_paragraphs(&ctx, para_id, MergeDirection::Forward)
            .map_err(EngineError::message)?;
        Ok(json!({ "revisionId": receipt.revision_ids.into_iter().next() }).to_string())
    }

    pub fn toggle_mark(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        mark_json: &str,
    ) -> EngineResult<()> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        apply_mark(
            self.engine.doc(),
            StoryRange::new(story, start, end),
            mark_json,
        )
    }

    pub fn format_range(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        delta_json: &str,
    ) -> EngineResult<()> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let delta = parse_inline_format_delta(delta_json)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .format_range(&ctx, StoryRange::new(story, start, end), &delta)
            .map(|_| ())
            .map_err(EngineError::message)
    }

    pub fn set_hyperlink(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        hyperlink_json: &str,
    ) -> EngineResult<()> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let value: Value = serde_json::from_str(hyperlink_json).map_err(EngineError::message)?;
        let hyperlink = if value.is_null() {
            None
        } else if value.is_object() {
            Some(json_to_any(&value)?)
        } else {
            return Err(EngineError::message(
                "set_hyperlink expects an object or null",
            ));
        };
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .set_hyperlink(&ctx, StoryRange::new(story, start, end), hyperlink)
            .map(|_| ())
            .map_err(EngineError::message)
    }

    pub fn clear_formatting(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
    ) -> EngineResult<()> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .clear_formatting(&ctx, StoryRange::new(story, start, end))
            .map(|_| ())
            .map_err(EngineError::message)
    }

    pub fn next_paragraph_style(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
    ) -> EngineResult<Option<String>> {
        let span = find_para_span(self.engine.doc(), story, para_id)?;
        if offset != span.pilcrow - span.start {
            return Ok(None);
        }
        let styles = self.paragraph_style_resolver()?;
        let Some(styles) = styles.as_ref() else {
            return Ok(None);
        };
        let current = self
            .engine
            .doc()
            .selection_context(&StoryRange::new(story, span.pilcrow, span.pilcrow))
            .map_err(EngineError::message)?
            .style_id;
        Ok(styles.next_paragraph_style(current.as_deref()))
    }

    #[allow(clippy::too_many_arguments)]
    pub fn apply_paragraph_style(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        style_id: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<()> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let selector = ParaSelector::Range(StoryRange::new(story, start, end));
        let ctx = edit_ctx(author_name, author_date)?;
        let styles = self.paragraph_style_resolver()?;
        if let Some(styles) = styles.as_ref() {
            self.engine
                .doc()
                .restyle_paragraphs(&ctx, &selector, style_id, styles)
                .map(|_| ())
                .map_err(EngineError::message)
        } else {
            let mut delta = ParaAttrDelta::default();
            delta
                .other
                .insert("pStyle".to_owned(), Some(Any::from(style_id)));
            self.engine
                .doc()
                .set_paragraph_attrs(&ctx, &selector, &delta)
                .map(|_| ())
                .map_err(EngineError::message)
        }
    }

    #[allow(clippy::too_many_arguments)]
    pub fn set_paragraph_attrs(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        attrs_json: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<()> {
        let start = loc_index(self.engine.doc(), story, start_para, start_offset)?;
        let end = loc_index(self.engine.doc(), story, end_para, end_offset)?;
        let selector = ParaSelector::Range(StoryRange::new(story, start, end));
        let delta = parse_para_attr_delta(attrs_json)?;
        let ctx = edit_ctx(author_name, author_date)?;
        self.engine
            .doc()
            .set_paragraph_attrs(&ctx, &selector, &delta)
            .map(|_| ())
            .map_err(EngineError::message)
    }
}
