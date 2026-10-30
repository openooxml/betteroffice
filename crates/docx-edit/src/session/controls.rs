use serde_json::Value;
use yrs::{Any, Assoc, Transact};

use super::args::{edit_ctx, image_geometry_entries, loc_index, parse_payload, text_receipt_json};
use super::{EditSessionCore, EngineError, EngineResult};
use crate::batch::outcome_json;
use crate::content_controls::{ContentControlQuery, ContentControlsOptions};
use crate::presence::encode_sticky;
use crate::{EditCtx, Position};

impl EditSessionCore {
    pub fn insert_image(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        payload_json: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<String> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let Any::Map(payload) = Any::from_json(payload_json).map_err(EngineError::message)? else {
            return Err(EngineError::message("insert_image expects a JSON object"));
        };
        let mut entries: Vec<_> = payload
            .iter()
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect();
        if !payload.contains_key("noChangeAspect") {
            entries.push(("noChangeAspect".to_owned(), Any::Bool(true)));
        }
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .insert_embed(&ctx, Position::new(story, index), "image", entries)
            .map_err(EngineError::message)?;
        Ok(text_receipt_json(receipt))
    }

    pub fn move_embed(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        to_story: &str,
        to_para_id: &str,
        to_offset: u32,
    ) -> EngineResult<String> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let to_index = loc_index(self.engine.doc(), to_story, to_para_id, to_offset)?;
        let ctx = EditCtx::local(String::new(), String::new());
        let receipt = self
            .engine
            .doc()
            .move_embed(
                &ctx,
                Position::new(story, index),
                Position::new(to_story, to_index),
            )
            .map_err(EngineError::message)?;
        Ok(text_receipt_json(receipt))
    }

    pub fn set_content_control_value(&self, embed_id: &str, value_json: &str) -> EngineResult<()> {
        let value = Any::from_json(value_json).map_err(EngineError::message)?;
        self.select_embed_story(embed_id)?;
        let (story, index) = self
            .engine
            .doc()
            .embed_position(embed_id)
            .map_err(EngineError::message)?;
        if self.fill_text_control(&story, index, &value)? {
            return Ok(());
        }
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .set_embed_attrs_by_id(&ctx, embed_id, vec![("value".to_owned(), value)])
            .map(|_| ())
            .map_err(EngineError::message)
    }

    pub fn set_content_control_value_at(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        value_json: &str,
    ) -> EngineResult<()> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let value = Any::from_json(value_json).map_err(EngineError::message)?;
        if self.fill_text_control(story, index, &value)? {
            return Ok(());
        }
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .set_embed_attrs(
                &ctx,
                Position::new(story, index),
                vec![("value".to_owned(), value)],
            )
            .map(|_| ())
            .map_err(EngineError::message)
    }

    pub fn clear_content_control_value(&self, embed_id: &str) -> EngineResult<()> {
        self.select_embed_story(embed_id)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .set_embed_attrs_by_id(&ctx, embed_id, vec![("value".to_owned(), Any::Null)])
            .map(|_| ())
            .map_err(EngineError::message)
    }

    pub fn set_image_geometry(&self, embed_id: &str, geometry_json: &str) -> EngineResult<()> {
        let entries = image_geometry_entries(geometry_json)?;
        self.select_embed_story(embed_id)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .set_embed_attrs_by_id(&ctx, embed_id, entries)
            .map(|_| ())
            .map_err(EngineError::message)
    }

    pub fn set_image_geometry_at(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        geometry_json: &str,
    ) -> EngineResult<()> {
        let entries = image_geometry_entries(geometry_json)?;
        let doc = self.engine.doc();
        let at = Position::new(story, loc_index(doc, story, para_id, offset)?);
        if doc
            .embed_kind(&at)
            .map_err(EngineError::message)?
            .as_deref()
            != Some("image")
        {
            return Err(EngineError::message(format!(
                "offset {offset} of paragraph {para_id:?} holds no image"
            )));
        }
        let ctx = EditCtx::local(String::new(), String::new());
        doc.set_embed_attrs(&ctx, at, entries)
            .map(|_| ())
            .map_err(EngineError::message)
    }

    pub fn insert_page_break(&self, story: &str, para_id: &str, offset: u32) -> EngineResult<()> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .insert_embed(&ctx, Position::new(story, index), "pageBreak", vec![])
            .map(|_| ())
            .map_err(EngineError::message)
    }

    pub fn insert_section_break(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        break_type: &str,
    ) -> EngineResult<()> {
        if !matches!(
            break_type,
            "nextPage" | "continuous" | "oddPage" | "evenPage"
        ) {
            return Err(EngineError::message(format!(
                "unknown section break type {break_type:?}"
            )));
        }
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .insert_embed(
                &ctx,
                Position::new(story, index),
                "sectionBreak",
                vec![("type".to_owned(), Any::from(break_type))],
            )
            .map(|_| ())
            .map_err(EngineError::message)
    }

    pub fn insert_watermark(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        watermark_json: &str,
    ) -> EngineResult<()> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let value: Value = serde_json::from_str(watermark_json).map_err(EngineError::message)?;
        if !value.is_object() {
            return Err(EngineError::message(
                "insert_watermark expects a JSON object",
            ));
        }
        let payload = parse_payload(Some(&value))?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .insert_embed(&ctx, Position::new(story, index), "watermark", payload)
            .map(|_| ())
            .map_err(EngineError::message)
    }

    pub fn set_paragraph_attr(
        &self,
        para_id: &str,
        key: &str,
        value_json: &str,
    ) -> EngineResult<()> {
        let value = Any::from_json(value_json).map_err(EngineError::message)?;
        self.engine
            .doc()
            .set_paragraph_attr(para_id, key, value)
            .map_err(EngineError::message)
    }

    fn select_embed_story(&self, embed_id: &str) -> EngineResult<()> {
        let story = self
            .engine
            .doc()
            .embed_story(embed_id)
            .map_err(EngineError::message)?;
        self.undo.select_story(&story);
        Ok(())
    }

    /// Fills the text control at `index` of `story` with a string value; `false` when the embed
    /// there is not a text control.
    pub(crate) fn fill_text_control(
        &self,
        story: &str,
        index: u32,
        value: &Any,
    ) -> EngineResult<bool> {
        let doc = self.engine.doc();
        if !doc
            .text_control_at(story, index)
            .map_err(EngineError::message)?
        {
            return Ok(false);
        }
        let Any::String(text) = value else {
            return Err(EngineError::message(
                "a plain- or rich-text content control takes a string value",
            ));
        };
        match doc
            .fill_text_control_at(story, index, text, &self.undo)
            .map_err(EngineError::message)?
        {
            Ok(_) => Ok(true),
            Err(refusal) => Err(EngineError::message(format!(
                "{}: {}",
                serde_json::to_value(refusal.failure.code)
                    .ok()
                    .and_then(|code| code.as_str().map(str::to_owned))
                    .unwrap_or_default(),
                refusal.failure.message
            ))),
        }
    }

    pub fn content_control_targets_json(&self) -> EngineResult<String> {
        let txn = self.engine.doc().yrs_doc().transact();
        let inventory = crate::content_controls::Inventory::build(self.engine.doc(), &txn)
            .map_err(|failure| EngineError::message(failure.message))?;
        drop(txn);
        let mut values = Vec::new();
        for record in inventory
            .records
            .into_iter()
            .filter(|record| record.listed && !record.alias)
        {
            let raw = match record.site {
                crate::content_controls::Site::Inline { raw, .. }
                | crate::content_controls::Site::Block { raw, .. } => raw,
                _ => continue,
            };
            let position = self
                .engine
                .doc()
                .cached_sticky_index(&record.story, raw, Assoc::After)
                .map_err(EngineError::message)?
                .ok_or_else(|| EngineError::message("control could not be made sticky"))?;
            values.push(serde_json::json!({"control":record.control,"at":{"story":record.story,"index":raw,"encoded":encode_sticky(&position)}}));
        }
        serde_json::to_string(&values).map_err(EngineError::message)
    }

    pub fn repeat_content_control_at(
        &self,
        story: &str,
        index: u32,
        remove: bool,
    ) -> EngineResult<bool> {
        self.engine
            .doc()
            .repeat_control_at(&EditCtx::local("", ""), story, index, remove)
            .map_err(EngineError::message)
    }

    pub fn set_content_control_value_index(
        &self,
        story: &str,
        index: u32,
        json: &str,
    ) -> EngineResult<()> {
        let value = Any::from_json(json).map_err(EngineError::message)?;
        if self.fill_text_control(story, index, &value)? {
            return Ok(());
        }
        if !matches!(
            self.engine
                .doc()
                .embed_kind(&Position::new(story, index))
                .map_err(EngineError::message)?
                .as_deref(),
            Some("sdt" | "blockSdt")
        ) {
            return Err(EngineError::message("unknown-control"));
        }
        self.engine
            .doc()
            .set_embed_attrs(
                &EditCtx::local("", ""),
                Position::new(story, index),
                vec![("value".to_owned(), value)],
            )
            .map(|_| ())
            .map_err(EngineError::message)
    }

    /// The content controls of the committed state:
    /// `{"stories"?,"maxControls"?,"maxBytes"?}` -> `{"ok":true,"version","content"}` or
    /// `{"ok":false,"version","failure"}`. Control ids and anchors are scoped to the returned
    /// version. Reads only.
    pub fn list_content_controls_json(&self, options: &str) -> EngineResult<String> {
        let options: ContentControlsOptions =
            serde_json::from_str(options).map_err(EngineError::message)?;
        outcome_json(&self.engine.doc().list_content_controls(&options))
            .map_err(EngineError::message)
    }

    /// [`EditSessionCore::list_content_controls_json`] keeping the controls that match `query`
    /// (`{"kind":"id","controlId"}`, `{"kind":"tag","tag"}`, `{"kind":"ooxmlId","ooxmlId"}` or
    /// `{"kind":"alias","alias"}`) exactly.
    pub fn find_content_controls_json(&self, query: &str, options: &str) -> EngineResult<String> {
        let query: ContentControlQuery =
            serde_json::from_str(query).map_err(EngineError::message)?;
        let options: ContentControlsOptions =
            serde_json::from_str(options).map_err(EngineError::message)?;
        outcome_json(&self.engine.doc().find_content_controls(&query, &options))
            .map_err(EngineError::message)
    }
}
