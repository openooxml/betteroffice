use serde_json::json;

use super::{EditSessionCore, EngineError, EngineResult};
use crate::{EditCtx, ListKind};

impl EditSessionCore {
    pub fn part_metadata_json(&self) -> String {
        self.engine.doc().part_metadata_json()
    }

    pub fn set_part_metadata(&self, json: &str, remove_json: &str) -> EngineResult<bool> {
        let remove: Vec<String> =
            serde_json::from_str(remove_json).map_err(EngineError::message)?;
        let changed = self
            .engine
            .doc()
            .set_part_metadata(&EditCtx::local("", ""), json, &remove)
            .map_err(EngineError::message)?;
        if changed && !remove.is_empty() {
            self.cell_selection.borrow_mut().take();
        }
        Ok(changed)
    }

    pub fn document_properties_json(&self) -> String {
        self.engine.doc().document_properties_json()
    }

    pub fn set_document_property(&self, key: &str, json: &str) -> EngineResult<bool> {
        self.engine
            .doc()
            .set_document_property(&EditCtx::local("", ""), key, json)
            .map_err(EngineError::message)
    }

    pub fn style_languages_json(&self) -> String {
        self.engine.doc().style_languages_json()
    }

    pub fn set_style_language(
        &self,
        style_id: Option<&str>,
        slots_json: &str,
    ) -> EngineResult<bool> {
        self.engine
            .doc()
            .set_style_language(&EditCtx::local("", ""), style_id, slots_json)
            .map_err(EngineError::message)
    }

    pub fn patch_style_languages_xml(&self, source: &str) -> EngineResult<String> {
        self.engine
            .doc()
            .patch_style_languages_xml(source)
            .map_err(EngineError::message)
    }

    /// The paragraph attributes, as JSON, that put a paragraph on a list of `kind`, `bullet` or
    /// `numbered`; see [`crate::EditingDoc::ensure_list_definition`].
    pub fn ensure_list_definition(
        &self,
        story: &str,
        para_id: &str,
        kind: &str,
    ) -> EngineResult<String> {
        let kind = match kind {
            "bullet" => ListKind::Bullet,
            "numbered" => ListKind::Numbered,
            _ => return Err(EngineError::message(format!("unknown list kind {kind:?}"))),
        };
        let attrs = self
            .engine
            .doc()
            .ensure_list_definition(&EditCtx::local("", ""), story, para_id, kind)
            .map_err(EngineError::from)?;
        serde_json::to_string(&attrs).map_err(EngineError::message)
    }

    /// The list definitions the editor created, by key, as JSON
    /// `[{"key","kind","levelTexts","levelFormats"}]`.
    pub fn created_list_definitions_json(&self) -> String {
        let definitions: Vec<_> = self
            .engine
            .doc()
            .created_list_definitions()
            .into_iter()
            .map(|definition| {
                let formats: Vec<_> = (0..definition.level_texts.len())
                    .map(|level| definition.kind.level_format(level))
                    .collect();
                json!({
                    "key": definition.key,
                    "kind": definition.kind.name(),
                    "levelTexts": definition.level_texts,
                    "levelFormats": formats,
                })
            })
            .collect();
        serde_json::Value::Array(definitions).to_string()
    }

    pub fn list_sections_json(&self) -> EngineResult<String> {
        let sections = self
            .engine
            .doc()
            .list_sections()
            .map_err(EngineError::message)?;
        serde_json::to_string(&sections).map_err(EngineError::message)
    }

    pub fn section_margins_json(&self) -> String {
        self.engine.doc().section_margins_json()
    }

    pub fn set_section_margins(&self, section: &str, json: &str) -> EngineResult<bool> {
        self.engine
            .doc()
            .set_section_margins(&EditCtx::local("", ""), section, json)
            .map_err(EngineError::message)
    }

    /// Adds a story holding one paragraph with `initial_text` (which must not
    /// contain paragraph breaks), `p_style` and `alignment`. Receipt:
    /// `{"paraId"}`: the paragraph ending at the story's pilcrow. Errors when
    /// the story id already exists.
    pub fn create_story(
        &self,
        story_id: &str,
        initial_text: &str,
        p_style: &str,
        alignment: &str,
    ) -> EngineResult<String> {
        let para_id = self
            .engine
            .doc()
            .create_story(story_id, initial_text, p_style, alignment)
            .map_err(EngineError::from)?;
        Ok(json!({ "paraId": para_id }).to_string())
    }

    /// Removes one complete story and its content. Errors on an unknown story.
    pub fn delete_story(&self, story_id: &str) -> EngineResult<()> {
        self.engine
            .doc()
            .delete_story(story_id)
            .map_err(EngineError::from)
    }
}
