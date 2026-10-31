use serde_json::json;

use super::{EditSessionCore, EngineError, EngineResult};
use crate::EditCtx;

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
            .map_err(EngineError::message)?;
        Ok(json!({ "paraId": para_id }).to_string())
    }

    /// Removes one complete story and its content. Errors on an unknown story.
    pub fn delete_story(&self, story_id: &str) -> EngineResult<()> {
        self.engine
            .doc()
            .delete_story(story_id)
            .map_err(EngineError::message)
    }
}
