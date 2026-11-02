use ooxml_opc::PackageBytes;
use serde_json::json;

use super::json::parse_render_env;
use super::{EditSessionCore, EngineError, EngineResult};
use crate::batch::outcome_json;
use crate::compare::CompareOutcome;
use crate::engine::RelayoutTrigger;
use crate::structured::ExportOptions;

impl EditSessionCore {
    /// Structured export of the committed document state:
    /// `{"revisionView","stories"?,"includeFormatting"?,"maxBlocks"?,"maxBytes"?}` ->
    /// `{"ok":true,"version","content"}` or `{"ok":false,"version","failure"}`. Anchors are scoped
    /// to the returned version. Reads only: nothing is committed, minted or published.
    pub fn export_structured_json(&self, options: &str) -> EngineResult<String> {
        let options: ExportOptions = serde_json::from_str(options).map_err(EngineError::message)?;
        outcome_json(&self.engine.doc().export_structured(&options)).map_err(EngineError::message)
    }

    /// [`EditSessionCore::export_structured_json`] rendered as Markdown from the same read:
    /// `{"ok":true,"version","content":{"markdown","anchors","diagnostics","truncated"}}`.
    pub fn export_markdown_json(&self, options: &str) -> EngineResult<String> {
        let options: ExportOptions = serde_json::from_str(options).map_err(EngineError::message)?;
        outcome_json(&self.engine.doc().export_markdown(&options)).map_err(EngineError::message)
    }

    /// [`EditSessionCore::export_structured_json`] with the page map of the retained region
    /// layout. `options` adds `"includeGeometry"?`, `"expectLayoutVersion"?`,
    /// `"maxFragments"?` and `"maxLayoutBytes"?`; the reply is
    /// `{"ok":true,"version","content":{"structured","layout"}}` or a refusal. An editor passes
    /// the region layout request it would lay the document out with now as `current_request`,
    /// and the layout must have been computed from the same inputs. Lays nothing out and changes
    /// nothing.
    pub fn export_structured_with_pages_json(
        &self,
        options: &str,
        current_request: Option<String>,
    ) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        let options: crate::structured::PageExportOptions =
            serde_json::from_str(options).map_err(EngineError::message)?;
        let read = match current_request {
            Some(current) => self
                .engine
                .export_structured_with_pages_for(&options, &current),
            None => self.engine.export_structured_with_pages(&options),
        };
        outcome_json(&read).map_err(EngineError::message)
    }

    /// Lays this private session out with its own fonts and exports it with pages as a
    /// snapshot. `fonts` holds the font files back to back, `font_lengths` their byte lengths;
    /// `request` is a region layout request whose font chains name fonts by their index. The
    /// reply is `{"ok":true,"content"}` or `{"ok":false,"failure"}`; a rejected font or an
    /// unusable request throws. The session's measurement fonts are left untouched.
    pub fn export_snapshot_with_private_fonts_json(
        &self,
        fonts: &[u8],
        font_lengths: &[u32],
        request: &str,
        options: &str,
    ) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        let options: crate::structured::PageExportOptions =
            serde_json::from_str(options).map_err(EngineError::message)?;
        let mut files = Vec::with_capacity(font_lengths.len());
        let mut rest = fonts;
        for length in font_lengths {
            if rest.len() < *length as usize {
                return Err(EngineError::message("font lengths exceed the font bytes"));
            }
            let (file, tail) = rest.split_at(*length as usize);
            files.push(file);
            rest = tail;
        }
        let outcome = self
            .engine
            .export_snapshot_with_private_fonts(&files, request, &options)
            .map_err(EngineError::message)?
            .map(|content| json!({ "content": content }))
            .map_err(|failure| json!({ "failure": failure }));
        outcome_json(&outcome).map_err(EngineError::message)
    }

    /// Compares two DOCX packages into this session, which must hold no document: the original
    /// is opened here and the revised body text differences are applied as tracked changes,
    /// outside undo history.
    ///
    /// `options` is `{"author","date","granularity"?,"unsupported"?,"limits"?}`. Returns a final
    /// `{"ok":false,"diagnostics"}` or `{"ok":true,"noop":true,"changes":[],"diagnostics"}`, or
    /// `{"ok":true,"noop":false,"save"}` naming what a save of the applied changes may write:
    /// then pass the saved bytes to [`EditSessionCore::finish_compared_docx_json`], or the save
    /// failure to [`EditSessionCore::fail_compared_docx_json`]. Malformed options and internal
    /// failures throw.
    pub fn compare_docx_json(
        &self,
        original: &[u8],
        revised: &[u8],
        options: &str,
    ) -> EngineResult<String> {
        self.awaiting_comment_baseline.set(false);
        let options = match crate::compare::parse_options(options).map_err(EngineError::message)? {
            Ok(options) => options,
            Err(diagnostic) => {
                return crate::compare::refused_json(diagnostic).map_err(EngineError::message);
            }
        };
        let outcome = crate::compare::compare_into(
            self.engine.doc(),
            &self.undo,
            original,
            revised,
            &options,
        )
        .map_err(EngineError::message)?;
        let json = outcome
            .to_json(&options.limits)
            .map_err(EngineError::message)?;
        if let CompareOutcome::Applied(applied) = outcome {
            self.engine.set_relayout_trigger(RelayoutTrigger::Open);
            self.docx_source.replace(Some(PackageBytes::from(original)));
            self.paragraph_styles.replace(None);
            self.docx_digest.replace(None);
            self.compared.replace(Some((applied, options.limits)));
        }
        Ok(json)
    }

    /// The final comparison result for `bytes`, the saved applied changes, verified against both
    /// inputs: `{"ok":true,"changes","diagnostics"}` or `{"ok":false,"diagnostics"}`. Throws
    /// when no comparison awaits its saved bytes.
    pub fn finish_compared_docx_json(&self, bytes: &[u8]) -> EngineResult<String> {
        let (applied, limits) = self
            .compared
            .take()
            .ok_or_else(|| EngineError::message("no comparison awaits its saved bytes"))?;
        applied.finish(bytes, &limits).map_err(EngineError::message)
    }

    /// The final comparison result when saving the applied changes failed with `message`.
    pub fn fail_compared_docx_json(&self, message: &str) -> EngineResult<String> {
        let (applied, limits) = self
            .compared
            .take()
            .ok_or_else(|| EngineError::message("no comparison awaits its saved bytes"))?;
        applied.fail(message, &limits).map_err(EngineError::message)
    }

    /// Lowers one story to a `LayoutBlock[]` JSON array: the block, run and
    /// table vocabulary the layout engine consumes. `env_json` supplies the
    /// document-level values lowering cannot read off the story:
    /// `{"themeColors":{slot: hex},"defaultTabStopTwips":number|null,
    /// "pageContentHeight":number|null,"numericIds":{yrsId: number},
    /// "tocStyleIds":[styleId],"showHiddenText":bool,
    /// "defaultParagraphStyleId":string,
    /// "revisionPreview":{revisionId: "accepted"|"rejected"}}`, all
    /// optional. Errors when the story does not end in a pilcrow, holds a
    /// malformed table, references itself through a cell story, or contains an
    /// embed lowering does not support.
    pub fn yrs_blocks_for_story(&self, story: &str, env_json: &str) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        let env = parse_render_env(env_json)?;
        self.engine
            .lower_story_json(story, &env)
            .map_err(EngineError::message)
    }
}
