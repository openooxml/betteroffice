use std::sync::Arc;

use serde_json::{Value, json};

use super::json::{
    origin_json, paragraph_anchor_json, paragraph_origin_json, parse_paragraph_anchor,
    persisted_anchor_json, persisted_receipt_json, source_anchor_json,
};
use super::{EditSessionCore, EngineError, EngineResult, ParagraphIdentitiesCache};
use crate::{AnchorResolution, AnchorUnsupported, ParagraphIdRefusal, ParagraphRef, SpliceAnchor};

impl EditSessionCore {
    pub(crate) fn paragraph_identities_uncached(&self) -> EngineResult<String> {
        let identities = self.engine.doc().paragraph_identities();
        let paragraphs: Vec<Value> = identities
            .paragraphs
            .iter()
            .map(|paragraph| {
                json!({
                    "session": match &paragraph.paragraph {
                        ParagraphRef::Session { .. } => {
                            paragraph_anchor_json(&identities.session_id, &paragraph.paragraph)
                        }
                        ParagraphRef::Source(_) => Value::Null,
                    },
                    "origin": paragraph_origin_json(paragraph.origin),
                    "ooxmlParaId": paragraph.ooxml_para_id,
                    "idOrigin": origin_json(paragraph.id_origin),
                    "persisted": persisted_anchor_json(
                        paragraph.source_story.as_ref(),
                        paragraph.ooxml_para_id.as_deref(),
                    ),
                    "source": paragraph.source.as_ref().map_or(Value::Null, source_anchor_json),
                })
            })
            .collect();
        serde_json::to_string(&json!({
            "sessionId": identities.session_id,
            "packageSha256": identities.package_sha256,
            "paragraphs": paragraphs,
        }))
        .map_err(EngineError::message)
    }

    /// Gives every paragraph that saves without a Word paragraph ID a fresh
    /// one and repairs duplicates; see [`crate::EditingDoc::persist_paragraph_ids`].
    /// Receipt: `{"status":"applied","assignments":[{"paragraph",
    /// "replacedParaId","previousOoxmlParaId","ooxmlParaId","idOrigin",
    /// "persisted"}],"diagnostics":[…]}`, or `{"status":"refused","refusal"}`
    /// when nothing was changed.
    pub fn persist_paragraph_ids(&self) -> EngineResult<String> {
        let doc = self.engine.doc();
        let session_id = doc.session_id();
        let receipt = match doc.persist_paragraph_ids() {
            Ok(persisted) => persisted_receipt_json(&session_id, &persisted),
            Err(ParagraphIdRefusal::AmbiguousCommentReference {
                ooxml_para_id,
                comment_ids,
            }) => json!({
                "status": "refused",
                "refusal": {
                    "kind": "ambiguous-comment-reference",
                    "ooxmlParaId": ooxml_para_id,
                    "commentIds": comment_ids,
                },
            }),
        };
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    /// The Word paragraph ID each paragraph of `story` saves with, as a JSON
    /// array of strings or `null` in document order. Errors on an unknown story.
    pub fn story_paragraph_ids(&self, story: &str) -> EngineResult<String> {
        let ids = self
            .engine
            .doc()
            .story_paragraph_ids(story)
            .map_err(EngineError::message)?;
        serde_json::to_string(&ids).map_err(EngineError::message)
    }

    /// Every paragraph's identities: `{"sessionId","packageSha256",
    /// "paragraphs":[{"session","origin","ooxmlParaId","idOrigin",
    /// "persisted","source"}]}`, session paragraphs with stories sorted and
    /// in document order, then source paragraphs outside the stories, whose
    /// `session` is `null`. Reads only.
    pub fn paragraph_identities(&self) -> EngineResult<String> {
        let doc = self.engine.doc();
        let version = doc.version();
        let source = doc.source_index();
        if let Some(cached) = self.paragraph_identities_cache.borrow().as_ref()
            && cached.version == version
            && match (&cached.source, &source) {
                (Some(cached), Some(source)) => Arc::ptr_eq(cached, source),
                (None, None) => true,
                _ => false,
            }
        {
            return Ok(cached.json.clone());
        }
        let json = self.paragraph_identities_uncached()?;
        self.paragraph_identities_cache
            .replace(Some(ParagraphIdentitiesCache {
                version,
                source,
                json: json.clone(),
            }));
        Ok(json)
    }

    /// Resolves a `session`, `source` or `persisted` anchor JSON against the
    /// current state: `{"status":"found","anchor"}`, `{"status":"missing"}`,
    /// `{"status":"ambiguous","candidates"}` or `{"status":"unsupported",
    /// "reason"}` with a `foreign-session`, `foreign-package` or
    /// `no-source-package` reason. A found source paragraph outside the
    /// session stories comes back as its source anchor. Errors on a malformed
    /// anchor.
    pub fn resolve_paragraph_anchor(&self, anchor_json: &str) -> EngineResult<String> {
        let anchor = parse_paragraph_anchor(anchor_json)?;
        let doc = self.engine.doc();
        let session_id = doc.session_id();
        let result = match doc.resolve_paragraph_anchor(&anchor) {
            AnchorResolution::Found(paragraph) => json!({
                "status": "found",
                "anchor": paragraph_anchor_json(&session_id, &paragraph),
            }),
            AnchorResolution::Missing => json!({ "status": "missing" }),
            AnchorResolution::Ambiguous(candidates) => json!({
                "status": "ambiguous",
                "candidates": candidates
                    .iter()
                    .map(|paragraph| paragraph_anchor_json(&session_id, paragraph))
                    .collect::<Vec<_>>(),
            }),
            AnchorResolution::Unsupported(reason) => json!({
                "status": "unsupported",
                "reason": match reason {
                    AnchorUnsupported::ForeignSession => "foreign-session",
                    AnchorUnsupported::ForeignPackage => "foreign-package",
                    AnchorUnsupported::NoSourcePackage => "no-source-package",
                },
            }),
        };
        serde_json::to_string(&result).map_err(EngineError::message)
    }

    /// The paragraph IDs a save applies, as the package writer's
    /// `paragraphIds` request field: `{"assignments":[{"part","ordinal",
    /// "paraId"}],"patchedParts":[{"part","paraIds":[[ordinal,"ID"]]}]}`,
    /// plus `"splicedParts":[{"part","sha256","paragraphs":[[ordinal,"key"]],
    /// "changed":[ordinal],"inserted":[{"before"|"after":ordinal}],"removed":[ordinal]}]`,
    /// whose session keys the caller resolves to the model paragraphs it marks
    /// with their ordinals.
    pub fn paragraph_save_plan(&self) -> EngineResult<String> {
        let plan = self.engine.doc().paragraph_save_plan();
        serde_json::to_string(&json!({
            "assignments": plan
                .assignments
                .iter()
                .map(|(part, ordinal, para_id)| {
                    json!({ "part": part, "ordinal": ordinal, "paraId": para_id })
                })
                .collect::<Vec<_>>(),
            "patchedParts": plan
                .patched_parts
                .iter()
                .map(|(part, para_ids)| json!({ "part": part, "paraIds": para_ids }))
                .collect::<Vec<_>>(),
            "splicedParts": plan
                .spliced_parts
                .iter()
                .map(|part| {
                    json!({
                        "part": part.part,
                        "sha256": part.sha256,
                        "paragraphs": part.paragraphs,
                        "changed": part.changed,
                        "sectionOnly": part.section_only,
                        "paragraphPatches": part.paragraph_patches,
                        "legacyParagraphs": part.legacy_paragraphs,
                        "tablePatches": part.table_patches,
                        "inserted": part
                            .inserted
                            .iter()
                            .map(|anchor| match anchor {
                                SpliceAnchor::Before(ordinal) => json!({ "before": ordinal }),
                                SpliceAnchor::After(ordinal) => json!({ "after": ordinal }),
                            })
                            .collect::<Vec<_>>(),
                        "removed": part.removed,
                    })
                })
                .collect::<Vec<_>>(),
            "sectionSplicedParts": plan.section_spliced_parts.iter().map(|part| json!({
                "part": part.part,
                "sha256": part.sha256,
                "paragraphs": part.paragraphs,
                "changed": part.changed,
                "sectionOnly": part.section_only,
                "tablePatches": part.table_patches,
                "inserted": part.inserted.iter().map(|anchor| match anchor {
                    SpliceAnchor::Before(ordinal) => json!({ "before": ordinal }),
                    SpliceAnchor::After(ordinal) => json!({ "after": ordinal }),
                }).collect::<Vec<_>>(),
                "removed": part.removed,
            })).collect::<Vec<_>>(),
        }))
        .map_err(EngineError::message)
    }

    /// Reconciles and publishes the `[owner, paraId]` pairs a save captured,
    /// an owner being a session key or a source occurrence's
    /// `{partUri}#{ordinal}`; returns the stale pairs as the same JSON shape.
    /// See [`crate::EditingDoc::record_saved_paragraph_ids`].
    pub fn record_saved_paragraph_ids(&self, saved_json: &str) -> EngineResult<String> {
        let saved: Vec<(String, String)> =
            serde_json::from_str(saved_json).map_err(EngineError::message)?;
        let stale = self.engine.doc().record_saved_paragraph_ids(&saved);
        serde_json::to_string(&stale).map_err(EngineError::message)
    }

    /// The Word paragraph IDs a DOCX package holds, as JSON mapping each XML
    /// part URI to its paragraphs' IDs in document order, in canonical form.
    pub fn written_paragraph_ids(&self, bytes: &[u8]) -> EngineResult<String> {
        let ids = docx_parse::paragraph_identity::paragraph_ids_by_part(bytes)
            .map_err(EngineError::message)?;
        serde_json::to_string(&ids).map_err(EngineError::message)
    }

    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub(crate) fn run_inheritance_styles(&self, styles_json: &str) -> EngineResult<Value> {
        serde_json::from_str(styles_json).map_err(EngineError::message)
    }
}
