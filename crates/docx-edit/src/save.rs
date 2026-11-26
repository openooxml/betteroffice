use docx_parse::serializer::paragraph_ids::{S13TouchedPart, S13WholePart};
use docx_parse::serializer::{
    S13ParagraphId, S13ParagraphIds, S13PatchedPart, S13SpliceAnchor, S13SplicedPart,
};

use crate::{ParagraphSavePlan, SpliceAnchor};

#[doc(hidden)]
pub use docx_parse::serializer::model_save::write_model_with_report;
#[doc(hidden)]
pub use docx_parse::{
    write_docx_s13_json as write_docx_json,
    write_docx_s13_with_warnings_json as write_docx_with_warnings_json,
};

/// `plan` as the package writer's request field, every planned part kept.
#[doc(hidden)]
pub fn session_paragraph_ids(plan: ParagraphSavePlan) -> S13ParagraphIds {
    S13ParagraphIds {
        removed_comment_ids: Vec::new(),
        section_edits: None,
        assignments: plan
            .assignments
            .into_iter()
            .map(|(part, ordinal, para_id)| S13ParagraphId {
                part,
                ordinal,
                para_id,
            })
            .collect(),
        patched_parts: plan
            .patched_parts
            .into_iter()
            .map(|(part, para_ids)| S13PatchedPart { part, para_ids })
            .collect(),
        spliced_parts: plan
            .spliced_parts
            .into_iter()
            .map(|part| S13SplicedPart {
                part: part.part,
                sha256: part.sha256,
                paragraphs: part
                    .paragraphs
                    .into_iter()
                    .map(|(ordinal, _)| ordinal)
                    .collect(),
                changed: part.changed,
                inserted: part
                    .inserted
                    .into_iter()
                    .map(|anchor| match anchor {
                        SpliceAnchor::Before(ordinal) => S13SpliceAnchor::Before(ordinal),
                        SpliceAnchor::After(ordinal) => S13SpliceAnchor::After(ordinal),
                    })
                    .collect(),
                removed: part.removed,
                paragraph_patches: part.paragraph_patches,
                legacy_paragraphs: part.legacy_paragraphs,
                table_patches: part.table_patches,
                format_only: part.format_only,
            })
            .collect(),
        whole_parts: plan
            .whole_parts
            .into_iter()
            .map(|part| S13WholePart {
                part: part.part,
                sha256: part.sha256,
                changed: part.changed,
            })
            .collect(),
        touched_parts: plan
            .touched_parts
            .into_iter()
            .map(|(part, touched)| S13TouchedPart { part, touched })
            .collect(),
    }
}
