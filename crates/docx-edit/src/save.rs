use std::collections::{HashMap, HashSet};

use docx_parse::header_footer::HeaderFooter;
use docx_parse::notes::Note;
use docx_parse::relationships::{
    Relationship, TargetMode, office_document_path, relationship_types, resolve_relative_path,
};
use docx_parse::s9::{S9ParseOptions, parse_docx_s9_wire_from_parts};
use docx_parse::serializer::paragraph_ids::{S13TouchedPart, S13WholePart};
use docx_parse::serializer::{
    S13ParagraphId, S13ParagraphIds, S13PatchedPart, S13SaveReport, S13SaveRequest,
    S13SpliceAnchor, S13SplicedPart, write_docx_s13_parts_with_report,
};
use docx_parse::xml::{ParseBudget, ParseError, ParseLimits};
use serde::{Deserialize, Serialize, de::IgnoredAny};
use serde_json::Value;

use crate::seed::{SeedMedia, package_digest, seed_parsed_docx};
use crate::{EditingDoc, ParagraphSavePlan, SpliceAnchor};

const COMMENTS: &str = "word/comments.xml";
const FOOTNOTES: &str = "word/footnotes.xml";
const ENDNOTES: &str = "word/endnotes.xml";

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct NoteCoverage {
    footnotes: Option<IgnoredAny>,
    endnotes: Option<IgnoredAny>,
    footnote_separators: Option<IgnoredAny>,
    endnote_separators: Option<IgnoredAny>,
}

pub fn write_docx_json(request_json: &str, original_docx: &[u8]) -> Result<Vec<u8>, String> {
    write_json(request_json, original_docx).map(|(bytes, _)| bytes)
}

pub fn write_docx_with_warnings_json(
    request_json: &str,
    original_docx: &[u8],
) -> Result<Vec<u8>, String> {
    let (bytes, report) = write_json(request_json, original_docx)?;
    let warnings = serde_json::to_vec(&report.warnings).map_err(|error| error.to_string())?;
    let mut framed = (warnings.len() as u64).to_le_bytes().to_vec();
    framed.extend_from_slice(&warnings);
    framed.extend_from_slice(&bytes);
    Ok(framed)
}

fn write_json(
    request_json: &str,
    original_docx: &[u8],
) -> Result<(Vec<u8>, S13SaveReport), String> {
    let request: S13SaveRequest =
        serde_json::from_str(request_json).map_err(|error| error.to_string())?;
    let coverage = if request.paragraph_ids.is_none() && request.selective.is_none() {
        Some(serde_json::from_str(request_json).map_err(|error| error.to_string())?)
    } else {
        None
    };
    let parts = ooxml_opc::unzip_parts(original_docx).map_err(ParseError::Container);
    parts
        .and_then(|parts| {
            write_model(request, original_docx, &parts, &ParseLimits::default(), coverage)
        })
        .map_err(|error| error.to_string())
}

pub fn write_model_with_report(
    request: S13SaveRequest,
    original_docx: &[u8],
    parts: &[(String, Vec<u8>)],
    limits: &ParseLimits,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    write_model(request, original_docx, parts, limits, None)
}

fn write_model(
    mut request: S13SaveRequest,
    original_docx: &[u8],
    parts: &[(String, Vec<u8>)],
    limits: &ParseLimits,
    coverage: Option<NoteCoverage>,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    if request.paragraph_ids.is_none() && request.selective.is_none() {
        request.determinism.validate()?;
        request.paragraph_ids = Some(guarded_plan(
            &mut request,
            original_docx,
            parts,
            limits,
            coverage,
        )?);
    }
    write_docx_s13_parts_with_report(request, parts, Some(original_docx))
}

fn guarded_plan(
    request: &mut S13SaveRequest,
    bytes: &[u8],
    parts: &[(String, Vec<u8>)],
    limits: &ParseLimits,
    coverage: Option<NoteCoverage>,
) -> Result<S13ParagraphIds, ParseError> {
    let digest = package_digest(bytes);
    let envelope = parse_docx_s9_wire_from_parts(
        bytes,
        parts,
        S9ParseOptions {
            detect_variables: false,
            determinism_seed: Some(digest.clone()),
            source_ordinals: true,
            ..S9ParseOptions::default()
        },
        limits,
    )?;
    let source = &envelope.document.package;
    let body_path = office_document_path(parts, &mut ParseBudget::new(limits))?;
    let mut changed = HashSet::new();
    if !same_model(&request.document.content, &source.document.content)?
        || request.document.final_section_properties != source.document.final_section_properties
        || request.document.custom_root_bindings != source.document.custom_root_bindings
    {
        changed.insert(body_path.clone());
    }
    if request.document.comments.is_some()
        && !same_model(&request.document.comments, &source.document.comments)?
    {
        changed.insert(COMMENTS.to_owned());
    }

    let relationships: HashMap<_, _> = request
        .relationship_entries
        .iter()
        .map(|(id, relationship)| (id.as_str(), relationship))
        .collect();
    for (id, source) in &source.relationship_entries {
        if let Some(relationship) = relationships.get(id.as_str())
            && (relationship.target != source.target
                || relationship.target_mode != source.target_mode
                || relationship.relationship_type != source.relationship_type)
        {
            return Err(save_error(format!(
                "cannot replace document relationship {id}"
            )));
        }
    }
    guard_header_footers(
        &mut request.header_entries,
        source.header_entries.as_deref().unwrap_or_default(),
        &relationships,
        &body_path,
        relationship_types::HEADER,
        &mut changed,
    )?;
    guard_header_footers(
        &mut request.footer_entries,
        source.footer_entries.as_deref().unwrap_or_default(),
        &relationships,
        &body_path,
        relationship_types::FOOTER,
        &mut changed,
    )?;

    let covered = |present: fn(&NoteCoverage) -> bool| coverage.as_ref().is_none_or(present);
    let note_parts = [
        (
            FOOTNOTES,
            &mut request.footnotes,
            &mut request.footnote_separators,
            source.footnotes.as_deref().unwrap_or_default(),
            source.footnote_separators.as_deref().unwrap_or_default(),
            covered(|coverage| coverage.footnotes.is_some()),
            covered(|coverage| coverage.footnote_separators.is_some()),
        ),
        (
            ENDNOTES,
            &mut request.endnotes,
            &mut request.endnote_separators,
            source.endnotes.as_deref().unwrap_or_default(),
            source.endnote_separators.as_deref().unwrap_or_default(),
            covered(|coverage| coverage.endnotes.is_some()),
            covered(|coverage| coverage.endnote_separators.is_some()),
        ),
    ];
    for (path, notes, separators, source_notes, source_separators, notes_read, separators_read) in
        note_parts
    {
        let notes_changed = notes_read && !same_model(notes.as_slice(), source_notes)?;
        let separators_changed =
            separators_read && !same_model(separators.as_slice(), source_separators)?;
        if notes_changed || separators_changed {
            if !notes_read {
                notes.extend_from_slice(source_notes);
            }
            if !separators_read {
                separators.extend_from_slice(source_separators);
            }
            if notes.is_empty() && separators.is_empty() {
                return Err(save_error(format!("cannot remove every note from {path}")));
            }
            clear_stale_note_xml(notes, source_notes)?;
            clear_stale_note_xml(separators, source_separators)?;
            changed.insert(path.to_owned());
        }
    }

    if let Some(numbering) = &request.numbering
        && numbering != &source.numbering
        && (parts.iter().any(|(path, _)| path == "word/numbering.xml")
            || source.relationship_entries.iter().any(|(_, relationship)| {
                relationship.relationship_type == relationship_types::NUMBERING
            }))
    {
        return Err(save_error("cannot replace existing numbering definitions"));
    }

    let session = EditingDoc::new(1);
    seed_parsed_docx(
        &session,
        envelope,
        parts.to_vec(),
        bytes.into(),
        digest,
        SeedMedia::AsParsed,
    )
    .map_err(save_error)?;
    session.begin_opening(None);
    session.mark_seeded();
    let mut plan = session_paragraph_ids(session.paragraph_save_plan());
    let keep = |part: &str| {
        let path = if part == "word/document.xml" {
            body_path.as_str()
        } else {
            part
        };
        !changed.contains(path)
    };
    plan.assignments.retain(|entry| keep(&entry.part));
    plan.patched_parts.retain(|entry| keep(&entry.part));
    plan.spliced_parts.retain(|entry| keep(&entry.part));
    plan.whole_parts.retain(|entry| keep(&entry.part));
    plan.touched_parts.retain(|entry| keep(&entry.part));
    if !keep(&body_path) {
        plan.section_edits = None;
    }
    Ok(plan)
}

fn guard_header_footers(
    entries: &mut Vec<(String, HeaderFooter)>,
    source: &[(String, HeaderFooter)],
    relationships: &HashMap<&str, &Relationship>,
    body_path: &str,
    kind: &str,
    changed: &mut HashSet<String>,
) -> Result<(), ParseError> {
    let source: HashMap<_, _> = source
        .iter()
        .map(|(id, story)| (id.as_str(), story))
        .collect();
    let mut edits = HashMap::<String, usize>::new();
    let mut targets = Vec::with_capacity(entries.len());
    for (index, (id, story)) in entries.iter().enumerate() {
        let target = relationships
            .get(id.as_str())
            .filter(|relationship| {
                relationship.relationship_type == kind
                    && relationship.target_mode != Some(TargetMode::External)
                    && !relationship.target.is_empty()
            })
            .map(|relationship| resolve_relative_path(body_path, &relationship.target))
            .transpose()?;
        if let Some(path) = &target {
            let equal = source
                .get(id.as_str())
                .map(|original| same_model(story, *original))
                .transpose()?
                .unwrap_or(false);
            if !equal {
                edits.insert(path.clone(), index);
                changed.insert(path.clone());
            }
        }
        targets.push(target);
    }
    let mut index = 0;
    entries.retain(|_| {
        let current = index;
        index += 1;
        targets[current]
            .as_ref()
            .and_then(|path| edits.get(path))
            .is_none_or(|&owner| owner == current)
    });
    Ok(())
}

fn clear_stale_note_xml(notes: &mut [Note], source: &[Note]) -> Result<(), ParseError> {
    for note in notes {
        let Some(original) = source
            .iter()
            .find(|original| original.id == note.id)
            .or_else(|| {
                let xml = note.verbatim_xml.as_ref()?;
                source
                    .iter()
                    .find(|original| original.verbatim_xml.as_ref() == Some(xml))
            })
        else {
            continue;
        };
        let content_changed = note.id != original.id
            || note.story_type != original.story_type
            || note.note_type != original.note_type
            || note.custom_root_bindings != original.custom_root_bindings
            || !same_model(&note.content, &original.content)?;
        if content_changed && note.verbatim_xml.is_some() {
            if note.verbatim_xml != original.verbatim_xml {
                return Err(save_error(format!(
                    "conflicting XML and model edits to note {}",
                    note.id
                )));
            }
            note.verbatim_xml = None;
            note.source_ordinal = None;
        }
    }
    Ok(())
}

fn same_model<T: Serialize + PartialEq + ?Sized>(left: &T, right: &T) -> Result<bool, ParseError> {
    if left == right {
        return Ok(true);
    }
    let normalized = |value| {
        let mut value =
            serde_json::to_value(value).map_err(|error| save_error(error.to_string()))?;
        omit_source_ordinals(&mut value);
        Ok::<_, ParseError>(value)
    };
    Ok(normalized(left)? == normalized(right)?)
}

fn omit_source_ordinals(value: &mut Value) {
    match value {
        Value::Object(fields) => {
            fields.remove("sourceOrdinal");
            fields.values_mut().for_each(omit_source_ordinals);
        }
        Value::Array(items) => items.iter_mut().for_each(omit_source_ordinals),
        _ => {}
    }
}

fn save_error(message: impl Into<String>) -> ParseError {
    ParseError::Canonical(format!("S13 package save: {}", message.into()))
}

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

