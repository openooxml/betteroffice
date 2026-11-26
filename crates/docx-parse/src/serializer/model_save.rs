use std::collections::{BTreeSet, HashMap};

use serde::Deserialize;
use serde_json::{Map, Value};
use sha2::{Digest, Sha256};

use crate::document::DocumentBody;
use crate::header_footer::HeaderFooter;
use crate::notes::Note;
use crate::numbering::NumberingDefinitions;
use crate::relationships::{
    Relationship, RelationshipMap, RelationshipTarget, parse_relationships, relationship_part_path,
    relationship_types, resolve_relationship_target,
};
use crate::s9::{S9PackageWire, S9ParseOptions, parse_save_source};
use crate::xml::{ParseBudget, ParseError, ParseLimits};

use super::paragraph_ids::{S13ParagraphIds, S13PatchedPart};
use super::s13::{
    S13SaveOptions, S13SaveRequest, S13SelectiveSave, WriteSettings, save_core_only, write_parts,
};
use super::{S13SaveReport, SerializerDeterminism};

const COMMENTS: &str = "word/comments.xml";
const FOOTNOTES: &str = "word/footnotes.xml";
const ENDNOTES: &str = "word/endnotes.xml";
const NUMBERING: &str = "word/numbering.xml";

pub const WHOLE_MODEL_WARNING: &str = "DOCX_WHOLE_MODEL_WRITE ";
pub(super) const NO_BINDING_CHANGES: &str = "DOCX_SAVE_BINDINGS 0";

#[derive(Clone, Copy)]
struct Coverage {
    source_ordinals: bool,
    notes: [bool; 4],
    rewrite_body: bool,
    unchanged_body: bool,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct SaveWire {
    determinism: SerializerDeterminism,
    document: Option<DocumentBody>,
    #[serde(default)]
    header_entries: Vec<(String, HeaderFooter)>,
    #[serde(default)]
    footer_entries: Vec<(String, HeaderFooter)>,
    #[serde(default, deserialize_with = "present_notes")]
    footnotes: Option<Vec<Note>>,
    #[serde(default, deserialize_with = "present_notes")]
    endnotes: Option<Vec<Note>>,
    #[serde(default, deserialize_with = "present_notes")]
    footnote_separators: Option<Vec<Note>>,
    #[serde(default, deserialize_with = "present_notes")]
    endnote_separators: Option<Vec<Note>>,
    #[serde(default)]
    relationship_entries: Vec<(String, Relationship)>,
    numbering: Option<NumberingDefinitions>,
    #[serde(default)]
    options: S13SaveOptions,
    selective: Option<S13SelectiveSave>,
    paragraph_ids: Option<S13ParagraphIds>,
    style_languages: Option<serde_json::Value>,
    #[serde(default)]
    fresh_package: bool,
    #[serde(default)]
    report_bindings: bool,
    #[serde(default)]
    rewrite_body: bool,
    unchanged_source_sha256: Option<String>,
    unchanged_body_sha256: Option<String>,
}

fn present_notes<'de, D: serde::Deserializer<'de>>(
    deserializer: D,
) -> Result<Option<Vec<Note>>, D::Error> {
    Vec::deserialize(deserializer).map(Some)
}

pub(crate) fn write_json(
    json: &str,
    original_docx: &[u8],
    wire_diagnostics: Option<&mut Vec<String>>,
) -> Result<(Vec<u8>, S13SaveReport), String> {
    let wire: SaveWire = serde_json::from_str(json).map_err(|error| error.to_string())?;
    let wire_diagnostics = wire_diagnostics.filter(|_| wire.report_bindings);
    if let Some(digest) = wire.unchanged_source_sha256 {
        if wire.document.is_some()
            || wire.fresh_package
            || wire.style_languages.is_some()
            || wire.paragraph_ids.is_some()
            || wire.selective.is_some()
            || wire.rewrite_body
        {
            return Err("unchangedSourceSha256 requires a metadata-only save request".to_owned());
        }
        if digest != format!("{:x}", Sha256::digest(original_docx)) {
            return Err("DOCX_SOURCE_CHANGED".to_owned());
        }
        wire.determinism
            .validate()
            .map_err(|error| error.to_string())?;
        let bytes =
            super::s13::save_unchanged_archive(&wire.options, &wire.determinism.now, original_docx)
                .map_err(|error| error.to_string())?;
        if let Some(wire) = wire_diagnostics {
            wire.push(NO_BINDING_CHANGES.to_owned());
        }
        return Ok((bytes, S13SaveReport::default()));
    }
    let source_digest = wire.unchanged_body_sha256.map(|expected| {
        let digest = format!("{:x}", Sha256::digest(original_docx));
        (digest == expected, digest)
    });
    let coverage = Coverage {
        unchanged_body: source_digest.as_ref().is_some_and(|(matched, _)| *matched),
        source_ordinals: json.contains("\"sourceOrdinal\""),
        rewrite_body: wire.rewrite_body,
        notes: [
            wire.footnotes.is_some(),
            wire.endnotes.is_some(),
            wire.footnote_separators.is_some(),
            wire.endnote_separators.is_some(),
        ],
    };
    let request = S13SaveRequest {
        determinism: wire.determinism,
        document: wire.document.ok_or("missing field `document`")?,
        header_entries: wire.header_entries,
        footer_entries: wire.footer_entries,
        footnotes: wire.footnotes.unwrap_or_default(),
        endnotes: wire.endnotes.unwrap_or_default(),
        footnote_separators: wire.footnote_separators.unwrap_or_default(),
        endnote_separators: wire.endnote_separators.unwrap_or_default(),
        relationship_entries: wire.relationship_entries,
        numbering: wire.numbering,
        options: wire.options,
        selective: wire.selective,
        paragraph_ids: wire.paragraph_ids,
    };
    let parts = ooxml_opc::unzip_parts(original_docx)
        .map_err(|error| ParseError::Container(error).to_string())?;
    if wire.style_languages.is_some()
        && (request.paragraph_ids.is_none() || request.selective.is_some())
    {
        return Err("styleLanguages requires a session save plan".to_owned());
    }
    if wire.fresh_package || request.paragraph_ids.is_some() || request.selective.is_some() {
        return write_parts(
            request,
            &parts,
            Some(original_docx),
            WriteSettings {
                style_languages: wire.style_languages.as_ref(),
                product_save: wire.fresh_package,
                ..WriteSettings::default()
            },
        )
        .map_err(|error| error.to_string());
    }
    write_model(
        request,
        original_docx,
        &parts,
        &ParseLimits::default(),
        coverage,
        wire_diagnostics,
        None,
        source_digest.map(|(_, digest)| digest),
    )
    .map_err(|error| error.to_string())
}

pub fn write_model_with_report(
    request: S13SaveRequest,
    original_docx: &[u8],
    parts: &[(String, Vec<u8>)],
    limits: &ParseLimits,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    write_model_with_body_policy(request, original_docx, parts, limits, false)
}

#[doc(hidden)]
pub fn write_model_with_body_policy(
    request: S13SaveRequest,
    original_docx: &[u8],
    parts: &[(String, Vec<u8>)],
    limits: &ParseLimits,
    rewrite_body: bool,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    write_native_model(request, original_docx, parts, limits, rewrite_body, None)
}

#[doc(hidden)]
pub fn write_native_model(
    request: S13SaveRequest,
    original_docx: &[u8],
    parts: &[(String, Vec<u8>)],
    limits: &ParseLimits,
    rewrite_body: bool,
    whole_model_reason: Option<&str>,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    write_model(
        request,
        original_docx,
        parts,
        limits,
        Coverage {
            source_ordinals: true,
            notes: [true; 4],
            rewrite_body,
            unchanged_body: false,
        },
        None,
        whole_model_reason,
        None,
    )
}

#[allow(clippy::too_many_arguments)]
fn write_model(
    mut request: S13SaveRequest,
    bytes: &[u8],
    parts: &[(String, Vec<u8>)],
    limits: &ParseLimits,
    coverage: Coverage,
    wire: Option<&mut Vec<String>>,
    whole_model_reason: Option<&str>,
    source_digest: Option<String>,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    if request.paragraph_ids.is_some() || request.selective.is_some() {
        return write_parts(request, parts, Some(bytes), WriteSettings::default());
    }
    request.determinism.validate()?;
    let parse_options = S9ParseOptions {
        detect_variables: false,
        parse_headers_footers: !request.header_entries.is_empty()
            || !request.footer_entries.is_empty(),
        parse_notes: coverage.notes.iter().any(|covered| *covered),
        source_ordinals: coverage.source_ordinals,
        determinism_seed: source_digest,
        ..S9ParseOptions::default()
    };
    let (body_path, source) = match parse_save_source(bytes, parts, parse_options, limits) {
        Ok(source) => source,
        Err(error) => {
            if let Some(requested) = &request.numbering
                && let Some((path, xml)) = crate::s8::find_part(parts, NUMBERING)
                && let Ok(original) = crate::numbering::parse_numbering(
                    Some(xml),
                    path,
                    &mut ParseBudget::new(limits),
                )
                && requested != &original.definitions
            {
                return Err(save_error("cannot replace existing numbering definitions"));
            }
            return write_fallback(request, bytes, parts, coverage, &error.to_string());
        }
    };
    let part_paths = PartPaths::new(parts);
    let mut numbering_added = false;
    if let Some(numbering) = &request.numbering
        && numbering != &source.numbering
    {
        if part_paths.contains(NUMBERING)
            || source.relationship_entries.iter().any(|(_, relationship)| {
                relationship.relationship_type == relationship_types::NUMBERING
            })
        {
            return Err(save_error("cannot replace existing numbering definitions"));
        }
        numbering_added = !numbering.abstract_nums.is_empty() || !numbering.nums.is_empty();
    }
    if let Some(reason) = whole_model_reason {
        return write_fallback(request, bytes, parts, coverage, reason);
    }
    if let Err(error) = validate_relationships(
        &body_path,
        &request.relationship_entries,
        &source.relationship_entries,
    ) {
        return write_fallback(request, bytes, parts, coverage, &error.to_string());
    }
    let aliases = match header_footer_paths(
        &request,
        &source,
        &body_path,
        &part_paths,
        parts,
        limits,
    ) {
        Ok(aliases) => aliases,
        Err(error @ ParseError::Canonical(_)) => return Err(error),
        Err(error) => return write_fallback(request, bytes, parts, coverage, &error.to_string()),
    };
    if aliases.edits.values().any(|edit| {
        edit.path.eq_ignore_ascii_case(&body_path)
            || [COMMENTS, FOOTNOTES, ENDNOTES, NUMBERING]
                .iter()
                .any(|reserved| edit.path.eq_ignore_ascii_case(reserved))
    }) {
        return write_fallback(
            request,
            bytes,
            parts,
            coverage,
            "a header or footer targets another story part",
        );
    }
    let mut changed = BTreeSet::new();
    let reuse_document_root =
        request.document.custom_root_bindings == source.document.custom_root_bindings;
    if coverage.rewrite_body
        || !reuse_document_root
        || (!coverage.unchanged_body
            && (!same_model(&request.document.content, &source.document.content)
                || request.document.final_section_properties
                    != source.document.final_section_properties))
    {
        changed.insert(body_path.clone());
    }
    if let Some(comments) = &request.document.comments
        && !same_model(
            comments.as_slice(),
            source.document.comments.as_deref().unwrap_or_default(),
        )
    {
        changed.insert(part_paths.resolve(COMMENTS).to_owned());
    }
    changed.extend(aliases.edits.values().map(|edit| edit.path.clone()));
    let [footnotes, endnotes, footnote_separators, endnote_separators] = coverage.notes;
    let note_parts = [
        NotePart {
            path: part_paths.resolve(FOOTNOTES),
            notes: &mut request.footnotes,
            separators: &mut request.footnote_separators,
            source_notes: source.footnotes.as_deref().unwrap_or_default(),
            source_separators: source.footnote_separators.as_deref().unwrap_or_default(),
            notes_read: footnotes,
            separators_read: footnote_separators,
        },
        NotePart {
            path: part_paths.resolve(ENDNOTES),
            notes: &mut request.endnotes,
            separators: &mut request.endnote_separators,
            source_notes: source.endnotes.as_deref().unwrap_or_default(),
            source_separators: source.endnote_separators.as_deref().unwrap_or_default(),
            notes_read: endnotes,
            separators_read: endnote_separators,
        },
    ];
    for notes in &note_parts {
        if notes.changed() {
            changed.insert(notes.path.to_owned());
        }
    }

    if numbering_added {
        changed.insert(NUMBERING.to_owned());
    }
    let added_story_relationship = request
        .relationship_entries
        .iter()
        .any(|(id, relationship)| {
            matches!(
                relationship.relationship_type.as_str(),
                relationship_types::HEADER | relationship_types::FOOTER
            ) && source
                .relationship_entries
                .iter()
                .all(|(source_id, _)| source_id != id)
                && matches!(
                    resolve_relationship_target(&body_path, relationship),
                    Ok(RelationshipTarget::Internal(ref path)) if part_paths.contains(path)
                )
        });
    if changed.is_empty() && !added_story_relationship {
        if let Some(wire) = wire {
            wire.push(NO_BINDING_CHANGES.to_owned());
        }
        return save_core_only(&request, &body_path, parts, bytes);
    }

    let mut empty_notes = [false; 2];
    for (index, notes) in note_parts.into_iter().enumerate() {
        if changed.contains(notes.path) {
            notes.prepare(limits)?;
            empty_notes[index] = true;
        }
    }
    let mut index = 0;
    for entries in [&mut request.header_entries, &mut request.footer_entries] {
        entries.retain_mut(|(_, story)| {
            let edit = aliases.targets[index]
                .as_ref()
                .and_then(|key| aliases.edits.get(key));
            let keep = edit.is_none_or(|edit| edit.owner == index);
            if let Some(edit) = edit.filter(|_| keep) {
                *story = edit.story.clone();
            }
            index += 1;
            keep
        });
    }

    request.paragraph_ids = Some(source_part_plan(&source, &body_path, &part_paths, &changed));
    let (saved, mut report) = write_parts(
        request,
        parts,
        Some(bytes),
        WriteSettings {
            reuse_document_root,
            empty_notes,
            wire,
            legacy_diagnostics: false,
            product_save: true,
            ..WriteSettings::default()
        },
    )?;
    report.warnings.extend(changed.into_iter().map(|part| format!(
        "DOCX_PART_REWRITTEN {part} was rewritten from the document model. Content the model does not carry may be lost. Saving edits through a session (createYrsSession, saveYrsDocx) patches changed paragraphs in place where it can."
    )));
    Ok((saved, report))
}

/// Fresh sessions retain authored IDs and have no paragraph ID claims to apply.
fn source_part_plan(
    source: &S9PackageWire,
    body_path: &str,
    paths: &PartPaths<'_>,
    changed: &BTreeSet<String>,
) -> S13ParagraphIds {
    let mut preserved = BTreeSet::from([body_path]);
    preserved.extend(source_story_paths(source, body_path, paths));
    preserved.extend([COMMENTS, FOOTNOTES, ENDNOTES].map(|path| paths.resolve(path)));
    S13ParagraphIds {
        patched_parts: preserved
            .into_iter()
            .filter(|path| !changed.contains(*path) && paths.contains(path))
            .map(|part| S13PatchedPart {
                part: part.to_owned(),
                para_ids: Vec::new(),
            })
            .collect(),
        ..S13ParagraphIds::default()
    }
}

fn write_fallback(
    request: S13SaveRequest,
    bytes: &[u8],
    parts: &[(String, Vec<u8>)],
    coverage: Coverage,
    reason: &str,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    let empty_notes = [
        coverage.notes[0] && coverage.notes[2],
        coverage.notes[1] && coverage.notes[3],
    ];
    let empty_notes = std::array::from_fn(|index| {
        empty_notes[index]
            && parts
                .iter()
                .any(|(path, _)| path.eq_ignore_ascii_case([FOOTNOTES, ENDNOTES][index]))
    });
    let (saved, mut report) = write_parts(
        request,
        parts,
        Some(bytes),
        WriteSettings {
            empty_notes,
            product_save: true,
            ..WriteSettings::default()
        },
    )?;
    report.warnings.push(format!(
        "{WHOLE_MODEL_WARNING}could not verify which stories were unchanged ({reason}). Stories written from the model may lose unmodelled content."
    ));
    Ok((saved, report))
}

struct PartPaths<'a>(HashMap<String, &'a str>);

impl<'a> PartPaths<'a> {
    fn new(parts: &'a [(String, Vec<u8>)]) -> Self {
        Self(
            parts
                .iter()
                .map(|(path, _)| (path.to_ascii_lowercase(), path.as_str()))
                .collect(),
        )
    }

    fn resolve<'b>(&'b self, path: &'b str) -> &'b str {
        self.0
            .get(&path.to_ascii_lowercase())
            .copied()
            .unwrap_or(path)
    }

    fn contains(&self, path: &str) -> bool {
        self.0.contains_key(&path.to_ascii_lowercase())
    }
}

fn validate_relationships(
    body_path: &str,
    requested: &[(String, Relationship)],
    source: &[(String, Relationship)],
) -> Result<(), ParseError> {
    let relationships: HashMap<_, _> = requested.iter().map(|(id, rel)| (id, rel)).collect();
    for (id, source) in source {
        if let Some(relationship) = relationships.get(id)
            && (relationship.relationship_type != source.relationship_type
                || !same_relationship_target(body_path, relationship, source))
        {
            return Err(save_error(format!("cannot replace document relationship {id}")));
        }
    }
    Ok(())
}

fn same_relationship_target(body_path: &str, left: &Relationship, right: &Relationship) -> bool {
    if left.target == right.target && left.target_mode == right.target_mode {
        return true;
    }
    match (
        resolve_relationship_target(body_path, left),
        resolve_relationship_target(body_path, right),
    ) {
        (Ok(RelationshipTarget::Internal(left)), Ok(RelationshipTarget::Internal(right))) => {
            left.eq_ignore_ascii_case(&right)
        }
        (Ok(RelationshipTarget::External(left)), Ok(RelationshipTarget::External(right))) => {
            left == right
        }
        _ => false,
    }
}

fn story_path(
    body_path: &str,
    relationship: &Relationship,
    kind: &str,
) -> Result<Option<String>, ParseError> {
    if relationship.relationship_type != kind || relationship.target.is_empty() {
        return Ok(None);
    }
    match resolve_relationship_target(body_path, relationship)? {
        RelationshipTarget::Internal(path) => Ok(Some(path)),
        RelationshipTarget::External(_) => Ok(None),
    }
}

fn source_story_paths<'a>(
    source: &'a S9PackageWire,
    body_path: &'a str,
    paths: &'a PartPaths<'_>,
) -> impl Iterator<Item = &'a str> {
    source
        .relationship_entries
        .iter()
        .filter_map(move |(_, relationship)| {
            let kind = relationship.relationship_type.as_str();
            if !matches!(kind, relationship_types::HEADER | relationship_types::FOOTER) {
                return None;
            }
            let target = story_path(body_path, relationship, kind).ok()??;
            paths.0.get(&target.to_ascii_lowercase()).copied()
        })
}

struct AliasEdit {
    owner: usize,
    relationship_id: String,
    kind: &'static str,
    path: String,
    story: HeaderFooter,
}

struct HeaderFooterPlan {
    edits: HashMap<String, AliasEdit>,
    targets: Vec<Option<String>>,
}

fn header_footer_paths(
    request: &S13SaveRequest,
    source: &S9PackageWire,
    body_path: &str,
    paths: &PartPaths<'_>,
    parts: &[(String, Vec<u8>)],
    limits: &ParseLimits,
) -> Result<HeaderFooterPlan, ParseError> {
    let relationships: HashMap<_, _> = request
        .relationship_entries
        .iter()
        .map(|(id, rel)| (id.as_str(), rel))
        .collect();
    let originals: HashMap<_, _> = source
        .header_entries
        .iter()
        .chain(&source.footer_entries)
        .flatten()
        .map(|(id, story)| (id.as_str(), story))
        .collect();
    let mut original_targets = HashMap::new();
    for (relationship_id, relationship) in &source.relationship_entries {
        if let Some(original) = originals.get(relationship_id.as_str())
            && let Some(path) =
                story_path(body_path, relationship, &relationship.relationship_type)?
        {
            original_targets
                .entry((
                    path.to_ascii_lowercase(),
                    relationship.relationship_type.as_str(),
                ))
                .or_insert(*original);
        }
    }
    let mut edits = HashMap::<String, AliasEdit>::new();
    let mut targets = Vec::new();
    let mut story_relationships = HashMap::<String, RelationshipMap>::new();
    let mut budget = ParseBudget::new(limits);
    for (entries, kind) in [
        (&request.header_entries, relationship_types::HEADER),
        (&request.footer_entries, relationship_types::FOOTER),
    ] {
        for (id, story) in entries {
            let target = match relationships.get(id.as_str()) {
                Some(relationship) => story_path(body_path, relationship, kind)?,
                None => None,
            };
            let path = target.as_deref().map(|target| paths.resolve(target));
            let key = path.map(str::to_ascii_lowercase);
            if let (Some(path), Some(key)) = (path, &key) {
                let original = originals
                    .get(id.as_str())
                    .copied()
                    .or_else(|| original_targets.get(&(key.clone(), kind)).copied());
                let owner_relationships = if story.watermark.is_some()
                    || original.is_some_and(|original| original.watermark.is_some())
                {
                    if !story_relationships.contains_key(key) {
                        let rels_path = relationship_part_path(path);
                        let relationships = match crate::s8::find_part(parts, &rels_path) {
                            Some((name, bytes)) => parse_relationships(bytes, name, &mut budget)?,
                            None => RelationshipMap::new(),
                        };
                        story_relationships.insert(key.clone(), relationships);
                    }
                    story_relationships.get(key)
                } else {
                    None
                };
                if original.is_none_or(|original| {
                    !same_model_with_relationships(story, original, owner_relationships)
                }) {
                    if let Some(edit) = edits.get_mut(key) {
                        let owner_id = &edit.relationship_id;
                        let owner_base = originals
                            .get(owner_id.as_str())
                            .copied()
                            .or_else(|| original_targets.get(&(key.clone(), edit.kind)).copied());
                        if let Some(merged) = merge_alias(
                            &edit.story,
                            story,
                            owner_base,
                            original,
                            owner_relationships,
                        ) {
                            edit.story = merged;
                        } else {
                            return Err(save_error(format!(
                                "conflicting header/footer edits to {path} through {owner_id} and {id}"
                            )));
                        }
                    } else {
                        edits.insert(
                            key.clone(),
                            AliasEdit {
                                owner: targets.len(),
                                relationship_id: id.clone(),
                                kind,
                                path: path.to_owned(),
                                story: story.clone(),
                            },
                        );
                    }
                }
            }
            targets.push(key);
        }
    }
    Ok(HeaderFooterPlan { edits, targets })
}

fn merge_alias(
    left: &HeaderFooter,
    right: &HeaderFooter,
    source_left: Option<&HeaderFooter>,
    source_right: Option<&HeaderFooter>,
    relationships: Option<&RelationshipMap>,
) -> Option<HeaderFooter> {
    fn tree(story: &HeaderFooter) -> Option<Value> {
        let mut value = serde_json::to_value(story).ok()?;
        if let Some(object) = value.as_object_mut() {
            object.remove("type");
            object.remove("hdrFtrType");
            if let Some(Value::Array(bindings)) = object.remove("customRootBindings") {
                object.insert(
                    "customRootBindings".into(),
                    Value::Object(
                        bindings
                            .into_iter()
                            .filter_map(|binding| {
                                Some((
                                    binding["name"].as_str()?.to_owned(),
                                    binding["value"].clone(),
                                ))
                            })
                            .collect(),
                    ),
                );
            }
        }
        Some(value)
    }

    // Each alias has its own parsed baseline, including generated paragraph metadata.
    fn merge(
        left: &Value,
        right: &Value,
        source_left: &Value,
        source_right: &Value,
        relationships: Option<&RelationshipMap>,
    ) -> Option<Value> {
        let same = |a, b| same_value(a, b, relationships);
        if same(right, source_right) || same(left, right) {
            return Some(left.clone());
        }
        if same(left, source_left)
            && (same(source_left, source_right) || !right.is_object() && !right.is_array())
        {
            return Some(right.clone());
        }
        if let (Some(left), Some(right)) = (left.as_object(), right.as_object()) {
            let keys: BTreeSet<_> = left
                .keys()
                .chain(right.keys())
                .chain(source_left.as_object().into_iter().flatten().map(|(key, _)| key))
                .chain(source_right.as_object().into_iter().flatten().map(|(key, _)| key))
                .collect();
            let mut fields = serde_json::Map::new();
            for key in keys {
                let value = if ignores_model_field(key, left, right, relationships) {
                    left.get(key).cloned().unwrap_or_default()
                } else {
                    merge(
                        left.get(key).unwrap_or(&Value::Null),
                        right.get(key).unwrap_or(&Value::Null),
                        &source_left[key.as_str()],
                        &source_right[key.as_str()],
                        relationships,
                    )?
                };
                if !value.is_null() {
                    fields.insert(key.clone(), value);
                }
            }
            return Some(Value::Object(fields));
        }
        if let (Some(left), Some(right)) = (left.as_array(), right.as_array())
            && left.len() == right.len()
        {
            return left
                .iter()
                .zip(right)
                .enumerate()
                .map(|(index, (first, second))| {
                    merge(
                        first,
                        second,
                        &source_left[index],
                        &source_right[index],
                        relationships,
                    )
                })
                .collect::<Option<Vec<_>>>()
                .map(Value::Array);
        }
        if same(left, source_left) {
            return Some(right.clone());
        }
        // A previously absent object can be edited independently by both aliases.
        if source_left.is_null() && source_right.is_null() && (left.is_null() || right.is_null()) {
            return Some(if left.is_null() { right } else { left }.clone());
        }
        None
    }

    let left_tree = tree(left)?;
    let right_tree = tree(right)?;
    let source_left = source_left.map(tree).unwrap_or(Some(Value::Null))?;
    let source_right = source_right.map(tree).unwrap_or(Some(Value::Null))?;
    let mut result = merge(&left_tree, &right_tree, &source_left, &source_right, relationships)?;
    let object = result.as_object_mut()?;
    object.insert("type".into(), Value::String(left.story_type.clone()));
    object.insert("hdrFtrType".into(), Value::String(left.hdr_ftr_type.clone()));
    if let Some(Value::Object(bindings)) = object.remove("customRootBindings") {
        object.insert(
            "customRootBindings".into(),
            Value::Array(
                bindings
                    .into_iter()
                    .map(|(name, value)| serde_json::json!({ "name": name, "value": value }))
                    .collect(),
            ),
        );
    }
    serde_json::from_value(result).ok()
}

struct NotePart<'a> {
    path: &'a str,
    notes: &'a mut Vec<Note>,
    separators: &'a mut Vec<Note>,
    source_notes: &'a [Note],
    source_separators: &'a [Note],
    notes_read: bool,
    separators_read: bool,
}

impl NotePart<'_> {
    fn changed(&self) -> bool {
        (self.notes_read && !same_notes(self.notes, self.source_notes))
            || (self.separators_read && !same_notes(self.separators, self.source_separators))
    }

    fn prepare(self, limits: &ParseLimits) -> Result<(), ParseError> {
        if !self.notes_read {
            self.notes.extend_from_slice(self.source_notes);
        }
        if !self.separators_read {
            self.separators.extend_from_slice(self.source_separators);
        }
        let note_key = |note: &Note| if note.id == 0.0 { 0 } else { note.id.to_bits() };
        let mut originals = HashMap::<u64, Vec<&Note>>::new();
        for note in self.source_notes.iter().chain(self.source_separators) {
            if note.verbatim_xml.is_some() {
                originals.entry(note_key(note)).or_default().push(note);
            }
        }
        let source_xml: BTreeSet<_> = originals
            .values()
            .flatten()
            .filter_map(|note| note.verbatim_xml.as_deref())
            .collect();
        let mut budget = ParseBudget::new(limits);
        for (notes, supplied) in [
            (self.notes, self.notes_read),
            (self.separators, self.separators_read),
        ] {
            if supplied {
                for note in notes {
                    let originals = originals
                        .get(&note_key(note))
                        .map(Vec::as_slice)
                        .unwrap_or_default();
                    clear_stale_note_xml(note, originals, &source_xml, self.path, &mut budget)?;
                }
            }
        }
        Ok(())
    }
}

fn clear_stale_note_xml(
    note: &mut Note,
    originals: &[&Note],
    source_xml: &BTreeSet<&str>,
    path: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<(), ParseError> {
    let Some(xml) = note.verbatim_xml.as_deref() else {
        note.source_ordinal = None;
        return Ok(());
    };
    let matching = originals
        .iter()
        .copied()
        .find(|original| same_note_content(note, original));
    if matching.is_some_and(|original| original.verbatim_xml.as_deref() == Some(xml)) {
        note.source_ordinal = matching.and_then(|original| original.source_ordinal);
        return Ok(());
    }
    if source_xml.contains(xml)
        || (originals.is_empty() && verbatim_note_has_other_id(xml, note, path, budget))
    {
        note.verbatim_xml = None;
        note.source_ordinal = None;
        return Ok(());
    }
    let consistent = parse_verbatim_note(xml, note, path, budget).is_some_and(|edited| {
        if same_model(&edited, &note.content) {
            return true;
        }
        matching
            .and_then(|original| original.verbatim_xml.as_deref())
            .and_then(|source| parse_verbatim_note(source, note, path, budget))
            .is_some_and(|source| same_verbatim_note_content(&edited, &source, path, budget))
    });
    if !consistent {
        return Err(save_error(format!(
            "{} {} in {path} has verbatimXml that conflicts with its content",
            note.story_type, note.id
        )));
    }
    note.source_ordinal = None;
    Ok(())
}

fn verbatim_note_has_other_id(
    xml: &str,
    note: &Note,
    path: &str,
    budget: &mut ParseBudget<'_>,
) -> bool {
    let Ok(parsed) = crate::xml::parse_xml_strict(xml.as_bytes(), path, budget) else {
        return false;
    };
    let [element] = parsed.roots.as_slice() else {
        return false;
    };
    element.local_name() == note.story_type
        && element
            .attribute(Some("w"), "id")
            .and_then(|id| id.trim().parse::<i64>().ok())
            .is_some_and(|id| id as f64 != note.id)
}

fn parse_verbatim_note(
    xml: &str,
    note: &Note,
    path: &str,
    budget: &mut ParseBudget<'_>,
) -> Option<Vec<crate::block::BlockContent>> {
    let expected = if path.eq_ignore_ascii_case(FOOTNOTES) {
        "footnote"
    } else {
        "endnote"
    };
    super::raw::validate_raw_subtree(xml, "w", expected).ok()?;
    let parsed = crate::xml::parse_xml_strict(xml.as_bytes(), path, budget).ok()?;
    let [element] = parsed.roots.as_slice() else {
        return None;
    };
    if element.name != format!("w:{expected}")
        || note.story_type != expected
        || element.attribute(Some("w"), "id")?.trim().parse::<i64>().ok()? as f64 != note.id
        || element.attribute(Some("w"), "type").unwrap_or("normal") != note.note_type
    {
        return None;
    }
    crate::block::StoryParser {
        relationships: None,
        theme: None,
        styles: None,
        doc_defaults: None,
        numbering: None,
        media: &Default::default(),
        charts: &Default::default(),
        smart_art: &mut Default::default(),
        budget,
        ids: &mut crate::paragraph::HexIdAllocator::from_sha256(&"0".repeat(64)).ok()?,
        part: path,
    }
    .parse_blocks(element, 0, false)
    .ok()
}

fn same_verbatim_note_content(
    left: &[crate::block::BlockContent],
    right: &[crate::block::BlockContent],
    path: &str,
    budget: &mut ParseBudget<'_>,
) -> bool {
    use crate::block::{BlockContent, is_bookmark_block};

    left.len() == right.len()
        && left.iter().zip(right).all(|(left, right)| match (left, right) {
            (BlockContent::RawXml(a), BlockContent::RawXml(b))
                if is_bookmark_block(left) && is_bookmark_block(right) =>
            {
                let a = crate::xml::parse_xml_strict(a.xml.as_bytes(), path, budget);
                let b = crate::xml::parse_xml_strict(b.xml.as_bytes(), path, budget);
                let (Ok(a), Ok(b)) = (a, b) else {
                    return false;
                };
                let (Some(a), Some(b)) = (a.root(), b.root()) else {
                    return false;
                };
                a.name == b.name
                    && ["id", "name", "colFirst", "colLast"]
                        .iter()
                        .all(|name| a.attribute(Some("w"), name) == b.attribute(Some("w"), name))
            }
            _ => same_model(left, right),
        })
}

fn same_notes(notes: &[Note], source: &[Note]) -> bool {
    notes.len() == source.len()
        && notes.iter().zip(source).all(|(note, original)| {
            note == original
                || (note.verbatim_xml == original.verbatim_xml && same_note_content(note, original))
        })
}

fn same_note_content(note: &Note, original: &Note) -> bool {
    note.id == original.id
        && note.story_type == original.story_type
        && note.note_type == original.note_type
        && note.custom_root_bindings == original.custom_root_bindings
        && same_model(&note.content, &original.content)
}

fn same_model<T: serde::Serialize + PartialEq + ?Sized>(left: &T, right: &T) -> bool {
    same_model_with_relationships(left, right, None)
}

fn same_model_with_relationships<T: serde::Serialize + PartialEq + ?Sized>(
    left: &T,
    right: &T,
    relationships: Option<&RelationshipMap>,
) -> bool {
    left == right
        || match (serde_json::to_value(left), serde_json::to_value(right)) {
            (Ok(left), Ok(right)) => same_value(&left, &right, relationships),
            _ => false,
        }
}

fn ignores_model_field(
    key: &str,
    left: &Map<String, Value>,
    right: &Map<String, Value>,
    relationships: Option<&RelationshipMap>,
) -> bool {
    let linked_image = |fields: &Map<String, Value>| {
        fields.get("type").and_then(Value::as_str) == Some("image")
            && fields
                .get("rId")
                .and_then(Value::as_str)
                .is_some_and(|id| !id.is_empty())
    };
    let linked_watermark = |fields: &Map<String, Value>| {
        fields.get("kind").and_then(Value::as_str) == Some("picture")
            && (fields
                .get("mediaPath")
                .and_then(Value::as_str)
                .is_some_and(|path| !path.is_empty())
                || fields
                    .get("relId")
                    .and_then(Value::as_str)
                    .is_some_and(|id| {
                        relationships.is_some_and(|relationships| relationships.contains_key(id))
                    }))
    };
    match key {
        "sourceOrdinal" | "repeatedParaId" => true,
        "src" | "mimeType" => linked_image(left) && linked_image(right),
        "dataUrl" | "contentType" => linked_watermark(left) && linked_watermark(right),
        _ => false,
    }
}

fn same_value(left: &Value, right: &Value, relationships: Option<&RelationshipMap>) -> bool {
    match (left, right) {
        (Value::Object(left), Value::Object(right)) => {
            let ignore = |key| ignores_model_field(key, left, right, relationships);
            left.keys().filter(|key| !ignore(key.as_str())).count()
                == right.keys().filter(|key| !ignore(key.as_str())).count()
                && left
                    .iter()
                    .filter(|(key, _)| !ignore(key.as_str()))
                    .all(|(key, value)| {
                        right
                            .get(key)
                            .is_some_and(|other| same_value(value, other, relationships))
                    })
        }
        (Value::Array(left), Value::Array(right)) => {
            left.len() == right.len()
                && left
                    .iter()
                    .zip(right)
                    .all(|(left, right)| same_value(left, right, relationships))
        }
        _ => left == right,
    }
}

fn save_error(message: impl Into<String>) -> ParseError {
    ParseError::Canonical(format!("S13 package save: {}", message.into()))
}
