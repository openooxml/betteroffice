use std::collections::{BTreeSet, HashMap};

use serde::Deserialize;
use serde_json::{Map, Value};
use sha2::{Digest, Sha256};

use crate::block::StoryParser;
use crate::document::DocumentBody;
use crate::header_footer::HeaderFooter;
use crate::notes::Note;
use crate::numbering::{NumberingDefinitions, NumberingMap};
use crate::paragraph::HexIdAllocator;
use crate::relationships::{
    Relationship, RelationshipMap, RelationshipTarget, parse_relationships, relationship_part_path,
    relationship_types, resolve_relationship_target,
};
use crate::s8::note_part_relationships;
use crate::s9::{S9PackageWire, S9ParseOptions, parse_save_source};
use crate::xml::{ParseBudget, ParseError, ParseLimits};

use super::paragraph_ids::{S13ParagraphIds, S13PatchedPart};
use super::s13::{
    S13SaveOptions, S13SaveRequest, S13SelectiveSave, WriteSettings, save_core_only,
    write_docx_s13_parts_with_report, write_parts,
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
    #[serde(default)]
    fresh_package: bool,
    #[serde(default)]
    report_bindings: bool,
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
            || wire.paragraph_ids.is_some()
            || wire.selective.is_some()
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
        rewrite_body: false,
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
    if request.paragraph_ids.is_some() || request.selective.is_some() {
        return write_docx_s13_parts_with_report(request, &parts, Some(original_docx))
            .map_err(|error| error.to_string());
    }
    if wire.fresh_package {
        return write_parts(
            request,
            &parts,
            Some(original_docx),
            WriteSettings {
                product_save: true,
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
        None,
    )
}

fn write_model(
    request: S13SaveRequest,
    bytes: &[u8],
    parts: &[(String, Vec<u8>)],
    limits: &ParseLimits,
    coverage: Coverage,
    wire: Option<&mut Vec<String>>,
    source_digest: Option<String>,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    if request.paragraph_ids.is_some() || request.selective.is_some() {
        return write_parts(
            request,
            parts,
            Some(bytes),
            WriteSettings {
                legacy_diagnostics: false,
                ..WriteSettings::default()
            },
        );
    }
    request.determinism.validate()?;
    match write_guarded(
        &request,
        bytes,
        parts,
        limits,
        coverage,
        wire,
        source_digest,
    ) {
        Ok(saved) => Ok(saved),
        Err(error) => write_fallback(request, bytes, parts, &error.to_string()),
    }
}

fn write_guarded(
    request: &S13SaveRequest,
    bytes: &[u8],
    parts: &[(String, Vec<u8>)],
    limits: &ParseLimits,
    coverage: Coverage,
    wire: Option<&mut Vec<String>>,
    source_digest: Option<String>,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    let parse_options = S9ParseOptions {
        detect_variables: false,
        parse_headers_footers: !request.header_entries.is_empty()
            || !request.footer_entries.is_empty(),
        parse_notes: coverage.notes.iter().any(|covered| *covered),
        source_ordinals: coverage.source_ordinals,
        determinism_seed: source_digest,
        ..S9ParseOptions::default()
    };
    let unchanged_body = coverage.unchanged_body && !coverage.rewrite_body;
    let (body_path, source) =
        parse_save_source(bytes, parts, parse_options, limits, unchanged_body)?;
    let part_paths = PartPaths::new(parts);
    if part_paths.0.len() != parts.len() {
        return Err(save_error("ambiguous case-insensitive package paths"));
    }
    let mut numbering_added = false;
    if let Some(numbering) = &request.numbering
        && numbering != &source.numbering
    {
        if part_paths.contains(NUMBERING)
            || source.relationship_entries.iter().any(|(_, relationship)| {
                relationship.relationship_type == relationship_types::NUMBERING
            })
        {
            return Err(save_error("numbering differs from the source package"));
        }
        numbering_added = !numbering.abstract_nums.is_empty() || !numbering.nums.is_empty();
    }
    validate_relationships(
        &body_path,
        &request.relationship_entries,
        &source.relationship_entries,
    )?;
    let aliases = header_footer_paths(request, &source, &body_path, &part_paths, parts, limits)?;
    if aliases.changed.iter().any(|path| {
        path.eq_ignore_ascii_case(&body_path)
            || [COMMENTS, FOOTNOTES, ENDNOTES, NUMBERING]
                .iter()
                .any(|reserved| path.eq_ignore_ascii_case(reserved))
    }) {
        return Err(save_error("a header or footer targets another story part"));
    }
    if !unchanged_body && request.document.content.is_empty() && !source.document.content.is_empty()
    {
        return Err(save_error("the document body may be a placeholder"));
    }
    let mut changed = BTreeSet::new();
    let reuse_document_root = unchanged_body
        || request.document.custom_root_bindings == source.document.custom_root_bindings;
    if coverage.rewrite_body
        || !reuse_document_root
        || (!unchanged_body
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
    changed.extend(aliases.changed);
    let [footnotes, endnotes, footnote_separators, endnote_separators] = coverage.notes;
    let note_parts = [
        NotePart {
            path: part_paths.resolve(FOOTNOTES),
            source_notes: source.footnotes.as_deref().unwrap_or_default(),
            source_separators: source.footnote_separators.as_deref().unwrap_or_default(),
            notes_read: footnotes,
            separators_read: footnote_separators,
        },
        NotePart {
            path: part_paths.resolve(ENDNOTES),
            source_notes: source.endnotes.as_deref().unwrap_or_default(),
            source_separators: source.endnote_separators.as_deref().unwrap_or_default(),
            notes_read: endnotes,
            separators_read: endnote_separators,
        },
    ];
    for (part, notes, separators) in [
        (
            &note_parts[0],
            &request.footnotes,
            &request.footnote_separators,
        ),
        (
            &note_parts[1],
            &request.endnotes,
            &request.endnote_separators,
        ),
    ] {
        if part.changed(notes, separators)? {
            changed.insert(part.path.to_owned());
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
        let saved = save_core_only(request, &body_path, parts, bytes)?;
        if let Some(wire) = wire {
            wire.push(NO_BINDING_CHANGES.to_owned());
        }
        return Ok(saved);
    }

    let mut request = request.clone();
    let mut empty_notes = [false; 2];
    for (index, (part, notes, separators)) in [
        (
            note_parts[0],
            &mut request.footnotes,
            &mut request.footnote_separators,
        ),
        (
            note_parts[1],
            &mut request.endnotes,
            &mut request.endnote_separators,
        ),
    ]
    .into_iter()
    .enumerate()
    {
        if changed.contains(part.path) {
            part.prepare(notes, separators, parts, &source, limits)?;
            empty_notes[index] = true;
        }
    }
    let mut owners = aliases.owners.into_iter();
    for entries in [&mut request.header_entries, &mut request.footer_entries] {
        entries.retain(|_| owners.next().unwrap_or(true));
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
        },
    )?;
    report.warnings.extend(changed.into_iter().map(|part| format!(
        "DOCX_PART_REWRITTEN {part} was rewritten from the document model. Content the model does not carry may be lost. Saving edits through a session (createYrsSession, saveYrsDocx) patches changed paragraphs in place where it can."
    )));
    Ok((saved, report))
}

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
    reason: &str,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    let (saved, mut report) = write_parts(
        request,
        parts,
        Some(bytes),
        WriteSettings {
            legacy_diagnostics: false,
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
            return Err(save_error(format!(
                "cannot replace document relationship {id}"
            )));
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
            if !matches!(
                kind,
                relationship_types::HEADER | relationship_types::FOOTER
            ) {
                return None;
            }
            let target = story_path(body_path, relationship, kind).ok()??;
            paths.0.get(&target.to_ascii_lowercase()).copied()
        })
}

struct HeaderFooterPlan {
    changed: BTreeSet<String>,
    owners: Vec<bool>,
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
    let mut seen = HashMap::<String, (&str, &HeaderFooter)>::new();
    let mut changed = BTreeSet::new();
    let mut owners = Vec::new();
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
            let Some(target) = target else {
                owners.push(true);
                continue;
            };
            let path = paths.resolve(&target);
            let key = path.to_ascii_lowercase();
            let original = originals
                .get(id.as_str())
                .copied()
                .or_else(|| original_targets.get(&(key.clone(), kind)).copied());
            let previous = seen.get(&key).copied();
            let owner_relationships = if story.watermark.is_some()
                || original.is_some_and(|original| original.watermark.is_some())
                || previous.is_some_and(|(_, story)| story.watermark.is_some())
            {
                if let std::collections::hash_map::Entry::Vacant(entry) =
                    story_relationships.entry(key.clone())
                {
                    let rels_path = relationship_part_path(path);
                    let relationships = match crate::s8::find_part(parts, &rels_path) {
                        Some((name, bytes)) => parse_relationships(bytes, name, &mut budget)?,
                        None => RelationshipMap::new(),
                    };
                    entry.insert(relationships);
                }
                story_relationships.get(&key)
            } else {
                None
            };
            if let Some((owner_id, owner)) = previous {
                if !same_model_with_relationships(story, owner, owner_relationships) {
                    return Err(save_error(format!(
                        "conflicting header/footer models for {path} through {owner_id} and {id}"
                    )));
                }
            } else {
                seen.insert(key, (id.as_str(), story));
            }
            owners.push(previous.is_none());
            if story.content.is_empty()
                && original.is_some_and(|original| !original.content.is_empty())
            {
                return Err(save_error(format!("{path} may be a placeholder")));
            }
            if original.is_none_or(|original| {
                !same_model_with_relationships(story, original, owner_relationships)
            }) {
                changed.insert(path.to_owned());
            }
        }
    }
    Ok(HeaderFooterPlan { changed, owners })
}

#[derive(Clone, Copy)]
struct NotePart<'a> {
    path: &'a str,
    source_notes: &'a [Note],
    source_separators: &'a [Note],
    notes_read: bool,
    separators_read: bool,
}

impl NotePart<'_> {
    fn changed(&self, notes: &[Note], separators: &[Note]) -> Result<bool, ParseError> {
        let may_be_placeholder = |notes: &[Note], originals: &[Note]| {
            notes.iter().any(|note| {
                note.content.is_empty()
                    && note.verbatim_xml.is_none()
                    && originals
                        .iter()
                        .any(|original| original.id == note.id && !original.content.is_empty())
            })
        };
        if (self.notes_read && may_be_placeholder(notes, self.source_notes))
            || (self.separators_read && may_be_placeholder(separators, self.source_separators))
        {
            return Err(save_error(format!(
                "{} may contain note placeholders",
                self.path
            )));
        }
        Ok((self.notes_read && !same_notes(notes, self.source_notes))
            || (self.separators_read && !same_notes(separators, self.source_separators)))
    }

    fn prepare(
        self,
        notes: &mut Vec<Note>,
        separators: &mut Vec<Note>,
        parts: &[(String, Vec<u8>)],
        source: &S9PackageWire,
        limits: &ParseLimits,
    ) -> Result<(), ParseError> {
        if !self.notes_read {
            notes.extend_from_slice(self.source_notes);
        }
        if !self.separators_read {
            separators.extend_from_slice(self.source_separators);
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
        budget.compare_media_from(parts);
        let relationships = note_part_relationships(parts, self.path, &mut budget)?
            .unwrap_or_else(|| source.relationship_entries.iter().cloned().collect());
        let styles = source
            .styles
            .iter()
            .flat_map(|definitions| &definitions.styles)
            .map(|style| (style.style_id.clone(), style.clone()))
            .collect();
        let doc_defaults = source
            .styles
            .as_ref()
            .and_then(|definitions| definitions.doc_defaults.as_ref());
        let numbering = NumberingMap {
            definitions: source.numbering.clone(),
        };
        let media = crate::media::build_save_media_map(parts);
        let charts = source.chart_entries.iter().cloned().collect();
        let diagram_parts = parts
            .iter()
            .filter(|(path, _)| path.to_ascii_lowercase().starts_with("word/diagrams/"))
            .map(|(path, bytes)| (path.clone(), bytes.as_slice()))
            .collect();
        let mut smart_art = crate::smart_art::create_smart_art_context(&diagram_parts);
        let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64))?;
        let mut parser = StoryParser {
            relationships: Some(&relationships),
            theme: Some(&source.theme),
            styles: Some(&styles),
            doc_defaults,
            numbering: Some(&numbering),
            media: &media,
            charts: &charts,
            smart_art: &mut smart_art,
            budget: &mut budget,
            ids: &mut ids,
            part: self.path,
        };
        for (notes, supplied) in [(notes, self.notes_read), (separators, self.separators_read)] {
            if supplied {
                for note in notes {
                    let originals = originals
                        .get(&note_key(note))
                        .map(Vec::as_slice)
                        .unwrap_or_default();
                    clear_stale_note_xml(note, originals, &source_xml, &mut parser)?;
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
    parser: &mut StoryParser<'_, '_>,
) -> Result<(), ParseError> {
    let path = parser.part;
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
        || (originals.is_empty() && verbatim_note_has_other_id(xml, note, path, parser.budget))
    {
        note.verbatim_xml = None;
        note.source_ordinal = None;
        return Ok(());
    }
    let consistent = parse_verbatim_note(xml, note, parser).is_some_and(|edited| {
        if same_model(&edited, &note.content) {
            return true;
        }
        matching
            .and_then(|original| original.verbatim_xml.as_deref())
            .and_then(|source| parse_verbatim_note(source, note, parser))
            .is_some_and(|source| same_verbatim_note_content(&edited, &source, path, parser.budget))
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
    parser: &mut StoryParser<'_, '_>,
) -> Option<Vec<crate::block::BlockContent>> {
    let path = parser.part;
    let expected = if path.eq_ignore_ascii_case(FOOTNOTES) {
        "footnote"
    } else {
        "endnote"
    };
    super::raw::validate_raw_subtree(xml, "w", expected).ok()?;
    let parsed = crate::xml::parse_xml_strict(xml.as_bytes(), path, parser.budget).ok()?;
    let [element] = parsed.roots.as_slice() else {
        return None;
    };
    if element.name != format!("w:{expected}")
        || note.story_type != expected
        || element
            .attribute(Some("w"), "id")?
            .trim()
            .parse::<i64>()
            .ok()? as f64
            != note.id
        || element.attribute(Some("w"), "type").unwrap_or("normal") != note.note_type
    {
        return None;
    }
    parser.parse_blocks(element, 0, false).ok()
}

fn same_verbatim_note_content(
    left: &[crate::block::BlockContent],
    right: &[crate::block::BlockContent],
    path: &str,
    budget: &mut ParseBudget<'_>,
) -> bool {
    use crate::block::{BlockContent, is_bookmark_block};

    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(left, right)| match (left, right) {
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
                        && ["id", "name", "colFirst", "colLast"].iter().all(|name| {
                            a.attribute(Some("w"), name) == b.attribute(Some("w"), name)
                        })
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
            (Ok(left), Ok(right)) => same_value(&left, &right, relationships, false),
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
        "hdrFtrType" => [left, right].into_iter().all(|fields| {
            matches!(
                fields.get("type").and_then(Value::as_str),
                Some("header" | "footer")
            )
        }),
        "src" | "mimeType" => linked_image(left) && linked_image(right),
        "dataUrl" | "contentType" => linked_watermark(left) && linked_watermark(right),
        _ => false,
    }
}

fn same_value(
    left: &Value,
    right: &Value,
    relationships: Option<&RelationshipMap>,
    drawing_model: bool,
) -> bool {
    match (left, right) {
        (Value::Object(left), Value::Object(right)) => {
            let shape = [left, right]
                .into_iter()
                .all(|fields| fields.get("type").and_then(Value::as_str) == Some("shape"));
            let ignore = |key| {
                (drawing_model && matches!(key, "otherSha256" | "textBoxContentSha256"))
                    || ignores_model_field(key, left, right, relationships)
            };
            left.keys().filter(|key| !ignore(key.as_str())).count()
                == right.keys().filter(|key| !ignore(key.as_str())).count()
                && left
                    .iter()
                    .filter(|(key, _)| !ignore(key.as_str()))
                    .all(|(key, value)| {
                        right.get(key).is_some_and(|other| {
                            same_value(value, other, relationships, shape && key == "drawingModel")
                        })
                    })
        }
        (Value::Array(left), Value::Array(right)) => {
            left.len() == right.len()
                && left
                    .iter()
                    .zip(right)
                    .all(|(left, right)| same_value(left, right, relationships, false))
        }
        _ => left == right,
    }
}

fn save_error(message: impl Into<String>) -> ParseError {
    ParseError::Canonical(format!("S13 package save: {}", message.into()))
}
