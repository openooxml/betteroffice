use std::collections::{BTreeSet, HashMap};

use serde::Deserialize;

use crate::document::DocumentBody;
use crate::header_footer::HeaderFooter;
use crate::notes::Note;
use crate::numbering::NumberingDefinitions;
use crate::relationships::{
    Relationship, RelationshipTarget, relationship_types, resolve_relationship_target,
};
use crate::s9::{S9PackageWire, S9ParseOptions, parse_save_source};
use crate::xml::{ParseBudget, ParseError, ParseLimits};

use super::paragraph_ids::{S13ParagraphIds, S13PatchedPart};
use super::s13::{
    S13SaveOptions, S13SaveRequest, S13SelectiveSave, save_core_only,
    write_docx_s13_parts_with_report,
};
use super::{S13SaveReport, SerializerDeterminism};

const COMMENTS: &str = "word/comments.xml";
const FOOTNOTES: &str = "word/footnotes.xml";
const ENDNOTES: &str = "word/endnotes.xml";
const NUMBERING: &str = "word/numbering.xml";

pub const WHOLE_MODEL_WARNING: &str = "DOCX_WHOLE_MODEL_WRITE ";

#[derive(Clone, Copy)]
struct Coverage {
    source_ordinals: bool,
    notes: [bool; 4],
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct SaveWire {
    determinism: SerializerDeterminism,
    document: DocumentBody,
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
}

fn present_notes<'de, D: serde::Deserializer<'de>>(
    deserializer: D,
) -> Result<Option<Vec<Note>>, D::Error> {
    Vec::deserialize(deserializer).map(Some)
}

pub(crate) fn write_json(
    json: &str,
    original_docx: &[u8],
) -> Result<(Vec<u8>, S13SaveReport), String> {
    let wire: SaveWire = serde_json::from_str(json).map_err(|error| error.to_string())?;
    let coverage = Coverage {
        source_ordinals: json.contains("\"sourceOrdinal\""),
        notes: [
            wire.footnotes.is_some(),
            wire.endnotes.is_some(),
            wire.footnote_separators.is_some(),
            wire.endnote_separators.is_some(),
        ],
    };
    let request = S13SaveRequest {
        determinism: wire.determinism,
        document: wire.document,
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
    if let Some(languages) = wire.style_languages.as_ref() {
        if request.paragraph_ids.is_none() || request.selective.is_some() {
            return Err("styleLanguages requires a session save plan".to_owned());
        }
        return super::s13::write_parts(
            request,
            &parts,
            Some(original_docx),
            Some(languages),
            true,
        )
        .map_err(|error| error.to_string());
    }
    write_model(request, original_docx, &parts, &ParseLimits::default(), coverage)
        .map_err(|error| error.to_string())
}

pub fn write_model_with_report(
    request: S13SaveRequest,
    original_docx: &[u8],
    parts: &[(String, Vec<u8>)],
    limits: &ParseLimits,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    write_model(
        request,
        original_docx,
        parts,
        limits,
        Coverage {
            source_ordinals: true,
            notes: [true; 4],
        },
    )
}

fn write_model(
    mut request: S13SaveRequest,
    bytes: &[u8],
    parts: &[(String, Vec<u8>)],
    limits: &ParseLimits,
    coverage: Coverage,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    if request.paragraph_ids.is_some() || request.selective.is_some() {
        return write_docx_s13_parts_with_report(request, parts, Some(bytes));
    }
    request.determinism.validate()?;
    let parse_options = S9ParseOptions {
        detect_variables: false,
        parse_headers_footers: !request.header_entries.is_empty() || !request.footer_entries.is_empty(),
        parse_notes: coverage.notes.iter().any(|covered| *covered),
        source_ordinals: coverage.source_ordinals,
        ..S9ParseOptions::default()
    };
    let (body_path, source) = match parse_save_source(bytes, parts, parse_options, limits) {
        Ok(source) => source,
        Err(error) => return write_fallback(request, bytes, parts, &error.to_string()),
    };
    let part_paths = PartPaths::new(parts);
    let aliases = header_footer_paths(&request, &source, &body_path, &part_paths)?;
    if aliases.edits.values().any(|(_, path)| {
        path == &body_path
            || [COMMENTS, FOOTNOTES, ENDNOTES, NUMBERING]
                .iter()
                .any(|reserved| path.eq_ignore_ascii_case(reserved))
    }) {
        return write_fallback(request, bytes, parts, "a header or footer targets another story part");
    }
    let mut changed = BTreeSet::new();
    let reuse_document_root =
        request.document.custom_root_bindings == source.document.custom_root_bindings;
    if !same_model(&request.document.content, &source.document.content)
        || request.document.final_section_properties != source.document.final_section_properties
        || !reuse_document_root
    {
        changed.insert(body_path.clone());
    }
    if request.document.comments.is_some()
        && !same_model(&request.document.comments, &source.document.comments)
    {
        changed.insert(part_paths.resolve(COMMENTS).to_owned());
    }
    changed.extend(aliases.edits.values().map(|(_, path)| path.clone()));
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

    validate_relationships(&request.relationship_entries, &source.relationship_entries)?;
    if let Some(numbering) = &request.numbering
        && numbering != &source.numbering
    {
        if part_paths.contains(NUMBERING) || source.relationship_entries.iter().any(|(_, relationship)| {
            relationship.relationship_type == relationship_types::NUMBERING
        }) {
            return Err(save_error("cannot replace existing numbering definitions"));
        }
        if !numbering.abstract_nums.is_empty() || !numbering.nums.is_empty() {
            changed.insert(NUMBERING.to_owned());
        }
    }
    let added_story_relationship = request.relationship_entries.iter().any(|(id, relationship)| {
        matches!(relationship.relationship_type.as_str(), relationship_types::HEADER | relationship_types::FOOTER)
            && source.relationship_entries.iter().all(|(source_id, _)| source_id != id)
            && matches!(resolve_relationship_target(&body_path, relationship), Ok(RelationshipTarget::Internal(ref path))
                if part_paths.contains(path))
    });
    if changed.is_empty() && !added_story_relationship {
        return save_core_only(&request, &body_path, parts, bytes);
    }

    for notes in &note_parts {
        if changed.contains(notes.path) && notes.would_remove_all() {
            return write_fallback(request, bytes, parts, "the model omits every note in an existing note part");
        }
    }

    for notes in note_parts {
        if changed.contains(notes.path) {
            notes.prepare();
        }
    }
    let mut index = 0;
    for entries in [&mut request.header_entries, &mut request.footer_entries] {
        entries.retain(|_| {
            let key = &aliases.targets[index];
            let keep = key.as_ref().and_then(|key| aliases.edits.get(key))
                .is_none_or(|(owner, _)| *owner == index);
            index += 1;
            keep
        });
    }

    request.paragraph_ids = Some(source_part_plan(&source, &body_path, &part_paths, &changed));
    let (saved, mut report) = super::s13::write_parts(request, parts, Some(bytes), None, reuse_document_root)?;
    report.warnings.extend(changed.into_iter().map(|part| format!(
        "DOCX_PART_REWRITTEN {part} was written from the document model. Content the model does not carry may be lost. Edit through a session (createYrsSession, saveYrsDocx) to preserve that content."
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
    reason: &str,
) -> Result<(Vec<u8>, S13SaveReport), ParseError> {
    let (saved, mut report) = write_docx_s13_parts_with_report(request, parts, Some(bytes))?;
    report.warnings.push(format!(
        "{WHOLE_MODEL_WARNING}could not verify which stories were unchanged ({reason}). Stories written from the model may lose unmodelled content."
    ));
    Ok((saved, report))
}

struct PartPaths<'a>(HashMap<String, &'a str>);

impl<'a> PartPaths<'a> {
    fn new(parts: &'a [(String, Vec<u8>)]) -> Self {
        Self(parts.iter().map(|(path, _)| (path.to_ascii_lowercase(), path.as_str())).collect())
    }

    fn resolve<'b>(&'b self, path: &'b str) -> &'b str {
        self.0.get(&path.to_ascii_lowercase()).copied().unwrap_or(path)
    }

    fn contains(&self, path: &str) -> bool {
        self.0.contains_key(&path.to_ascii_lowercase())
    }
}

fn validate_relationships(
    requested: &[(String, Relationship)],
    source: &[(String, Relationship)],
) -> Result<(), ParseError> {
    let relationships: HashMap<_, _> = requested.iter().map(|(id, rel)| (id, rel)).collect();
    for (id, source) in source {
        if let Some(relationship) = relationships.get(id)
            && (relationship.target != source.target
                || relationship.target_mode != source.target_mode
                || relationship.relationship_type != source.relationship_type)
        {
            return Err(save_error(format!("cannot replace document relationship {id}")));
        }
    }
    Ok(())
}

fn story_path(body_path: &str, relationship: &Relationship, kind: &str) -> Result<Option<String>, ParseError> {
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
    source.relationship_entries.iter().filter_map(move |(_, relationship)| {
        let kind = relationship.relationship_type.as_str();
        if !matches!(kind, relationship_types::HEADER | relationship_types::FOOTER) {
            return None;
        }
        let target = story_path(body_path, relationship, kind).ok()??;
        paths.0.get(&target.to_ascii_lowercase()).copied()
    })
}

struct HeaderFooterPlan {
    edits: HashMap<String, (usize, String)>,
    targets: Vec<Option<String>>,
}

fn header_footer_paths(
    request: &S13SaveRequest,
    source: &S9PackageWire,
    body_path: &str,
    paths: &PartPaths<'_>,
) -> Result<HeaderFooterPlan, ParseError> {
    let relationships: HashMap<_, _> = request.relationship_entries.iter()
        .map(|(id, rel)| (id, rel)).collect();
    let originals: HashMap<_, _> = source.header_entries.iter().chain(&source.footer_entries)
        .flatten().map(|(id, story)| (id, story)).collect();
    let mut edits = HashMap::<String, (usize, String)>::new();
    let mut targets = Vec::new();
    let mut stories = Vec::<(&String, &HeaderFooter)>::new();
    for (entries, kind) in [
        (&request.header_entries, relationship_types::HEADER),
        (&request.footer_entries, relationship_types::FOOTER),
    ] {
        for (id, story) in entries {
            let target = relationships.get(id)
                .and_then(|rel| story_path(body_path, rel, kind).ok().flatten());
            let path = target.as_deref().map(|target| paths.resolve(target));
            let key = path.map(str::to_ascii_lowercase);
            if let (Some(path), Some(key)) = (path, &key)
                && originals.get(id).is_none_or(|original| !same_model(story, *original))
            {
                if let Some((previous, _)) = edits.get(key) {
                    let (prior_id, prior_story) = stories[*previous];
                    if !same_alias(story, prior_story, originals.get(id).copied(), originals.get(prior_id).copied()) {
                        return Err(save_error(format!(
                            "conflicting header/footer edits to {path} through {prior_id} and {id}"
                        )));
                    }
                } else {
                    edits.insert(key.clone(), (stories.len(), path.to_owned()));
                }
            }
            targets.push(key);
            stories.push((id, story));
        }
    }
    Ok(HeaderFooterPlan { edits, targets })
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

    fn would_remove_all(&self) -> bool {
        self.notes.is_empty() && self.separators.is_empty()
            && (self.notes_read || self.source_notes.is_empty())
            && (self.separators_read || self.source_separators.is_empty())
    }

    fn prepare(self) {
        if !self.notes_read {
            self.notes.extend_from_slice(self.source_notes);
        }
        if !self.separators_read {
            self.separators.extend_from_slice(self.source_separators);
        }
        let source = self.source_notes.iter().chain(self.source_separators);
        for note in self.notes.iter_mut().chain(self.separators.iter_mut()) {
            clear_stale_note_xml(note, source.clone());
        }
    }
}

fn clear_stale_note_xml<'a>(note: &mut Note, source: impl Iterator<Item = &'a Note>) {
    let original = source.filter(|original| original.verbatim_xml.is_some())
        .find(|original| {
            (note.verbatim_xml.is_none() || note.verbatim_xml == original.verbatim_xml)
                && same_note_content(note, original)
        });
    note.verbatim_xml = original.and_then(|original| original.verbatim_xml.clone());
    note.source_ordinal = original.and_then(|original| original.source_ordinal);
}

fn same_notes(notes: &[Note], source: &[Note]) -> bool {
    notes.len() == source.len() && notes.iter().zip(source).all(|(note, original)| {
        note == original || ((note.verbatim_xml.is_none() || note.verbatim_xml == original.verbatim_xml)
            && same_note_content(note, original))
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
    left == right || same_json(left, right, &["sourceOrdinal", "repeatedParaId"])
}

fn same_alias(
    left: &HeaderFooter,
    right: &HeaderFooter,
    source_left: Option<&HeaderFooter>,
    source_right: Option<&HeaderFooter>,
) -> bool {
    if same_model(left, right) {
        return true;
    }
    let (Some(source_left), Some(source_right)) = (source_left, source_right) else {
        return false;
    };
    let trees = [left, right, source_left, source_right].map(serde_json::to_value);
    let [Ok(left), Ok(right), Ok(source_left), Ok(source_right)] = trees else {
        return false;
    };
    same_alias_tree(&left, &right, &source_left, &source_right)
}

fn same_alias_tree(
    left: &serde_json::Value,
    right: &serde_json::Value,
    source_left: &serde_json::Value,
    source_right: &serde_json::Value,
) -> bool {
    use serde_json::Value;

    if left == right {
        return true;
    }
    match (left, right) {
        (Value::Object(left), Value::Object(right)) => {
            let considered = |key: &str| !matches!(key, "sourceOrdinal" | "repeatedParaId");
            left.keys().filter(|key| considered(key)).count()
                == right.keys().filter(|key| considered(key)).count()
                && left.iter().filter(|(key, _)| considered(key)).all(|(key, value)| {
                    let Some(other) = right.get(key) else { return false };
                    let before = &source_left[key.as_str()];
                    let other_before = &source_right[key.as_str()];
                    (key == "textId" && value == before && other == other_before)
                        || same_alias_tree(value, other, before, other_before)
                })
        }
        (Value::Array(left), Value::Array(right)) => {
            left.len() == right.len() && left.iter().zip(right).enumerate().all(|(index, (left, right))| {
                same_alias_tree(left, right, &source_left[index], &source_right[index])
            })
        }
        _ => false,
    }
}

fn same_json<T: serde::Serialize + ?Sized>(left: &T, right: &T, ignored: &[&str]) -> bool {
    fn normalize(mut json: String, ignored: &[&str]) -> String {
        for key in ignored {
            let field = format!(",\"{key}\":");
            if !json.contains(&field) {
                continue;
            }
            let mut output = String::with_capacity(json.len());
            let mut rest = json.as_str();
            while let Some(start) = rest.find(&field) {
                let value = &rest[start + field.len()..];
                let Some(end) = value.find([',', '}']) else { break };
                let scalar = &value[..end];
                if scalar == "true" || scalar == "false" || (!scalar.is_empty() && scalar.bytes().all(|byte| byte.is_ascii_digit())) {
                    output.push_str(&rest[..start]);
                    rest = &value[end..];
                } else {
                    output.push_str(&rest[..start + field.len()]);
                    rest = value;
                }
            }
            output.push_str(rest);
            json = output;
        }
        json
    }
    match (serde_json::to_string(left), serde_json::to_string(right)) {
        (Ok(left), Ok(right)) => normalize(left, ignored) == normalize(right, ignored),
        _ => false,
    }
}

fn save_error(message: impl Into<String>) -> ParseError {
    ParseError::Canonical(format!("S13 package save: {}", message.into()))
}
