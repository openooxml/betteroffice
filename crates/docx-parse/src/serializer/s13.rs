//! S13 package-save primitives.
//!
//! Selective document updates operate on exact UTF-8 spans. They never parse
//! and re-emit unchanged markup, so every byte outside an explicitly changed
//! paragraph remains authored exactly as it appeared in the source package.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use base64::Engine as _;
use indexmap::IndexMap;
use serde::{Deserialize, Serialize};

use crate::block::BlockContent;
use crate::document::DocumentBody;
use crate::header_footer::HeaderFooter;
use crate::image::Image;
use crate::inline::{Hyperlink, InlineNode, Run, RunContent};
use crate::notes::Note;
use crate::numbering::NumberingDefinitions;
use crate::paragraph::{Paragraph, ParagraphContent};
use crate::paragraph_identity::parse_paragraph_id;
use crate::relationships::{
    Relationship, TargetMode, relationship_part_path, relationship_types, resolve_relative_path,
};
use crate::table::{Table, TableCell};
use crate::vml::Watermark;
use crate::xml::ParseError;

use super::context::SerializerContext;
use super::numbering::serialize_numbering_xml;
use super::paragraph::serialize_paragraph;
use super::paragraph_ids::{
    COMMENTS_PART, S13ParagraphIds, apply_assignments, comment_companions, model_paragraph_ids,
    patch_comment_parts, patch_part,
};
use super::parts::{
    serialize_comments_extended_part, serialize_comments_extensible_part,
    serialize_comments_ids_part, serialize_comments_with_info, serialize_document_part,
    serialize_endnotes_part, serialize_footnotes_part, serialize_header_footer_part,
};
use super::raw::{validate_math_subtree, validate_raw_subtree, validate_replayed_fragment};
use super::s10::SerializerDeterminism;
use super::xml_writer::escape_xml;

const EMPTY_RELS_XML: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>";
const HEADER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FOOTER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const NUMBERING_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_EXTENDED_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
const COMMENTS_IDS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml";
const COMMENTS_EXTENSIBLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml";

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct S13SaveOptions {
    #[serde(default = "default_true")]
    pub update_modified_date: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub modified_by: Option<String>,
}

impl Default for S13SaveOptions {
    fn default() -> Self {
        Self {
            update_modified_date: true,
            modified_by: None,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct S13SelectiveSave {
    pub changed_para_ids: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct S13SaveRequest {
    pub determinism: SerializerDeterminism,
    pub document: DocumentBody,
    #[serde(default)]
    pub header_entries: Vec<(String, HeaderFooter)>,
    #[serde(default)]
    pub footer_entries: Vec<(String, HeaderFooter)>,
    #[serde(default)]
    pub footnotes: Vec<Note>,
    #[serde(default)]
    pub endnotes: Vec<Note>,
    #[serde(default)]
    pub footnote_separators: Vec<Note>,
    #[serde(default)]
    pub endnote_separators: Vec<Note>,
    #[serde(default)]
    pub relationship_entries: Vec<(String, Relationship)>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub numbering: Option<NumberingDefinitions>,
    #[serde(default)]
    pub options: S13SaveOptions,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selective: Option<S13SelectiveSave>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph_ids: Option<S13ParagraphIds>,
}

fn default_true() -> bool {
    true
}

/// Assemble a complete DOCX package and deflate it through the existing
/// `ooxml-opc` writer. Original entries seed the output in archive order;
/// only editor-owned parts are overwritten or appended.
pub fn write_docx_s13(
    request: S13SaveRequest,
    original_docx: &[u8],
) -> Result<Vec<u8>, ParseError> {
    let original_parts = ooxml_opc::unzip_parts(original_docx).map_err(ParseError::Container)?;
    write_docx_s13_parts(request, &original_parts, Some(original_docx))
}

/// [`write_docx_s13`] seeded from already-inflated parts. `source` is the
/// original container, when the caller still holds it, so unchanged members
/// re-emit verbatim instead of being re-deflated.
pub fn write_docx_s13_parts(
    mut request: S13SaveRequest,
    original_parts: &[(String, Vec<u8>)],
    source: Option<&[u8]>,
) -> Result<Vec<u8>, ParseError> {
    request.determinism.validate()?;
    let limits = crate::xml::ParseLimits::default();
    let mut budget = crate::xml::ParseBudget::new(&limits);
    let document_path = crate::relationships::office_document_path(original_parts, &mut budget)?;
    let mut package = Package::new(original_parts, document_path);
    let relationships: IndexMap<_, _> = request.relationship_entries.iter().cloned().collect();
    let paragraph_ids = request.paragraph_ids.take().unwrap_or_default();
    paragraph_ids.validate()?;
    apply_paragraph_id_assignments(&mut request, &paragraph_ids, &relationships, &package)?;
    let patched: HashMap<&str, &[(u32, String)]> = paragraph_ids
        .patched_parts
        .iter()
        .map(|part| (part.part.as_str(), part.para_ids.as_slice()))
        .collect();
    let patched_part = |package: &Package, path: &str| {
        let path = package.resolve_path(path);
        patched.get(path).and_then(|ids| {
            package
                .original_bytes(path)
                .and_then(|bytes| patch_part(bytes, ids))
        })
    };

    if request.selective.is_some() {
        validate_selective_header_footer_parts(&package, &relationships)?;
    } else {
        process_new_images(&mut request, &relationships, &mut package)?;
        process_new_watermark_images(&mut request, &relationships, &mut package)?;
        process_new_hyperlinks(&mut request, &relationships, &mut package)?;
    }

    let mut context = SerializerContext::new(&request.determinism)?;
    let document_xml = if let Some(patched) = patched_part(&package, "word/document.xml") {
        String::from_utf8(patched).map_err(|error| save_error(error.to_string()))?
    } else if let Some(selective) = request.selective.as_ref() {
        let original = package
            .text("word/document.xml")
            .ok_or_else(|| save_error("selective save has no word/document.xml"))?;
        match build_selective_document_xml(
            &request.document,
            &original,
            &selective.changed_para_ids,
            &mut context,
        )? {
            Some(patched) => patched,
            None => {
                let serialized = serialize_document_part(&request.document, &mut context)?;
                build_patched_document_xml(&original, &serialized, &selective.changed_para_ids)
                    .ok_or_else(|| save_error("selective document patch is unsafe"))?
            }
        }
    } else {
        serialize_document_part(&request.document, &mut context)?
    };
    package.set_text("word/document.xml", document_xml);

    serialize_header_footer_parts(
        &request.header_entries,
        &request.footer_entries,
        &relationships,
        &mut package,
        &mut context,
    )?;
    let story_parts: Vec<String> = relationships
        .values()
        .filter(|relationship| {
            matches!(
                relationship.relationship_type.as_str(),
                relationship_types::HEADER | relationship_types::FOOTER
            ) && relationship.target_mode != Some(TargetMode::External)
        })
        .filter_map(|relationship| {
            resolve_relative_path(&package.document_path, &relationship.target).ok()
        })
        .collect();
    for path in story_parts {
        if let Some(bytes) = patched_part(&package, &path) {
            package.set(path, bytes);
        }
    }

    if request.selective.is_none() {
        ensure_header_footer_parts(&relationships, &mut package)?;
        ensure_numbering_part(request.numbering.as_ref(), &mut package);
    }

    let patched_comments = patched.get(COMMENTS_PART).and_then(|ids| {
        let companions = comment_companions().map(|path| (path, package.original_bytes(path)));
        patch_comment_parts(package.original_bytes(COMMENTS_PART)?, &companions, ids)
    });
    if let Some(parts) = patched_comments {
        for (path, bytes) in parts {
            package.set(path, bytes);
        }
    } else {
        if comments_need_paragraph_ids(&request.document) {
            let mut reserved = crate::paragraph_identity::package_paragraph_ids(original_parts);
            model_paragraph_ids(&request.document, &mut reserved);
            model_paragraph_ids(&request.header_entries, &mut reserved);
            model_paragraph_ids(&request.footer_entries, &mut reserved);
            for notes in [
                &request.footnotes,
                &request.endnotes,
                &request.footnote_separators,
                &request.endnote_separators,
            ] {
                model_paragraph_ids(notes, &mut reserved);
            }
            context.reserve_paragraph_ids(reserved);
        }
        serialize_comment_parts(&request.document, &mut package, &mut context);
    }

    if request.selective.is_none() {
        let mut footnotes = request.footnote_separators;
        footnotes.extend(request.footnotes);
        if let Some(bytes) = patched_part(&package, "word/footnotes.xml") {
            package.set("word/footnotes.xml", bytes);
        } else if !footnotes.is_empty() {
            package.set_text(
                "word/footnotes.xml",
                serialize_footnotes_part(&footnotes, &mut context)?,
            );
        }
        let mut endnotes = request.endnote_separators;
        endnotes.extend(request.endnotes);
        if let Some(bytes) = patched_part(&package, "word/endnotes.xml") {
            package.set("word/endnotes.xml", bytes);
        } else if !endnotes.is_empty() {
            package.set_text(
                "word/endnotes.xml",
                serialize_endnotes_part(&endnotes, &mut context)?,
            );
        }
    }

    if request.options.update_modified_date || request.options.modified_by.is_some() {
        if let Some(core_xml) = package.text("docProps/core.xml") {
            let updated = update_core_properties(
                &core_xml,
                request.options.update_modified_date,
                request.options.modified_by.as_deref(),
                context.now(),
            );
            package.set_text("docProps/core.xml", updated);
        }
    }

    match source {
        Some(source) => ooxml_opc::rezip_parts_preserving(&package.refs(), source),
        None => ooxml_opc::rezip_parts_borrowed(&package.refs()),
    }
    .map_err(ParseError::Container)
}

/// Edits overlay borrowed `original` entries so unchanged parts are never copied.
#[derive(Debug)]
struct Package<'a> {
    original: &'a [(String, Vec<u8>)],
    appended: Vec<(String, Vec<u8>)>,
    overrides: HashMap<usize, Vec<u8>>,
    /// Maps a part path to an index into `original` followed by `appended`.
    positions: HashMap<String, usize>,
    document_path: String,
    document_relationships_path: String,
}

impl<'a> Package<'a> {
    fn new(original: &'a [(String, Vec<u8>)], document_path: String) -> Self {
        let positions = original
            .iter()
            .enumerate()
            .map(|(index, (path, _))| (path.clone(), index))
            .collect();
        let document_relationships_path =
            crate::relationships::relationship_part_path(&document_path);
        Self {
            original,
            appended: Vec::new(),
            overrides: HashMap::new(),
            positions,
            document_path,
            document_relationships_path,
        }
    }

    fn resolve_path<'b>(&'b self, path: &'b str) -> &'b str {
        match path {
            "word/document.xml" => &self.document_path,
            "word/_rels/document.xml.rels" => &self.document_relationships_path,
            _ => path,
        }
    }

    fn contains(&self, path: &str) -> bool {
        self.positions.contains_key(self.resolve_path(path))
    }

    fn bytes(&self, path: &str) -> Option<&[u8]> {
        let index = *self.positions.get(self.resolve_path(path))?;
        Some(if index < self.original.len() {
            self.overrides
                .get(&index)
                .map_or(self.original[index].1.as_slice(), Vec::as_slice)
        } else {
            self.appended[index - self.original.len()].1.as_slice()
        })
    }

    fn text(&self, path: &str) -> Option<String> {
        self.bytes(path)
            .map(|bytes| String::from_utf8_lossy(bytes).into_owned())
    }

    /// The part as the source package holds it, ignoring this save's edits.
    fn original_bytes(&self, path: &str) -> Option<&'a [u8]> {
        let index = *self.positions.get(self.resolve_path(path))?;
        self.original.get(index).map(|(_, bytes)| bytes.as_slice())
    }

    fn set(&mut self, path: impl Into<String>, bytes: Vec<u8>) {
        let path = path.into();
        let path = self.resolve_path(&path).to_owned();
        if let Some(index) = self.positions.get(&path).copied() {
            if index < self.original.len() {
                self.overrides.insert(index, bytes);
            } else {
                self.appended[index - self.original.len()].1 = bytes;
            }
        } else {
            self.positions
                .insert(path.clone(), self.original.len() + self.appended.len());
            self.appended.push((path, bytes));
        }
    }

    fn set_text(&mut self, path: impl Into<String>, xml: String) {
        self.set(path, xml.into_bytes());
    }

    fn paths(&self) -> impl Iterator<Item = &str> {
        self.original
            .iter()
            .map(|(path, _)| path.as_str())
            .chain(self.appended.iter().map(|(path, _)| path.as_str()))
    }

    /// Effective `(path, bytes)` entries in archive order: originals with
    /// overlays applied, then appends.
    fn refs(&self) -> Vec<(String, &[u8])> {
        let mut entries = Vec::with_capacity(self.original.len() + self.appended.len());
        for (index, (path, bytes)) in self.original.iter().enumerate() {
            let bytes = self
                .overrides
                .get(&index)
                .map_or(bytes.as_slice(), Vec::as_slice);
            entries.push((path.clone(), bytes));
        }
        entries.extend(
            self.appended
                .iter()
                .map(|(path, bytes)| (path.clone(), bytes.as_slice())),
        );
        entries
    }
}

fn save_error(message: impl Into<String>) -> ParseError {
    ParseError::Canonical(format!("S13 package save: {}", message.into()))
}

fn comments_need_paragraph_ids(document: &DocumentBody) -> bool {
    let valid = |id: Option<&str>| id.is_some_and(|id| parse_paragraph_id(id).is_some());
    document.comments.iter().flatten().any(|comment| {
        !valid(comment.content.last().and_then(|p| p.para_id.as_deref()))
            && !valid(comment.para_id.as_deref())
    })
}

/// Applies the plan's source-paragraph IDs to every serialized part's model.
fn apply_paragraph_id_assignments(
    request: &mut S13SaveRequest,
    paragraph_ids: &S13ParagraphIds,
    relationships: &IndexMap<String, Relationship>,
    package: &Package,
) -> Result<(), ParseError> {
    let parts = paragraph_ids.assignments_by_part();
    if parts.is_empty() {
        return Ok(());
    }
    if let Some(ids) = parts.get(package.document_path.as_str()) {
        let comments = request.document.comments.take();
        apply_assignments(&mut request.document, ids)?;
        request.document.comments = comments;
    }
    if let (Some(ids), Some(comments)) =
        (parts.get(COMMENTS_PART), request.document.comments.as_mut())
    {
        apply_assignments(comments, ids)?;
    }
    for entries in [&mut request.header_entries, &mut request.footer_entries] {
        for (relationship_id, story) in entries.iter_mut() {
            let path = relationships
                .get(relationship_id)
                .filter(|relationship| relationship.target_mode != Some(TargetMode::External))
                .and_then(|relationship| {
                    resolve_relative_path(&package.document_path, &relationship.target).ok()
                });
            if let Some(ids) = path.as_deref().and_then(|path| parts.get(path)) {
                apply_assignments(story, ids)?;
            }
        }
    }
    for (path, notes) in [
        ("word/footnotes.xml", &mut request.footnotes),
        ("word/footnotes.xml", &mut request.footnote_separators),
        ("word/endnotes.xml", &mut request.endnotes),
        ("word/endnotes.xml", &mut request.endnote_separators),
    ] {
        if let Some(ids) = parts.get(path) {
            apply_assignments(notes, ids)?;
        }
    }
    Ok(())
}

fn validate_selective_header_footer_parts(
    package: &Package,
    relationships: &IndexMap<String, Relationship>,
) -> Result<(), ParseError> {
    for relationship in relationships.values() {
        if relationship.relationship_type != relationship_types::HEADER
            && relationship.relationship_type != relationship_types::FOOTER
        {
            continue;
        }
        if relationship.target_mode == Some(TargetMode::External) {
            continue;
        }
        let path = resolve_relative_path(&package.document_path, &relationship.target)?;
        if !package.contains(&path) {
            return Err(save_error(format!(
                "selective save cannot register new header/footer part {path}"
            )));
        }
    }
    Ok(())
}

fn serialize_header_footer_parts(
    headers: &[(String, HeaderFooter)],
    footers: &[(String, HeaderFooter)],
    relationships: &IndexMap<String, Relationship>,
    package: &mut Package,
    context: &mut SerializerContext,
) -> Result<(), ParseError> {
    for (entries, relationship_type) in [
        (headers, relationship_types::HEADER),
        (footers, relationship_types::FOOTER),
    ] {
        for (relationship_id, story) in entries {
            let Some(relationship) = relationships.get(relationship_id) else {
                continue;
            };
            if relationship.relationship_type != relationship_type
                || relationship.target.is_empty()
                || relationship.target_mode == Some(TargetMode::External)
            {
                continue;
            }
            package.set_text(
                resolve_relative_path(&package.document_path, &relationship.target)?,
                serialize_header_footer_part(story, context)?,
            );
        }
    }
    Ok(())
}

fn ensure_header_footer_parts(
    relationships: &IndexMap<String, Relationship>,
    package: &mut Package,
) -> Result<(), ParseError> {
    let parts: Vec<_> = relationships
        .iter()
        .filter_map(|(relationship_id, relationship)| {
            if relationship.target_mode == Some(TargetMode::External) {
                return None;
            }
            let content_type = match relationship.relationship_type.as_str() {
                relationship_types::HEADER => HEADER_CONTENT_TYPE,
                relationship_types::FOOTER => FOOTER_CONTENT_TYPE,
                _ => return None,
            };
            let target = relationship.target.clone();
            Some((
                relationship_id.as_str(),
                relationship.relationship_type.as_str(),
                target,
                content_type,
            ))
        })
        .collect();
    if parts.is_empty() {
        return Ok(());
    }

    if let Some(mut content_types) = package.text("[Content_Types].xml") {
        let mut changed = false;
        for (_, _, target, content_type) in &parts {
            let part_name = format!(
                "/{}",
                resolve_relative_path(&package.document_path, target)?
            );
            if !content_types.contains(&format!("PartName=\"{part_name}\"")) {
                let entry =
                    format!("<Override PartName=\"{part_name}\" ContentType=\"{content_type}\"/>");
                if let Some(updated) = append_before(&content_types, "</Types>", &entry) {
                    content_types = updated;
                    changed = true;
                }
            }
        }
        if changed {
            package.set_text("[Content_Types].xml", content_types);
        }
    }

    let path = "word/_rels/document.xml.rels";
    let mut relationships = RelationshipsIndex::parse(read_rels_or_stub(package, path));
    for (relationship_id, relationship_type, target, _) in parts {
        if relationships.xml_contains(&format!("Id=\"{relationship_id}\"")) {
            continue;
        }
        relationships.push_new(relationship_id.to_owned(), relationship_type, target, None);
    }
    if let Some(updated) = relationships.serialize() {
        package.set_text(path, updated);
    }
    Ok(())
}

fn ensure_numbering_part(numbering: Option<&NumberingDefinitions>, package: &mut Package) {
    let Some(numbering) = numbering else { return };
    if numbering.abstract_nums.is_empty() && numbering.nums.is_empty() {
        return;
    }
    if package.contains("word/numbering.xml") {
        return;
    }
    package.set_text("word/numbering.xml", serialize_numbering_xml(numbering));

    if let Some(content_types) = package.text("[Content_Types].xml")
        && !content_types.contains("PartName=\"/word/numbering.xml\"")
    {
        let entry = format!(
            "<Override PartName=\"/word/numbering.xml\" ContentType=\"{NUMBERING_CONTENT_TYPE}\"/>"
        );
        if let Some(updated) = append_before(&content_types, "</Types>", &entry) {
            package.set_text("[Content_Types].xml", updated);
        }
    }

    let path = "word/_rels/document.xml.rels";
    let mut relationships = RelationshipsIndex::parse(read_rels_or_stub(package, path));
    if !relationships.xml_contains("Target=\"numbering.xml\"") {
        let relationship_id = relationships.next_id();
        relationships.push_new(
            relationship_id,
            relationship_types::NUMBERING,
            "numbering.xml".to_owned(),
            None,
        );
        if let Some(updated) = relationships.serialize() {
            package.set_text(path, updated);
        }
    }
}

fn serialize_comment_parts(
    document: &DocumentBody,
    package: &mut Package,
    context: &mut SerializerContext,
) {
    let Some(comments) = document
        .comments
        .as_ref()
        .filter(|comments| !comments.is_empty())
    else {
        return;
    };
    let (comments_xml, infos) = serialize_comments_with_info(comments, context);
    package.set_text("word/comments.xml", comments_xml);

    let companions = [
        (
            "word/commentsExtended.xml",
            serialize_comments_extended_part(&infos),
        ),
        ("word/commentsIds.xml", serialize_comments_ids_part(&infos)),
        (
            "word/commentsExtensible.xml",
            serialize_comments_extensible_part(&infos, comments),
        ),
    ];
    for (path, xml) in companions {
        if !xml.is_empty() {
            package.set_text(path, xml);
        }
    }
    ensure_comment_parts(package);
}

fn ensure_comment_parts(package: &mut Package) {
    let parts = [
        (
            "/word/comments.xml",
            COMMENTS_CONTENT_TYPE,
            "comments.xml",
            relationship_types::COMMENTS,
        ),
        (
            "/word/commentsExtended.xml",
            COMMENTS_EXTENDED_CONTENT_TYPE,
            "commentsExtended.xml",
            relationship_types::COMMENTS_EXTENDED,
        ),
        (
            "/word/commentsIds.xml",
            COMMENTS_IDS_CONTENT_TYPE,
            "commentsIds.xml",
            relationship_types::COMMENTS_IDS,
        ),
        (
            "/word/commentsExtensible.xml",
            COMMENTS_EXTENSIBLE_CONTENT_TYPE,
            "commentsExtensible.xml",
            relationship_types::COMMENTS_EXTENSIBLE,
        ),
    ];

    if let Some(mut content_types) = package.text("[Content_Types].xml") {
        let mut changed = false;
        for (part_name, content_type, _, _) in parts {
            if content_types.contains(part_name) {
                continue;
            }
            let entry =
                format!("<Override PartName=\"{part_name}\" ContentType=\"{content_type}\"/>");
            if let Some(updated) = append_before(&content_types, "</Types>", &entry) {
                content_types = updated;
                changed = true;
            }
        }
        if changed {
            package.set_text("[Content_Types].xml", content_types);
        }
    }

    let path = "word/_rels/document.xml.rels";
    let Some(relationships_xml) = package.text(path) else {
        return;
    };
    let mut relationships = RelationshipsIndex::parse(relationships_xml);
    for (_, _, target, relationship_type) in parts {
        if relationships.xml_contains(target) {
            continue;
        }
        let relationship_id = relationships.next_id();
        relationships.push_new(relationship_id, relationship_type, target.to_owned(), None);
    }
    if let Some(updated) = relationships.serialize() {
        package.set_text(path, updated);
    }
}

fn read_rels_or_stub(package: &Package, path: &str) -> String {
    normalize_relationships_root(
        &package
            .text(path)
            .unwrap_or_else(|| EMPTY_RELS_XML.to_owned()),
    )
}

fn normalize_relationships_root(xml: &str) -> String {
    let Some(start) = xml.find("<Relationships") else {
        return xml.to_owned();
    };
    let Some(end) = find_tag_end(xml.as_bytes(), start) else {
        return xml.to_owned();
    };
    let tag = &xml[start..=end];
    if !is_self_closing(tag.as_bytes()) {
        return xml.to_owned();
    }
    let mut opening = tag[..tag.len() - 1].trim_end().to_owned();
    opening.pop();
    opening.push('>');
    let replacement = format!("{opening}</Relationships>");
    let mut normalized = String::with_capacity(xml.len() + replacement.len() - tag.len());
    normalized.push_str(&xml[..start]);
    normalized.push_str(&replacement);
    normalized.push_str(&xml[end + 1..]);
    normalized
}

fn append_before(xml: &str, closing: &str, value: &str) -> Option<String> {
    let offset = xml.find(closing)?;
    let mut updated = String::with_capacity(xml.len() + value.len());
    updated.push_str(&xml[..offset]);
    updated.push_str(value);
    updated.push_str(&xml[offset..]);
    Some(updated)
}

fn find_max_relationship_id(xml: &str) -> u64 {
    let mut maximum = 0u64;
    let mut cursor = 0usize;
    while let Some(relative) = xml[cursor..].find("Id=\"rId") {
        let start = cursor + relative + 7;
        let digits: String = xml[start..]
            .chars()
            .take_while(char::is_ascii_digit)
            .collect();
        if let Ok(value) = digits.parse::<u64>() {
            maximum = maximum.max(value);
        }
        cursor = start + digits.len().max(1);
    }
    maximum
}

/// Attribute location: source XML span or staged `appended` element span.
#[derive(Clone, Copy, Debug)]
enum Attr {
    Span(usize, usize),
    Appended(usize, usize, usize),
}

#[derive(Clone, Copy, Debug)]
struct RelationshipEntry {
    id: Option<Attr>,
    relationship_type: Option<Attr>,
    target: Option<Attr>,
    target_mode: Option<Attr>,
}

impl RelationshipEntry {
    fn parse(xml: &str, tag: &str) -> Self {
        let attr = |name: &str| {
            xml_attribute(tag, name).map(|value| {
                let start = value.as_ptr() as usize - xml.as_ptr() as usize;
                Attr::Span(start, start + value.len())
            })
        };
        Self {
            id: attr("Id"),
            relationship_type: attr("Type"),
            target: attr("Target"),
            target_mode: attr("TargetMode"),
        }
    }

    fn staged(element_position: usize, element: &str) -> Self {
        let attr = |name: &str| {
            xml_attribute(element, name).map(|value| {
                let start = value.as_ptr() as usize - element.as_ptr() as usize;
                Attr::Appended(element_position, start, start + value.len())
            })
        };
        Self {
            id: attr("Id"),
            relationship_type: attr("Type"),
            target: attr("Target"),
            target_mode: attr("TargetMode"),
        }
    }
}

/// A `.rels` part's entries plus staged appends; lookup maps build lazily.
/// Staged entries stay invisible to `external_hyperlink_id`/`existing_position`.
#[derive(Debug, Default)]
struct RelationshipsIndex {
    xml: String,
    entries: Vec<RelationshipEntry>,
    existing_count: usize,
    next_number: u64,
    appended: Vec<String>,
    by_id: Option<HashMap<String, usize>>,
    external_hyperlink_ids: Option<HashMap<String, String>>,
    normalized_target_ids: Option<HashMap<String, String>>,
}

impl RelationshipsIndex {
    fn parse(xml: String) -> Self {
        let entries: Vec<RelationshipEntry> = relationship_tags(&xml)
            .map(|tag| RelationshipEntry::parse(&xml, tag))
            .collect();
        Self {
            next_number: find_max_relationship_id(&xml) + 1,
            existing_count: entries.len(),
            entries,
            xml,
            ..Self::default()
        }
    }

    fn attr<'a>(&'a self, attr: &Option<Attr>) -> Option<&'a str> {
        match attr {
            Some(Attr::Span(start, end)) => self.xml.get(*start..*end),
            Some(Attr::Appended(entry, start, end)) => self.appended.get(*entry)?.get(*start..*end),
            None => None,
        }
    }

    fn id<'a>(&'a self, entry: &RelationshipEntry) -> Option<&'a str> {
        self.attr(&entry.id)
    }

    fn target<'a>(&'a self, entry: &RelationshipEntry) -> Option<&'a str> {
        self.attr(&entry.target)
    }

    fn target_matches(&self, entry: &RelationshipEntry, href: &str) -> bool {
        self.target(entry)
            .is_some_and(|target| decode_xml_entities(target) == href)
    }

    fn is_external_hyperlink(&self, entry: &RelationshipEntry) -> bool {
        self.attr(&entry.relationship_type) == Some(relationship_types::HYPERLINK)
            && self.attr(&entry.target_mode) == Some("External")
    }

    fn ids(&mut self) -> &HashMap<String, usize> {
        if self.by_id.is_none() {
            let mut map = HashMap::with_capacity(self.entries.len());
            for (position, entry) in self.entries.iter().enumerate() {
                if let Some(id) = self.id(entry) {
                    map.entry(id.to_owned()).or_insert(position);
                }
            }
            self.by_id = Some(map);
        }
        self.by_id.as_ref().unwrap()
    }

    fn position(&mut self, id: &str) -> Option<usize> {
        self.ids().get(id).copied()
    }

    fn contains_id(&mut self, id: &str) -> bool {
        self.ids().contains_key(id)
    }

    fn existing_position(&mut self, id: &str) -> Option<usize> {
        let position = self.position(id)?;
        (position < self.existing_count).then_some(position)
    }

    fn entry(&self, position: usize) -> &RelationshipEntry {
        &self.entries[position]
    }

    fn external_hyperlink_id(&mut self, href: &str) -> Option<String> {
        if self.external_hyperlink_ids.is_none() {
            let mut map = HashMap::new();
            for position in 0..self.existing_count {
                let entry = self.entries[position];
                if self.is_external_hyperlink(&entry)
                    && let (Some(target), Some(id)) = (self.target(&entry), self.id(&entry))
                {
                    map.entry(decode_xml_entities(target))
                        .or_insert_with(|| id.to_owned());
                }
            }
            self.external_hyperlink_ids = Some(map);
        }
        self.external_hyperlink_ids
            .as_ref()
            .unwrap()
            .get(href)
            .cloned()
    }

    fn id_for_target(&mut self, target: &str) -> Option<String> {
        if self.normalized_target_ids.is_none() {
            let mut map = HashMap::new();
            for position in 0..self.entries.len() {
                let entry = self.entries[position];
                if let (Some(candidate), Some(id)) = (self.target(&entry), self.id(&entry)) {
                    map.entry(normalize_media_target(candidate).to_owned())
                        .or_insert_with(|| id.to_owned());
                }
            }
            self.normalized_target_ids = Some(map);
        }
        self.normalized_target_ids
            .as_ref()
            .unwrap()
            .get(normalize_media_target(target))
            .cloned()
    }

    fn xml_contains(&self, needle: &str) -> bool {
        self.xml.contains(needle) || self.appended.iter().any(|entry| entry.contains(needle))
    }

    fn next_id(&self) -> String {
        format!("rId{}", self.next_number)
    }

    fn push_new(
        &mut self,
        id: String,
        relationship_type: &str,
        target: String,
        target_mode: Option<&str>,
    ) {
        let element = format!(
            "<Relationship Id=\"{id}\" Type=\"{relationship_type}\" Target=\"{target}\"{}/>",
            target_mode
                .map(|mode| format!(" TargetMode=\"{mode}\""))
                .unwrap_or_default()
        );
        if let Some(number) = id.strip_prefix("rId").and_then(|suffix| {
            suffix
                .chars()
                .take_while(char::is_ascii_digit)
                .collect::<String>()
                .parse::<u64>()
                .ok()
        }) {
            self.next_number = self.next_number.max(number + 1);
        }
        let position = self.entries.len();
        let entry = RelationshipEntry::staged(self.appended.len(), &element);
        self.entries.push(entry);
        self.appended.push(element);
        if self.by_id.is_none() && self.normalized_target_ids.is_none() {
            return;
        }
        let entry = &self.entries[position];
        let id = self.id(entry).map(str::to_owned);
        let target = self.target(entry).map(str::to_owned);
        if let Some(map) = self.by_id.as_mut()
            && let Some(id) = id.clone()
        {
            map.entry(id).or_insert(position);
        }
        if let Some(map) = self.normalized_target_ids.as_mut()
            && let (Some(target), Some(id)) = (target, id)
        {
            map.entry(normalize_media_target(&target).to_owned())
                .or_insert(id);
        }
    }

    fn try_append_new(
        &mut self,
        id: String,
        relationship_type: &str,
        target: String,
        target_mode: Option<&str>,
    ) -> Option<String> {
        if !self.xml.contains("</Relationships>") {
            return None;
        }
        self.push_new(id, relationship_type, target, target_mode);
        append_before(&self.xml, "</Relationships>", &self.appended.concat())
    }

    fn serialize(&self) -> Option<String> {
        if self.appended.is_empty() {
            return None;
        }
        append_before(&self.xml, "</Relationships>", &self.appended.concat())
    }
}

fn process_new_images(
    request: &mut S13SaveRequest,
    relationships: &IndexMap<String, Relationship>,
    package: &mut Package,
) -> Result<(), ParseError> {
    let mut image_number = find_max_image_number(package);
    let mut extensions = HashSet::new();

    process_image_part(
        package,
        "word/_rels/document.xml.rels",
        std::iter::once(&mut request.document.content),
        &mut image_number,
        &mut extensions,
    )?;
    process_image_part(
        package,
        "word/_rels/footnotes.xml.rels",
        request
            .footnote_separators
            .iter_mut()
            .chain(request.footnotes.iter_mut())
            .map(|note| &mut note.content),
        &mut image_number,
        &mut extensions,
    )?;
    process_image_part(
        package,
        "word/_rels/endnotes.xml.rels",
        request
            .endnote_separators
            .iter_mut()
            .chain(request.endnotes.iter_mut())
            .map(|note| &mut note.content),
        &mut image_number,
        &mut extensions,
    )?;

    for (entries, relationship_type) in [
        (&mut request.header_entries, relationship_types::HEADER),
        (&mut request.footer_entries, relationship_types::FOOTER),
    ] {
        for (relationship_id, story) in entries {
            let Some(relationship) = relationships.get(relationship_id) else {
                continue;
            };
            if relationship.relationship_type != relationship_type
                || relationship.target_mode == Some(TargetMode::External)
            {
                continue;
            }
            process_image_part(
                package,
                &owner_relationships_path(package, &relationship.target)?,
                std::iter::once(&mut story.content),
                &mut image_number,
                &mut extensions,
            )?;
        }
    }

    register_image_extensions(package, &extensions);
    Ok(())
}

fn process_image_part<'a>(
    package: &mut Package,
    relationships_path: &str,
    stories: impl Iterator<Item = &'a mut Vec<BlockContent>>,
    image_number: &mut u64,
    extensions: &mut HashSet<String>,
) -> Result<(), ParseError> {
    let mut relationships =
        RelationshipsIndex::parse(read_rels_or_stub(package, relationships_path));
    for blocks in stories {
        visit_new_images(blocks, &mut |image| {
            let Some(source) = image
                .src
                .as_deref()
                .filter(|source| source.starts_with("data:") && image.relationship_id.is_empty())
            else {
                return Ok(());
            };
            let (bytes, extension) = decode_image_data_url(source)?;
            *image_number += 1;
            let filename = format!("image{image_number}.{extension}");
            let new_relationship_id = relationships.next_id();
            package.set(format!("word/media/{filename}"), bytes);
            relationships.push_new(
                new_relationship_id.clone(),
                relationship_types::IMAGE,
                format!("media/{filename}"),
                None,
            );
            extensions.insert(extension);
            image.relationship_id = new_relationship_id;
            Ok(())
        })?;
    }
    if let Some(updated) = relationships.serialize() {
        package.set_text(relationships_path, updated);
    }
    Ok(())
}

fn run_has_drawing_image(run: &Run) -> bool {
    run.content
        .iter()
        .any(|content| matches!(content, RunContent::Drawing { .. }))
}

fn blocks_have_drawing_image(blocks: &[BlockContent]) -> bool {
    blocks.iter().any(|block| match block {
        BlockContent::Paragraph(paragraph) => {
            paragraph.content.iter().any(|content| match content {
                ParagraphContent::Inline(InlineNode::Run(run)) => run_has_drawing_image(run),
                ParagraphContent::Tracked(tracked) => tracked.content.iter().any(
                    |inline| matches!(inline, InlineNode::Run(run) if run_has_drawing_image(run)),
                ),
                _ => false,
            })
        }
        BlockContent::Table(table) => table.rows.iter().any(|row| {
            row.cells
                .iter()
                .any(|cell| blocks_have_drawing_image(&cell.content))
        }),
        BlockContent::BlockSdt(_) | BlockContent::RawXml(_) => false,
    })
}

fn visit_new_images(
    blocks: &mut [BlockContent],
    visit: &mut impl FnMut(&mut Image) -> Result<(), ParseError>,
) -> Result<(), ParseError> {
    for block in blocks {
        if !blocks_have_drawing_image(std::slice::from_ref(block)) {
            continue;
        }
        match block {
            BlockContent::Paragraph(paragraph) => {
                for content in &mut Arc::make_mut(paragraph).content {
                    match content {
                        ParagraphContent::Inline(InlineNode::Run(run)) => {
                            visit_run_images(run, visit)?
                        }
                        ParagraphContent::Tracked(tracked) => {
                            for inline in &mut tracked.content {
                                if let InlineNode::Run(run) = inline {
                                    visit_run_images(run, visit)?;
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            BlockContent::Table(table) => {
                for row in &mut Arc::make_mut(table).rows {
                    for cell in &mut row.cells {
                        visit_new_images(&mut cell.content, visit)?;
                    }
                }
            }
            // New images inside block SDTs remain on the selective-patch path.
            BlockContent::BlockSdt(_) => {}
            BlockContent::RawXml(_) => {}
        }
    }
    Ok(())
}

fn visit_run_images(
    run: &mut Run,
    visit: &mut impl FnMut(&mut Image) -> Result<(), ParseError>,
) -> Result<(), ParseError> {
    for content in &mut run.content {
        if let RunContent::Drawing { image } = content {
            visit(image)?;
        }
    }
    Ok(())
}

fn decode_image_data_url(source: &str) -> Result<(Vec<u8>, String), ParseError> {
    let encoded = source
        .strip_prefix("data:")
        .and_then(|source| source.split_once(";base64,"))
        .ok_or_else(|| save_error("invalid image data URL"))?;
    if encoded.0.is_empty() || encoded.1.is_empty() {
        return Err(save_error("invalid image data URL"));
    }
    let extension = match encoded.0 {
        "image/png" => "png",
        "image/jpeg" => "jpeg",
        "image/gif" => "gif",
        "image/bmp" => "bmp",
        "image/tiff" => "tiff",
        "image/webp" => "webp",
        "image/svg+xml" => "svg",
        _ => "png",
    };
    let bytes = base64::engine::general_purpose::STANDARD
        .decode(encoded.1)
        .map_err(|_| save_error("invalid image data URL base64"))?;
    Ok((bytes, extension.to_owned()))
}

fn find_max_image_number(package: &Package) -> u64 {
    package
        .paths()
        .filter_map(|path| {
            let suffix = path.strip_prefix("word/media/image")?;
            let digits: String = suffix.chars().take_while(char::is_ascii_digit).collect();
            (!digits.is_empty() && suffix[digits.len()..].starts_with('.'))
                .then(|| digits.parse::<u64>().ok())
                .flatten()
        })
        .max()
        .unwrap_or(0)
}

fn register_image_extensions(package: &mut Package, extensions: &HashSet<String>) {
    if extensions.is_empty() {
        return;
    }
    let Some(mut content_types) = package.text("[Content_Types].xml") else {
        return;
    };
    let mut changed = false;
    // Preserve package discovery order for newly appended image parts.
    let mut ordered = Vec::new();
    for path in package.paths() {
        let Some(extension) = path
            .strip_prefix("word/media/")
            .and_then(|name| name.rsplit_once('.').map(|(_, extension)| extension))
        else {
            continue;
        };
        if extensions.contains(extension) && !ordered.iter().any(|seen| seen == extension) {
            ordered.push(extension.to_owned());
        }
    }
    for extension in ordered {
        if content_types.contains(&format!("Extension=\"{extension}\"")) {
            continue;
        }
        let content_type = image_content_type(&extension);
        let entry = format!("<Default Extension=\"{extension}\" ContentType=\"{content_type}\"/>");
        if let Some(updated) = append_before(&content_types, "</Types>", &entry) {
            content_types = updated;
            changed = true;
        }
    }
    if changed {
        package.set_text("[Content_Types].xml", content_types);
    }
}

fn image_content_type(extension: &str) -> &'static str {
    match extension {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        "svg" => "image/svg+xml",
        "webp" => "image/webp",
        "wmf" => "image/x-wmf",
        "emf" => "image/x-emf",
        _ => "application/octet-stream",
    }
}

fn owner_relationships_path(package: &Package, target: &str) -> Result<String, ParseError> {
    let path = resolve_relative_path(&package.document_path, target)?;
    Ok(relationship_part_path(&path))
}

fn process_new_watermark_images(
    request: &mut S13SaveRequest,
    relationships: &IndexMap<String, Relationship>,
    package: &mut Package,
) -> Result<(), ParseError> {
    let mut image_number = find_max_image_number(package);
    let mut extensions = HashSet::new();
    let mut written_media = HashMap::<String, String>::new();
    let mut rels_parts: IndexMap<String, RelationshipsIndex> = IndexMap::new();

    for (relationship_id, story) in &mut request.header_entries {
        let Some(Watermark::Picture {
            relationship_id: watermark_relationship_id,
            media_path,
            data_url,
            ..
        }) = story.watermark.as_mut()
        else {
            continue;
        };
        let Some(relationship) = relationships.get(relationship_id) else {
            continue;
        };
        if relationship.target_mode == Some(TargetMode::External) {
            continue;
        }
        let story_path = resolve_relative_path(&package.document_path, &relationship.target)?;
        let relationships_path = relationship_part_path(&story_path);
        let rels = rels_parts
            .entry(relationships_path.clone())
            .or_insert_with(|| {
                RelationshipsIndex::parse(read_rels_or_stub(package, &relationships_path))
            });

        if watermark_relationship_id
            .as_ref()
            .is_some_and(|id| rels.contains_id(id))
        {
            continue;
        }

        let filename = if let Some(path) = media_path.as_deref() {
            path.rsplit('/')
                .next()
                .filter(|name| !name.is_empty())
                .map(str::to_owned)
        } else if let Some(source) = data_url
            .as_deref()
            .filter(|source| source.starts_with("data:"))
        {
            if let Some(filename) = written_media.get(source) {
                Some(filename.clone())
            } else {
                let (bytes, extension) = decode_image_data_url(source)?;
                image_number += 1;
                let filename = format!("image{image_number}.{extension}");
                package.set(format!("word/media/{filename}"), bytes);
                extensions.insert(extension);
                written_media.insert(source.to_owned(), filename.clone());
                Some(filename)
            }
        } else {
            None
        };
        let Some(filename) = filename else { continue };
        let target = format!("media/{filename}");

        if let Some(existing) = rels.id_for_target(&target) {
            *watermark_relationship_id = Some(existing);
            continue;
        }

        let new_relationship_id = rels.next_id();
        if let Some(updated) = rels.try_append_new(
            new_relationship_id.clone(),
            relationship_types::IMAGE,
            target,
            None,
        ) {
            package.set_text(&relationships_path, updated);
            *watermark_relationship_id = Some(new_relationship_id);
        }
    }

    register_image_extensions(package, &extensions);
    Ok(())
}

fn normalize_media_target(target: &str) -> &str {
    target
        .strip_prefix("./")
        .or_else(|| target.strip_prefix('/'))
        .unwrap_or(target)
        .strip_prefix("word/")
        .unwrap_or_else(|| {
            target
                .strip_prefix("./")
                .or_else(|| target.strip_prefix('/'))
                .unwrap_or(target)
        })
}

fn relationship_tags(xml: &str) -> impl Iterator<Item = &str> {
    XmlTagIter::new(xml, "Relationship")
}

struct XmlTagIter<'a> {
    xml: &'a str,
    name: &'static str,
    cursor: usize,
}

impl<'a> XmlTagIter<'a> {
    fn new(xml: &'a str, name: &'static str) -> Self {
        Self {
            xml,
            name,
            cursor: 0,
        }
    }
}

impl<'a> Iterator for XmlTagIter<'a> {
    type Item = &'a str;

    fn next(&mut self) -> Option<Self::Item> {
        let opening = format!("<{}", self.name);
        while let Some(relative) = self.xml[self.cursor..].find(&opening) {
            let start = self.cursor + relative;
            let boundary = self.xml.as_bytes().get(start + opening.len()).copied();
            if !matches!(boundary, Some(b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>')) {
                self.cursor = start + opening.len();
                continue;
            }
            let end = find_tag_end(self.xml.as_bytes(), start)?;
            self.cursor = end + 1;
            return Some(&self.xml[start..=end]);
        }
        None
    }
}

fn xml_attribute<'a>(tag: &'a str, name: &str) -> Option<&'a str> {
    let needle = format!("{name}=\"");
    let start = tag.find(&needle)? + needle.len();
    let end = tag[start..].find('"')? + start;
    Some(&tag[start..end])
}

fn process_new_hyperlinks(
    request: &mut S13SaveRequest,
    relationships: &IndexMap<String, Relationship>,
    package: &mut Package,
) -> Result<(), ParseError> {
    process_hyperlink_part(
        package,
        "word/_rels/document.xml.rels",
        std::iter::once(&mut request.document.content),
    );
    process_hyperlink_part(
        package,
        "word/_rels/footnotes.xml.rels",
        request
            .footnote_separators
            .iter_mut()
            .chain(request.footnotes.iter_mut())
            .map(|note| &mut note.content),
    );
    process_hyperlink_part(
        package,
        "word/_rels/endnotes.xml.rels",
        request
            .endnote_separators
            .iter_mut()
            .chain(request.endnotes.iter_mut())
            .map(|note| &mut note.content),
    );
    for (entries, relationship_type) in [
        (&mut request.header_entries, relationship_types::HEADER),
        (&mut request.footer_entries, relationship_types::FOOTER),
    ] {
        for (relationship_id, story) in entries {
            let Some(relationship) = relationships.get(relationship_id) else {
                continue;
            };
            if relationship.relationship_type != relationship_type
                || relationship.target_mode == Some(TargetMode::External)
            {
                continue;
            }
            process_hyperlink_part(
                package,
                &owner_relationships_path(package, &relationship.target)?,
                std::iter::once(&mut story.content),
            );
        }
    }
    Ok(())
}

fn process_hyperlink_part<'a>(
    package: &mut Package,
    relationships_path: &str,
    stories: impl Iterator<Item = &'a mut Vec<BlockContent>>,
) {
    let mut relationships =
        RelationshipsIndex::parse(read_rels_or_stub(package, relationships_path));
    for blocks in stories {
        visit_hyperlinks(blocks, &mut |hyperlink| {
            // Bookmark anchors resolve inside the owning story and never need
            // an OPC relationship, even though the parser also exposes their
            // convenient `#anchor` form through `href`.
            if hyperlink
                .anchor
                .as_deref()
                .is_some_and(|value| !value.is_empty())
            {
                return;
            }
            let current = hyperlink
                .relationship_id
                .as_deref()
                .and_then(|id| relationships.existing_position(id))
                .map(|position| *relationships.entry(position));
            let Some(href) = hyperlink.href.as_deref() else {
                if current.is_none() {
                    hyperlink.relationship_id = None;
                }
                return;
            };

            if current
                .as_ref()
                .is_some_and(|entry| relationships.target_matches(entry, href))
            {
                return;
            }
            if let Some(existing) = relationships.external_hyperlink_id(href) {
                hyperlink.relationship_id = Some(existing);
                return;
            }

            let new_relationship_id = relationships.next_id();
            relationships.push_new(
                new_relationship_id.clone(),
                relationship_types::HYPERLINK,
                escape_xml(href),
                Some("External"),
            );
            hyperlink.relationship_id = Some(new_relationship_id);
        });
    }
    if let Some(updated) = relationships.serialize() {
        package.set_text(relationships_path, updated);
    }
}

fn block_has_hyperlink(block: &BlockContent) -> bool {
    match block {
        BlockContent::Paragraph(paragraph) => paragraph
            .content
            .iter()
            .any(|content| matches!(content, ParagraphContent::Inline(InlineNode::Hyperlink(_)))),
        BlockContent::Table(table) => table.rows.iter().any(|row| {
            row.cells
                .iter()
                .any(|cell| cell.content.iter().any(block_has_hyperlink))
        }),
        BlockContent::BlockSdt(sdt) => sdt.content.iter().any(block_has_hyperlink),
        BlockContent::RawXml(_) => false,
    }
}

fn visit_hyperlinks(blocks: &mut [BlockContent], visit: &mut impl FnMut(&mut Hyperlink)) {
    for block in blocks {
        if !block_has_hyperlink(block) {
            continue;
        }
        match block {
            BlockContent::Paragraph(paragraph) => {
                for content in &mut Arc::make_mut(paragraph).content {
                    if let ParagraphContent::Inline(InlineNode::Hyperlink(hyperlink)) = content {
                        visit(hyperlink);
                    }
                }
            }
            BlockContent::Table(table) => {
                for row in &mut Arc::make_mut(table).rows {
                    for cell in &mut row.cells {
                        visit_hyperlinks(&mut cell.content, visit);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => visit_hyperlinks(&mut Arc::make_mut(sdt).content, visit),
            BlockContent::RawXml(_) => {}
        }
    }
}

fn decode_xml_entities(value: &str) -> String {
    let mut decoded = String::with_capacity(value.len());
    let mut cursor = 0usize;
    while let Some(relative) = value[cursor..].find('&') {
        let start = cursor + relative;
        decoded.push_str(&value[cursor..start]);
        let Some(relative_end) = value[start..].find(';') else {
            decoded.push_str(&value[start..]);
            return decoded;
        };
        let end = start + relative_end;
        let entity = &value[start + 1..end];
        let replacement = match entity {
            "lt" => Some('<'),
            "gt" => Some('>'),
            "quot" => Some('"'),
            "apos" => Some('\''),
            "amp" => Some('&'),
            value if value.starts_with("#x") => u32::from_str_radix(&value[2..], 16)
                .ok()
                .and_then(char::from_u32),
            value if value.starts_with('#') => value[1..].parse().ok().and_then(char::from_u32),
            _ => None,
        };
        if let Some(replacement) = replacement {
            decoded.push(replacement);
        } else {
            decoded.push_str(&value[start..=end]);
        }
        cursor = end + 1;
    }
    decoded.push_str(&value[cursor..]);
    decoded
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Span {
    start: usize,
    end: usize,
}

#[derive(Clone, Debug, Default)]
struct ParagraphIndex {
    count: usize,
    by_id: HashMap<String, Vec<Span>>,
}

/// Patch only the requested `w14:paraId` paragraphs from a complete serialized
/// document part. Returns `None` whenever the two documents cannot be proven
/// structurally compatible with a selective save.
pub fn build_patched_document_xml(
    original_xml: &str,
    serialized_xml: &str,
    changed_ids: &[String],
) -> Option<String> {
    if changed_ids.is_empty() {
        return Some(original_xml.to_owned());
    }

    let original = index_paragraphs(original_xml)?;
    let serialized = index_paragraphs(serialized_xml)?;
    if original.count != serialized.count {
        return None;
    }

    let mut replacements = Vec::with_capacity(changed_ids.len());
    for id in changed_ids {
        let [original_span] = original.by_id.get(id)?.as_slice() else {
            return None;
        };
        let [serialized_span] = serialized.by_id.get(id)?.as_slice() else {
            return None;
        };
        replacements.push((
            *original_span,
            &serialized_xml[serialized_span.start..serialized_span.end],
        ));
    }
    replacements.sort_unstable_by(|(left, _), (right, _)| right.start.cmp(&left.start));

    let mut patched = original_xml.to_owned();
    for (span, replacement) in replacements {
        patched.replace_range(span.start..span.end, replacement);
    }
    Some(patched)
}

/// Census of every `w:p` a full serialize would emit for this model. The walk
/// mirrors the serializer's emission rules (table-cell fallbacks included) so
/// a count mismatch proves the model no longer lines up with the source part;
/// `allocates_ids` flags generated `wp:docPr/@id` draws that would move the
/// seeded id allocator relative to a whole-document serialize.
#[derive(Default)]
struct SelectiveParagraphIndex<'a> {
    count: usize,
    by_id: HashMap<String, Vec<Paragraph>>,
    changed: HashSet<&'a str>,
    allocates_ids: bool,
}

impl<'a> SelectiveParagraphIndex<'a> {
    fn new(changed_ids: &'a [String]) -> Self {
        Self {
            changed: changed_ids.iter().map(String::as_str).collect(),
            ..Self::default()
        }
    }

    fn story(&mut self, blocks: &[BlockContent]) -> Option<()> {
        for block in blocks {
            match block {
                BlockContent::Paragraph(paragraph) => self.paragraph(paragraph)?,
                BlockContent::Table(table) => self.table(table)?,
                BlockContent::BlockSdt(sdt) => {
                    self.raw_subtree(sdt.properties.raw_properties_xml.as_deref(), "sdtPr")?;
                    self.raw_subtree(sdt.properties.raw_end_properties_xml.as_deref(), "sdtEndPr")?;
                    self.story(&sdt.content)?;
                }
                BlockContent::RawXml(raw) => self.fragment(&raw.xml)?,
            }
        }
        Some(())
    }

    fn table(&mut self, table: &Table) -> Option<()> {
        for row in &table.rows {
            for cell in &row.cells {
                self.cell(cell)?;
            }
        }
        Some(())
    }

    fn cell(&mut self, cell: &TableCell) -> Option<()> {
        // Mirrors `serialize_table_cell`: block SDTs emit nothing and an
        // otherwise empty cell still emits a `<w:p/>` fallback.
        let mut emitted = false;
        for block in &cell.content {
            match block {
                BlockContent::Paragraph(paragraph) => {
                    self.paragraph(paragraph)?;
                    emitted = true;
                }
                BlockContent::Table(table) => {
                    self.table(table)?;
                    emitted = true;
                }
                BlockContent::BlockSdt(_) => {}
                BlockContent::RawXml(raw) => {
                    self.fragment(&raw.xml)?;
                    emitted = true;
                }
            }
        }
        if !emitted {
            self.count += 1;
        }
        Some(())
    }

    fn paragraph(&mut self, paragraph: &Paragraph) -> Option<()> {
        self.count += 1;
        if let Some(id) = emitted_paragraph_id(paragraph)
            && self.changed.contains(id)
        {
            self.by_id
                .entry(id.to_owned())
                .or_default()
                .push(paragraph.clone());
        }
        for content in &paragraph.content {
            match content {
                ParagraphContent::Inline(node) => self.inline(node)?,
                ParagraphContent::Tracked(change) => {
                    if matches!(
                        change.node_type.as_str(),
                        "insertion" | "deletion" | "moveFrom" | "moveTo"
                    ) {
                        for item in &change.content {
                            match item {
                                InlineNode::Run(run) => self.run(run)?,
                                InlineNode::Hyperlink(hyperlink) => self.hyperlink(hyperlink)?,
                                _ => {}
                            }
                        }
                    }
                }
                _ => {}
            }
        }
        Some(())
    }

    fn inline(&mut self, node: &InlineNode) -> Option<()> {
        match node {
            InlineNode::Run(run) => self.run(run),
            InlineNode::Hyperlink(hyperlink) => self.hyperlink(hyperlink),
            InlineNode::BookmarkStart(_) | InlineNode::BookmarkEnd(_) => Some(()),
            InlineNode::SimpleField(field) => {
                for run in &field.content {
                    self.run(run)?;
                }
                Some(())
            }
            InlineNode::ComplexField(field) => {
                for run in &field.field_code {
                    self.run(run)?;
                }
                match field
                    .structured_result
                    .as_ref()
                    .filter(|result| result.blocks.is_none())
                    .and_then(|result| result.inline.as_ref())
                {
                    Some(nodes) => {
                        for node in nodes {
                            self.inline(node)?;
                        }
                    }
                    None => {
                        for run in &field.field_result {
                            self.run(run)?;
                        }
                    }
                }
                Some(())
            }
            InlineNode::InlineSdt(sdt) => {
                self.raw_subtree(sdt.properties.raw_properties_xml.as_deref(), "sdtPr")?;
                self.raw_subtree(sdt.properties.raw_end_properties_xml.as_deref(), "sdtEndPr")?;
                for item in &sdt.content {
                    self.inline(item)?;
                }
                Some(())
            }
            InlineNode::Math(math) => {
                if !math.omml_xml.is_empty() {
                    validate_math_subtree(&math.omml_xml).ok()?;
                    self.count += count_paragraph_elements(&math.omml_xml)?;
                }
                Some(())
            }
            InlineNode::RawXml(raw) => self.fragment(&raw.xml),
        }
    }

    /// Hyperlink serialization only emits Run/BookmarkStart/BookmarkEnd
    /// children; only runs can carry nested paragraphs or generated ids.
    fn hyperlink(&mut self, hyperlink: &Hyperlink) -> Option<()> {
        for child in &hyperlink.children {
            if let InlineNode::Run(run) = child {
                self.run(run)?;
            }
        }
        Some(())
    }

    fn run(&mut self, run: &Run) -> Option<()> {
        for content in &run.content {
            match content {
                RunContent::Drawing { image } => {
                    if image.id.as_deref().is_none_or(str::is_empty) {
                        self.allocates_ids = true;
                    }
                }
                RunContent::Shape { shape } => {
                    if shape.id.as_deref().is_none_or(str::is_empty) {
                        self.allocates_ids = true;
                    }
                    if let Some(text_body) = shape.text_body.as_ref() {
                        for value in &text_body.content {
                            let block: BlockContent = serde_json::from_value(value.clone()).ok()?;
                            self.story(std::slice::from_ref(&block))?;
                        }
                    }
                }
                RunContent::HorizontalRule { rule } => self.fragment(&rule.xml)?,
                RunContent::Chart { chart } => {
                    self.raw_subtree_required(chart.drawing_xml.as_deref(), "drawing")?;
                }
                RunContent::OpaqueDrawing { xml, .. } => self.fragment(xml)?,
                _ => {}
            }
        }
        Some(())
    }

    fn fragment(&mut self, xml: &str) -> Option<()> {
        validate_replayed_fragment(xml).ok()?;
        self.count += count_paragraph_elements(xml)?;
        Some(())
    }

    fn raw_subtree(&mut self, xml: Option<&str>, local_name: &'static str) -> Option<()> {
        let Some(xml) = xml else { return Some(()) };
        validate_raw_subtree(xml, "w", local_name).ok()?;
        self.count += count_paragraph_elements(xml)?;
        Some(())
    }

    fn raw_subtree_required(&mut self, xml: Option<&str>, local_name: &'static str) -> Option<()> {
        validate_raw_subtree(xml?, "w", local_name).ok()?;
        self.count += count_paragraph_elements(xml?)?;
        Some(())
    }
}

/// `w14:paraId` exactly as the paragraph serializer would emit it: the typed
/// id first, then the first replayed attribute of that name.
fn emitted_paragraph_id(paragraph: &Paragraph) -> Option<&str> {
    if let Some(id) = paragraph.para_id.as_deref().filter(|id| !id.is_empty()) {
        return Some(id);
    }
    paragraph
        .extra_attributes
        .iter()
        .find(|attribute| attribute.name == "w14:paraId")
        .map(|attribute| attribute.value.as_str())
}

/// Count `<w:p>` element starts in an already-validated fragment, skipping
/// comments and CDATA like `index_paragraphs` does.
fn count_paragraph_elements(xml: &str) -> Option<usize> {
    let bytes = xml.as_bytes();
    let mut count = 0usize;
    let mut cursor = 0usize;
    while cursor < bytes.len() {
        let Some(relative) = bytes[cursor..].iter().position(|byte| *byte == b'<') else {
            break;
        };
        let start = cursor + relative;
        if bytes[start..].starts_with(b"<!--") {
            cursor = find_bytes(bytes, start + 4, b"-->")? + 3;
            continue;
        }
        if bytes[start..].starts_with(b"<![CDATA[") {
            cursor = find_bytes(bytes, start + 9, b"]]>")? + 3;
            continue;
        }
        let end = find_tag_end(bytes, start)?;
        if is_open_paragraph_tag(&bytes[start..=end]) {
            count += 1;
        }
        cursor = end + 1;
    }
    Some(count)
}

/// Splice only the changed `w14:paraId` paragraphs into the untouched source
/// part. `None` whenever the result cannot be proven identical to the
/// serialize-then-patch path (whose behavior and errors the fallback preserves).
fn build_selective_document_xml(
    document: &DocumentBody,
    original_xml: &str,
    changed_ids: &[String],
    context: &mut SerializerContext,
) -> Result<Option<String>, ParseError> {
    let mut index = SelectiveParagraphIndex::new(changed_ids);
    if index.story(&document.content).is_none() || index.allocates_ids {
        return Ok(None);
    }
    if changed_ids.is_empty() {
        return Ok(Some(original_xml.to_owned()));
    }
    let Some(original) = index_paragraphs(original_xml) else {
        return Ok(None);
    };
    if index.count != original.count {
        return Ok(None);
    }

    let mut replacements = Vec::with_capacity(changed_ids.len());
    for id in changed_ids {
        let Some(&span) = original
            .by_id
            .get(id)
            .filter(|spans| spans.len() == 1)
            .and_then(|spans| spans.first())
        else {
            return Ok(None);
        };
        let Some([paragraph]) = index.by_id.get(id).map(Vec::as_slice) else {
            return Ok(None);
        };
        replacements.push((span, serialize_paragraph(paragraph, context)?));
    }
    replacements.sort_unstable_by(|(left, _), (right, _)| right.start.cmp(&left.start));

    let mut patched = original_xml.to_owned();
    for (span, replacement) in replacements {
        patched.replace_range(span.start..span.end, &replacement);
    }
    Ok(Some(patched))
}

/// Updates direct-child core-property text using a fixed clock.
pub fn update_core_properties(
    core_xml: &str,
    update_modified_date: bool,
    modified_by: Option<&str>,
    now: &str,
) -> String {
    let mut updated = core_xml.to_owned();
    if update_modified_date {
        let value = format!(
            "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{}</dcterms:modified>",
            escape_xml(now)
        );
        updated = replace_text_element(&updated, "dcterms:modified", &value).unwrap_or_else(|| {
            insert_before_closing(&updated, "cp:coreProperties", &value)
                .unwrap_or_else(|| updated.clone())
        });
    }
    if let Some(modified_by) = modified_by.filter(|value| !value.is_empty()) {
        let value = format!(
            "<cp:lastModifiedBy>{}</cp:lastModifiedBy>",
            escape_xml(modified_by)
        );
        updated =
            replace_text_element(&updated, "cp:lastModifiedBy", &value).unwrap_or_else(|| {
                insert_before_closing(&updated, "cp:coreProperties", &value)
                    .unwrap_or_else(|| updated.clone())
            });
    }
    updated
}

fn replace_text_element(xml: &str, name: &str, replacement: &str) -> Option<String> {
    let opening = format!("<{name}");
    let closing = format!("</{name}>");
    let start = xml.find(&opening)?;
    let opening_end = find_tag_end(xml.as_bytes(), start)?;
    let content = &xml[opening_end + 1..];
    let next_tag = content.find('<')?;
    if !content[next_tag..].starts_with(&closing) {
        return None;
    }
    let end = opening_end + 1 + next_tag + closing.len();
    let mut result = String::with_capacity(xml.len() - (end - start) + replacement.len());
    result.push_str(&xml[..start]);
    result.push_str(replacement);
    result.push_str(&xml[end..]);
    Some(result)
}

fn insert_before_closing(xml: &str, name: &str, value: &str) -> Option<String> {
    let closing = format!("</{name}>");
    let offset = xml.find(&closing)?;
    let mut result = String::with_capacity(xml.len() + value.len());
    result.push_str(&xml[..offset]);
    result.push_str(value);
    result.push_str(&xml[offset..]);
    Some(result)
}

fn index_paragraphs(xml: &str) -> Option<ParagraphIndex> {
    let bytes = xml.as_bytes();
    let mut index = ParagraphIndex::default();
    let mut open: Vec<(usize, Option<String>)> = Vec::new();
    let mut cursor = 0usize;

    while cursor < bytes.len() {
        let relative = bytes[cursor..].iter().position(|byte| *byte == b'<')?;
        let start = cursor + relative;
        if bytes[start..].starts_with(b"<!--") {
            cursor = find_bytes(bytes, start + 4, b"-->")? + 3;
            continue;
        }
        if bytes[start..].starts_with(b"<![CDATA[") {
            cursor = find_bytes(bytes, start + 9, b"]]>")? + 3;
            continue;
        }
        let end = find_tag_end(bytes, start)?;
        let tag = &bytes[start..=end];

        if is_open_paragraph_tag(tag) {
            index.count += 1;
            let id = paragraph_id(tag);
            if is_self_closing(tag) {
                if let Some(id) = id {
                    index.by_id.entry(id).or_default().push(Span {
                        start,
                        end: end + 1,
                    });
                }
            } else {
                open.push((start, id));
            }
        } else if is_close_paragraph_tag(tag) {
            let (paragraph_start, id) = open.pop()?;
            if let Some(id) = id {
                index.by_id.entry(id).or_default().push(Span {
                    start: paragraph_start,
                    end: end + 1,
                });
            }
        }
        cursor = end + 1;
    }

    open.is_empty().then_some(index)
}

fn is_open_paragraph_tag(tag: &[u8]) -> bool {
    tag.starts_with(b"<w:p")
        && matches!(
            tag.get(4),
            Some(b'>') | Some(b'/') | Some(b' ' | b'\t' | b'\r' | b'\n')
        )
}

fn is_close_paragraph_tag(tag: &[u8]) -> bool {
    tag == b"</w:p>"
}

fn is_self_closing(tag: &[u8]) -> bool {
    tag[..tag.len().saturating_sub(1)]
        .iter()
        .rev()
        .find(|byte| !byte.is_ascii_whitespace())
        == Some(&b'/')
}

fn paragraph_id(tag: &[u8]) -> Option<String> {
    let mut cursor = 4usize;
    while cursor < tag.len() {
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if matches!(tag.get(cursor), None | Some(b'>' | b'/')) {
            break;
        }
        let name_start = cursor;
        while tag
            .get(cursor)
            .is_some_and(|byte| !byte.is_ascii_whitespace() && !matches!(byte, b'=' | b'>' | b'/'))
        {
            cursor += 1;
        }
        let name = &tag[name_start..cursor];
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            continue;
        }
        cursor += 1;
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        let quote = *tag.get(cursor)?;
        if !matches!(quote, b'\'' | b'"') {
            return None;
        }
        cursor += 1;
        let value_start = cursor;
        while tag.get(cursor) != Some(&quote) {
            cursor += 1;
            if cursor >= tag.len() {
                return None;
            }
        }
        if name == b"w14:paraId" {
            return std::str::from_utf8(&tag[value_start..cursor])
                .ok()
                .map(str::to_owned);
        }
        cursor += 1;
    }
    None
}

fn find_tag_end(bytes: &[u8], start: usize) -> Option<usize> {
    let mut quote = None;
    for (offset, byte) in bytes.get(start + 1..)?.iter().copied().enumerate() {
        match (quote, byte) {
            (None, b'\'' | b'"') => quote = Some(byte),
            (Some(current), byte) if current == byte => quote = None,
            (None, b'>') => return Some(start + 1 + offset),
            _ => {}
        }
    }
    None
}

fn find_bytes(haystack: &[u8], start: usize, needle: &[u8]) -> Option<usize> {
    haystack
        .get(start..)?
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|offset| start + offset)
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn determinism() -> serde_json::Value {
        json!({
            "seed": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
            "now": "2030-01-02T03:04:05.006Z"
        })
    }

    fn text_paragraph(text: &str, para_id: Option<&str>) -> serde_json::Value {
        let mut paragraph = json!({
            "type": "paragraph",
            "content": [{
                "type": "run",
                "content": [{ "type": "text", "text": text }]
            }]
        });
        if let Some(para_id) = para_id {
            paragraph["paraId"] = json!(para_id);
        }
        paragraph
    }

    fn image_paragraph(data_url: &str) -> serde_json::Value {
        json!({
            "type": "paragraph",
            "content": [{
                "type": "run",
                "content": [{
                    "type": "drawing",
                    "image": {
                        "type": "image",
                        "rId": "",
                        "src": data_url,
                        "size": { "width": 9525, "height": 9525 },
                        "wrap": { "type": "inline" }
                    }
                }]
            }]
        })
    }

    fn base_package(document_xml: &str) -> Vec<u8> {
        ooxml_opc::rezip_parts(&[
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
            ),
            (
                "_rels/.rels".to_owned(),
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
            ),
            ("word/document.xml".to_owned(), document_xml.as_bytes().to_vec()),
            (
                "word/_rels/document.xml.rels".to_owned(),
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.to_vec(),
            ),
            (
                "docProps/core.xml".to_owned(),
                br#"<cp:coreProperties><dcterms:modified>past</dcterms:modified></cp:coreProperties>"#.to_vec(),
            ),
            ("word/media/keep.bin".to_owned(), vec![0, 1, 2, 255]),
            ("custom/opaque.dat".to_owned(), b"opaque\0bytes".to_vec()),
        ])
        .expect("base package")
    }

    fn part_map(bytes: &[u8]) -> IndexMap<String, Vec<u8>> {
        ooxml_opc::unzip_parts(bytes)
            .expect("unzip")
            .into_iter()
            .collect()
    }

    #[test]
    fn selective_patch_preserves_every_untouched_byte() {
        let original = concat!(
            "<?xml version=\"1.0\"?><w:document><w:body>",
            "<w:p w14:paraId=\"AAAA\"><w:r><w:t> old A </w:t></w:r></w:p>",
            "<!-- authored spacing -->",
            "<w:p custom=\"x\" w14:paraId='BBBB'><w:r><w:t>old B</w:t></w:r></w:p>",
            "</w:body></w:document>"
        );
        let serialized = concat!(
            "<w:document><w:body>",
            "<w:p w14:paraId=\"AAAA\"><w:r><w:t>new A</w:t></w:r></w:p>",
            "<w:p w14:paraId='BBBB'><w:r><w:t>new B</w:t></w:r></w:p>",
            "</w:body></w:document>"
        );
        let patched = build_patched_document_xml(original, serialized, &["BBBB".to_owned()])
            .expect("safe patch");
        assert!(patched.contains("<w:t> old A </w:t>"));
        assert!(patched.contains("<!-- authored spacing -->"));
        assert!(patched.contains("<w:t>new B</w:t>"));
        let original_b = "<w:p custom=\"x\" w14:paraId='BBBB'><w:r><w:t>old B</w:t></w:r></w:p>";
        let serialized_b = "<w:p w14:paraId='BBBB'><w:r><w:t>new B</w:t></w:r></w:p>";
        assert_eq!(patched, original.replace(original_b, serialized_b));
    }

    #[test]
    fn selective_patch_rejects_duplicates_and_structural_changes() {
        let duplicate = "<w:p w14:paraId=\"A\"/><w:p w14:paraId=\"A\"/>";
        let single = "<w:p w14:paraId=\"A\"/>";
        assert!(build_patched_document_xml(duplicate, duplicate, &["A".to_owned()]).is_none());
        assert!(build_patched_document_xml(single, duplicate, &["A".to_owned()]).is_none());
    }

    #[test]
    fn core_properties_use_fixed_clock_and_escape_modifier() {
        let xml = "<cp:coreProperties><dcterms:modified old=\"1\">past</dcterms:modified></cp:coreProperties>";
        let updated =
            update_core_properties(xml, true, Some("A & <B>\""), "2030-01-02T03:04:05.006Z");
        assert_eq!(
            updated,
            "<cp:coreProperties><dcterms:modified xsi:type=\"dcterms:W3CDTF\">2030-01-02T03:04:05.006Z</dcterms:modified><cp:lastModifiedBy>A &amp; &lt;B&gt;&quot;</cp:lastModifiedBy></cp:coreProperties>"
        );
    }

    #[test]
    fn malformed_core_text_does_not_scan_across_markup() {
        let xml =
            "<cp:coreProperties><dcterms:modified><bad/></dcterms:modified></cp:coreProperties>";
        let updated = update_core_properties(xml, true, None, "2030-01-02T03:04:05.006Z");
        assert!(updated.contains("<dcterms:modified><bad/></dcterms:modified>"));
        assert!(updated.contains("2030-01-02T03:04:05.006Z"));
    }

    #[test]
    fn package_save_reuses_container_and_preserves_unowned_parts() {
        let original = base_package(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>old</w:t></w:r></w:p></w:body></w:document>",
        );
        let request: S13SaveRequest = serde_json::from_value(json!({
            "determinism": determinism(),
            "document": { "content": [text_paragraph("new", None)] },
            "options": { "updateModifiedDate": true }
        }))
        .expect("request");
        let saved = write_docx_s13(request, &original).expect("save");
        let before = part_map(&original);
        let after = part_map(&saved);

        assert_eq!(
            before.keys().collect::<Vec<_>>(),
            after.keys().collect::<Vec<_>>()
        );
        assert_eq!(after["custom/opaque.dat"], before["custom/opaque.dat"]);
        assert_eq!(after["word/media/keep.bin"], before["word/media/keep.bin"]);
        assert!(String::from_utf8_lossy(&after["word/document.xml"]).contains("<w:t>new</w:t>"));
        assert!(
            String::from_utf8_lossy(&after["docProps/core.xml"])
                .contains("2030-01-02T03:04:05.006Z")
        );
    }

    #[test]
    fn selective_package_save_keeps_unchanged_document_spans_exact() {
        let original_document = concat!(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"><w:body>",
            "<w:p w14:paraId=\"AAAAAAAA\"><w:r><w:t xml:space=\"preserve\"> keep me </w:t></w:r></w:p>",
            "<!-- opaque authored gap -->",
            "<w:p w14:paraId=\"BBBBBBBB\"><w:r><w:t>old</w:t></w:r></w:p>",
            "</w:body></w:document>"
        );
        let original = base_package(original_document);
        let request: S13SaveRequest = serde_json::from_value(json!({
            "determinism": determinism(),
            "document": {
                "content": [
                    text_paragraph("model copy", Some("AAAAAAAA")),
                    text_paragraph("edited", Some("BBBBBBBB"))
                ]
            },
            "options": { "updateModifiedDate": false },
            "selective": { "changedParaIds": ["BBBBBBBB"] }
        }))
        .expect("request");
        let saved = write_docx_s13(request, &original).expect("selective save");
        let parts = part_map(&saved);
        let document = String::from_utf8(parts["word/document.xml"].clone()).unwrap();
        assert!(document.contains("<w:t xml:space=\"preserve\"> keep me </w:t>"));
        assert!(document.contains("<!-- opaque authored gap -->"));
        assert!(document.contains("<w:t>edited</w:t>"));
        assert!(!document.contains("model copy"));
        assert_eq!(parts["custom/opaque.dat"], b"opaque\0bytes");
    }

    #[test]
    fn bookmark_anchor_does_not_create_external_relationship() {
        let original = base_package(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>",
        );
        let request: S13SaveRequest = serde_json::from_value(json!({
            "determinism": determinism(),
            "document": { "content": [{
                "type": "paragraph",
                "content": [{
                    "type": "hyperlink",
                    "href": "#inside",
                    "anchor": "inside",
                    "children": [{
                        "type": "run",
                        "content": [{ "type": "text", "text": "jump" }]
                    }]
                }]
            }] },
            "options": { "updateModifiedDate": false }
        }))
        .expect("request");
        let saved = write_docx_s13(request, &original).expect("save");
        let parts = part_map(&saved);
        let document = String::from_utf8_lossy(&parts["word/document.xml"]);
        let relationships = String::from_utf8_lossy(&parts["word/_rels/document.xml.rels"]);

        assert!(document.contains("w:anchor=\"inside\""));
        assert!(!relationships.contains(relationship_types::HYPERLINK));
    }

    #[test]
    fn package_ids_and_media_names_are_scoped_across_body_and_headers() {
        let original = base_package(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>",
        );
        let request: S13SaveRequest = serde_json::from_value(json!({
            "determinism": determinism(),
            "document": {
                "content": [image_paragraph("data:image/png;base64,AQID")]
            },
            "headerEntries": [["rIdHeader", {
                "type": "header",
                "hdrFtrType": "default",
                "content": [image_paragraph("data:image/png;base64,BAUG")]
            }]],
            "relationshipEntries": [["rIdHeader", {
                "id": "rIdHeader",
                "type": relationship_types::HEADER,
                "target": "header1.xml"
            }]],
            "options": { "updateModifiedDate": false }
        }))
        .expect("request");
        let saved = write_docx_s13(request, &original).expect("save");
        let parts = part_map(&saved);

        assert_eq!(parts["word/media/image1.png"], [1, 2, 3]);
        assert_eq!(parts["word/media/image2.png"], [4, 5, 6]);
        let document = String::from_utf8_lossy(&parts["word/document.xml"]);
        let header = String::from_utf8_lossy(&parts["word/header1.xml"]);
        let document_id = XmlTagIter::new(&document, "wp:docPr")
            .next()
            .and_then(|tag| xml_attribute(tag, "id"))
            .unwrap();
        let header_id = XmlTagIter::new(&header, "wp:docPr")
            .next()
            .and_then(|tag| xml_attribute(tag, "id"))
            .unwrap();
        assert_ne!(document_id, header_id);
        assert!(
            String::from_utf8_lossy(&parts["[Content_Types].xml"])
                .contains("PartName=\"/word/header1.xml\"")
        );
        assert!(
            String::from_utf8_lossy(&parts["word/_rels/document.xml.rels"])
                .contains("Id=\"rIdHeader\"")
        );
    }

    #[test]
    fn multiple_new_relationships_in_one_save_get_unique_ids() {
        let original = base_package(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>",
        );
        let request: S13SaveRequest = serde_json::from_value(json!({
            "determinism": determinism(),
            "document": { "content": [{
                "type": "paragraph",
                "content": [
                    {
                        "type": "hyperlink",
                        "href": "https://a.example",
                        "children": [{
                            "type": "run",
                            "content": [{ "type": "text", "text": "first" }]
                        }]
                    },
                    {
                        "type": "hyperlink",
                        "href": "https://b.example",
                        "children": [{
                            "type": "run",
                            "content": [{ "type": "text", "text": "second" }]
                        }]
                    }
                ]
            }, image_paragraph("data:image/png;base64,AQID"), image_paragraph("data:image/png;base64,BAUG")] },
            "options": { "updateModifiedDate": false }
        }))
        .expect("request");
        let saved = write_docx_s13(request, &original).expect("save");
        let parts = part_map(&saved);
        let document = String::from_utf8_lossy(&parts["word/document.xml"]);
        let relationships = String::from_utf8_lossy(&parts["word/_rels/document.xml.rels"]);

        let hyperlink_ids: Vec<String> = XmlTagIter::new(&document, "w:hyperlink")
            .filter_map(|tag| xml_attribute(tag, "r:id").map(str::to_owned))
            .collect();
        assert_eq!(hyperlink_ids.len(), 2);
        assert_ne!(hyperlink_ids[0], hyperlink_ids[1]);
        let media_ids: Vec<String> = XmlTagIter::new(&document, "a:blip")
            .filter_map(|tag| xml_attribute(tag, "r:embed").map(str::to_owned))
            .collect();
        assert_eq!(media_ids.len(), 2);
        assert_ne!(media_ids[0], media_ids[1]);

        let declared: Vec<&str> = relationship_tags(&relationships)
            .filter_map(|tag| xml_attribute(tag, "Id"))
            .collect();
        for relationship_id in hyperlink_ids.iter().chain(&media_ids) {
            assert!(
                declared.iter().any(|id| id == relationship_id),
                "missing relationship {relationship_id}"
            );
        }
        let mut unique = declared.clone();
        unique.sort_unstable();
        unique.dedup();
        assert_eq!(declared.len(), unique.len());
        assert!(relationships.contains("Target=\"https://a.example\""));
        assert!(relationships.contains("Target=\"https://b.example\""));
        assert_eq!(parts["word/media/image1.png"], [1, 2, 3]);
        assert_eq!(parts["word/media/image2.png"], [4, 5, 6]);
    }

    #[test]
    fn shared_header_rels_dedupe_watermark_media_relationship() {
        let original = base_package(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>",
        );
        let watermark = json!({
            "kind": "picture",
            "dataUrl": "data:image/png;base64,AQID",
            "scale": 1.0,
            "washout": true
        });
        let header_entry = |rel: &str, kind: &str| {
            json!([rel, {
                "type": "header",
                "hdrFtrType": kind,
                "content": [],
                "watermark": watermark.clone()
            }])
        };
        let request: S13SaveRequest = serde_json::from_value(json!({
            "determinism": determinism(),
            "document": { "content": [text_paragraph("x", None)] },
            "headerEntries": [
                header_entry("rIdHeaderA", "default"),
                header_entry("rIdHeaderB", "first"),
            ],
            "relationshipEntries": [
                ["rIdHeaderA", {
                    "id": "rIdHeaderA",
                    "type": relationship_types::HEADER,
                    "target": "header1.xml"
                }],
                ["rIdHeaderB", {
                    "id": "rIdHeaderB",
                    "type": relationship_types::HEADER,
                    "target": "header1.xml"
                }],
            ],
            "options": { "updateModifiedDate": false }
        }))
        .expect("request");
        let saved = write_docx_s13(request, &original).expect("save");
        let parts = part_map(&saved);

        let rels = String::from_utf8_lossy(&parts["word/_rels/header1.xml.rels"]);
        let media_rels: Vec<&str> = relationship_tags(&rels)
            .filter(|tag| xml_attribute(tag, "Target") == Some("media/image1.png"))
            .collect();
        assert_eq!(media_rels.len(), 1, "{rels}");
        assert_eq!(parts["word/media/image1.png"], [1, 2, 3]);
    }
}
