//! S13 package-save primitives.
//!
//! Selective document updates operate on exact UTF-8 spans. They never parse
//! and re-emit unchanged markup, so every byte outside an explicitly changed
//! paragraph remains authored exactly as it appeared in the source package.

use std::collections::{HashMap, HashSet};

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
use crate::paragraph::ParagraphContent;
use crate::relationships::{Relationship, relationship_types};
use crate::vml::Watermark;
use crate::xml::ParseError;

use super::context::SerializerContext;
use super::numbering::serialize_numbering_xml;
use super::parts::{
    serialize_comments_extended_part, serialize_comments_extensible_part,
    serialize_comments_ids_part, serialize_comments_with_info, serialize_document_part,
    serialize_endnotes_part, serialize_footnotes_part, serialize_header_footer_part,
};
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
}

fn default_true() -> bool {
    true
}

/// Assemble a complete DOCX package and deflate it through the existing
/// `ooxml-opc` writer. Original entries seed the output in archive order;
/// only editor-owned parts are overwritten or appended.
pub fn write_docx_s13(
    mut request: S13SaveRequest,
    original_docx: &[u8],
) -> Result<Vec<u8>, ParseError> {
    request.determinism.validate()?;
    let original_parts =
        ooxml_opc::unzip_parts(original_docx).map_err(ParseError::Container)?;
    let mut package = Package::new(original_parts);
    let relationships: IndexMap<_, _> = request.relationship_entries.iter().cloned().collect();

    if request.selective.is_some() {
        validate_selective_header_footer_parts(&package, &relationships)?;
    } else {
        process_new_images(&mut request, &relationships, &mut package)?;
        process_new_watermark_images(&mut request, &relationships, &mut package)?;
        process_new_hyperlinks(&mut request, &relationships, &mut package);
    }

    let mut context = SerializerContext::new(&request.determinism)?;
    let serialized_document = serialize_document_part(&request.document, &mut context)?;
    let document_xml = if let Some(selective) = request.selective.as_ref() {
        let original = package
            .text("word/document.xml")
            .ok_or_else(|| save_error("selective save has no word/document.xml"))?;
        build_patched_document_xml(&original, &serialized_document, &selective.changed_para_ids)
            .ok_or_else(|| save_error("selective document patch is unsafe"))?
    } else {
        serialized_document
    };
    package.set_text("word/document.xml", document_xml);

    serialize_header_footer_parts(
        &request.header_entries,
        &request.footer_entries,
        &relationships,
        &mut package,
        &mut context,
    )?;

    if request.selective.is_none() {
        ensure_header_footer_parts(&relationships, &mut package);
        ensure_numbering_part(request.numbering.as_ref(), &mut package);
    }

    serialize_comment_parts(&request.document, &mut package, &mut context);

    if request.selective.is_none() {
        let mut footnotes = request.footnote_separators;
        footnotes.extend(request.footnotes);
        if !footnotes.is_empty() {
            package.set_text(
                "word/footnotes.xml",
                serialize_footnotes_part(&footnotes, &mut context)?,
            );
        }
        let mut endnotes = request.endnote_separators;
        endnotes.extend(request.endnotes);
        if !endnotes.is_empty() {
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

    ooxml_opc::rezip_parts(&package.parts).map_err(ParseError::Container)
}

#[derive(Clone, Debug)]
struct Package {
    parts: Vec<(String, Vec<u8>)>,
    positions: HashMap<String, usize>,
}

impl Package {
    fn new(parts: Vec<(String, Vec<u8>)>) -> Self {
        let positions = parts
            .iter()
            .enumerate()
            .map(|(index, (path, _))| (path.clone(), index))
            .collect();
        Self { parts, positions }
    }

    fn contains(&self, path: &str) -> bool {
        self.positions.contains_key(path)
    }

    fn bytes(&self, path: &str) -> Option<&[u8]> {
        self.positions
            .get(path)
            .map(|index| self.parts[*index].1.as_slice())
    }

    fn text(&self, path: &str) -> Option<String> {
        self.bytes(path)
            .map(|bytes| String::from_utf8_lossy(bytes).into_owned())
    }

    fn set(&mut self, path: impl Into<String>, bytes: Vec<u8>) {
        let path = path.into();
        if let Some(index) = self.positions.get(&path).copied() {
            self.parts[index].1 = bytes;
        } else {
            self.positions.insert(path.clone(), self.parts.len());
            self.parts.push((path, bytes));
        }
    }

    fn set_text(&mut self, path: impl Into<String>, xml: String) {
        self.set(path, xml.into_bytes());
    }
}

fn save_error(message: impl Into<String>) -> ParseError {
    ParseError::Canonical(format!("S13 package save: {}", message.into()))
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
        let path = header_footer_filename(&relationship.target);
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
            if relationship.relationship_type != relationship_type || relationship.target.is_empty()
            {
                continue;
            }
            package.set_text(
                header_footer_filename(&relationship.target),
                serialize_header_footer_part(story, context)?,
            );
        }
    }
    Ok(())
}

fn ensure_header_footer_parts(
    relationships: &IndexMap<String, Relationship>,
    package: &mut Package,
) {
    let parts: Vec<_> = relationships
        .iter()
        .filter_map(|(relationship_id, relationship)| {
            let content_type = match relationship.relationship_type.as_str() {
                relationship_types::HEADER => HEADER_CONTENT_TYPE,
                relationship_types::FOOTER => FOOTER_CONTENT_TYPE,
                _ => return None,
            };
            let target = relationship
                .target
                .trim_start_matches('/')
                .strip_prefix("word/")
                .unwrap_or(relationship.target.trim_start_matches('/'))
                .to_owned();
            Some((
                relationship_id.as_str(),
                relationship.relationship_type.as_str(),
                target,
                content_type,
            ))
        })
        .collect();
    if parts.is_empty() {
        return;
    }

    if let Some(mut content_types) = package.text("[Content_Types].xml") {
        let mut changed = false;
        for (_, _, target, content_type) in &parts {
            let part_name = format!("/word/{target}");
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
    let mut relationships_xml = read_rels_or_stub(package, path);
    let mut changed = false;
    for (relationship_id, relationship_type, target, _) in parts {
        if relationships_xml.contains(&format!("Id=\"{relationship_id}\"")) {
            continue;
        }
        let entry = format!(
            "<Relationship Id=\"{relationship_id}\" Type=\"{relationship_type}\" Target=\"{target}\"/>"
        );
        if let Some(updated) = append_before(&relationships_xml, "</Relationships>", &entry) {
            relationships_xml = updated;
            changed = true;
        }
    }
    if changed {
        package.set_text(path, relationships_xml);
    }
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
    let relationships_xml = read_rels_or_stub(package, path);
    if !relationships_xml.contains("Target=\"numbering.xml\"") {
        let relationship_id = format!("rId{}", find_max_relationship_id(&relationships_xml) + 1);
        let entry = format!(
            "<Relationship Id=\"{relationship_id}\" Type=\"{}\" Target=\"numbering.xml\"/>",
            relationship_types::NUMBERING
        );
        if let Some(updated) = append_before(&relationships_xml, "</Relationships>", &entry) {
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
    let Some(mut relationships_xml) = package.text(path) else {
        return;
    };
    let mut changed = false;
    for (_, _, target, relationship_type) in parts {
        if relationships_xml.contains(target) {
            continue;
        }
        let relationship_id = format!("rId{}", find_max_relationship_id(&relationships_xml) + 1);
        let entry = format!(
            "<Relationship Id=\"{relationship_id}\" Type=\"{relationship_type}\" Target=\"{target}\"/>"
        );
        if let Some(updated) = append_before(&relationships_xml, "</Relationships>", &entry) {
            relationships_xml = updated;
            changed = true;
        }
    }
    if changed {
        package.set_text(path, relationships_xml);
    }
}

fn header_footer_filename(target: &str) -> String {
    if target.starts_with('/') {
        target.trim_start_matches('/').to_owned()
    } else {
        format!("word/{target}")
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
            if relationship.relationship_type != relationship_type {
                continue;
            }
            process_image_part(
                package,
                &owner_relationships_path(&relationship.target),
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
    let relationships_xml = read_rels_or_stub(package, relationships_path);
    let mut relationship_id = find_max_relationship_id(&relationships_xml);
    let mut entries = Vec::new();
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
            relationship_id += 1;
            let filename = format!("image{image_number}.{extension}");
            let new_relationship_id = format!("rId{relationship_id}");
            package.set(format!("word/media/{filename}"), bytes);
            entries.push(format!(
                "<Relationship Id=\"{new_relationship_id}\" Type=\"{}\" Target=\"media/{filename}\"/>",
                relationship_types::IMAGE
            ));
            extensions.insert(extension);
            image.relationship_id = new_relationship_id;
            Ok(())
        })?;
    }
    if !entries.is_empty()
        && let Some(updated) =
            append_before(&relationships_xml, "</Relationships>", &entries.concat())
    {
        package.set_text(relationships_path, updated);
    }
    Ok(())
}

fn visit_new_images(
    blocks: &mut [BlockContent],
    visit: &mut impl FnMut(&mut Image) -> Result<(), ParseError>,
) -> Result<(), ParseError> {
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => {
                for content in &mut paragraph.content {
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
                for row in &mut table.rows {
                    for cell in &mut row.cells {
                        visit_new_images(&mut cell.content, visit)?;
                    }
                }
            }
            // Pinned incumbent behavior: newly inserted images inside a block
            // SDT do not take the full-repack fast path yet.
            BlockContent::BlockSdt(_) => {}
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
        .parts
        .iter()
        .filter_map(|(path, _)| {
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
    // TypeScript preserves discovery order through Set. Sorting here would
    // perturb canonical child order, so recover package discovery order from
    // newly appended image parts.
    let mut ordered = Vec::new();
    for (path, _) in &package.parts {
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

fn owner_relationships_path(target: &str) -> String {
    let path = header_footer_filename(target);
    let filename = path.strip_prefix("word/").unwrap_or(&path).to_owned();
    format!("word/_rels/{filename}.rels")
}

fn process_new_watermark_images(
    request: &mut S13SaveRequest,
    relationships: &IndexMap<String, Relationship>,
    package: &mut Package,
) -> Result<(), ParseError> {
    let mut image_number = find_max_image_number(package);
    let mut extensions = HashSet::new();
    let mut written_media = HashMap::<String, String>::new();

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
        let relationships_path = owner_relationships_path(&relationship.target);
        let relationships_xml = read_rels_or_stub(package, &relationships_path);

        if watermark_relationship_id
            .as_ref()
            .is_some_and(|id| relationship_element_for_id(&relationships_xml, id).is_some())
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

        if let Some(existing) = relationship_tags(&relationships_xml).find_map(|tag| {
            let candidate = xml_attribute(tag, "Target")?;
            (normalize_media_target(candidate) == normalize_media_target(&target))
                .then(|| xml_attribute(tag, "Id").map(str::to_owned))
                .flatten()
        }) {
            *watermark_relationship_id = Some(existing);
            continue;
        }

        let new_relationship_id =
            format!("rId{}", find_max_relationship_id(&relationships_xml) + 1);
        let entry = format!(
            "<Relationship Id=\"{new_relationship_id}\" Type=\"{}\" Target=\"{target}\"/>",
            relationship_types::IMAGE
        );
        if let Some(updated) = append_before(&relationships_xml, "</Relationships>", &entry) {
            package.set_text(relationships_path, updated);
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

fn relationship_element_for_id<'a>(xml: &'a str, relationship_id: &str) -> Option<&'a str> {
    relationship_tags(xml).find(|tag| xml_attribute(tag, "Id") == Some(relationship_id))
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
) {
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
            if relationship.relationship_type != relationship_type {
                continue;
            }
            process_hyperlink_part(
                package,
                &owner_relationships_path(&relationship.target),
                std::iter::once(&mut story.content),
            );
        }
    }
}

fn process_hyperlink_part<'a>(
    package: &mut Package,
    relationships_path: &str,
    stories: impl Iterator<Item = &'a mut Vec<BlockContent>>,
) {
    let relationships_xml = read_rels_or_stub(package, relationships_path);
    let mut relationship_id = find_max_relationship_id(&relationships_xml);
    let mut entries = Vec::new();
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
                .and_then(|id| relationship_element_for_id(&relationships_xml, id));
            let Some(href) = hyperlink.href.as_deref() else {
                if current.is_none() {
                    hyperlink.relationship_id = None;
                }
                return;
            };

            if relationship_targets_href(current, href)
                || current.is_some_and(|tag| {
                    !relationship_is_external_hyperlink(tag)
                        && relationship_target_matches(Some(tag), href)
                })
            {
                return;
            }
            if let Some(existing) = relationship_tags(&relationships_xml).find_map(|tag| {
                relationship_targets_href(Some(tag), href)
                    .then(|| xml_attribute(tag, "Id").map(str::to_owned))
                    .flatten()
            }) {
                hyperlink.relationship_id = Some(existing);
                return;
            }

            relationship_id += 1;
            let new_relationship_id = format!("rId{relationship_id}");
            entries.push(format!(
                "<Relationship Id=\"{new_relationship_id}\" Type=\"{}\" Target=\"{}\" TargetMode=\"External\"/>",
                relationship_types::HYPERLINK,
                escape_xml(href)
            ));
            hyperlink.relationship_id = Some(new_relationship_id);
        });
    }
    if !entries.is_empty()
        && let Some(updated) =
            append_before(&relationships_xml, "</Relationships>", &entries.concat())
    {
        package.set_text(relationships_path, updated);
    }
}

fn visit_hyperlinks(blocks: &mut [BlockContent], visit: &mut impl FnMut(&mut Hyperlink)) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => {
                for content in &mut paragraph.content {
                    if let ParagraphContent::Inline(InlineNode::Hyperlink(hyperlink)) = content {
                        visit(hyperlink);
                    }
                }
            }
            BlockContent::Table(table) => {
                for row in &mut table.rows {
                    for cell in &mut row.cells {
                        visit_hyperlinks(&mut cell.content, visit);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => visit_hyperlinks(&mut sdt.content, visit),
        }
    }
}

fn relationship_is_external_hyperlink(tag: &str) -> bool {
    xml_attribute(tag, "Type") == Some(relationship_types::HYPERLINK)
        && xml_attribute(tag, "TargetMode") == Some("External")
}

fn relationship_target_matches(tag: Option<&str>, href: &str) -> bool {
    tag.and_then(|tag| xml_attribute(tag, "Target"))
        .is_some_and(|target| decode_xml_entities(target) == href)
}

fn relationship_targets_href(tag: Option<&str>, href: &str) -> bool {
    tag.is_some_and(relationship_is_external_hyperlink) && relationship_target_matches(tag, href)
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

/// Update core properties using the caller's fixed clock. The scan is linear
/// and intentionally mirrors the incumbent's narrow direct-child text update:
/// malformed/nested content is not consumed across a later opening tag.
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
}
