//! S8 complete content projection: sections, headers/footers, notes, comments,
//! and recursively dispatched text-box stories.

use base64::Engine as _;
use indexmap::IndexMap;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::block::StoryParser;
use crate::canonical::{canonical_sha256, from_serializable, to_canonical_bytes};
use crate::chart::parse_chart_parts;
use crate::comments::{Comment, parse_comments, remove_orphan_comment_ranges};
use crate::document::{DocumentBody, extract_all_template_variables, parse_document_body};
use crate::header_footer::{HeaderFooter, parse_related_header_footers};
use crate::media::build_media_map;
use crate::notes::{Note, parse_notes};
use crate::numbering::parse_numbering;
use crate::paragraph::HexIdAllocator;
use crate::relationships::{RelationshipMap, parse_relationships};
use crate::settings::parse_settings;
use crate::smart_art::create_smart_art_context;
use crate::styles::{StyleMap, parse_style_definitions};
use crate::theme::{apply_theme_font_lang, parse_theme};
use crate::xml::{ParseBudget, ParseError, ParseLimits, parse_xml};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S8Projection {
    pub body: DocumentBody,
    pub header_entries: Vec<(String, HeaderFooter)>,
    pub footer_entries: Vec<(String, HeaderFooter)>,
    pub footnotes: Vec<Note>,
    pub endnotes: Vec<Note>,
    pub footnote_separators: Vec<Note>,
    pub endnote_separators: Vec<Note>,
    pub template_variables: Vec<String>,
    pub smart_art_warnings: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S8WireEnvelope {
    pub wire_version: u8,
    pub projection: S8Projection,
    pub canonical_base64: String,
    pub canonical_sha256: String,
}

pub fn parse_docx_s8_projection(data: &[u8]) -> Result<S8Projection, ParseError> {
    let parts = ooxml_opc::unzip_parts(data).map_err(ParseError::Container)?;
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let settings = parse_settings(
        find_part(&parts, "word/settings.xml").map(|(_, bytes)| bytes),
        "word/settings.xml",
        &mut budget,
    )?;
    let mut theme = parse_theme(
        find_part(&parts, "word/theme/theme1.xml").map(|(_, bytes)| bytes),
        "word/theme/theme1.xml",
        &mut budget,
    )?;
    apply_theme_font_lang(&mut theme, settings.theme_font_lang.as_ref());
    let style_definitions = find_part(&parts, "word/styles.xml")
        .map(|(path, xml)| parse_style_definitions(xml, Some(&theme), path, &mut budget))
        .transpose()?;
    let styles: StyleMap = style_definitions
        .as_ref()
        .map(|definitions| {
            definitions
                .styles
                .iter()
                .map(|style| (style.style_id.clone(), style.clone()))
                .collect()
        })
        .unwrap_or_default();
    let doc_defaults = style_definitions
        .as_ref()
        .and_then(|definitions| definitions.doc_defaults.as_ref());
    let numbering = parse_numbering(
        find_part(&parts, "word/numbering.xml").map(|(_, bytes)| bytes),
        "word/numbering.xml",
        &mut budget,
    )?;
    let document_relationships = match find_part(&parts, "word/_rels/document.xml.rels") {
        Some((path, xml)) => parse_relationships(xml, path, &mut budget)?,
        None => RelationshipMap::new(),
    };
    let media = build_media_map(&parts);
    let all_xml: IndexMap<_, _> = parts
        .iter()
        .filter(|(path, _)| {
            let lower = path.to_ascii_lowercase();
            lower.ends_with(".xml") || lower.ends_with(".rels")
        })
        .cloned()
        .collect();
    let charts = parse_chart_parts(&all_xml, &mut budget)?;
    let mut smart_art = create_smart_art_context(&all_xml);
    let digest = format!("{:x}", Sha256::digest(data));
    let mut ids = HexIdAllocator::from_sha256(&digest)?;

    let mut body = match find_part(&parts, "word/document.xml") {
        Some((path, xml)) => {
            let document = parse_xml(xml, path, &mut budget)?;
            match document.root() {
                Some(root) => {
                    let mut parser = StoryParser {
                        relationships: Some(&document_relationships),
                        theme: Some(&theme),
                        styles: Some(&styles),
                        doc_defaults,
                        numbering: Some(&numbering),
                        media: &media,
                        charts: &charts,
                        smart_art: &mut smart_art,
                        budget: &mut budget,
                        ids: &mut ids,
                        part: path,
                    };
                    parse_document_body(root, &mut parser)?
                }
                None => DocumentBody::default(),
            }
        }
        None => DocumentBody::default(),
    };

    let (mut headers, mut footers) = parse_related_header_footers(
        &parts,
        &document_relationships,
        Some(&theme),
        Some(&styles),
        doc_defaults,
        Some(&numbering),
        &media,
        &charts,
        &mut smart_art,
        &mut budget,
        &mut ids,
    )?;

    let all_footnotes = parse_note_part(
        &parts,
        "word/footnotes.xml",
        true,
        &document_relationships,
        Some(&theme),
        Some(&styles),
        doc_defaults,
        Some(&numbering),
        &media,
        &charts,
        &mut smart_art,
        &mut budget,
        &mut ids,
    )?;
    let all_endnotes = parse_note_part(
        &parts,
        "word/endnotes.xml",
        false,
        &document_relationships,
        Some(&theme),
        Some(&styles),
        doc_defaults,
        Some(&numbering),
        &media,
        &charts,
        &mut smart_art,
        &mut budget,
        &mut ids,
    )?;
    let (mut footnotes, mut footnote_separators) = partition_notes(all_footnotes);
    let (mut endnotes, mut endnote_separators) = partition_notes(all_endnotes);

    let comments = parse_comment_part(
        &parts,
        &document_relationships,
        Some(&theme),
        Some(&styles),
        doc_defaults,
        &media,
        &charts,
        &mut smart_art,
        &mut budget,
        &mut ids,
    )?;
    let comment_ids: Vec<_> = comments.iter().map(|comment| comment.id).collect();
    if !comments.is_empty() {
        body.comments = Some(comments);
    }

    remove_orphan_comment_ranges(&mut body.content, &comment_ids);
    if let Some(sections) = &mut body.sections {
        for section in sections {
            remove_orphan_comment_ranges(&mut section.content, &comment_ids);
        }
    }
    for story in headers.values_mut().chain(footers.values_mut()) {
        remove_orphan_comment_ranges(&mut story.content, &comment_ids);
    }
    for note in footnotes
        .iter_mut()
        .chain(endnotes.iter_mut())
        .chain(footnote_separators.iter_mut())
        .chain(endnote_separators.iter_mut())
    {
        remove_orphan_comment_ranges(&mut note.content, &comment_ids);
    }

    let template_variables = extract_all_template_variables(&body.content);
    Ok(S8Projection {
        body,
        header_entries: headers.into_iter().collect(),
        footer_entries: footers.into_iter().collect(),
        footnotes,
        endnotes,
        footnote_separators,
        endnote_separators,
        template_variables,
        smart_art_warnings: smart_art.warnings,
    })
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn parse_note_part(
    parts: &[(String, Vec<u8>)],
    owner_path: &str,
    footnotes: bool,
    document_relationships: &RelationshipMap,
    theme: Option<&crate::theme::Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&crate::styles::DocDefaults>,
    numbering: Option<&crate::numbering::NumberingMap>,
    media: &crate::media::MediaMap,
    charts: &crate::chart::ChartPartsMap,
    smart_art: &mut crate::smart_art::SmartArtContext,
    budget: &mut ParseBudget<'_>,
    ids: &mut HexIdAllocator,
) -> Result<Vec<Note>, ParseError> {
    let Some((path, xml)) = find_part(parts, owner_path) else {
        return Ok(Vec::new());
    };
    let relationship_path = relationship_part_path(path);
    let part_relationships = find_part(parts, &relationship_path)
        .map(|(relationship_path, xml)| parse_relationships(xml, relationship_path, budget))
        .transpose()?;
    let relationships = part_relationships
        .as_ref()
        .unwrap_or(document_relationships);
    let document = parse_xml(xml, path, budget)?;
    let Some(root) = document.root() else {
        return Ok(Vec::new());
    };
    let mut parser = StoryParser {
        relationships: Some(relationships),
        theme,
        styles,
        doc_defaults,
        numbering,
        media,
        charts,
        smart_art,
        budget,
        ids,
        part: path,
    };
    parse_notes(root, footnotes, &mut parser)
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn parse_comment_part(
    parts: &[(String, Vec<u8>)],
    document_relationships: &RelationshipMap,
    theme: Option<&crate::theme::Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&crate::styles::DocDefaults>,
    media: &crate::media::MediaMap,
    charts: &crate::chart::ChartPartsMap,
    smart_art: &mut crate::smart_art::SmartArtContext,
    budget: &mut ParseBudget<'_>,
    ids: &mut HexIdAllocator,
) -> Result<Vec<Comment>, ParseError> {
    let Some((path, xml)) = find_part(parts, "word/comments.xml") else {
        return Ok(Vec::new());
    };
    let relationship_path = relationship_part_path(path);
    let part_relationships = find_part(parts, &relationship_path)
        .map(|(relationship_path, xml)| parse_relationships(xml, relationship_path, budget))
        .transpose()?;
    let relationships = part_relationships
        .as_ref()
        .unwrap_or(document_relationships);
    let document = parse_xml(xml, path, budget)?;
    let Some(root) = document.root() else {
        return Ok(Vec::new());
    };
    let mut parser = StoryParser {
        relationships: Some(relationships),
        theme,
        styles,
        doc_defaults,
        // Incumbent comment parsing deliberately does not resolve numbering
        // even though other story owners do.
        numbering: None,
        media,
        charts,
        smart_art,
        budget,
        ids,
        part: path,
    };
    parse_comments(
        root,
        find_part(parts, "word/commentsExtensible.xml").map(|(_, bytes)| bytes),
        find_part(parts, "word/commentsExtended.xml").map(|(_, bytes)| bytes),
        &mut parser,
    )
}

pub(crate) fn partition_notes(notes: Vec<Note>) -> (Vec<Note>, Vec<Note>) {
    let mut normal = Vec::new();
    let mut separators = Vec::new();
    for note in notes {
        if note.is_separator() {
            separators.push(note);
        } else {
            normal.push(note);
        }
    }
    (normal, separators)
}

fn relationship_part_path(owner_path: &str) -> String {
    match owner_path.rsplit_once('/') {
        Some((directory, filename)) => format!("{directory}/_rels/{filename}.rels"),
        None => format!("_rels/{owner_path}.rels"),
    }
}

pub(crate) fn find_part<'a>(
    parts: &'a [(String, Vec<u8>)],
    path: &str,
) -> Option<(&'a str, &'a [u8])> {
    parts
        .iter()
        .find(|(candidate, _)| candidate.eq_ignore_ascii_case(path))
        .map(|(path, bytes)| (path.as_str(), bytes.as_slice()))
}

pub fn s8_wire_envelope(projection: S8Projection) -> Result<S8WireEnvelope, ParseError> {
    let canonical =
        from_serializable(&projection).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let bytes =
        to_canonical_bytes(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let sha =
        canonical_sha256(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    Ok(S8WireEnvelope {
        wire_version: 1,
        projection,
        canonical_base64: base64::engine::general_purpose::STANDARD.encode(bytes),
        canonical_sha256: sha,
    })
}

pub fn parse_docx_s8_wire(data: &[u8]) -> Result<S8WireEnvelope, ParseError> {
    s8_wire_envelope(parse_docx_s8_projection(data)?)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::block::BlockContent;
    use crate::paragraph::ParagraphContent;

    #[test]
    fn package_owns_all_story_parts_and_prunes_orphan_anchors() {
        let parts = vec![
            (
                "word/document.xml".to_owned(),
                br#"<w:document xmlns:w="w" xmlns:r="r"><w:body><w:p><w:commentRangeStart w:id="1"/><w:commentRangeEnd w:id="9"/></w:p><w:sectPr><w:headerReference w:type="first" r:id="rHeader"/><w:titlePg/></w:sectPr></w:body></w:document>"#.to_vec(),
            ),
            (
                "word/_rels/document.xml.rels".to_owned(),
                format!(r#"<Relationships><Relationship Id="rHeader" Type="{}" Target="header1.xml"/></Relationships>"#, crate::relationships::relationship_types::HEADER).into_bytes(),
            ),
            (
                "word/header1.xml".to_owned(),
                br#"<w:hdr xmlns:w="w"><w:p><w:r><w:t>header</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:hdr>"#.to_vec(),
            ),
            (
                "word/footnotes.xml".to_owned(),
                br#"<w:footnotes xmlns:w="w"><w:footnote w:id="-1" w:type="separator"><w:p/></w:footnote><w:footnote w:id="1"><w:p/></w:footnote></w:footnotes>"#.to_vec(),
            ),
            (
                "word/endnotes.xml".to_owned(),
                br#"<w:endnotes xmlns:w="w"><w:endnote w:id="2"><w:p/></w:endnote></w:endnotes>"#.to_vec(),
            ),
            (
                "word/comments.xml".to_owned(),
                br#"<w:comments xmlns:w="w"><w:comment w:id="1" w:author="Ada"><w:p/></w:comment></w:comments>"#.to_vec(),
            ),
        ];
        let package = ooxml_opc::rezip_parts(&parts).unwrap();
        let parsed = parse_docx_s8_projection(&package).unwrap();
        assert_eq!(parsed.header_entries.len(), 1);
        assert!(matches!(
            parsed.header_entries[0].1.content[1],
            BlockContent::Table(_)
        ));
        assert_eq!(parsed.footnotes.len(), 1);
        assert_eq!(parsed.footnote_separators.len(), 1);
        assert_eq!(parsed.endnotes.len(), 1);
        assert_eq!(parsed.body.comments.as_ref().unwrap()[0].author, "Ada");
        let BlockContent::Paragraph(paragraph) = &parsed.body.content[0] else {
            panic!("paragraph")
        };
        assert_eq!(
            paragraph
                .content
                .iter()
                .map(ParagraphContent::node_type)
                .collect::<Vec<_>>(),
            ["commentRangeStart"]
        );
        let properties = parsed.body.final_section_properties.as_ref().unwrap();
        assert_eq!(properties.title_pg, Some(true));
        assert_eq!(
            properties.header_references.as_ref().unwrap()[0].reference_type,
            "first"
        );
    }
}
