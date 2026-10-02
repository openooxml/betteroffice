//! S9 full package orchestration and the versioned read-facade wire model.

use std::collections::HashSet;
use std::sync::Arc;

use base64::Engine as _;
use indexmap::IndexMap;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::block::{BlockContent, StoryParser};
use crate::canonical::{CanonicalValue, canonical_sha256, from_serializable, to_canonical_bytes};
use crate::chart::{Chart, parse_chart_parts};
use crate::comments::remove_orphan_comment_ranges;
use crate::document::{DocumentBody, extract_all_template_variables, parse_document_body_compact};
use crate::fonts::{FontTable, parse_font_table};
use crate::header_footer::{HeaderFooter, parse_related_header_footers};
use crate::media::{MediaFile, MediaScan, MediaTable, build_media_map_with_warnings};
use crate::notes::Note;
use crate::numbering::{NumberingDefinitions, parse_numbering};
use crate::paragraph::{HexIdAllocator, Paragraph};
use crate::paragraph_identity::parse_paragraph_id;
use crate::relationships::{Relationship, RelationshipMap, parse_relationships};
use crate::s8::{find_part, parse_comment_part, parse_note_part, partition_notes};
use crate::settings::{DocumentSettings, is_valid_utf8_xml_text, parse_settings};
use crate::smart_art::create_smart_art_context;
use crate::styles::{StyleDefinitions, StyleMap, parse_style_definitions};
use crate::theme::{Theme, apply_theme_font_lang, parse_theme};
use crate::xml::{ParseBudget, ParseError, ParseLimits, parse_xml};

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default, rename_all = "camelCase", deny_unknown_fields)]
pub struct S9ParseOptions {
    pub parse_headers_footers: bool,
    pub parse_notes: bool,
    pub detect_variables: bool,
    pub determinism_seed: Option<String>,
    pub include_canonical: bool,
    /// Records each paragraph's `w:p` occurrence in its part as `sourceOrdinal`.
    pub source_ordinals: bool,
}

impl Default for S9ParseOptions {
    fn default() -> Self {
        Self {
            parse_headers_footers: true,
            parse_notes: true,
            detect_variables: true,
            determinism_seed: None,
            include_canonical: false,
            source_ordinals: false,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S9PackageWire {
    pub document: S9DocumentBodyWire,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub styles: Option<StyleDefinitions>,
    pub theme: Theme,
    pub numbering: NumberingDefinitions,
    pub settings: DocumentSettings,
    pub font_table: FontTable,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_entries: Option<Vec<(String, HeaderFooter)>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footer_entries: Option<Vec<(String, HeaderFooter)>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnotes: Option<Vec<Note>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub endnotes: Option<Vec<Note>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnote_separators: Option<Vec<Note>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub endnote_separators: Option<Vec<Note>>,
    pub relationship_entries: Vec<(String, Relationship)>,
    pub media_entries: Vec<(String, Arc<MediaFile>)>,
    pub chart_entries: Vec<(String, Chart)>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S9DocumentBodyWire {
    pub content: Vec<BlockContent>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sections: Option<Vec<S9SectionWire>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub final_section_properties: Option<crate::section::SectionProperties>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub custom_root_bindings: Vec<crate::paragraph::RawAttribute>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comments: Option<Vec<crate::comments::Comment>>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S9SectionWire {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    pub properties: crate::section::SectionProperties,
    pub content_start: usize,
    pub content_end: usize,
}

impl From<DocumentBody> for S9DocumentBodyWire {
    fn from(body: DocumentBody) -> Self {
        let mut section_ends = Vec::new();
        for (index, block) in body.content.iter().enumerate() {
            if matches!(block, BlockContent::Paragraph(paragraph) if paragraph.section_properties.is_some())
            {
                section_ends.push(index + 1);
            }
        }
        if section_ends.last().copied() != Some(body.content.len()) || section_ends.is_empty() {
            section_ends.push(body.content.len());
        }
        let mut section_index = 0usize;
        let mut offset = 0usize;
        let sections = body.sections.map(|sections| {
            sections
                .into_iter()
                .map(|section| {
                    let content_start = offset;
                    offset = section_ends
                        .get(section_index)
                        .copied()
                        .unwrap_or(body.content.len());
                    section_index += 1;
                    S9SectionWire {
                        id: section.id,
                        properties: section.properties,
                        content_start,
                        content_end: offset,
                    }
                })
                .collect()
        });
        Self {
            content: body.content,
            sections,
            final_section_properties: body.final_section_properties,
            custom_root_bindings: body.custom_root_bindings,
            comments: body.comments,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S9DocumentWire {
    pub package: S9PackageWire,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub template_variables: Option<Vec<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub warnings: Option<Vec<String>>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BinaryPartWire {
    pub path: String,
    pub base64: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S9WireEnvelope {
    pub wire_version: u8,
    pub document: S9DocumentWire,
    pub embedded_font_parts: Vec<BinaryPartWire>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_table_relationships_xml: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub canonical_base64: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub canonical_sha256: Option<String>,
}

pub fn parse_docx_s9_wire(
    data: &[u8],
    options: S9ParseOptions,
) -> Result<S9WireEnvelope, ParseError> {
    parse_docx_s9_wire_with_limits(data, options, &ParseLimits::default())
}

/// [`parse_docx_s9_wire`] with a caller-supplied resource budget.
pub fn parse_docx_s9_wire_with_limits(
    data: &[u8],
    options: S9ParseOptions,
    limits: &ParseLimits,
) -> Result<S9WireEnvelope, ParseError> {
    Ok(parse_docx_s9_wire_parts_with_limits(data, options, limits)?.0)
}

/// [`parse_docx_s9_wire_with_limits`] plus the inflated parts in archive order.
pub fn parse_docx_s9_wire_parts_with_limits(
    data: &[u8],
    options: S9ParseOptions,
    limits: &ParseLimits,
) -> Result<(S9WireEnvelope, Vec<(String, Vec<u8>)>), ParseError> {
    let parts = ooxml_opc::unzip_parts(data).map_err(ParseError::Container)?;
    let envelope = parse_s9_package(&parts, data, options, limits, None, None)?
        .expect("a whole body is never refused");
    Ok((envelope, parts))
}

/// [`parse_docx_s9_wire_parts_with_limits`] leaving the media in the package:
/// an image names its part by a [`MediaTable`] token where it would carry a
/// `data:` URL, and the envelope lists no media entries. Watermarks keep
/// their `data:` URLs. Returns the parts the parse inflated and the media
/// table.
pub fn parse_docx_s9_wire_with_media_table(
    data: Arc<[u8]>,
    options: S9ParseOptions,
    limits: &ParseLimits,
) -> Result<(S9WireEnvelope, Vec<(String, Vec<u8>)>, MediaTable), ParseError> {
    let (parts, table) = media_table_parts(&data)?;
    let envelope = parse_s9_package(&parts, &data, options, limits, None, Some(&table))?
        .expect("a whole body is never refused");
    Ok((envelope, parts, table))
}

/// The package's parts other than its media, inflated, and its media table.
pub fn media_table_parts(
    data: &Arc<[u8]>,
) -> Result<(Vec<(String, Vec<u8>)>, MediaTable), ParseError> {
    media_table_parts_within(data, ooxml_opc::MAX_TOTAL_UNCOMPRESSED_BYTES)
}

/// Bounded extraction before transcoding, reserving compressed images' sizes.
#[doc(hidden)]
pub fn media_table_parts_within(
    data: &Arc<[u8]>,
    budget: u64,
) -> Result<(Vec<(String, Vec<u8>)>, MediaTable), ParseError> {
    let package =
        ooxml_opc::RetainedPackage::new(Arc::clone(data)).map_err(ParseError::Container)?;
    let scan = MediaScan::new(package, budget).map_err(ParseError::Container)?;
    let mut parts = ooxml_opc::unzip_parts_where(data, scan.remaining_budget(), |path| {
        !scan.keeps_compressed(path)
    })
    .map_err(ParseError::Container)?;
    let table = scan.finish(&mut parts).map_err(ParseError::Container)?;
    Ok((parts, table))
}

/// Parses already inflated package parts with only the body's first
/// `blocks` blocks, for a preview of its first pages: at least that many, and
/// on until no field spans past the last one. `None` when the document holds
/// a float placed from outside the text or a section with columns, which the
/// layout may apply to the body's first pages, so no cut of the body lays out
/// like the whole. The parts alone cannot be hashed into generated IDs, so
/// `options.determinism_seed` is required.
pub fn parse_docx_s9_preview_from_parts(
    parts: &[(String, Vec<u8>)],
    blocks: usize,
    options: S9ParseOptions,
    limits: &ParseLimits,
) -> Result<Option<S9WireEnvelope>, ParseError> {
    if options.determinism_seed.is_none() || options.include_canonical {
        return Err(ParseError::Canonical(
            "parsing parts needs a determinism seed and no canonical envelope".to_owned(),
        ));
    }
    parse_s9_package(parts, &[], options, limits, Some(blocks), None)
}

#[doc(hidden)]
pub fn parse_docx_s9_preview_from_parts_full_dom(
    parts: &[(String, Vec<u8>)],
    blocks: usize,
    options: S9ParseOptions,
    limits: &ParseLimits,
) -> Result<Option<S9WireEnvelope>, ParseError> {
    if options.determinism_seed.is_none() || options.include_canonical {
        return Err(ParseError::Canonical(
            "parsing parts needs a determinism seed and no canonical envelope".to_owned(),
        ));
    }
    parse_s9_package_impl(parts, &[], options, limits, Some(blocks), None, false)
}

/// [`parse_docx_s9_preview_from_parts`] with images naming their parts by
/// `table`'s tokens, as [`parse_docx_s9_wire_with_media_table`] does; `parts`
/// need not hold the media.
pub fn parse_docx_s9_preview_with_media_table(
    parts: &[(String, Vec<u8>)],
    table: &MediaTable,
    blocks: usize,
    options: S9ParseOptions,
    limits: &ParseLimits,
) -> Result<Option<S9WireEnvelope>, ParseError> {
    if options.determinism_seed.is_none() || options.include_canonical {
        return Err(ParseError::Canonical(
            "parsing parts needs a determinism seed and no canonical envelope".to_owned(),
        ));
    }
    parse_s9_package(parts, &[], options, limits, Some(blocks), Some(table))
}

/// `None` only when `body_blocks` would cut a body the preview refuses; see
/// [`parse_docx_s9_preview_from_parts`].
fn parse_s9_package(
    parts: &[(String, Vec<u8>)],
    data: &[u8],
    options: S9ParseOptions,
    limits: &ParseLimits,
    body_blocks: Option<usize>,
    media_table: Option<&MediaTable>,
) -> Result<Option<S9WireEnvelope>, ParseError> {
    parse_s9_package_impl(parts, data, options, limits, body_blocks, media_table, true)
}

fn parse_s9_package_impl(
    parts: &[(String, Vec<u8>)],
    data: &[u8],
    options: S9ParseOptions,
    limits: &ParseLimits,
    body_blocks: Option<usize>,
    media_table: Option<&MediaTable>,
    prefix_preview: bool,
) -> Result<Option<S9WireEnvelope>, ParseError> {
    if media_table.is_some() && options.include_canonical {
        return Err(ParseError::Canonical(
            "a canonical envelope needs the media inflated".to_owned(),
        ));
    }
    let mut budget = ParseBudget::new(limits);
    if let Some(table) = media_table {
        budget.read_media_from(table);
    }
    if options.source_ordinals {
        budget.record_source_ordinals();
    }
    let document_path = crate::relationships::office_document_path(parts, &mut budget)?;
    let document_relationships_path = crate::relationships::relationship_part_path(&document_path);

    let settings = parse_settings(
        find_part(parts, "word/settings.xml").map(|(_, bytes)| bytes),
        "word/settings.xml",
        &mut budget,
    )?;
    let mut theme = parse_theme(
        find_part(parts, "word/theme/theme1.xml").map(|(_, bytes)| bytes),
        "word/theme/theme1.xml",
        &mut budget,
    )?;
    apply_theme_font_lang(&mut theme, settings.theme_font_lang.as_ref());

    let styles = find_part(parts, "word/styles.xml")
        .filter(|(_, xml)| !xml.is_empty())
        .map(|(path, xml)| parse_style_definitions(xml, Some(&theme), path, &mut budget))
        .transpose()?;
    let style_map: StyleMap = styles
        .as_ref()
        .map(|definitions| {
            definitions
                .styles
                .iter()
                .map(|style| (style.style_id.clone(), style.clone()))
                .collect()
        })
        .unwrap_or_default();
    let doc_defaults = styles
        .as_ref()
        .and_then(|definitions| definitions.doc_defaults.as_ref());
    let numbering = parse_numbering(
        find_part(parts, "word/numbering.xml").map(|(_, bytes)| bytes),
        "word/numbering.xml",
        &mut budget,
    )?;
    let font_table = parse_font_table(
        find_part(parts, "word/fontTable.xml").map(|(_, bytes)| bytes),
        "word/fontTable.xml",
        &mut budget,
    )?;
    let relationships = match find_part(parts, &document_relationships_path) {
        Some((path, xml)) => parse_relationships(xml, path, &mut budget)?,
        None => RelationshipMap::new(),
    };
    let (media, media_warnings) = match media_table {
        Some(table) => (table.media_map(), table.warnings().to_vec()),
        None => build_media_map_with_warnings(parts),
    };
    let all_xml: IndexMap<_, _> = parts
        .iter()
        .filter(|(path, _)| {
            let lower = path.to_ascii_lowercase();
            lower.ends_with(".xml") || lower.ends_with(".rels")
        })
        .map(|(path, bytes)| (path.clone(), bytes.as_slice()))
        .collect::<IndexMap<String, &[u8]>>();
    let charts = parse_chart_parts(&all_xml, limits);
    let mut smart_art = create_smart_art_context(&all_xml);
    let digest = options
        .determinism_seed
        .clone()
        .unwrap_or_else(|| format!("{:x}", Sha256::digest(data)));
    let mut ids = HexIdAllocator::from_sha256(&digest)?;

    let document_part = find_part(parts, &document_path);
    let mut warnings = Vec::new();
    let mut body = match document_part.filter(|(_, xml)| !xml.is_empty()) {
        Some((path, xml)) => {
            let mut scanned_budget = body_blocks
                .filter(|_| prefix_preview)
                .map(|_| budget.clone());
            let prefix = if let Some(blocks) = body_blocks.filter(|_| prefix_preview) {
                match crate::document::streaming_body_cut(xml, scanned_budget.as_mut().unwrap()) {
                    Ok(refused) => {
                        if refused {
                            return Ok(None);
                        }
                        let keep = blocks.saturating_mul(2).max(blocks.saturating_add(64));
                        crate::document::body_prefix(xml, keep).map(|xml| (xml, keep))
                    }
                    _ => None,
                }
            } else {
                None
            };
            let parsed = parse_xml(
                prefix.as_ref().map_or(xml, |(xml, _)| xml.as_slice()),
                path,
                &mut budget,
            )?;
            if prefix.is_some() {
                budget = scanned_budget.unwrap();
            }
            match parsed.root() {
                Some(root) => {
                    if body_blocks.is_some() && crate::document::refuses_a_body_cut(root) {
                        return Ok(None);
                    }
                    let mut parser = StoryParser {
                        relationships: Some(&relationships),
                        theme: Some(&theme),
                        styles: Some(&style_map),
                        doc_defaults,
                        numbering: Some(&numbering),
                        media: &media,
                        charts: &charts,
                        smart_art: &mut smart_art,
                        budget: &mut budget,
                        ids: &mut ids,
                        part: path,
                    };
                    if let Some((_, keep)) = &prefix {
                        let kept_children = root.child("w", "body").map_or(0, |body| {
                            body.child_elements()
                                .take(*keep)
                                .map(|child| {
                                    if matches!(child.local_name(), "customXml" | "smartTag") {
                                        crate::block::transparent_children(child, false).len()
                                    } else {
                                        1
                                    }
                                })
                                .sum::<usize>()
                        });
                        let (body, read) = crate::document::parse_document_body_compact_with_read(
                            root,
                            &mut parser,
                            body_blocks,
                            kept_children,
                        )?;
                        if read >= kept_children {
                            return parse_s9_package_impl(
                                parts,
                                data,
                                options,
                                limits,
                                body_blocks,
                                media_table,
                                false,
                            );
                        }
                        body
                    } else {
                        parse_document_body_compact(root, &mut parser, body_blocks)?
                    }
                }
                None => DocumentBody::default(),
            }
        }
        None => {
            warnings.push("No document.xml found in DOCX".to_owned());
            DocumentBody::default()
        }
    };

    let (mut headers, mut footers) = if options.parse_headers_footers {
        let (headers, footers) = parse_related_header_footers(
            parts,
            &relationships,
            Some(&theme),
            Some(&style_map),
            doc_defaults,
            Some(&numbering),
            &media,
            &charts,
            &mut smart_art,
            &mut budget,
            &mut ids,
        )?;
        (Some(headers), Some(footers))
    } else {
        (None, None)
    };

    let (mut footnotes, mut endnotes, mut footnote_separators, mut endnote_separators) =
        if options.parse_notes {
            let all_footnotes = parse_note_part(
                parts,
                "word/footnotes.xml",
                true,
                &relationships,
                Some(&theme),
                Some(&style_map),
                doc_defaults,
                Some(&numbering),
                &media,
                &charts,
                &mut smart_art,
                &mut budget,
                &mut ids,
            )?;
            let all_endnotes = parse_note_part(
                parts,
                "word/endnotes.xml",
                false,
                &relationships,
                Some(&theme),
                Some(&style_map),
                doc_defaults,
                Some(&numbering),
                &media,
                &charts,
                &mut smart_art,
                &mut budget,
                &mut ids,
            )?;
            let (footnotes, footnote_separators) = partition_notes(all_footnotes);
            let (endnotes, endnote_separators) = partition_notes(all_endnotes);
            (
                Some(footnotes),
                Some(endnotes),
                Some(footnote_separators),
                Some(endnote_separators),
            )
        } else {
            (None, None, None, None)
        };

    let comments = parse_comment_part(
        parts,
        &relationships,
        Some(&theme),
        Some(&style_map),
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
    for story in headers
        .iter_mut()
        .flat_map(IndexMap::values_mut)
        .chain(footers.iter_mut().flat_map(IndexMap::values_mut))
    {
        remove_orphan_comment_ranges(&mut story.content, &comment_ids);
    }
    for note in footnotes
        .iter_mut()
        .flat_map(|notes| notes.iter_mut())
        .chain(endnotes.iter_mut().flat_map(|notes| notes.iter_mut()))
        .chain(
            footnote_separators
                .iter_mut()
                .flat_map(|notes| notes.iter_mut()),
        )
        .chain(
            endnote_separators
                .iter_mut()
                .flat_map(|notes| notes.iter_mut()),
        )
    {
        remove_orphan_comment_ranges(&mut note.content, &comment_ids);
    }

    dedupe_package_paragraph_ids(
        &mut body,
        headers.as_mut(),
        footers.as_mut(),
        footnotes.as_mut(),
        endnotes.as_mut(),
        footnote_separators.as_mut(),
        endnote_separators.as_mut(),
    );

    let template_variables = options
        .detect_variables
        .then(|| extract_all_template_variables(&body.content));
    warnings.extend(media_warnings);
    warnings.extend(smart_art.warnings);
    let warnings = (!warnings.is_empty()).then_some(warnings);

    let mut document = S9DocumentWire {
        package: S9PackageWire {
            document: body.into(),
            styles,
            theme,
            numbering: numbering.definitions,
            settings,
            font_table,
            header_entries: headers.map(|stories| stories.into_iter().collect()),
            footer_entries: footers.map(|stories| stories.into_iter().collect()),
            footnotes,
            endnotes,
            footnote_separators,
            endnote_separators,
            relationship_entries: relationships.into_iter().collect(),
            media_entries: match media_table {
                Some(_) => Vec::new(),
                None => media.into_iter().collect(),
            },
            chart_entries: charts.into_iter().collect(),
        },
        template_variables,
        warnings,
    };
    if let Some(error) = budget.media_read_error() {
        return Err(ParseError::Container(error));
    }
    if let Some(table) = media_table {
        resolve_host_media(&mut document.package, table)?;
    }

    let (canonical_base64, canonical_sha256) = if options.include_canonical {
        let canonical = canonical_document(&document, data)?;
        let canonical_bytes = to_canonical_bytes(&canonical)
            .map_err(|error| ParseError::Canonical(error.to_string()))?;
        let sha = canonical_sha256(&canonical)
            .map_err(|error| ParseError::Canonical(error.to_string()))?;
        (
            Some(base64::engine::general_purpose::STANDARD.encode(canonical_bytes)),
            Some(sha),
        )
    } else {
        (None, None)
    };
    let embedded_font_parts = parts
        .iter()
        .filter(|(path, _)| path.to_ascii_lowercase().starts_with("word/fonts/"))
        .map(|(path, bytes)| BinaryPartWire {
            path: path.clone(),
            base64: base64::engine::general_purpose::STANDARD.encode(bytes),
        })
        .collect();
    let font_table_relationships_xml = find_part(parts, "word/_rels/fontTable.xml.rels")
        .filter(|(_, xml)| is_valid_utf8_xml_text(xml))
        .map(|(_, xml)| String::from_utf8_lossy(xml).into_owned());

    Ok(Some(S9WireEnvelope {
        wire_version: 1,
        document,
        embedded_font_parts,
        font_table_relationships_xml,
        canonical_base64,
        canonical_sha256,
    }))
}

/// Flags every paragraph whose ID repeats one earlier in the package,
/// comparing values numerically; each keeps its authored ID.
fn dedupe_package_paragraph_ids(
    body: &mut DocumentBody,
    headers: Option<&mut IndexMap<String, HeaderFooter>>,
    footers: Option<&mut IndexMap<String, HeaderFooter>>,
    footnotes: Option<&mut Vec<Note>>,
    endnotes: Option<&mut Vec<Note>>,
    footnote_separators: Option<&mut Vec<Note>>,
    endnote_separators: Option<&mut Vec<Note>>,
) {
    let mut seen = HashSet::new();
    dedupe_blocks(&mut body.content, &mut seen);
    for story in headers.into_iter().flat_map(IndexMap::values_mut) {
        dedupe_blocks(&mut story.content, &mut seen);
    }
    for story in footers.into_iter().flat_map(IndexMap::values_mut) {
        dedupe_blocks(&mut story.content, &mut seen);
    }
    for note in footnotes
        .into_iter()
        .flat_map(|notes| notes.iter_mut())
        .chain(endnotes.into_iter().flat_map(|notes| notes.iter_mut()))
        .chain(
            footnote_separators
                .into_iter()
                .flat_map(|notes| notes.iter_mut()),
        )
        .chain(
            endnote_separators
                .into_iter()
                .flat_map(|notes| notes.iter_mut()),
        )
    {
        dedupe_blocks(&mut note.content, &mut seen);
    }
}

fn dedupe_blocks(blocks: &mut [BlockContent], seen: &mut HashSet<u32>) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => dedupe_paragraph(Arc::make_mut(paragraph), seen),
            BlockContent::Table(table) => {
                for row in &mut Arc::make_mut(table).rows {
                    for cell in &mut row.cells {
                        dedupe_blocks(&mut cell.content, seen);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => dedupe_blocks(&mut Arc::make_mut(sdt).content, seen),
            BlockContent::RawXml(_) => {}
        }
    }
}

fn dedupe_paragraph(paragraph: &mut Paragraph, seen: &mut HashSet<u32>) {
    if let Some(id) = paragraph.para_id.as_deref().and_then(parse_paragraph_id)
        && !seen.insert(id)
    {
        paragraph.repeated_para_id = Some(true);
    }
}

fn canonical_document(
    document: &S9DocumentWire,
    original: &[u8],
) -> Result<CanonicalValue, ParseError> {
    let package = &document.package;
    let mut package_entries = vec![
        (
            "document".to_owned(),
            canonical_document_body(&package.document)?,
        ),
        serializable_entry("theme", &package.theme)?,
        serializable_entry("numbering", &package.numbering)?,
        serializable_entry("settings", &package.settings)?,
        serializable_entry("fontTable", &package.font_table)?,
        (
            "relationships".to_owned(),
            ordered_map(&package.relationship_entries)?,
        ),
        ("media".to_owned(), canonical_media(&package.media_entries)?),
        ("charts".to_owned(), ordered_map(&package.chart_entries)?),
    ];
    if let Some(styles) = &package.styles {
        package_entries.push(serializable_entry("styles", styles)?);
    }
    for (name, stories) in [
        ("headers", package.header_entries.as_ref()),
        ("footers", package.footer_entries.as_ref()),
    ] {
        if let Some(stories) = stories {
            package_entries.push((name.to_owned(), ordered_map(stories)?));
        }
    }
    for (name, notes) in [
        ("footnotes", package.footnotes.as_ref()),
        ("endnotes", package.endnotes.as_ref()),
        ("footnoteSeparators", package.footnote_separators.as_ref()),
        ("endnoteSeparators", package.endnote_separators.as_ref()),
    ] {
        if let Some(notes) = notes {
            package_entries.push(serializable_entry(name, notes)?);
        }
    }

    let mut document_entries = vec![
        (
            "package".to_owned(),
            CanonicalValue::Object(package_entries),
        ),
        (
            "originalBuffer".to_owned(),
            CanonicalValue::Binary(original.to_vec()),
        ),
    ];
    if let Some(variables) = &document.template_variables {
        document_entries.push(serializable_entry("templateVariables", variables)?);
    }
    if let Some(warnings) = &document.warnings {
        document_entries.push(serializable_entry("warnings", warnings)?);
    }
    Ok(CanonicalValue::Object(document_entries))
}

fn canonical_document_body(body: &S9DocumentBodyWire) -> Result<CanonicalValue, ParseError> {
    let mut entries = vec![serializable_entry("content", &body.content)?];
    if let Some(sections) = &body.sections {
        let sections = sections
            .iter()
            .map(|section| {
                if section.content_start > section.content_end
                    || section.content_end > body.content.len()
                {
                    return Err(ParseError::Canonical(
                        "S9 section content range is invalid".to_owned(),
                    ));
                }
                let mut section_entries = vec![
                    serializable_entry("properties", &section.properties)?,
                    serializable_entry(
                        "content",
                        &body.content[section.content_start..section.content_end],
                    )?,
                ];
                if let Some(id) = &section.id {
                    section_entries.push(serializable_entry("id", id)?);
                }
                Ok(CanonicalValue::Object(section_entries))
            })
            .collect::<Result<Vec<_>, ParseError>>()?;
        entries.push(("sections".to_owned(), CanonicalValue::Array(sections)));
    }
    if let Some(properties) = &body.final_section_properties {
        entries.push(serializable_entry("finalSectionProperties", properties)?);
    }
    if let Some(comments) = &body.comments {
        entries.push(serializable_entry("comments", comments)?);
    }
    Ok(CanonicalValue::Object(entries))
}

fn serializable_entry<T: Serialize + ?Sized>(
    name: &str,
    value: &T,
) -> Result<(String, CanonicalValue), ParseError> {
    Ok((
        name.to_owned(),
        from_serializable(value).map_err(|error| ParseError::Canonical(error.to_string()))?,
    ))
}

fn ordered_map<T: Serialize>(entries: &[(String, T)]) -> Result<CanonicalValue, ParseError> {
    entries
        .iter()
        .map(|(key, value)| {
            Ok((
                key.clone(),
                from_serializable(value)
                    .map_err(|error| ParseError::Canonical(error.to_string()))?,
            ))
        })
        .collect::<Result<Vec<_>, ParseError>>()
        .map(CanonicalValue::OrderedMap)
}

/// Gives the watermarks of a package parsed against `table` the `data:` URLs
/// their tokens stand for, since they reach the host as parsed.
fn resolve_host_media(package: &mut S9PackageWire, table: &MediaTable) -> Result<(), ParseError> {
    for (_, story) in package
        .header_entries
        .iter_mut()
        .chain(package.footer_entries.iter_mut())
        .flatten()
    {
        if let Some(crate::vml::Watermark::Picture {
            data_url: Some(data_url),
            ..
        }) = story.watermark.as_mut()
            && let Some(index) =
                crate::media::media_token_index(data_url).filter(|index| *index < table.len())
        {
            *data_url = table.data_url(index).map_err(ParseError::Container)?;
        }
    }
    Ok(())
}

fn canonical_media(entries: &[(String, Arc<MediaFile>)]) -> Result<CanonicalValue, ParseError> {
    entries
        .iter()
        .map(|(key, file)| {
            let data = base64::engine::general_purpose::STANDARD
                .decode(&file.base64)
                .map_err(|error| ParseError::Canonical(error.to_string()))?;
            let mut values = vec![
                ("path".to_owned(), CanonicalValue::String(file.path.clone())),
                (
                    "mimeType".to_owned(),
                    CanonicalValue::String(file.mime_type.clone()),
                ),
                ("data".to_owned(), CanonicalValue::Binary(data)),
                (
                    "dataUrl".to_owned(),
                    CanonicalValue::String(file.data_url.clone()),
                ),
            ];
            if let Some(filename) = &file.filename {
                values.push((
                    "filename".to_owned(),
                    CanonicalValue::String(filename.clone()),
                ));
            }
            Ok((key.clone(), CanonicalValue::Object(values)))
        })
        .collect::<Result<Vec<_>, ParseError>>()
        .map(CanonicalValue::OrderedMap)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn full_wire_applies_options_and_package_wide_paragraph_ids() {
        let parts = vec![
            (
                "word/document.xml".to_owned(),
                br#"<w:document xmlns:w="w" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>{name}</w:t></w:r></w:p><w:p w14:paraId="00000001"/></w:body></w:document>"#.to_vec(),
            ),
            (
                "word/header1.xml".to_owned(),
                br#"<w:hdr xmlns:w="w"><w:p/></w:hdr>"#.to_vec(),
            ),
        ];
        let package = ooxml_opc::rezip_parts(&parts).unwrap();
        let parsed = parse_docx_s9_wire(
            &package,
            S9ParseOptions {
                parse_headers_footers: false,
                parse_notes: false,
                detect_variables: true,
                determinism_seed: None,
                include_canonical: true,
                source_ordinals: false,
            },
        )
        .unwrap();
        assert!(parsed.document.package.header_entries.is_none());
        assert!(parsed.document.package.footnotes.is_none());
        assert_eq!(
            parsed.document.template_variables.as_deref(),
            Some(["name".to_owned()].as_slice())
        );
        let BlockContent::Paragraph(first) = &parsed.document.package.document.content[0] else {
            panic!("paragraph")
        };
        let BlockContent::Paragraph(second) = &parsed.document.package.document.content[1] else {
            panic!("paragraph")
        };
        assert_eq!(first.para_id.as_deref(), Some("00000001"));
        assert_eq!(second.para_id.as_deref(), Some("00000001"));
        assert_eq!(
            (first.repeated_para_id, second.repeated_para_id),
            (None, Some(true))
        );
        assert_eq!(parsed.canonical_sha256.as_deref().unwrap().len(), 64);
    }
}
