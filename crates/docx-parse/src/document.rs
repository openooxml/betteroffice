//! Body-level document assembly, sections, and read-only paragraph queries.

use std::collections::HashSet;

use indexmap::IndexMap;
use quick_xml::Reader;
use quick_xml::events::{BytesDecl, Event};
use serde::{Deserialize, Serialize};

use crate::block::{BlockContent, StoryParser, transparent_children};
use crate::comments::Comment;
use crate::inline::{InlineNode, RunContent};
use crate::paragraph::RawAttribute;
use crate::paragraph::{Paragraph, ParagraphContent};
use crate::section::{
    SectionProperties, apply_section_inheritance, default_section_properties,
    parse_section_properties,
};
use crate::xml::{ParseError, XmlElement};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct Section {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    pub properties: SectionProperties,
    pub content: Vec<BlockContent>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentBody {
    pub content: Vec<BlockContent>,
    /// Sections with header/footer inheritance resolved forward.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sections: Option<Vec<Section>>,
    /// The body-level `w:sectPr` as authored. Resolution lives in
    /// `sections`; re-emitting a resolved copy would rehome inherited
    /// header and footer references onto the final section on save.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub final_section_properties: Option<SectionProperties>,
    /// Root namespace bindings and attributes retained outside the standard boilerplate.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub custom_root_bindings: Vec<RawAttribute>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comments: Option<Vec<Comment>>,
}

/// Whether `document` holds what the layout may apply to the body's first
/// pages from anywhere in it, so that no cut of the body lays out like the
/// whole: a float placed from outside the text (a drawing anchored to the page
/// or margin, a table or frame not anchored to the text, VML positioned from
/// the page or margin), or a section with columns, whose settings sections
/// without their own may take.
pub(crate) fn refuses_a_body_cut(document: &XmlElement) -> bool {
    let mut pending = vec![document];
    while let Some(element) = pending.pop() {
        let refused = match element.local_name() {
            "anchor" => {
                element.attribute(None, "simplePos") == Some("1")
                    || element
                        .child_by_local_name("positionV")
                        .and_then(|position| position.attribute(None, "relativeFrom"))
                        .is_some_and(|from| !matches!(from, "paragraph" | "line"))
            }
            "tblpPr" => element.attribute(Some("w"), "vertAnchor") != Some("text"),
            "framePr" => element.attribute(Some("w"), "vAnchor") != Some("text"),
            "cols" => {
                element
                    .attribute(Some("w"), "num")
                    .is_some_and(|count| !matches!(count.trim(), "" | "0" | "1"))
                    || element.children_named("w", "col").nth(1).is_some()
            }
            _ => crate::vml::placed_off_the_text(element),
        };
        if refused {
            return true;
        }
        pending.extend(element.child_elements());
    }
    false
}

#[derive(Default)]
struct CutFrame {
    anchor: bool,
    position_seen: bool,
    columns: bool,
    column_seen: bool,
    namespace_declarations: Vec<(String, String)>,
    drawing: bool,
}

#[derive(Default)]
struct CutDrawing {
    bindings: IndexMap<String, String>,
    used_prefixes: HashSet<String>,
    attribute_names: HashSet<String>,
    attribute_bytes: usize,
}

impl CutDrawing {
    fn use_attribute(&mut self, name: &str, value: &str) {
        if matches!(
            crate::xml::local_name(name),
            "Ignorable"
                | "MustUnderstand"
                | "Requires"
                | "ProcessContent"
                | "PreserveElements"
                | "PreserveAttributes"
        ) {
            for token in value.split_whitespace() {
                self.used_prefixes.insert(
                    token
                        .split_once(':')
                        .map_or(token, |(prefix, _)| prefix)
                        .to_owned(),
                );
            }
        }
        if let Some((prefix, _)) = name.split_once(':')
            && prefix != "xmlns"
        {
            self.used_prefixes.insert(prefix.to_owned());
        }
    }
}

fn finish_cut_drawing(
    drawings: &mut Vec<CutDrawing>,
    part: &str,
    budget: &mut crate::xml::ParseBudget<'_>,
) -> Result<(), ParseError> {
    let mut drawing = drawings.pop().unwrap();
    for (name, value) in drawing.bindings {
        let prefix = name.strip_prefix("xmlns:").unwrap_or("");
        if !drawing.used_prefixes.contains(prefix) || drawing.attribute_names.contains(&name) {
            continue;
        }
        drawing.attribute_bytes += name.len() + value.len();
        if drawing.attribute_names.len() >= budget.limits().max_attributes_per_element
            || drawing.attribute_bytes > budget.limits().max_attribute_bytes
        {
            return Err(ParseError::ResourceLimit {
                kind: "drawingNamespaceAliases",
                part: part.to_owned(),
            });
        }
        budget.charge_text(name.len() + value.len(), part)?;
        for ancestor in drawings.iter_mut() {
            ancestor.use_attribute(&name, &value);
        }
        drawing.attribute_names.insert(name);
    }
    Ok(())
}

pub(crate) fn streaming_refuses_a_body_cut(xml: &[u8]) -> Result<bool, ParseError> {
    let limits = crate::xml::ParseLimits::default();
    streaming_body_cut(xml, &mut crate::xml::ParseBudget::new(&limits))
}

pub(crate) fn streaming_body_cut(
    xml: &[u8],
    budget: &mut crate::xml::ParseBudget<'_>,
) -> Result<bool, ParseError> {
    let repaired = crate::xml::escape_stray_ampersands(xml);
    let xml = repaired.as_ref();
    let part = "word/document.xml";
    let error = |offset, message: String| ParseError::MalformedXml {
        part: part.to_owned(),
        offset,
        message,
    };
    if starts_with_a_byte_order_mark(xml) {
        return Err(error(0, "byte order mark".to_owned()));
    }
    budget.charge_xml_bytes(xml.len(), part)?;
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<CutFrame> = Vec::new();
    let mut drawings: Vec<CutDrawing> = Vec::new();
    let mut roots = 0;
    let mut refused = false;
    loop {
        let event = reader
            .read_event()
            .map_err(|err| error(reader.buffer_position(), err.to_string()))?;
        budget.charge_event(part)?;
        match event {
            Event::Start(ref start) | Event::Empty(ref start) => {
                if stack.len() + 1 > budget.limits().max_xml_depth {
                    return Err(ParseError::ResourceLimit {
                        kind: "xmlDepth",
                        part: part.to_owned(),
                    });
                }
                if stack.is_empty() {
                    roots += 1;
                    if roots > 1 {
                        return Err(error(0, "multiple root elements".to_owned()));
                    }
                }
                let name = start.name();
                let name = reader
                    .decoder()
                    .decode(name.as_ref())
                    .map_err(|err| error(reader.buffer_position(), err.to_string()))?;
                budget.charge_text(name.len(), part)?;
                let local = crate::xml::local_name(&name);
                let mut bindings = IndexMap::new();
                if matches!(local, "pict" | "object") {
                    for ancestor in &stack {
                        for (name, value) in &ancestor.namespace_declarations {
                            let prefix = name.strip_prefix("xmlns:").unwrap_or("");
                            if crate::xml::canonical_namespace(prefix) == Some(value.as_str()) {
                                bindings.shift_remove(name.as_str());
                            } else {
                                bindings.insert(name.clone(), value.clone());
                            }
                        }
                    }
                }
                let mut frame = CutFrame {
                    anchor: local == "anchor",
                    columns: local == "cols",
                    drawing: !bindings.is_empty(),
                    ..CutFrame::default()
                };
                if frame.drawing {
                    drawings.push(CutDrawing {
                        bindings,
                        ..CutDrawing::default()
                    });
                }
                for drawing in &mut drawings {
                    drawing
                        .used_prefixes
                        .insert(crate::xml::namespace_prefix(&name).unwrap_or("").to_owned());
                }
                let position = stack.last_mut().is_some_and(|parent| {
                    if parent.anchor && !parent.position_seen && local == "positionV" {
                        parent.position_seen = true;
                        true
                    } else {
                        false
                    }
                });
                if let Some(parent) = stack.last_mut()
                    && parent.columns
                    && local == "col"
                {
                    refused |= parent.column_seen;
                    parent.column_seen = true;
                }
                let mut unprefixed = None;
                let mut prefixed = None;
                let mut attribute_bytes = 0usize;
                for (index, attribute) in start.attributes().enumerate() {
                    if index >= budget.limits().max_attributes_per_element {
                        return Err(ParseError::ResourceLimit {
                            kind: "attributesPerElement",
                            part: part.to_owned(),
                        });
                    }
                    let attribute = attribute
                        .map_err(|err| error(reader.buffer_position(), err.to_string()))?;
                    if attribute.value.contains(&b'<') {
                        return Err(error(
                            reader.buffer_position(),
                            "unescaped '<' in attribute value".to_owned(),
                        ));
                    }
                    let key = reader
                        .decoder()
                        .decode(attribute.key.as_ref())
                        .map_err(|err| error(reader.buffer_position(), err.to_string()))?;
                    #[allow(deprecated)]
                    let value = attribute
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(|err| error(reader.buffer_position(), err.to_string()))?;
                    attribute_bytes = attribute_bytes
                        .checked_add(key.len() + value.len())
                        .ok_or_else(|| ParseError::ResourceLimit {
                            kind: "attributeBytes",
                            part: part.to_owned(),
                        })?;
                    if attribute_bytes > budget.limits().max_attribute_bytes {
                        return Err(ParseError::ResourceLimit {
                            kind: "attributeBytes",
                            part: part.to_owned(),
                        });
                    }
                    budget.charge_text(key.len() + value.len(), part)?;
                    if key == "xmlns" || key.starts_with("xmlns:") {
                        frame
                            .namespace_declarations
                            .push((key.to_string(), value.to_string()));
                    }
                    for drawing in &mut drawings {
                        drawing.use_attribute(&key, &value);
                    }
                    if frame.drawing {
                        let drawing = drawings.last_mut().unwrap();
                        drawing.attribute_names.insert(key.to_string());
                        drawing.attribute_bytes = attribute_bytes;
                    }
                    if local == "anchor" && key == "simplePos" {
                        refused |= value == "1";
                    }
                    if position && key == "relativeFrom" {
                        refused |= !matches!(value.as_ref(), "paragraph" | "line");
                    }
                    let wanted = match local {
                        "tblpPr" => "vertAnchor",
                        "framePr" => "vAnchor",
                        "cols" => "num",
                        _ => "",
                    };
                    let test = || {
                        if local == "cols" {
                            !matches!(value.trim(), "" | "0" | "1")
                        } else {
                            value != "text"
                        }
                    };
                    if !wanted.is_empty() {
                        if key == wanted {
                            unprefixed = Some(test());
                        } else if key.strip_prefix("w:") == Some(wanted) {
                            prefixed = Some(test());
                        }
                    } else if !matches!(local, "anchor") && key == "style" {
                        refused |= crate::vml::style_placed_off_the_text(Some(&value));
                    }
                }
                match local {
                    "tblpPr" | "framePr" => refused |= prefixed.or(unprefixed).unwrap_or(true),
                    "cols" => refused |= prefixed.or(unprefixed).unwrap_or(false),
                    _ => {}
                }
                if matches!(event, Event::Start(_)) {
                    stack.push(frame);
                } else if frame.drawing {
                    finish_cut_drawing(&mut drawings, part, budget)?;
                }
            }
            Event::End(_) => {
                let frame = stack.pop().ok_or_else(|| {
                    error(
                        reader.buffer_position(),
                        "unexpected closing element".to_owned(),
                    )
                })?;
                if frame.drawing {
                    finish_cut_drawing(&mut drawings, part, budget)?;
                }
            }
            Event::Text(text) => {
                let decoded = text
                    .decode()
                    .map_err(|err| error(reader.buffer_position(), err.to_string()))?;
                let value = quick_xml::escape::unescape(&decoded)
                    .map_err(|err| error(reader.buffer_position(), err.to_string()))?;
                budget.charge_text(value.len(), part)?;
                if stack.is_empty() && !value.trim().is_empty() {
                    return Err(error(
                        reader.buffer_position(),
                        "text outside the root element".to_owned(),
                    ));
                }
            }
            Event::CData(text) => {
                let value = text
                    .decode()
                    .map_err(|err| error(reader.buffer_position(), err.to_string()))?;
                budget.charge_text(value.len(), part)?;
                if stack.is_empty() && !value.trim().is_empty() {
                    return Err(error(
                        reader.buffer_position(),
                        "text outside the root element".to_owned(),
                    ));
                }
            }
            Event::GeneralRef(reference) => {
                let decoded = reference
                    .decode()
                    .map_err(|err| error(reader.buffer_position(), err.to_string()))?;
                let character = if reference.is_char_ref() {
                    reference
                        .resolve_char_ref()
                        .map_err(|err| error(reader.buffer_position(), err.to_string()))?
                        .filter(|character| crate::xml::is_legal_xml_character(*character))
                } else {
                    quick_xml::escape::resolve_predefined_entity(&decoded)
                        .and_then(|value| value.chars().next())
                };
                let Some(character) = character else {
                    return Err(ParseError::UnsafeXml {
                        kind: "non-predefined or illegal entity reference",
                        part: part.to_owned(),
                    });
                };
                budget.charge_text(character.len_utf8(), part)?;
                if stack.is_empty() && !character.is_whitespace() {
                    return Err(error(
                        reader.buffer_position(),
                        "entity outside the root element".to_owned(),
                    ));
                }
            }
            Event::DocType(_) => {
                return Err(ParseError::UnsafeXml {
                    kind: "DTD/entity declarations are forbidden",
                    part: part.to_owned(),
                });
            }
            Event::Decl(ref declaration) if declares_an_encoding_other_than_utf8(declaration) => {
                return Err(error(0, "declared encoding other than UTF-8".to_owned()));
            }
            Event::Decl(_) | Event::PI(_) | Event::Comment(_) => {}
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(error(
            reader.buffer_position(),
            "unclosed element".to_owned(),
        ));
    }
    Ok(refused)
}

// Reader offsets index the input and raw attribute names are the decoded ones only for
// UTF-8 without a byte order mark.
fn starts_with_a_byte_order_mark(xml: &[u8]) -> bool {
    [&b"\xEF\xBB\xBF"[..], b"\xFE\xFF", b"\xFF\xFE"]
        .iter()
        .any(|bom| xml.starts_with(bom))
}

fn declares_an_encoding_other_than_utf8(declaration: &BytesDecl<'_>) -> bool {
    match declaration.encoding() {
        Some(Ok(encoding)) => !encoding.eq_ignore_ascii_case(b"utf-8"),
        Some(Err(_)) => true,
        None => false,
    }
}

pub(crate) fn body_prefix(xml: &[u8], keep: usize) -> Option<Vec<u8>> {
    if starts_with_a_byte_order_mark(xml) {
        return None;
    }
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().allow_dangling_amp = true;
    let mut stack = Vec::new();
    let mut body_start = None;
    let mut body_end = None;
    let mut children = 0;
    let mut child_start = 0;
    let mut kept_end = 0;
    let mut top_section = None;
    let mut final_section = None;
    let mut section_in_child = false;
    let mut ppr_seen = false;
    let mut ppr_depth = None;
    let mut wrapped_section = false;
    let mut child_is_wrapper = false;
    loop {
        let begin = reader.buffer_position() as usize;
        let event = reader.read_event().ok()?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref start) | Event::Empty(ref start) => {
                let name = start.name();
                let name = std::str::from_utf8(name.as_ref()).ok()?;
                let local = crate::xml::local_name(name);
                if stack.is_empty() && local != "document" {
                    return None;
                }
                if stack.len() == 1 && local == "body" {
                    if body_start.is_some() || matches!(event, Event::Empty(_)) {
                        return None;
                    }
                    body_start = Some(end);
                    kept_end = end;
                }
                if body_start.is_some() && body_end.is_none() {
                    if stack.len() == 2 {
                        children += 1;
                        child_start = begin;
                        section_in_child = false;
                        ppr_seen = false;
                        ppr_depth = None;
                        child_is_wrapper = matches!(local, "sdt" | "customXml" | "smartTag");
                    }
                    if stack.len() == 3 && stack.last() == Some(&"p") && local == "pPr" && !ppr_seen
                    {
                        ppr_seen = true;
                        if matches!(event, Event::Start(_)) {
                            ppr_depth = Some(4);
                        }
                    }
                    if local == "sectPr" {
                        if stack.len() == 4 && ppr_depth == Some(4) {
                            section_in_child = true;
                        }
                        if children > keep
                            && top_section.is_none()
                            && child_is_wrapper
                            && stack.last() == Some(&"pPr")
                            && stack.get(stack.len().saturating_sub(2)) == Some(&"p")
                        {
                            wrapped_section = true;
                        }
                    }
                    if matches!(event, Event::Empty(_)) && stack.len() == 2 {
                        if children <= keep {
                            kept_end = end;
                        }
                        if local == "sectPr" && final_section.is_none() {
                            final_section = Some(begin..end);
                        }
                    }
                }
                if matches!(event, Event::Start(_)) {
                    stack.push(match local {
                        "document" => "document",
                        "body" => "body",
                        "p" => "p",
                        "pPr" => "pPr",
                        "sectPr" => "sectPr",
                        _ => "",
                    });
                }
            }
            Event::End(_) => {
                if stack.len() == 3 && body_end.is_none() {
                    let local = *stack.last()?;
                    if children <= keep {
                        kept_end = end;
                    } else if local == "p" && section_in_child && top_section.is_none() {
                        top_section = Some(child_start..end);
                    }
                    if local == "sectPr" && final_section.is_none() {
                        final_section = Some(child_start..end);
                    }
                }
                if stack.len() == 4 && ppr_depth == Some(4) {
                    ppr_depth = None;
                }
                if stack.len() == 2 && stack.last() == Some(&"body") {
                    body_end = Some(begin);
                }
                stack.pop()?;
            }
            Event::Decl(ref declaration) if declares_an_encoding_other_than_utf8(declaration) => {
                return None;
            }
            Event::DocType(_) => return None,
            Event::Eof => break,
            _ => {}
        }
    }
    if !stack.is_empty() || children <= keep || wrapped_section {
        return None;
    }
    let mut prefix = xml[..kept_end].to_vec();
    if let Some(section) = top_section {
        prefix.extend_from_slice(&xml[section]);
    }
    if let Some(section) = final_section
        && section.start >= kept_end
    {
        prefix.extend_from_slice(&xml[section]);
    }
    prefix.extend_from_slice(&xml[body_end?..]);
    streaming_refuses_a_body_cut(&prefix).ok()?;
    Some(prefix)
}

/// Assemble the body below an already safe-parsed `w:document` root.
pub fn parse_document_body(
    document: &XmlElement,
    parser: &mut StoryParser<'_, '_>,
) -> Result<DocumentBody, ParseError> {
    parse_document_body_impl(document, parser, true, None, None).map(|(body, _)| body)
}

/// Parses a body without cloning blocks into section content.
pub(crate) fn parse_document_body_compact(
    document: &XmlElement,
    parser: &mut StoryParser<'_, '_>,
    body_blocks: Option<usize>,
) -> Result<DocumentBody, ParseError> {
    parse_document_body_impl(document, parser, false, body_blocks, None).map(|(body, _)| body)
}

pub(crate) fn parse_document_body_compact_with_read(
    document: &XmlElement,
    parser: &mut StoryParser<'_, '_>,
    body_blocks: Option<usize>,
    kept_children: usize,
) -> Result<(DocumentBody, usize), ParseError> {
    parse_document_body_impl(document, parser, false, body_blocks, Some(kept_children))
}

fn parse_document_body_impl(
    document: &XmlElement,
    parser: &mut StoryParser<'_, '_>,
    clone_section_content: bool,
    body_blocks: Option<usize>,
    read_limit: Option<usize>,
) -> Result<(DocumentBody, usize), ParseError> {
    if document.local_name() != "document" {
        return Ok((DocumentBody::default(), 0));
    }
    let Some(body) = document.child("w", "body") else {
        return Ok((DocumentBody::default(), 0));
    };
    let (content, read) =
        parser.parse_blocks_until_with_read_limit(body, 0, false, body_blocks, read_limit)?;
    // A body cut short ends inside the section whose properties the next
    // section-ending paragraph carries.
    let cut_section = body_blocks.and_then(|_| {
        transparent_children(body, false)
            .into_iter()
            .skip(read)
            .filter(|child| child.matches_name("w", "p"))
            .find_map(|paragraph| paragraph.child("w", "pPr")?.child("w", "sectPr"))
    });
    let final_section_properties = cut_section
        .or_else(|| body.child("w", "sectPr"))
        .map(|element| parse_section_properties(Some(element)));
    let mut sections = build_sections(
        &content,
        final_section_properties.as_ref(),
        clone_section_content,
    );
    let mut properties: Vec<_> = sections
        .iter()
        .map(|section| section.properties.clone())
        .collect();
    apply_section_inheritance(&mut properties);
    for (section, properties) in sections.iter_mut().zip(properties) {
        section.properties = properties;
    }
    Ok((
        DocumentBody {
            content,
            sections: Some(sections),
            final_section_properties,
            custom_root_bindings: custom_root_bindings(document),
            comments: None,
        },
        read,
    ))
}

/// Root bindings and attributes not regenerated by the standard boilerplate.
pub(crate) fn custom_root_bindings(document: &XmlElement) -> Vec<RawAttribute> {
    document
        .attributes
        .iter()
        .filter(|(name, _)| {
            name.strip_prefix("xmlns:")
                .map_or(name.as_str() != "xmlns", |prefix| {
                    !crate::serializer::parts::is_story_root_prefix(prefix)
                })
        })
        .map(|(name, value)| RawAttribute {
            name: name.clone(),
            value: value.clone(),
        })
        .collect()
}

fn build_sections(
    content: &[BlockContent],
    final_properties: Option<&SectionProperties>,
    clone_content: bool,
) -> Vec<Section> {
    let mut sections = Vec::new();
    let mut current = Vec::new();
    let mut current_len = 0usize;
    for block in content {
        current_len += 1;
        if clone_content {
            current.push(block.clone());
        }
        if let BlockContent::Paragraph(paragraph) = block
            && let Some(properties) = &paragraph.section_properties
        {
            sections.push(Section {
                id: None,
                properties: properties.clone(),
                content: std::mem::take(&mut current),
            });
            current_len = 0;
        }
    }
    if current_len > 0 || sections.is_empty() {
        sections.push(Section {
            id: None,
            properties: final_properties
                .cloned()
                .unwrap_or_else(default_section_properties),
            content: current,
        });
    }
    sections
}

pub fn get_paragraph_text(paragraph: &Paragraph) -> String {
    let mut text = String::new();
    for content in &paragraph.content {
        let ParagraphContent::Inline(content) = content else {
            continue;
        };
        match content {
            InlineNode::Run(run) => append_run_text(&run.content, &mut text, true),
            InlineNode::Hyperlink(link) => {
                for child in &link.children {
                    if let InlineNode::Run(run) = child {
                        append_run_text(&run.content, &mut text, false);
                    }
                }
            }
            InlineNode::SimpleField(field) => {
                for run in &field.content {
                    append_run_text(&run.content, &mut text, false);
                }
            }
            InlineNode::ComplexField(field) => {
                for run in &field.field_result {
                    append_run_text(&run.content, &mut text, false);
                }
            }
            _ => {}
        }
    }
    text
}

fn append_run_text(content: &[RunContent], output: &mut String, include_separators: bool) {
    for content in content {
        match content {
            RunContent::Text { text, .. } => output.push_str(text),
            RunContent::Tab if include_separators => output.push('\t'),
            RunContent::Break { break_type, .. } if include_separators => {
                output.push(if break_type.as_deref() == Some("page") {
                    '\u{000c}'
                } else {
                    '\n'
                });
            }
            _ => {}
        }
    }
}

pub fn is_empty_paragraph(paragraph: &Paragraph) -> bool {
    get_paragraph_text(paragraph).trim().is_empty()
        && !paragraph.content.iter().any(|content| {
            matches!(
                content,
                ParagraphContent::Inline(InlineNode::Run(run))
                    if run.content.iter().any(|content| matches!(
                        content,
                        RunContent::Drawing { .. } | RunContent::Shape { .. }
                    ))
            )
        })
}

pub fn extract_template_variables(text: &str) -> Vec<String> {
    let bytes = text.as_bytes();
    let mut variables = Vec::new();
    let mut cursor = 0usize;
    while cursor < bytes.len() {
        let Some(relative) = bytes[cursor..].iter().position(|byte| *byte == b'{') else {
            break;
        };
        let start = cursor + relative + 1;
        let Some(relative_end) = bytes[start..].iter().position(|byte| *byte == b'}') else {
            break;
        };
        let end = start + relative_end;
        let candidate = &text[start..end];
        let valid = candidate
            .bytes()
            .enumerate()
            .all(|(index, byte)| match byte {
                b'A'..=b'Z' | b'a'..=b'z' | b'_' => true,
                b'0'..=b'9' if index > 0 => true,
                b'-' | b'.' if index > 0 => true,
                _ => false,
            });
        if valid && !candidate.is_empty() && !variables.iter().any(|value| value == candidate) {
            variables.push(candidate.to_owned());
            cursor = end.saturating_add(1);
        } else {
            // A failed match restarts at the second brace.
            cursor = start;
        }
    }
    variables
}

pub fn extract_all_template_variables(content: &[BlockContent]) -> Vec<String> {
    let mut variables = Vec::new();
    for block in content {
        match block {
            BlockContent::Paragraph(paragraph) => {
                for variable in extract_template_variables(&get_paragraph_text(paragraph)) {
                    if !variables.contains(&variable) {
                        variables.push(variable);
                    }
                }
            }
            BlockContent::Table(table) => {
                for variable in extract_table_template_variables(table) {
                    if !variables.contains(&variable) {
                        variables.push(variable);
                    }
                }
            }
            // Block SDTs are not traversed.
            BlockContent::BlockSdt(_) => {}
            BlockContent::RawXml(_) => {}
        }
    }
    variables
}

fn extract_table_template_variables(table: &crate::table::Table) -> Vec<String> {
    let mut variables = Vec::new();
    for row in &table.rows {
        for cell in &row.cells {
            for block in &cell.content {
                match block {
                    BlockContent::Paragraph(paragraph) => {
                        for variable in extract_template_variables(&get_paragraph_text(paragraph)) {
                            if !variables.contains(&variable) {
                                variables.push(variable);
                            }
                        }
                    }
                    BlockContent::Table(table) => {
                        for variable in extract_table_template_variables(table) {
                            if !variables.contains(&variable) {
                                variables.push(variable);
                            }
                        }
                    }
                    BlockContent::BlockSdt(_) => {}
                    BlockContent::RawXml(_) => {}
                }
            }
        }
    }
    variables
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart::ChartPartsMap;
    use crate::media::MediaMap;
    use crate::paragraph::HexIdAllocator;
    use crate::s9::{
        S9ParseOptions, parse_docx_s9_preview_from_parts, parse_docx_s9_preview_from_parts_full_dom,
    };
    use crate::smart_art::SmartArtContext;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn parse(xml: &str) -> DocumentBody {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(xml.as_bytes(), "word/document.xml", &mut budget).unwrap();
        let media = MediaMap::new();
        let charts = ChartPartsMap::new();
        let mut smart_art = SmartArtContext::default();
        let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64)).unwrap();
        let mut parser = StoryParser {
            relationships: None,
            theme: None,
            styles: None,
            doc_defaults: None,
            numbering: None,
            media: &media,
            charts: &charts,
            smart_art: &mut smart_art,
            budget: &mut budget,
            ids: &mut ids,
            part: "word/document.xml",
        };
        parse_document_body(document.root().unwrap(), &mut parser).unwrap()
    }

    #[test]
    fn builds_sections_inherits_story_refs_and_keeps_the_final_section_as_authored() {
        let body = parse(
            r#"<w:document xmlns:w="w" xmlns:r="r"><w:body>
              <w:p><w:r><w:t>{first}</w:t></w:r><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rH"/><w:titlePg/></w:sectPr></w:pPr></w:p>
              <w:p><w:r><w:t>{first} {second-name}</w:t></w:r></w:p>
              <w:sectPr><w:pgMar w:left="720"/></w:sectPr>
            </w:body></w:document>"#,
        );
        let sections = body.sections.as_ref().unwrap();
        assert_eq!(sections.len(), 2);
        assert_eq!(sections[1].properties.margin_left, Some(720.0));
        assert_eq!(sections[1].properties.title_pg, None);
        assert_eq!(
            sections[1].properties.header_references.as_ref().unwrap()[0].relationship_id,
            "rH"
        );
        // Kept as authored; header inheritance is resolved through `sections`.
        let final_properties = body.final_section_properties.as_ref().unwrap();
        assert_eq!(final_properties.margin_left, Some(720.0));
        assert!(final_properties.header_references.is_none());
        assert!(final_properties.title_pg.is_none());
        assert_eq!(
            extract_all_template_variables(&body.content),
            ["first", "second-name"]
        );
    }

    #[test]
    fn refuses_a_cut_of_a_body_with_floats_off_the_text_or_columns() {
        let refused = |body: &str| {
            let limits = ParseLimits::default();
            let mut budget = ParseBudget::new(&limits);
            let xml = format!(
                r#"<w:document xmlns:w="w" xmlns:wp="wp" xmlns:v="v"><w:body>{body}</w:body></w:document>"#
            );
            let document = parse_xml(xml.as_bytes(), "word/document.xml", &mut budget).unwrap();
            let refused = refuses_a_body_cut(document.root().unwrap());
            assert_eq!(
                streaming_refuses_a_body_cut(xml.as_bytes()).unwrap(),
                refused,
                "{body}"
            );
            refused
        };
        let anchor = |attributes: &str, vertical: &str| {
            format!(
                "<w:p><w:r><w:drawing><wp:anchor {attributes}>{vertical}</wp:anchor></w:drawing></w:r></w:p>"
            )
        };
        let table = |attributes: &str| {
            format!("<w:tbl><w:tblPr><w:tblpPr {attributes}/></w:tblPr></w:tbl>")
        };
        let frame =
            |attributes: &str| format!("<w:p><w:pPr><w:framePr {attributes}/></w:pPr></w:p>");
        let section =
            |columns: &str| format!("<w:p><w:pPr><w:sectPr>{columns}</w:sectPr></w:pPr></w:p>");
        let vml = |style: &str| {
            format!(r#"<w:p><w:r><w:pict><v:shape style="{style}"/></w:pict></w:r></w:p>"#)
        };
        for body in [
            anchor("", r#"<wp:positionV relativeFrom="margin"/>"#),
            anchor("", r#"<wp:positionV relativeFrom="page"/>"#),
            anchor("", r#"<wp:positionV relativeFrom="pag&#101;"/>"#),
            anchor(
                "",
                r#"<wp:positionV relativeFrom="margin"/><wp:positionV relativeFrom="line"/>"#,
            ),
            anchor("", r#"<wp:positionV relativeFrom="topMargin"/>"#),
            anchor(
                r#"simplePos="1""#,
                r#"<wp:positionV relativeFrom="paragraph"/>"#,
            ),
            table(r#"w:vertAnchor="margin""#),
            table(r#"w:tblpY="60""#),
            frame(r#"w:vAnchor="page""#),
            frame(r#"w:w="2000""#),
            vml("position:absolute;mso-position-vertical-relative:margin"),
            vml("position:absolute;mso-position-vertical-relative:page"),
            section(r#"<w:cols w:num="2"/>"#),
            section(r#"<w:cols w:num="1"/><w:cols w:num="2"/>"#),
            section(r#"<w:cols><x:col xmlns:x="x"/><x:col xmlns:x="x"/></w:cols>"#),
            section(r#"<w:cols num="2" w:num="1"/><w:cols num="1" w:num="2"/>"#),
            section(r#"<w:cols><w:col w:w="3000"/><w:col w:w="3000"/></w:cols>"#),
        ] {
            assert!(refused(&body), "{body}");
        }
        for body in [
            anchor("", r#"<wp:positionV relativeFrom="paragraph"/>"#),
            anchor("", r#"<wp:positionV relativeFrom="line"/>"#),
            anchor(
                "",
                r#"<wp:wrap><wp:positionV relativeFrom="page"/></wp:wrap>"#,
            ),
            anchor("", r#"<wp:positionV/><wp:positionV relativeFrom="page"/>"#),
            anchor(
                "",
                r#"<wp:positionV relativeFrom="line"/><wp:positionV relativeFrom="page"/>"#,
            ),
            section(r#"<w:cols num="2" w:num="1"/>"#),
            vml("POSITION:absolute;mso-position-vertical-relative:page;position:static"),
            vml(
                "position:absolute;mso-position-vertical-relative:page;mso-position-vertical-relative:text",
            ),
            anchor("", ""),
            table(r#"w:vertAnchor="text""#),
            frame(r#"w:dropCap="drop" w:vAnchor="text""#),
            vml("position:absolute;mso-position-vertical-relative:text"),
            vml("width:300pt;height:165pt"),
            section(r#"<w:cols w:num="1" w:space="720"/>"#),
            section(r#"<w:cols w:space="720"/>"#),
        ] {
            assert!(!refused(&body), "{body}");
        }
    }

    #[test]
    fn streaming_cut_rejects_malformed_xml_even_after_a_refusal() {
        for xml in [
            "<w:document><w:body><w:p>",
            "<w:document><w:body></w:document>",
            "<w:document><w:body><wp:anchor simplePos=\"1\"/>",
            "<w:document a=\"1\" a=\"2\"/>",
            "<w:document a=\"<\"/>",
            "<w:document/><other/>",
            "<w:document>&unknown;</w:document>",
            "<!DOCTYPE document><w:document/>",
        ] {
            let limits = ParseLimits::default();
            assert!(
                parse_xml(
                    xml.as_bytes(),
                    "word/document.xml",
                    &mut ParseBudget::new(&limits)
                )
                .is_err(),
                "{xml}"
            );
            assert!(
                streaming_refuses_a_body_cut(xml.as_bytes()).is_err(),
                "{xml}"
            );
            assert!(body_prefix(xml.as_bytes(), 0).is_none(), "{xml}");
        }
    }

    fn drawing_preview_xml(off_text: bool) -> Vec<u8> {
        let body: String = (0..620)
            .map(|index| {
                let drawing = match index {
                    0 => r#"<w:pict><legacy:shape style="position:absolute;mso-position-vertical-relative:text;width:12pt;height:12pt"/></w:pict>"#,
                    1 => r#"<w:object><legacy:shape style="width:12pt;height:12pt"/><office:OLEObject Type="Embed" ProgID="Word.Document.12"/></w:object>"#,
                    550 if off_text => r#"<w:pict><legacy:shape style="position:absolute;mso-position-vertical-relative:page"/></w:pict>"#,
                    _ => "",
                };
                format!("<w:p><w:r>{drawing}<w:t>paragraph {index}</w:t></w:r></w:p>")
            })
            .collect();
        format!(
            r#"<w:document xmlns:w="{}" xmlns:legacy="{}" xmlns:office="{}"><w:body>{body}</w:body></w:document>"#,
            crate::xml::namespaces::W,
            crate::xml::namespaces::V,
            crate::xml::namespaces::O,
        )
        .into_bytes()
    }

    #[test]
    fn drawing_aliases_use_the_prefix_preview_and_match_the_full_dom() {
        let xml = drawing_preview_xml(false);
        let limits = ParseLimits::default();
        assert!(!streaming_body_cut(&xml, &mut ParseBudget::new(&limits)).unwrap());
        let parts = vec![("word/document.xml".to_owned(), xml.clone())];
        for blocks in [1usize, 50, 200] {
            let keep = blocks.saturating_mul(2).max(blocks.saturating_add(64));
            let prefix = body_prefix(&xml, keep).unwrap();
            assert!(prefix.len() < xml.len());
            let parsed =
                parse_xml(&prefix, "word/document.xml", &mut ParseBudget::new(&limits)).unwrap();
            let root = parsed.root().unwrap();
            assert_eq!(
                root.child("w", "body").unwrap().child_elements().count(),
                keep
            );
            assert_eq!(
                root.find_deep("w", "pict")
                    .unwrap()
                    .attributes
                    .get("xmlns:legacy")
                    .map(String::as_str),
                Some(crate::xml::namespaces::V)
            );
            assert_eq!(
                root.find_deep("w", "object")
                    .unwrap()
                    .attributes
                    .get("xmlns:office")
                    .map(String::as_str),
                Some(crate::xml::namespaces::O)
            );
            for source_ordinals in [false, true] {
                let options = S9ParseOptions {
                    determinism_seed: Some("7".repeat(64)),
                    source_ordinals,
                    ..S9ParseOptions::default()
                };
                let actual =
                    parse_docx_s9_preview_from_parts(&parts, blocks, options.clone(), &limits)
                        .unwrap();
                let expected =
                    parse_docx_s9_preview_from_parts_full_dom(&parts, blocks, options, &limits)
                        .unwrap();
                assert!(actual.is_some());
                assert_eq!(
                    serde_json::to_value(actual).unwrap(),
                    serde_json::to_value(expected).unwrap(),
                    "blocks={blocks}, source_ordinals={source_ordinals}"
                );
            }
        }
    }

    #[test]
    fn drawing_alias_attribute_bytes_match_the_full_dom_error() {
        let max_attribute_bytes = format!(
            "xmlns:w{}xmlns:legacy{}",
            crate::xml::namespaces::W,
            crate::xml::namespaces::V,
        )
        .len();
        let limits = ParseLimits {
            max_attribute_bytes,
            ..ParseLimits::default()
        };
        let padding = "x".repeat(max_attribute_bytes - "id".len());
        let xml = format!(
            r#"<w:document xmlns:w="{}" xmlns:legacy="{}"><w:body><w:p><w:r><w:pict id="{padding}"><legacy:shape/></w:pict></w:r></w:p></w:body></w:document>"#,
            crate::xml::namespaces::W,
            crate::xml::namespaces::V,
        );
        let expected = ParseError::ResourceLimit {
            kind: "drawingNamespaceAliases",
            part: "word/document.xml".to_owned(),
        };
        assert_eq!(
            streaming_body_cut(xml.as_bytes(), &mut ParseBudget::new(&limits)).unwrap_err(),
            expected
        );
        assert_eq!(
            parse_xml(
                xml.as_bytes(),
                "word/document.xml",
                &mut ParseBudget::new(&limits)
            )
            .unwrap_err(),
            expected
        );
        let parts = vec![("word/document.xml".to_owned(), xml.into_bytes())];
        let options = S9ParseOptions {
            determinism_seed: Some("7".repeat(64)),
            ..S9ParseOptions::default()
        };
        assert_eq!(
            parse_docx_s9_preview_from_parts(&parts, 1, options.clone(), &limits).unwrap_err(),
            expected
        );
        assert_eq!(
            parse_docx_s9_preview_from_parts_full_dom(&parts, 1, options, &limits).unwrap_err(),
            expected
        );
    }

    #[test]
    fn a_vml_shape_off_the_text_after_the_cut_refuses_both_preview_paths() {
        let xml = drawing_preview_xml(true);
        let limits = ParseLimits::default();
        assert!(streaming_body_cut(&xml, &mut ParseBudget::new(&limits)).unwrap());
        let parts = vec![("word/document.xml".to_owned(), xml.clone())];
        for blocks in [1usize, 50, 200] {
            let keep = blocks.saturating_mul(2).max(blocks.saturating_add(64));
            let prefix = body_prefix(&xml, keep).unwrap();
            assert!(!streaming_refuses_a_body_cut(&prefix).unwrap());
            let options = S9ParseOptions {
                determinism_seed: Some("7".repeat(64)),
                ..S9ParseOptions::default()
            };
            assert!(
                parse_docx_s9_preview_from_parts(&parts, blocks, options.clone(), &limits)
                    .unwrap()
                    .is_none()
            );
            assert!(
                parse_docx_s9_preview_from_parts_full_dom(&parts, blocks, options, &limits)
                    .unwrap()
                    .is_none()
            );
        }
    }

    #[test]
    fn nested_drawing_alias_budgets_match_the_full_dom() {
        let xml = format!(
            r#"<w:document xmlns:w="{}" xmlns:v="urn:old-v" xmlns:mc="{}" xmlns="urn:default" xmlns:ext="urn:outer" xmlns:extra="urn:extra" xmlns:Requires="extra" xmlns:a="urn:a" xmlns:unused="urn:unused"><w:body xmlns:v="{}" xmlns:ext="urn:inner"><w:pict xmlns:ext="urn:own" mc:Ignorable="ext" mc:MustUnderstand="a" mc:Requires="a:item" mc:ProcessContent="ext:item" mc:PreserveElements="a:*" mc:PreserveAttributes="ext:*"><w:object><Requires:marker ext:value="x"><leaf/></Requires:marker><w:pict xmlns:a="urn:own-a" a:value="x"><v:shape/></w:pict><w:object ext:value="x"/></w:object></w:pict></w:body></w:document>"#,
            crate::xml::namespaces::W,
            crate::xml::namespaces::MC,
            crate::xml::namespaces::V,
        );
        let defaults = ParseLimits::default();
        let document = parse_xml(
            xml.as_bytes(),
            "word/document.xml",
            &mut ParseBudget::new(&defaults),
        )
        .unwrap();
        let outer = document.root().unwrap().find_deep("w", "pict").unwrap();
        assert_eq!(
            outer.attributes.get("xmlns:extra").map(String::as_str),
            Some("urn:extra")
        );
        assert!(!outer.attributes.contains_key("xmlns:v"));
        assert!(!outer.attributes.contains_key("xmlns:unused"));
        assert!(
            !outer
                .child("w", "object")
                .unwrap()
                .attributes
                .contains_key("xmlns:extra")
        );
        let mut text_bytes = 0usize;
        let mut pending = vec![document.root().unwrap()];
        while let Some(element) = pending.pop() {
            text_bytes += element.name.len();
            text_bytes += element
                .attributes
                .iter()
                .map(|(name, value)| name.len() + value.len())
                .sum::<usize>();
            for node in &element.children {
                match node {
                    crate::xml::XmlNode::Element(child) => pending.push(child),
                    crate::xml::XmlNode::Text(value) | crate::xml::XmlNode::CData(value) => {
                        text_bytes += value.len();
                    }
                }
            }
        }
        let mut cases = Vec::new();
        for max_attributes_per_element in 0..=16 {
            cases.push(ParseLimits {
                max_attributes_per_element,
                ..ParseLimits::default()
            });
        }
        for max_attribute_bytes in [80, 160, 240, 320, 500, 1000] {
            cases.push(ParseLimits {
                max_attribute_bytes,
                ..ParseLimits::default()
            });
        }
        for max_xml_text_bytes in [text_bytes - 1, text_bytes, text_bytes + 1] {
            cases.push(ParseLimits {
                max_xml_text_bytes,
                ..ParseLimits::default()
            });
        }
        for limits in cases {
            let mut streaming_budget = ParseBudget::new(&limits);
            let mut dom_budget = ParseBudget::new(&limits);
            let actual = streaming_body_cut(xml.as_bytes(), &mut streaming_budget);
            let expected = parse_xml(xml.as_bytes(), "word/document.xml", &mut dom_budget)
                .map(|document| refuses_a_body_cut(document.root().unwrap()));
            assert_eq!(actual, expected, "{limits:?}");
            assert_eq!(streaming_budget.xml_events(), dom_budget.xml_events());
            if actual.is_ok() {
                assert_eq!(
                    streaming_budget
                        .charge_text(limits.max_xml_text_bytes - text_bytes, "word/document.xml"),
                    dom_budget
                        .charge_text(limits.max_xml_text_bytes - text_bytes, "word/document.xml")
                );
                assert!(
                    streaming_budget
                        .charge_text(1, "word/document.xml")
                        .is_err()
                );
            }
        }
    }

    #[test]
    fn body_prefix_keeps_the_next_section_and_the_authored_final_section() {
        let xml = br#"<?xml version="1.0"?><w:document xmlns:w="w"><w:body a="&amp;">
            <w:p><w:r><w:t>first & stray</w:t></w:r></w:p><w:customXml><w:p/></w:customXml>
            <w:p/><w:p><w:pPr><w:sectPr><w:pgSz w:w="10000"/></w:sectPr></w:pPr></w:p>
            <w:p/><w:sectPr><w:pgSz w:w="20000"/></w:sectPr>
            </w:body></w:document>"#;
        let prefix = body_prefix(xml, 2).unwrap();
        let limits = ParseLimits::default();
        let document =
            parse_xml(&prefix, "word/document.xml", &mut ParseBudget::new(&limits)).unwrap();
        let body = document.root().unwrap().child("w", "body").unwrap();
        assert_eq!(body.child_elements().count(), 4);
        assert_eq!(body.children_named("w", "p").count(), 2);
        assert!(String::from_utf8(prefix).unwrap().contains("first & stray"));
        assert!(body_prefix(xml, 6).is_none());
        assert!(body_prefix(xml, 7).is_none());
    }

    #[test]
    fn the_streaming_cut_falls_back_for_a_bom_or_a_declared_encoding_other_than_utf8() {
        let body = "<w:body><w:p/><w:p/><w:p/><w:sectPr/></w:body></w:document>";
        let plain = format!("<w:document xmlns:w=\"w\">{body}");
        assert!(body_prefix(plain.as_bytes(), 1).is_some());
        assert!(!streaming_refuses_a_body_cut(plain.as_bytes()).unwrap());
        let declared = format!("<?xml version=\"1.0\" encoding=\"UTF-8\"?>{plain}");
        assert!(body_prefix(declared.as_bytes(), 1).is_some());
        assert!(!streaming_refuses_a_body_cut(declared.as_bytes()).unwrap());
        let bom = [b"\xEF\xBB\xBF".as_slice(), plain.as_bytes()].concat();
        assert!(body_prefix(&bom, 1).is_none());
        assert!(streaming_refuses_a_body_cut(&bom).is_err());
        let other = format!("<?xml version=\"1.0\" encoding=\"ISO-2022-JP\"?>{plain}");
        assert!(body_prefix(other.as_bytes(), 1).is_none());
        assert!(streaming_refuses_a_body_cut(other.as_bytes()).is_err());
        let refusing = "<?xml version='1.0' encoding='ISO-2022-JP'?><w:document xmlns:w='w' \
            xmlns:wp='wp'><w:body><wp:anchor simplePos='1'/></w:body></w:document>";
        assert!(streaming_refuses_a_body_cut(refusing.as_bytes()).is_err());
    }

    #[test]
    fn body_prefix_falls_back_for_wrapped_section_breaks_before_the_next_top_level_break() {
        for wrapper in ["sdt", "customXml", "smartTag"] {
            let xml = format!(
                "<w:document><w:body><w:p/><w:{wrapper}><w:sdtContent><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:sdtContent></w:{wrapper}><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:body></w:document>"
            );
            assert!(body_prefix(xml.as_bytes(), 1).is_none(), "{wrapper}");
        }
        let xml = b"<w:document><w:body><w:p/><w:p><w:pPr><w:sectPr/></w:pPr></w:p><w:sdt><w:sdtContent><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:sdtContent></w:sdt></w:body></w:document>";
        assert!(body_prefix(xml, 1).is_some());
    }

    #[test]
    fn normalizes_an_empty_body_to_one_default_empty_section() {
        let body = parse(r#"<w:document xmlns:w="w"><w:body/></w:document>"#);
        assert!(body.content.is_empty());
        let sections = body.sections.unwrap();
        assert_eq!(sections.len(), 1);
        assert!(sections[0].content.is_empty());
        assert_eq!(sections[0].properties.page_width, Some(12_240.0));
        assert!(body.final_section_properties.is_none());
    }

    #[test]
    fn paragraph_text_and_empty_query_use_shallow_run_grammar() {
        let body = parse(
            r#"<w:document xmlns:w="w"><w:body><w:p><w:r><w:t> a </w:t><w:tab/><w:br/></w:r></w:p><w:p/></w:body></w:document>"#,
        );
        let BlockContent::Paragraph(first) = &body.content[0] else {
            panic!("paragraph")
        };
        assert_eq!(get_paragraph_text(first), " a \t\n");
        let BlockContent::Paragraph(empty) = &body.content[1] else {
            panic!("paragraph")
        };
        assert!(is_empty_paragraph(empty));
    }
}
