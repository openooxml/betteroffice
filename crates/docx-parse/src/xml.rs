//! Bounded, resolver-free OOXML event/tree core.

use indexmap::IndexMap;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use thiserror::Error;

/// Incumbent block/inline container limit from `parseLimits.ts`.
pub const MAX_NESTING_DEPTH: usize = 64;

/// OOXML namespaces mirrored from `xmlParser.ts` for later parser slices.
pub mod namespaces {
    pub const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    pub const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    pub const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    pub const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    pub const WP14: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
    pub const WPS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    pub const WPC: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
    pub const WPG: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    pub const PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    pub const M: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    pub const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    pub const V: &str = "urn:schemas-microsoft-com:vml";
    pub const O: &str = "urn:schemas-microsoft-com:office:office";
    pub const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
    pub const W15: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
    pub const CT: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
    pub const PR: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
}

/// Conservative defaults for the S1 XML/relationship foundation.
#[derive(Clone, Debug)]
pub struct ParseLimits {
    pub max_xml_bytes: usize,
    pub max_xml_events: usize,
    pub max_xml_text_bytes: usize,
    pub max_xml_depth: usize,
    pub max_attributes_per_element: usize,
    pub max_attribute_bytes: usize,
    pub max_relationships: usize,
    pub max_leaf_values: usize,
    pub max_blocks: usize,
    pub max_paragraphs: usize,
    pub max_tables: usize,
    pub max_table_rows: usize,
    pub max_table_cells: usize,
    pub max_notes: usize,
    pub max_comments: usize,
    pub max_nesting_depth: usize,
}

impl Default for ParseLimits {
    fn default() -> Self {
        Self {
            max_xml_bytes: 128 * 1024 * 1024,
            max_xml_events: 4_000_000,
            max_xml_text_bytes: 128 * 1024 * 1024,
            max_xml_depth: 256,
            max_attributes_per_element: 1024,
            max_attribute_bytes: 4 * 1024 * 1024,
            max_relationships: 250_000,
            max_leaf_values: 1_000_000,
            max_blocks: 500_000,
            max_paragraphs: 500_000,
            max_tables: 100_000,
            max_table_rows: 500_000,
            max_table_cells: 2_000_000,
            max_notes: 100_000,
            max_comments: 100_000,
            max_nesting_depth: MAX_NESTING_DEPTH,
        }
    }
}

/// One package-wide budget shared across every XML part.
#[derive(Debug)]
pub struct ParseBudget<'a> {
    limits: &'a ParseLimits,
    xml_bytes: usize,
    xml_events: usize,
    xml_text_bytes: usize,
    relationships: usize,
    leaf_values: usize,
    blocks: usize,
    paragraphs: usize,
    tables: usize,
    table_rows: usize,
    table_cells: usize,
    notes: usize,
    comments: usize,
}

impl<'a> ParseBudget<'a> {
    pub fn new(limits: &'a ParseLimits) -> Self {
        Self {
            limits,
            xml_bytes: 0,
            xml_events: 0,
            xml_text_bytes: 0,
            relationships: 0,
            leaf_values: 0,
            blocks: 0,
            paragraphs: 0,
            tables: 0,
            table_rows: 0,
            table_cells: 0,
            notes: 0,
            comments: 0,
        }
    }

    pub(crate) fn charge_xml_bytes(&mut self, amount: usize, part: &str) -> Result<(), ParseError> {
        charge(
            &mut self.xml_bytes,
            amount,
            self.limits.max_xml_bytes,
            "xmlBytes",
            part,
        )
    }

    pub(crate) fn charge_event(&mut self, part: &str) -> Result<(), ParseError> {
        charge(
            &mut self.xml_events,
            1,
            self.limits.max_xml_events,
            "xmlEvents",
            part,
        )
    }

    pub(crate) fn charge_text(&mut self, amount: usize, part: &str) -> Result<(), ParseError> {
        charge(
            &mut self.xml_text_bytes,
            amount,
            self.limits.max_xml_text_bytes,
            "xmlTextBytes",
            part,
        )
    }

    pub(crate) fn charge_relationship(&mut self, part: &str) -> Result<(), ParseError> {
        charge(
            &mut self.relationships,
            1,
            self.limits.max_relationships,
            "relationships",
            part,
        )
    }

    pub fn charge_leaf_value(&mut self, part: &str) -> Result<(), ParseError> {
        charge(
            &mut self.leaf_values,
            1,
            self.limits.max_leaf_values,
            "leafValues",
            part,
        )
    }

    /// Charge one block before allocating its model node.
    pub fn charge_block(&mut self, part: &str) -> Result<(), ParseError> {
        charge(&mut self.blocks, 1, self.limits.max_blocks, "blocks", part)
    }

    /// Charge one paragraph before parsing its inline subtree.
    pub fn charge_paragraph(&mut self, part: &str) -> Result<(), ParseError> {
        charge(
            &mut self.paragraphs,
            1,
            self.limits.max_paragraphs,
            "paragraphs",
            part,
        )
    }

    /// Charge one table before parsing its grid and rows.
    pub fn charge_table(&mut self, part: &str) -> Result<(), ParseError> {
        charge(&mut self.tables, 1, self.limits.max_tables, "tables", part)
    }

    /// Charge one table row before allocating its model node.
    pub fn charge_table_row(&mut self, part: &str) -> Result<(), ParseError> {
        charge(
            &mut self.table_rows,
            1,
            self.limits.max_table_rows,
            "tableRows",
            part,
        )
    }

    /// Charge one table cell before descending into its story.
    pub fn charge_table_cell(&mut self, part: &str) -> Result<(), ParseError> {
        charge(
            &mut self.table_cells,
            1,
            self.limits.max_table_cells,
            "tableCells",
            part,
        )
    }

    /// Charge one footnote/endnote before allocating its story model.
    pub fn charge_note(&mut self, part: &str) -> Result<(), ParseError> {
        charge(&mut self.notes, 1, self.limits.max_notes, "notes", part)
    }

    /// Charge one comment before allocating its metadata and story model.
    pub fn charge_comment(&mut self, part: &str) -> Result<(), ParseError> {
        charge(
            &mut self.comments,
            1,
            self.limits.max_comments,
            "comments",
            part,
        )
    }

    /// Check the shared story/container depth before descending.
    pub fn check_nesting_depth(&self, depth: usize, part: &str) -> Result<(), ParseError> {
        if depth > self.limits.max_nesting_depth {
            return Err(ParseError::ResourceLimit {
                kind: "nestingDepth",
                part: part.to_owned(),
            });
        }
        Ok(())
    }
}

fn charge(
    used: &mut usize,
    amount: usize,
    limit: usize,
    kind: &'static str,
    part: &str,
) -> Result<(), ParseError> {
    let next = used
        .checked_add(amount)
        .ok_or_else(|| ParseError::ResourceLimit {
            kind,
            part: part.to_owned(),
        })?;
    if next > limit {
        return Err(ParseError::ResourceLimit {
            kind,
            part: part.to_owned(),
        });
    }
    *used = next;
    Ok(())
}

#[derive(Debug, Error, PartialEq, Eq)]
pub enum ParseError {
    #[error("container error: {0}")]
    Container(String),
    #[error("unsafe XML in {part}: {kind}")]
    UnsafeXml { kind: &'static str, part: String },
    #[error("resource limit {kind} exceeded in {part}")]
    ResourceLimit { kind: &'static str, part: String },
    #[error("malformed XML in {part} at byte {offset}: {message}")]
    MalformedXml {
        part: String,
        offset: u64,
        message: String,
    },
    #[error("invalid relationship in {part}: {message}")]
    Relationship { part: String, message: String },
    #[error("canonical encoding failed: {0}")]
    Canonical(String),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlDocument {
    pub roots: Vec<XmlElement>,
}

impl XmlDocument {
    pub fn root(&self) -> Option<&XmlElement> {
        self.roots.first()
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XmlNode {
    Element(XmlElement),
    Text(String),
    CData(String),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlElement {
    pub name: String,
    pub attributes: IndexMap<String, String>,
    pub children: Vec<XmlNode>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ParsedColor {
    pub value: Option<String>,
    pub theme_color: Option<String>,
    pub theme_tint: Option<String>,
    pub theme_shade: Option<String>,
}

impl XmlElement {
    pub fn local_name(&self) -> &str {
        local_name(&self.name)
    }

    pub fn namespace_prefix(&self) -> Option<&str> {
        namespace_prefix(&self.name)
    }

    pub fn matches_name(&self, namespace: &str, local: &str) -> bool {
        self.name == format!("{namespace}:{local}") || self.local_name() == local
    }

    pub fn child(&self, namespace: &str, local: &str) -> Option<&XmlElement> {
        self.child_elements()
            .find(|child| child.matches_name(namespace, local))
    }

    pub fn children_named<'a>(
        &'a self,
        namespace: &'a str,
        local: &'a str,
    ) -> impl Iterator<Item = &'a XmlElement> + 'a {
        self.child_elements()
            .filter(move |child| child.matches_name(namespace, local))
    }

    pub fn child_by_local_name(&self, local: &str) -> Option<&XmlElement> {
        self.child_elements()
            .find(|child| child.local_name() == local)
    }

    pub fn children_by_local_name<'a>(
        &'a self,
        local: &'a str,
    ) -> impl Iterator<Item = &'a XmlElement> + 'a {
        self.child_elements()
            .filter(move |child| child.local_name() == local)
    }

    pub fn child_by_full_name(&self, full_name: &str) -> Option<&XmlElement> {
        self.child_elements().find(|child| child.name == full_name)
    }

    pub fn child_elements(&self) -> impl Iterator<Item = &XmlElement> {
        self.children.iter().filter_map(|node| match node {
            XmlNode::Element(element) => Some(element),
            XmlNode::Text(_) | XmlNode::CData(_) => None,
        })
    }

    pub fn attribute(&self, namespace: Option<&str>, name: &str) -> Option<&str> {
        namespace
            .and_then(|namespace| self.attributes.get(&format!("{namespace}:{name}")))
            .or_else(|| self.attributes.get(name))
            .map(String::as_str)
    }

    pub fn attribute_any<'a>(&'a self, names: &[&str]) -> Option<&'a str> {
        names
            .iter()
            .find_map(|name| self.attributes.get(*name).map(String::as_str))
    }

    pub fn text_content(&self) -> String {
        let mut text = String::new();
        self.append_text(&mut text);
        text
    }

    fn append_text(&self, output: &mut String) {
        for child in &self.children {
            match child {
                XmlNode::Text(text) | XmlNode::CData(text) => output.push_str(text),
                XmlNode::Element(element) => element.append_text(output),
            }
        }
    }

    pub fn has_flag(&self, namespace: Option<&str>, name: &str) -> bool {
        !matches!(
            self.attribute(namespace, name),
            None | Some("0" | "false" | "off")
        )
    }

    pub fn has_child(&self, namespace: &str, local: &str) -> bool {
        self.child(namespace, local).is_some()
    }

    pub fn parse_color(&self) -> ParsedColor {
        ParsedColor {
            value: self.attribute(Some("w"), "val").map(str::to_owned),
            theme_color: self.attribute(Some("w"), "themeColor").map(str::to_owned),
            theme_tint: self.attribute(Some("w"), "themeTint").map(str::to_owned),
            theme_shade: self.attribute(Some("w"), "themeShade").map(str::to_owned),
        }
    }

    pub fn parse_numeric_attribute(
        &self,
        namespace: Option<&str>,
        name: &str,
        scale: f64,
    ) -> Option<f64> {
        parse_javascript_integer_prefix(self.attribute(namespace, name)?).map(|value| value * scale)
    }

    pub fn parse_boolean(&self, namespace: &str) -> bool {
        !matches!(
            self.attribute(Some(namespace), "val"),
            Some("0" | "false" | "off")
        )
    }

    pub fn find_deep(&self, namespace: &str, local: &str) -> Option<&XmlElement> {
        if self.matches_name(namespace, local) {
            return Some(self);
        }
        self.child_elements()
            .find_map(|child| child.find_deep(namespace, local))
    }

    pub fn find_all_deep<'a>(
        &'a self,
        namespace: &str,
        local: &str,
        output: &mut Vec<&'a XmlElement>,
    ) {
        if self.matches_name(namespace, local) {
            output.push(self);
        }
        for child in self.child_elements() {
            child.find_all_deep(namespace, local, output);
        }
    }

    /// Structural XML serialization with text/attribute escaping by construction.
    pub fn to_xml(&self) -> String {
        let mut output = String::new();
        self.write_xml(&mut output);
        output
    }

    /// Match xml-js/js2xml's incumbent non-compact serialization used by
    /// public raw inline XML fields. Its attribute encoder only protects the
    /// surrounding double quote; decoded ampersands, angle brackets, and
    /// apostrophes are emitted verbatim, even when that makes invalid XML.
    pub fn to_incumbent_xml(&self) -> String {
        let mut output = String::new();
        self.write_incumbent_xml(&mut output);
        output
    }

    fn write_xml(&self, output: &mut String) {
        output.push('<');
        output.push_str(&self.name);
        for (name, value) in &self.attributes {
            output.push(' ');
            output.push_str(name);
            output.push_str("=\"");
            escape_attribute(value, output);
            output.push('"');
        }
        if self.children.is_empty() {
            output.push_str("/>");
            return;
        }
        output.push('>');
        for child in &self.children {
            match child {
                XmlNode::Element(element) => element.write_xml(output),
                XmlNode::Text(text) => escape_text(text, output),
                XmlNode::CData(text) => {
                    // Re-emit CDATA safely even when its payload contains the terminator.
                    output.push_str("<![CDATA[");
                    output.push_str(&text.replace("]]>", "]]]]><![CDATA[>"));
                    output.push_str("]]>");
                }
            }
        }
        output.push_str("</");
        output.push_str(&self.name);
        output.push('>');
    }

    fn write_incumbent_xml(&self, output: &mut String) {
        output.push('<');
        output.push_str(&self.name);
        for (name, value) in &self.attributes {
            output.push(' ');
            output.push_str(name);
            output.push_str("=\"");
            for character in value.chars() {
                if character == '"' {
                    output.push_str("&quot;");
                } else {
                    output.push(character);
                }
            }
            output.push('"');
        }
        if self.children.is_empty() {
            output.push_str("/>");
            return;
        }
        output.push('>');
        for child in &self.children {
            match child {
                XmlNode::Element(element) => element.write_incumbent_xml(output),
                XmlNode::Text(text) => escape_text(text, output),
                XmlNode::CData(text) => {
                    output.push_str("<![CDATA[");
                    output.push_str(&text.replace("]]>", "]]]]><![CDATA[>"));
                    output.push_str("]]>");
                }
            }
        }
        output.push_str("</");
        output.push_str(&self.name);
        output.push('>');
    }
}

/// Parse one XML part without a DTD/entity resolver and under a shared budget.
pub fn parse_xml(
    xml: &[u8],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<XmlDocument, ParseError> {
    budget.charge_xml_bytes(xml.len(), part)?;
    let repaired = escape_stray_ampersands(xml);
    let mut reader = Reader::from_reader(repaired.as_ref());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;

    let mut roots = Vec::new();
    let mut stack: Vec<XmlElement> = Vec::new();
    loop {
        let event = reader
            .read_event()
            .map_err(|error| malformed(&reader, part, error))?;
        budget.charge_event(part)?;
        match event {
            Event::Start(start) => {
                if stack.len() + 1 > budget.limits.max_xml_depth {
                    return Err(ParseError::ResourceLimit {
                        kind: "xmlDepth",
                        part: part.to_owned(),
                    });
                }
                stack.push(decode_element(&reader, start, part, budget)?);
            }
            Event::Empty(start) => {
                if stack.len() + 1 > budget.limits.max_xml_depth {
                    return Err(ParseError::ResourceLimit {
                        kind: "xmlDepth",
                        part: part.to_owned(),
                    });
                }
                append_element(
                    decode_element(&reader, start, part, budget)?,
                    &mut stack,
                    &mut roots,
                    part,
                )?;
            }
            Event::End(_) => {
                let element = stack.pop().ok_or_else(|| ParseError::MalformedXml {
                    part: part.to_owned(),
                    offset: reader.buffer_position(),
                    message: "unexpected closing element".to_owned(),
                })?;
                append_element(element, &mut stack, &mut roots, part)?;
            }
            Event::Text(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| malformed(&reader, part, error))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| malformed(&reader, part, error))?;
                budget.charge_text(unescaped.len(), part)?;
                append_text_node(
                    XmlNode::Text(unescaped.into_owned()),
                    &mut stack,
                    part,
                    reader.buffer_position(),
                )?;
            }
            Event::CData(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| malformed(&reader, part, error))?
                    .into_owned();
                budget.charge_text(decoded.len(), part)?;
                append_text_node(
                    XmlNode::CData(decoded),
                    &mut stack,
                    part,
                    reader.buffer_position(),
                )?;
            }
            Event::DocType(_) => {
                return Err(ParseError::UnsafeXml {
                    kind: "DTD/entity declarations are forbidden",
                    part: part.to_owned(),
                });
            }
            Event::Decl(_) | Event::PI(_) | Event::Comment(_) => {}
            Event::GeneralRef(reference) => {
                // quick-xml deliberately surfaces every reference separately.
                // Resolve only the five predefined XML entities and legal
                // numeric character references; there is no DTD entity table.
                let decoded = reference
                    .decode()
                    .map_err(|error| malformed(&reader, part, error))?;
                let resolved = if reference.is_char_ref() {
                    reference
                        .resolve_char_ref()
                        .map_err(|error| malformed(&reader, part, error))?
                        .filter(|character| is_legal_xml_character(*character))
                        .map(|character| character.to_string())
                } else {
                    quick_xml::escape::resolve_predefined_entity(&decoded).map(str::to_owned)
                };
                let Some(resolved) = resolved else {
                    return Err(ParseError::UnsafeXml {
                        kind: "non-predefined or illegal entity reference",
                        part: part.to_owned(),
                    });
                };
                budget.charge_text(resolved.len(), part)?;
                append_text_node(
                    XmlNode::Text(resolved),
                    &mut stack,
                    part,
                    reader.buffer_position(),
                )?;
            }
            Event::Eof => break,
        }
    }

    if !stack.is_empty() {
        return Err(ParseError::MalformedXml {
            part: part.to_owned(),
            offset: reader.buffer_position(),
            message: "unclosed element".to_owned(),
        });
    }
    Ok(XmlDocument { roots })
}

fn decode_element(
    reader: &Reader<&[u8]>,
    start: BytesStart<'_>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<XmlElement, ParseError> {
    let name = reader
        .decoder()
        .decode(start.name().as_ref())
        .map_err(|error| malformed(reader, part, error))?
        .into_owned();
    budget.charge_text(name.len(), part)?;

    let mut attributes = IndexMap::new();
    let mut attribute_bytes = 0usize;
    for (index, attribute) in start.attributes().enumerate() {
        if index >= budget.limits.max_attributes_per_element {
            return Err(ParseError::ResourceLimit {
                kind: "attributesPerElement",
                part: part.to_owned(),
            });
        }
        let attribute = attribute.map_err(|error| malformed(reader, part, error))?;
        let key = reader
            .decoder()
            .decode(attribute.key.as_ref())
            .map_err(|error| malformed(reader, part, error))?
            .into_owned();
        // Deliberately NOT decoded_and_normalized_value(): normalization folds
        // whitespace in attribute values, which would break byte-faithful round-trips.
        #[allow(deprecated)]
        let value = attribute
            .decode_and_unescape_value(reader.decoder())
            .map_err(|error| malformed(reader, part, error))?
            .into_owned();
        attribute_bytes = attribute_bytes
            .checked_add(key.len() + value.len())
            .ok_or_else(|| ParseError::ResourceLimit {
                kind: "attributeBytes",
                part: part.to_owned(),
            })?;
        if attribute_bytes > budget.limits.max_attribute_bytes {
            return Err(ParseError::ResourceLimit {
                kind: "attributeBytes",
                part: part.to_owned(),
            });
        }
        budget.charge_text(key.len() + value.len(), part)?;
        if attributes.insert(key.clone(), value).is_some() {
            return Err(ParseError::MalformedXml {
                part: part.to_owned(),
                offset: reader.buffer_position(),
                message: format!("duplicate attribute {key:?}"),
            });
        }
    }
    Ok(XmlElement {
        name,
        attributes,
        children: Vec::new(),
    })
}

fn append_element(
    element: XmlElement,
    stack: &mut [XmlElement],
    roots: &mut Vec<XmlElement>,
    part: &str,
) -> Result<(), ParseError> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlNode::Element(element));
    } else {
        if !roots.is_empty() {
            return Err(ParseError::MalformedXml {
                part: part.to_owned(),
                offset: 0,
                message: "multiple root elements".to_owned(),
            });
        }
        roots.push(element);
    }
    Ok(())
}

fn append_text_node(
    node: XmlNode,
    stack: &mut [XmlElement],
    part: &str,
    offset: u64,
) -> Result<(), ParseError> {
    if let Some(parent) = stack.last_mut() {
        if let XmlNode::Text(text) = &node
            && let Some(XmlNode::Text(previous)) = parent.children.last_mut()
        {
            // quick-xml surfaces entity references as separate events. A DOM
            // text node does not retain those lexical boundaries, so match
            // the incumbent parser by coalescing adjacent decoded text.
            previous.push_str(text);
            return Ok(());
        }
        parent.children.push(node);
        return Ok(());
    }
    let text = match &node {
        XmlNode::Text(text) | XmlNode::CData(text) => text,
        XmlNode::Element(_) => unreachable!(),
    };
    if text.trim().is_empty() {
        return Ok(());
    }
    Err(ParseError::MalformedXml {
        part: part.to_owned(),
        offset,
        message: "text outside the root element".to_owned(),
    })
}

fn malformed(reader: &Reader<&[u8]>, part: &str, error: impl ToString) -> ParseError {
    ParseError::MalformedXml {
        part: part.to_owned(),
        offset: reader.buffer_position(),
        message: error.to_string(),
    }
}

fn escape_stray_ampersands(xml: &[u8]) -> std::borrow::Cow<'_, [u8]> {
    // The incumbent parser repairs literal stray ampersands. Keep that narrow
    // compatibility only for UTF-8/ASCII input; declared entities are left for
    // the resolver-free parser to reject.
    let mut output: Option<Vec<u8>> = None;
    let mut index = 0;
    while index < xml.len() {
        if xml[index] != b'&' || is_entity_reference(&xml[index..]) {
            if let Some(output) = &mut output {
                output.push(xml[index]);
            }
            index += 1;
            continue;
        }
        let output = output.get_or_insert_with(|| xml[..index].to_vec());
        output.extend_from_slice(b"&amp;");
        index += 1;
    }
    output.map_or(std::borrow::Cow::Borrowed(xml), std::borrow::Cow::Owned)
}

fn is_legal_xml_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(character as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

fn is_entity_reference(input: &[u8]) -> bool {
    if input.len() < 3 || input[0] != b'&' {
        return false;
    }
    let Some(end) = input.iter().position(|byte| *byte == b';') else {
        return false;
    };
    if end < 2 {
        return false;
    }
    let body = &input[1..end];
    if let Some(decimal) = body.strip_prefix(b"#") {
        if let Some(hex) = decimal
            .strip_prefix(b"x")
            .or_else(|| decimal.strip_prefix(b"X"))
        {
            return !hex.is_empty() && hex.iter().all(u8::is_ascii_hexdigit);
        }
        return !decimal.is_empty() && decimal.iter().all(u8::is_ascii_digit);
    }
    body[0].is_ascii_alphabetic()
        && body
            .iter()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.' | b':'))
}

pub fn local_name(name: &str) -> &str {
    name.split_once(':').map_or(name, |(_, local)| local)
}

pub fn namespace_prefix(name: &str) -> Option<&str> {
    name.split_once(':').map(|(prefix, _)| prefix)
}

pub(crate) fn parse_javascript_integer_prefix(value: &str) -> Option<f64> {
    let value = value.trim_start();
    let bytes = value.as_bytes();
    let mut end = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    let start = end;
    while matches!(bytes.get(end), Some(byte) if byte.is_ascii_digit()) {
        end += 1;
    }
    if end == start {
        return None;
    }
    value[..end]
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
}

fn escape_text(value: &str, output: &mut String) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(character),
        }
    }
}

fn escape_attribute(value: &str, output: &mut String) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => output.push(character),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(input: &str) -> Result<XmlDocument, ParseError> {
        let limits = ParseLimits::default();
        parse_xml(
            input.as_bytes(),
            "word/test.xml",
            &mut ParseBudget::new(&limits),
        )
    }

    #[test]
    fn preserves_whitespace_repairs_stray_ampersands_and_supports_helpers() {
        let doc =
            parse("<w:p w:val=\"12px\"><w:r> left & right </w:r><w:r><![CDATA[<x>]]></w:r></w:p>")
                .unwrap();
        let root = doc.root().unwrap();
        assert!(root.matches_name("w", "p"));
        assert_eq!(
            root.parse_numeric_attribute(Some("w"), "val", 2.0),
            Some(24.0)
        );
        assert_eq!(root.text_content(), " left & right <x>");
        assert_eq!(root.children_named("w", "r").count(), 2);
        assert!(root.to_xml().contains(" left &amp; right "));
    }

    #[test]
    fn incumbent_raw_xml_pins_xml_js_attribute_escaping_quirk() {
        let doc = parse(r#"<x a="x&apos;y&gt;z&quot;q&lt;l&amp;m"/>"#).unwrap();
        assert_eq!(
            doc.root().unwrap().to_incumbent_xml(),
            r#"<x a="x'y>z&quot;q<l&m"/>"#
        );
        assert_eq!(
            doc.root().unwrap().to_xml(),
            r#"<x a="x&apos;y&gt;z&quot;q&lt;l&amp;m"/>"#
        );
    }

    #[test]
    fn parses_the_incumbent_color_helper_shape() {
        let doc = parse("<w:color w:val=\"112233\" w:themeColor=\"accent1\" w:themeTint=\"80\"/>")
            .unwrap();
        assert_eq!(
            doc.root().unwrap().parse_color(),
            ParsedColor {
                value: Some("112233".to_owned()),
                theme_color: Some("accent1".to_owned()),
                theme_tint: Some("80".to_owned()),
                theme_shade: None,
            }
        );
    }

    #[test]
    fn rejects_dtd_xxe_and_entity_declarations() {
        for xml in [
            "<!DOCTYPE x [<!ENTITY xxe SYSTEM 'file:///etc/passwd'>]><x>&xxe;</x>",
            "<!DOCTYPE x SYSTEM 'https://attacker.invalid/evil.dtd'><x/>",
            "<!DOCTYPE x [<!ENTITY a 'ha'><!ENTITY b '&a;&a;'>]><x>&b;</x>",
            "<!DOCTYPE x [<!ENTITY % p SYSTEM 'file:///tmp/p'>%p;]><x/>",
        ] {
            assert!(
                matches!(parse(xml), Err(ParseError::UnsafeXml { .. })),
                "{xml}"
            );
        }
    }

    #[test]
    fn accepts_only_predefined_and_numeric_references() {
        assert_eq!(
            parse("<x>&amp;&#65;&#x42;</x>")
                .unwrap()
                .root()
                .unwrap()
                .text_content(),
            "&AB"
        );
        assert!(matches!(
            parse("<x>&notDeclared;</x>"),
            Err(ParseError::UnsafeXml { .. } | ParseError::MalformedXml { .. })
        ));
    }

    #[test]
    fn enforces_depth_event_and_text_limits() {
        let limits = ParseLimits {
            max_xml_depth: 2,
            ..ParseLimits::default()
        };
        let error = parse_xml(
            b"<a><b><c/></b></a>",
            "deep.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap_err();
        assert!(matches!(
            error,
            ParseError::ResourceLimit {
                kind: "xmlDepth",
                ..
            }
        ));

        let limits = ParseLimits {
            max_xml_text_bytes: 1,
            ..ParseLimits::default()
        };
        let error =
            parse_xml(b"<a>xx</a>", "text.xml", &mut ParseBudget::new(&limits)).unwrap_err();
        assert!(matches!(
            error,
            ParseError::ResourceLimit {
                kind: "xmlTextBytes",
                ..
            }
        ));
    }
}
