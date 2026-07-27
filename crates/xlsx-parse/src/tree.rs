//! a bounded element tree for the few parts that need random access instead of
//! one streaming pass. every element records the byte span of its content, so a
//! rewrite splices into the source instead of reserializing it.

use std::ops::Range;
use std::rc::Rc;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::xml::{resolve_entity, xml_err};
use crate::{MAX_DEPTH, MAX_TREE_BYTES, MAX_TREE_NODES, MAX_TREE_TEXT_BYTES, ParseError};

pub(crate) enum Node {
    Element(Element),
    Text(String),
}

pub(crate) struct Attribute {
    /// the attribute as authored, prefix included.
    pub(crate) name: String,
    /// the namespace its prefix was bound to; unprefixed attributes have none.
    pub(crate) namespace: Option<Rc<str>>,
    pub(crate) value: String,
}

pub(crate) struct Element {
    /// the tag as authored, prefix included.
    pub(crate) name: String,
    /// the namespace the tag's prefix (or the default declaration) was bound
    /// to, so a part is read by expanded name rather than by literal prefix.
    pub(crate) namespace: Option<Rc<str>>,
    pub(crate) attributes: Vec<Attribute>,
    pub(crate) children: Vec<Node>,
    /// byte span of the content between the start and end tags; empty for a
    /// self-closing tag, which [`Element::splice_target`] therefore refuses.
    pub(crate) content: Range<usize>,
    self_closing: bool,
}

impl Element {
    pub(crate) fn local_name(&self) -> &str {
        self.name.rsplit(':').next().unwrap_or(&self.name)
    }

    pub(crate) fn namespace(&self) -> Option<&str> {
        self.namespace.as_deref()
    }

    /// whether this is the expanded name `{namespace}local`.
    pub(crate) fn is(&self, namespace: &str, local: &str) -> bool {
        self.local_name() == local && self.namespace() == Some(namespace)
    }

    /// the `prefix:name` attribute when `prefix` is given, else the bare
    /// `name`; both forms fall back to the other.
    pub(crate) fn attribute(&self, prefix: Option<&str>, name: &str) -> Option<&str> {
        let qualified = prefix.map(|prefix| format!("{prefix}:{name}"));
        self.attributes
            .iter()
            .find(|attribute| Some(attribute.name.as_str()) == qualified.as_deref())
            .or_else(|| {
                self.attributes
                    .iter()
                    .find(|attribute| attribute.name == name)
            })
            .map(|attribute| attribute.value.as_str())
    }

    /// the attribute with this expanded name, whatever prefix carries it.
    pub(crate) fn attribute_ns(&self, namespace: &str, local: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace.as_deref() == Some(namespace)
                    && attribute.name.rsplit(':').next() == Some(local)
            })
            .map(|attribute| attribute.value.as_str())
    }

    /// the attribute whose local name matches, whatever its prefix.
    pub(crate) fn attribute_local(&self, name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| attribute.name.rsplit(':').next().unwrap_or(&attribute.name) == name)
            .map(|attribute| attribute.value.as_str())
    }

    pub(crate) fn child_elements(&self) -> impl Iterator<Item = &Element> {
        self.children.iter().filter_map(|node| match node {
            Node::Element(element) => Some(element),
            Node::Text(_) => None,
        })
    }

    pub(crate) fn child(&self, local: &str) -> Option<&Element> {
        self.child_elements()
            .find(|child| child.local_name() == local)
    }

    pub(crate) fn text_content(&self) -> String {
        let mut text = String::new();
        self.append_text(&mut text);
        text
    }

    fn append_text(&self, out: &mut String) {
        for child in &self.children {
            match child {
                Node::Text(text) => out.push_str(text),
                Node::Element(element) => element.append_text(out),
            }
        }
    }

    /// the span a replacement may be written into. `None` for a self-closing
    /// tag, whose content span is a position rather than a range.
    pub(crate) fn splice_target(&self) -> Option<Range<usize>> {
        (!self.self_closing).then(|| self.content.clone())
    }
}

struct Budget {
    nodes: usize,
    text: usize,
}

/// the in-scope prefix bindings, as a stack that mirrors element nesting.
#[derive(Default)]
struct Namespaces {
    /// `("", uri)` is the default declaration.
    bindings: Vec<(String, Rc<str>)>,
    scopes: Vec<usize>,
}

impl Namespaces {
    fn open(&mut self, start: &BytesStart<'_>) -> Result<(), ParseError> {
        self.scopes.push(self.bindings.len());
        for attribute in start.attributes() {
            let attribute = attribute.map_err(xml_err)?;
            let Some(prefix) = declared_prefix(attribute.key.as_ref()) else {
                continue;
            };
            let value = attribute
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .map_err(xml_err)?;
            self.bindings
                .push((prefix.to_owned(), Rc::from(value.as_ref())));
        }
        Ok(())
    }

    fn close(&mut self) {
        if let Some(len) = self.scopes.pop() {
            self.bindings.truncate(len);
        }
    }

    fn resolve(&self, prefix: &str) -> Option<Rc<str>> {
        self.bindings
            .iter()
            .rev()
            .find(|(bound, _)| bound == prefix)
            .map(|(_, uri)| uri.clone())
            .filter(|uri| !uri.is_empty())
    }
}

/// the prefix an `xmlns` attribute declares, `""` for the default one.
fn declared_prefix(key: &[u8]) -> Option<&str> {
    let key = std::str::from_utf8(key).ok()?;
    if key == "xmlns" {
        return Some("");
    }
    key.strip_prefix("xmlns:").filter(|rest| !rest.is_empty())
}

fn split_prefix(name: &str) -> (&str, &str) {
    match name.split_once(':') {
        Some((prefix, local)) if !prefix.is_empty() && !local.is_empty() => (prefix, local),
        _ => ("", name),
    }
}

/// how a part's bytes carry its text. opc permits utf-8 and utf-16 only, and a
/// rewrite is written back in the encoding it was read in.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum Encoding {
    Utf8 { bom: bool },
    Utf16 { big_endian: bool, bom: bool },
}

impl Encoding {
    fn encode(self, text: &str) -> Vec<u8> {
        match self {
            Encoding::Utf8 { bom } => {
                let mut out = Vec::with_capacity(text.len() + 3);
                if bom {
                    out.extend_from_slice(&[0xEF, 0xBB, 0xBF]);
                }
                out.extend_from_slice(text.as_bytes());
                out
            }
            Encoding::Utf16 { big_endian, bom } => {
                let mut out = Vec::with_capacity(text.len() * 2 + 2);
                if bom {
                    out.extend_from_slice(if big_endian {
                        &[0xFE, 0xFF]
                    } else {
                        &[0xFF, 0xFE]
                    });
                }
                for unit in text.encode_utf16() {
                    out.extend_from_slice(&if big_endian {
                        unit.to_be_bytes()
                    } else {
                        unit.to_le_bytes()
                    });
                }
                out
            }
        }
    }
}

/// a part decoded to utf-8, remembering the encoding its bytes were in. every
/// span an [`Element`] records indexes this decoded text, so a rewrite splices
/// there and is re-encoded on the way out.
pub(crate) struct Part {
    text: String,
    encoding: Encoding,
}

impl Part {
    pub(crate) fn decode(data: &[u8]) -> Result<Self, ParseError> {
        if data.len() > MAX_TREE_BYTES {
            return Err(ParseError::TreeTooLarge);
        }
        let (encoding, body) = detect_encoding(data)?;
        let text = match encoding {
            Encoding::Utf8 { .. } => std::str::from_utf8(body)
                .map(str::to_owned)
                .map_err(|_| ParseError::Malformed("part is not valid utf-8".into()))?,
            Encoding::Utf16 { big_endian, .. } => decode_utf16(body, big_endian)?,
        };
        if text.len() > MAX_TREE_BYTES {
            return Err(ParseError::TreeTooLarge);
        }
        reject_foreign_declared_encoding(&text)?;
        Ok(Self { text, encoding })
    }

    pub(crate) fn tree(&self) -> Result<Element, ParseError> {
        parse_text(&self.text)
    }

    /// rewrite disjoint spans of the decoded text and re-encode the result.
    pub(crate) fn splice(&self, edits: &[Edit]) -> Result<Vec<u8>, ParseError> {
        let spliced = splice_text(&self.text, edits)?;
        Ok(self.encoding.encode(&spliced))
    }
}

/// what replaces one span of a part.
pub(crate) enum Replacement {
    /// element content, escaped on the way in.
    Text(String),
    /// markup this crate generated, written verbatim.
    Markup(String),
}

pub(crate) type Edit = (Range<usize>, Replacement);

/// The byte-order mark, or the shape of the first characters when there is
/// none. `FF FE 00 00` is utf-32, which opc does not permit.
fn detect_encoding(data: &[u8]) -> Result<(Encoding, &[u8]), ParseError> {
    Ok(match data {
        [0xFF, 0xFE, 0x00, 0x00, ..] | [0x00, 0x00, 0xFE, 0xFF, ..] => {
            return Err(ParseError::Malformed("part is not utf-8 or utf-16".into()));
        }
        [0xEF, 0xBB, 0xBF, rest @ ..] => (Encoding::Utf8 { bom: true }, rest),
        [0xFF, 0xFE, rest @ ..] => (
            Encoding::Utf16 {
                big_endian: false,
                bom: true,
            },
            rest,
        ),
        [0xFE, 0xFF, rest @ ..] => (
            Encoding::Utf16 {
                big_endian: true,
                bom: true,
            },
            rest,
        ),
        [0x3C, 0x00, ..] => (
            Encoding::Utf16 {
                big_endian: false,
                bom: false,
            },
            data,
        ),
        [0x00, 0x3C, ..] => (
            Encoding::Utf16 {
                big_endian: true,
                bom: false,
            },
            data,
        ),
        _ => (Encoding::Utf8 { bom: false }, data),
    })
}

fn decode_utf16(data: &[u8], big_endian: bool) -> Result<String, ParseError> {
    if !data.len().is_multiple_of(2) {
        return Err(ParseError::Malformed(
            "utf-16 part has an odd byte length".into(),
        ));
    }
    let units = data.chunks_exact(2).map(|pair| {
        let pair = [pair[0], pair[1]];
        if big_endian {
            u16::from_be_bytes(pair)
        } else {
            u16::from_le_bytes(pair)
        }
    });
    char::decode_utf16(units)
        .collect::<Result<String, _>>()
        .map_err(|_| ParseError::Malformed("utf-16 part has an unpaired surrogate".into()))
}

/// An xml declaration naming something other than utf-8 or utf-16 would make
/// the bytes we write back mean something else, so it is refused.
fn reject_foreign_declared_encoding(text: &str) -> Result<(), ParseError> {
    let Some(declaration) = text
        .strip_prefix("<?xml")
        .and_then(|rest| rest.split_once("?>"))
        .map(|(declaration, _)| declaration)
    else {
        return Ok(());
    };
    let Some(value) = declaration.split_once("encoding").and_then(|(_, rest)| {
        let rest = rest.trim_start().strip_prefix('=')?.trim_start();
        let quote = rest.chars().next()?;
        matches!(quote, '"' | '\'')
            .then(|| rest[1..].split(quote).next())
            .flatten()
    }) else {
        return Ok(());
    };
    if ["utf-8", "utf8", "utf-16", "utf16", "utf-16le", "utf-16be"]
        .iter()
        .any(|known| value.eq_ignore_ascii_case(known))
    {
        return Ok(());
    }
    Err(ParseError::Malformed(format!(
        "part declares the unsupported encoding {value}"
    )))
}

/// parse a whole part into an element tree. rejects doctypes outright and caps
/// depth, node count and total text so hostile markup cannot exhaust memory.
pub(crate) fn parse_tree(data: &[u8]) -> Result<Element, ParseError> {
    Part::decode(data)?.tree()
}

fn parse_text(text: &str) -> Result<Element, ParseError> {
    let mut reader = Reader::from_str(text);
    let config = reader.config_mut();
    config.expand_empty_elements = false;
    config.check_end_names = true;

    let mut budget = Budget {
        nodes: MAX_TREE_NODES,
        text: MAX_TREE_TEXT_BYTES,
    };
    let mut stack: Vec<Element> = Vec::new();
    let mut namespaces = Namespaces::default();
    let mut root: Option<Element> = None;
    loop {
        let opened = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(xml_err)?;
        let closed = reader.buffer_position() as usize;
        match event {
            Event::Start(start) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(ParseError::DepthExceeded);
                }
                namespaces.open(&start)?;
                let element = open_element(&start, closed, false, &mut budget, &namespaces)?;
                stack.push(element);
            }
            Event::Empty(start) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(ParseError::DepthExceeded);
                }
                namespaces.open(&start)?;
                let element = open_element(&start, closed, true, &mut budget, &namespaces);
                namespaces.close();
                place(element?, &mut stack, &mut root)?;
            }
            Event::End(_) => {
                let mut done = stack
                    .pop()
                    .ok_or_else(|| ParseError::Malformed("unbalanced end tag".into()))?;
                namespaces.close();
                done.content.end = opened.max(done.content.start);
                place(done, &mut stack, &mut root)?;
            }
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_err)?;
                push_text(&decoded, &mut stack, &mut budget)?;
            }
            Event::CData(text) => {
                let decoded = text.decode().map_err(xml_err)?;
                push_text(&decoded, &mut stack, &mut budget)?;
            }
            Event::GeneralRef(entity) => {
                let name = entity.decode().map_err(xml_err)?;
                let resolved = resolve_entity(&name)?;
                push_text(&resolved, &mut stack, &mut budget)?;
            }
            Event::DocType(_) => {
                return Err(ParseError::Malformed("doctype is not accepted".into()));
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if !stack.is_empty() {
        return Err(ParseError::Malformed("unclosed element".into()));
    }
    root.ok_or_else(|| ParseError::Malformed("no root element".into()))
}

fn open_element(
    start: &BytesStart<'_>,
    content_start: usize,
    self_closing: bool,
    budget: &mut Budget,
    namespaces: &Namespaces,
) -> Result<Element, ParseError> {
    budget.nodes = budget
        .nodes
        .checked_sub(1)
        .ok_or(ParseError::TreeTooLarge)?;
    let name = decode_name(start.name().as_ref())?;
    let namespace = namespaces.resolve(split_prefix(&name).0);
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute.map_err(xml_err)?;
        budget.nodes = budget
            .nodes
            .checked_sub(1)
            .ok_or(ParseError::TreeTooLarge)?;
        let name = decode_name(attribute.key.as_ref())?;
        let prefix = split_prefix(&name).0;
        attributes.push(Attribute {
            namespace: (!prefix.is_empty())
                .then(|| namespaces.resolve(prefix))
                .flatten(),
            name,
            value: attribute
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .map_err(xml_err)?
                .into_owned(),
        });
    }
    Ok(Element {
        name,
        namespace,
        attributes,
        children: Vec::new(),
        content: content_start..content_start,
        self_closing,
    })
}

/// an element or attribute name. an xml name is valid utf-8 by definition, so
/// anything else is malformed rather than something to guess at.
fn decode_name(raw: &[u8]) -> Result<String, ParseError> {
    std::str::from_utf8(raw)
        .map(str::to_owned)
        .map_err(|_| ParseError::Malformed("element name is not valid utf-8".into()))
}

fn place(
    element: Element,
    stack: &mut [Element],
    root: &mut Option<Element>,
) -> Result<(), ParseError> {
    match stack.last_mut() {
        Some(parent) => parent.children.push(Node::Element(element)),
        None if root.is_some() => {
            return Err(ParseError::Malformed("more than one root element".into()));
        }
        None => *root = Some(element),
    }
    Ok(())
}

fn push_text(text: &str, stack: &mut [Element], budget: &mut Budget) -> Result<(), ParseError> {
    let Some(parent) = stack.last_mut() else {
        return Ok(());
    };
    budget.text = budget
        .text
        .checked_sub(text.len())
        .ok_or(ParseError::TreeTooLarge)?;
    match parent.children.last_mut() {
        Some(Node::Text(existing)) => existing.push_str(text),
        _ => parent.children.push(Node::Text(text.to_owned())),
    }
    Ok(())
}

/// rewrite disjoint spans of `source` in one pass. spans must be sorted and
/// non-overlapping, which every caller derives from document order.
fn splice_text(source: &str, edits: &[Edit]) -> Result<String, ParseError> {
    let mut out = String::with_capacity(source.len());
    let mut cursor = 0usize;
    for (span, replacement) in edits {
        if span.start < cursor || span.end > source.len() || span.start > span.end {
            return Err(ParseError::Malformed("overlapping rewrite span".into()));
        }
        let head = source
            .get(cursor..span.start)
            .ok_or_else(|| ParseError::Malformed("rewrite span split a character".into()))?;
        out.push_str(head);
        match replacement {
            Replacement::Text(text) => out.push_str(&escape_text(text)?),
            Replacement::Markup(markup) => out.push_str(markup),
        }
        cursor = span.end;
    }
    let tail = source
        .get(cursor..)
        .ok_or_else(|| ParseError::Malformed("rewrite span split a character".into()))?;
    out.push_str(tail);
    Ok(out)
}

/// escape element content. `>` is escaped too so a replacement can never
/// reopen the `]]>` sequence, and `\r` numerically so parsing it back does not
/// silently turn it into `\n`. a character xml 1.0 cannot carry is refused
/// rather than written raw into a part no consumer could then read.
pub(crate) fn escape_text(value: &str) -> Result<String, ParseError> {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '\r' => out.push_str("&#13;"),
            '\t' | '\n' => out.push(ch),
            _ if ch < ' ' || ch == '\u{fffe}' || ch == '\u{ffff}' => {
                return Err(ParseError::UnsupportedEdit(format!(
                    "a rewrite carries U+{:04X}, which xml cannot express",
                    ch as u32
                )));
            }
            _ => out.push(ch),
        }
    }
    Ok(out)
}
