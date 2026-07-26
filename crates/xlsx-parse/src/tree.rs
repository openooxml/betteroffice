//! a bounded element tree for the few parts that need random access instead of
//! one streaming pass. every element records the byte span of its content, so a
//! rewrite splices into the source instead of reserializing it.

use std::ops::Range;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::xml::{resolve_entity, xml_err};
use crate::{MAX_DEPTH, MAX_TREE_BYTES, MAX_TREE_NODES, MAX_TREE_TEXT_BYTES, ParseError};

pub(crate) enum Node {
    Element(Element),
    Text(String),
}

pub(crate) struct Element {
    /// the tag as authored, prefix included.
    pub(crate) name: String,
    pub(crate) attributes: Vec<(String, String)>,
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

    /// the `prefix:name` attribute when `prefix` is given, else the bare
    /// `name`; both forms fall back to the other.
    pub(crate) fn attribute(&self, prefix: Option<&str>, name: &str) -> Option<&str> {
        let qualified = prefix.map(|prefix| format!("{prefix}:{name}"));
        self.attributes
            .iter()
            .find(|(key, _)| Some(key.as_str()) == qualified.as_deref())
            .or_else(|| self.attributes.iter().find(|(key, _)| key == name))
            .map(|(_, value)| value.as_str())
    }

    /// the attribute whose local name matches, whatever its prefix.
    pub(crate) fn attribute_local(&self, name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|(key, _)| key.rsplit(':').next().unwrap_or(key) == name)
            .map(|(_, value)| value.as_str())
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

/// parse a whole part into an element tree. rejects doctypes outright and caps
/// depth, node count and total text so hostile markup cannot exhaust memory.
pub(crate) fn parse_tree(data: &[u8]) -> Result<Element, ParseError> {
    if data.len() > MAX_TREE_BYTES {
        return Err(ParseError::TreeTooLarge);
    }
    let mut reader = Reader::from_reader(data);
    let config = reader.config_mut();
    config.expand_empty_elements = false;
    config.check_end_names = true;

    let mut budget = Budget {
        nodes: MAX_TREE_NODES,
        text: MAX_TREE_TEXT_BYTES,
    };
    let mut stack: Vec<Element> = Vec::new();
    let mut root: Option<Element> = None;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        let opened = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buf).map_err(xml_err)?;
        let closed = reader.buffer_position() as usize;
        match event {
            Event::Start(start) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(ParseError::DepthExceeded);
                }
                let element = open_element(&start, closed, false, &mut budget)?;
                stack.push(element);
            }
            Event::Empty(start) => {
                let element = open_element(&start, closed, true, &mut budget)?;
                place(element, &mut stack, &mut root)?;
            }
            Event::End(_) => {
                let mut done = stack
                    .pop()
                    .ok_or_else(|| ParseError::Malformed("unbalanced end tag".into()))?;
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
) -> Result<Element, ParseError> {
    budget.nodes = budget
        .nodes
        .checked_sub(1)
        .ok_or(ParseError::TreeTooLarge)?;
    let name = String::from_utf8_lossy(start.name().as_ref()).into_owned();
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute.map_err(xml_err)?;
        budget.nodes = budget
            .nodes
            .checked_sub(1)
            .ok_or(ParseError::TreeTooLarge)?;
        attributes.push((
            String::from_utf8_lossy(attribute.key.as_ref()).into_owned(),
            attribute
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .map_err(xml_err)?
                .into_owned(),
        ));
    }
    Ok(Element {
        name,
        attributes,
        children: Vec::new(),
        content: content_start..content_start,
        self_closing,
    })
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

/// rewrite disjoint byte spans of `source` in one pass. spans must be sorted
/// and non-overlapping, which every caller derives from document order.
pub(crate) fn splice(
    source: &[u8],
    edits: &[(Range<usize>, String)],
) -> Result<Vec<u8>, ParseError> {
    let mut out = Vec::with_capacity(source.len());
    let mut cursor = 0usize;
    for (span, replacement) in edits {
        if span.start < cursor || span.end > source.len() || span.start > span.end {
            return Err(ParseError::Malformed("overlapping rewrite span".into()));
        }
        out.extend_from_slice(&source[cursor..span.start]);
        out.extend_from_slice(escape_text(replacement).as_bytes());
        cursor = span.end;
    }
    out.extend_from_slice(&source[cursor..]);
    Ok(out)
}

/// escape element content. `>` is escaped too so a replacement can never
/// reopen the `]]>` sequence.
fn escape_text(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(ch),
        }
    }
    out
}
