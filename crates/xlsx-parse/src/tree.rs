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

impl Attribute {
    pub(crate) fn local_name(&self) -> &str {
        self.name.rsplit(':').next().unwrap_or(&self.name)
    }

    /// A default declaration never reaches an attribute, so an unprefixed one
    /// is unqualified however its element was declared.
    fn vocabulary(&self) -> Vocabulary<'_> {
        match self.namespace.as_deref() {
            Some(namespace) => Vocabulary::Bound(namespace),
            None => Vocabulary::Absent,
        }
    }
}

/// What a name's prefix, or the default declaration over it, resolved to.
/// Three ways to carry no namespace have to be told apart and resolution
/// collapses all of them, so every reader reports which one it saw.
#[derive(Clone, Copy)]
pub(crate) enum Vocabulary<'a> {
    Bound(&'a str),
    /// nothing in scope declares one, which is how a part authored without
    /// namespaces spells the vocabulary it looks like
    Absent,
    /// an `xmlns=""` in scope: this name declared its way out of every
    /// vocabulary, which is not the same as never having declared one
    Cleared,
}

/// Whether an unqualified name is read as a vocabulary's own.
#[derive(Clone, Copy, PartialEq)]
pub(crate) enum Unqualified {
    /// spreadsheetml, whose parts are sometimes authored bare
    Owned,
    /// relationships and friends, where a bare `id` is a different attribute
    /// rather than an unqualified `r:id`
    Foreign,
}

/// The one ownership rule in this crate. Every reader of a name — tree
/// elements, tree attributes, the streaming root sniff — decides here and
/// nowhere else, so a reader added later cannot answer it differently. Returns
/// the local name the QName answers to when it belongs to `namespaces`.
pub(crate) fn owned_local_name<'a>(
    qname: &'a str,
    vocabulary: Vocabulary<'_>,
    namespaces: &[&str],
    unqualified: Unqualified,
) -> Option<&'a str> {
    let (prefix, local) = valid_qname(qname)?;
    let ours = match vocabulary {
        Vocabulary::Bound(namespace) => namespaces.contains(&namespace),
        Vocabulary::Absent => prefix.is_empty() && unqualified == Unqualified::Owned,
        Vocabulary::Cleared => false,
    };
    ours.then_some(local)
}

/// Whether every element and attribute name in a subtree is one namespace
/// resolution can be applied to. [`Element::is`] and its callers resolve a
/// prefix at the first colon while comparing a local name after the last, so a
/// name carrying a second colon can be read as one it is not. A reader that
/// matches expanded names that way asks this first and declines the part.
pub(crate) fn names_are_resolvable(element: &Element) -> bool {
    valid_qname(&element.name).is_some()
        && element
            .attributes
            .iter()
            .all(|attribute| valid_qname(&attribute.name).is_some())
        && element.child_elements().all(names_are_resolvable)
}

/// A QName namespace resolution can be applied to at all — at most one colon,
/// with neither side of it empty — split into its prefix and local name. A
/// name that is not well formed answers to nothing, however its prefix
/// resolves, so no reader may derive a local name any other way.
fn valid_qname(name: &str) -> Option<(&str, &str)> {
    match name.split_once(':') {
        None => (!name.is_empty()).then_some(("", name)),
        Some((prefix, local)) => (!prefix.is_empty() && !local.is_empty() && !local.contains(':'))
            .then_some((prefix, local)),
    }
}

pub(crate) struct Element {
    /// the tag as authored, prefix included.
    pub(crate) name: String,
    /// the namespace the tag's prefix (or the default declaration) was bound
    /// to, so a part is read by expanded name rather than by literal prefix.
    pub(crate) namespace: Option<Rc<str>>,
    /// whether an `xmlns=""` was in scope. Such a name carries no namespace
    /// because its subtree declared its way out of one, which is not the same
    /// as a part that never declared one.
    pub(crate) default_cleared: bool,
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

    /// the attribute with this expanded name, whatever prefix carries it. An
    /// unqualified attribute is a different one, never this vocabulary's.
    pub(crate) fn attribute_ns(&self, namespace: &str, local: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                owned_local_name(
                    &attribute.name,
                    attribute.vocabulary(),
                    &[namespace],
                    Unqualified::Foreign,
                ) == Some(local)
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

    /// the attributes answering to a local name in `namespace`. A foreign one
    /// spells the name the same way but is a different attribute, so it is
    /// never among them.
    pub(crate) fn attributes_in<'a>(
        &'a self,
        namespaces: &'a [&'a str],
        local: &'a str,
    ) -> impl Iterator<Item = &'a Attribute> {
        self.attributes.iter().filter(move |attribute| {
            owned_local_name(
                &attribute.name,
                attribute.vocabulary(),
                namespaces,
                Unqualified::Owned,
            ) == Some(local)
        })
    }

    /// whether any attribute answers to a local name, in whatever namespace.
    /// For a name that disqualifies rather than supplies, a foreign match costs
    /// only a refusal while a missed real one costs correctness, so the wider
    /// question is the safe one to ask.
    pub(crate) fn any_attribute_named(&self, local: &str) -> bool {
        self.attributes_named(local) > 0
    }

    /// how many attributes answer to a local name, in whatever namespace. A
    /// gate over a reader that matches local names counts these, because the
    /// reader will see every one of them and take the first.
    pub(crate) fn attributes_named(&self, local: &str) -> usize {
        self.attributes
            .iter()
            .filter(|attribute| attribute.local_name() == local)
            .count()
    }

    /// the child elements answering to a local name in one of `namespaces`.
    pub(crate) fn children_in<'a>(
        &'a self,
        namespaces: &'a [&'a str],
        local: &'a str,
    ) -> impl Iterator<Item = &'a Element> {
        self.child_elements()
            .filter(move |child| child.answers_to(namespaces, local))
    }

    /// whether this element answers to a local name in one of `namespaces`,
    /// counting the unqualified form a part that declares none is authored in.
    /// A foreign element spells the name the same way but is a different one.
    pub(crate) fn answers_to(&self, namespaces: &[&str], local: &str) -> bool {
        owned_local_name(
            &self.name,
            self.vocabulary(),
            namespaces,
            Unqualified::Owned,
        ) == Some(local)
    }

    fn vocabulary(&self) -> Vocabulary<'_> {
        match self.namespace() {
            Some(namespace) => Vocabulary::Bound(namespace),
            None if self.default_cleared => Vocabulary::Cleared,
            None => Vocabulary::Absent,
        }
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

    /// whether an `xmlns=""` is in scope, taking every unprefixed name under it
    /// out of the vocabulary rather than leaving it in none by omission.
    fn default_cleared(&self) -> bool {
        self.bindings
            .iter()
            .rev()
            .find(|(bound, _)| bound.is_empty())
            .is_some_and(|(_, uri)| uri.is_empty())
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
        parse_text(&self.text, MAX_TREE_NODES).map(|(root, _)| root)
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
    let (pairs, _) = data.as_chunks::<2>();
    let units = pairs.iter().map(|pair| {
        if big_endian {
            u16::from_be_bytes(*pair)
        } else {
            u16::from_le_bytes(*pair)
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

pub(crate) async fn charge_bytes(mut bytes: usize, work: &ooxml_opc::WorkBudget) {
    while bytes > 0 {
        bytes = bytes.saturating_sub(work.take(bytes.div_ceil(64).min(256)).await * 64);
    }
}

pub(crate) async fn copy_text(source: &str, work: &ooxml_opc::WorkBudget) -> String {
    let mut text = String::with_capacity(source.len());
    append_text_sliced(&mut text, source, work).await;
    text
}

async fn append_text_sliced(target: &mut String, source: &str, work: &ooxml_opc::WorkBudget) {
    let mut offset = 0;
    while offset < source.len() {
        let count = work
            .take((source.len() - offset).div_ceil(64).min(256))
            .await
            * 64;
        let mut end = (offset + count).min(source.len());
        while !source.is_char_boundary(end) {
            end -= 1;
        }
        target.push_str(&source[offset..end]);
        offset = end;
    }
}

pub(crate) async fn text_content_sliced(element: &Element, work: &ooxml_opc::WorkBudget) -> String {
    let mut text = String::new();
    let mut stack = vec![element.children.iter()];
    while let Some(children) = stack.last_mut() {
        work.step().await;
        match children.next() {
            Some(Node::Text(value)) => append_text_sliced(&mut text, value, work).await,
            Some(Node::Element(child)) => stack.push(child.children.iter()),
            None => {
                stack.pop();
            }
        }
    }
    text
}

#[cfg(test)]
thread_local! {
    pub(crate) static OPENED_ELEMENTS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    pub(crate) static RETIRED_ELEMENTS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

pub(crate) async fn retire_tree(root: Element, work: &ooxml_opc::WorkBudget) {
    let mut stack = vec![root];
    while let Some(element) = stack.last_mut() {
        if let Some(child) = element.children.pop() {
            match child {
                Node::Element(child) => stack.push(child),
                Node::Text(text) => {
                    work.step().await;
                    drop(text);
                }
            }
        } else if !element.attributes.is_empty() {
            work.step().await;
            element.attributes.pop();
        } else {
            work.step().await;
            #[cfg(test)]
            RETIRED_ELEMENTS.set(RETIRED_ELEMENTS.get() + 1);
            stack.pop();
        }
    }
}

pub(crate) async fn names_are_resolvable_sliced(
    element: &Element,
    work: &ooxml_opc::WorkBudget,
) -> bool {
    let mut stack = vec![element];
    while let Some(element) = stack.pop() {
        work.step().await;
        charge_bytes(element.name.len(), work).await;
        if valid_qname(&element.name).is_none() {
            return false;
        }
        for attribute in &element.attributes {
            work.step().await;
            charge_bytes(attribute.name.len(), work).await;
            if valid_qname(&attribute.name).is_none() {
                return false;
            }
        }
        for child in element.child_elements() {
            work.step().await;
            stack.push(child);
        }
    }
    true
}

impl Namespaces {
    async fn open_sliced(
        &mut self,
        start: &BytesStart<'_>,
        work: &ooxml_opc::WorkBudget,
    ) -> Result<(), ParseError> {
        self.scopes.push(self.bindings.len());
        for attribute in start.attributes() {
            work.step().await;
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
}

async fn find_bytes_sliced(
    source: &str,
    needle: &[u8],
    work: &ooxml_opc::WorkBudget,
) -> Option<usize> {
    let bytes = source.as_bytes();
    let mut offset = 0;
    while offset < bytes.len() {
        let count = work
            .take((bytes.len() - offset).div_ceil(64).min(256))
            .await
            * 64;
        let end = (offset + count + needle.len() - 1).min(bytes.len());
        if let Some(found) = bytes[offset..end]
            .windows(needle.len())
            .position(|window| window == needle)
        {
            return Some(offset + found);
        }
        offset = (offset + count).min(bytes.len());
    }
    None
}

async fn trim_start_sliced<'a>(source: &'a str, work: &ooxml_opc::WorkBudget) -> &'a str {
    for (index, ch) in source.char_indices() {
        work.step().await;
        if !ch.is_whitespace() {
            return &source[index..];
        }
    }
    ""
}

async fn reject_foreign_declared_encoding_sliced(
    text: &str,
    work: &ooxml_opc::WorkBudget,
) -> Result<(), ParseError> {
    let Some(rest) = text.strip_prefix("<?xml") else {
        return Ok(());
    };
    let Some(end) = find_bytes_sliced(rest, b"?>", work).await else {
        return Ok(());
    };
    let declaration = &rest[..end];
    let Some(index) = find_bytes_sliced(declaration, b"encoding", work).await else {
        return Ok(());
    };
    let rest = trim_start_sliced(&declaration[index + 8..], work).await;
    let Some(rest) = rest.strip_prefix('=') else {
        return Ok(());
    };
    let rest = trim_start_sliced(rest, work).await;
    let Some(quote) = rest
        .chars()
        .next()
        .filter(|quote| matches!(*quote, '"' | '\''))
    else {
        return Ok(());
    };
    let rest = &rest[1..];
    let end = find_bytes_sliced(rest, &[quote as u8], work)
        .await
        .unwrap_or(rest.len());
    let value = &rest[..end];
    if ["utf-8", "utf8", "utf-16", "utf16", "utf-16le", "utf-16be"]
        .iter()
        .any(|known| value.eq_ignore_ascii_case(known))
    {
        return Ok(());
    }
    charge_bytes(value.len(), work).await;
    Err(ParseError::Malformed(format!(
        "part declares the unsupported encoding {value}"
    )))
}

pub(crate) async fn parse_tree_sliced(
    data: &[u8],
    work: &ooxml_opc::WorkBudget,
) -> Result<Element, ParseError> {
    if data.len() > MAX_TREE_BYTES {
        return Err(ParseError::TreeTooLarge);
    }
    let (encoding, body) = detect_encoding(data)?;
    let text = match encoding {
        Encoding::Utf8 { .. } => {
            let mut text = String::with_capacity(body.len());
            let mut offset = 0;
            while offset < body.len() {
                let count = work.take((body.len() - offset).div_ceil(64).min(256)).await * 64;
                let end = (offset + count).min(body.len());
                match std::str::from_utf8(&body[offset..end]) {
                    Ok(value) => {
                        text.push_str(value);
                        offset = end;
                    }
                    Err(error) if error.error_len().is_none() && end < body.len() => {
                        let valid_end = offset + error.valid_up_to();
                        let value = std::str::from_utf8(&body[offset..valid_end])
                            .map_err(|_| ParseError::Malformed("part is not valid utf-8".into()))?;
                        text.push_str(value);
                        offset = valid_end;
                    }
                    Err(_) => return Err(ParseError::Malformed("part is not valid utf-8".into())),
                }
            }
            text
        }
        Encoding::Utf16 { big_endian, .. } => {
            if !body.len().is_multiple_of(2) {
                return Err(ParseError::Malformed(
                    "utf-16 part has an odd byte length".into(),
                ));
            }
            let units = body.as_chunks::<2>().0.iter().map(|&pair| {
                if big_endian {
                    u16::from_be_bytes(pair)
                } else {
                    u16::from_le_bytes(pair)
                }
            });
            let mut text = String::new();
            for ch in char::decode_utf16(units) {
                work.step().await;
                text.push(ch.map_err(|_| {
                    ParseError::Malformed("utf-16 part has an unpaired surrogate".into())
                })?);
            }
            text
        }
    };
    if text.len() > MAX_TREE_BYTES {
        return Err(ParseError::TreeTooLarge);
    }
    reject_foreign_declared_encoding_sliced(&text, work).await?;
    parse_text_sliced(&text, MAX_TREE_NODES, work)
        .await
        .map(|(root, _)| root)
}

async fn parse_text_sliced(
    text: &str,
    nodes: usize,
    work: &ooxml_opc::WorkBudget,
) -> Result<(Element, usize), ParseError> {
    let mut reader = Reader::from_str(text);
    let config = reader.config_mut();
    config.expand_empty_elements = false;
    config.check_end_names = true;

    let mut budget = Budget {
        nodes,
        text: MAX_TREE_TEXT_BYTES,
    };
    let mut stack: Vec<Element> = Vec::new();
    let mut namespaces = Namespaces::default();
    let mut root: Option<Element> = None;
    let result = async {
        loop {
            work.step().await;
            let opened = reader.buffer_position() as usize;
            let event = reader.read_event().map_err(xml_err)?;
            let closed = reader.buffer_position() as usize;
            charge_bytes(closed.saturating_sub(opened), work).await;
            match event {
                Event::Start(start) => {
                    if stack.len() >= MAX_DEPTH {
                        return Err(ParseError::DepthExceeded);
                    }
                    namespaces.open_sliced(&start, work).await?;
                    let element =
                        open_element_sliced(&start, closed, false, &mut budget, &namespaces, work)
                            .await?;
                    stack.push(element);
                }
                Event::Empty(start) => {
                    if stack.len() >= MAX_DEPTH {
                        return Err(ParseError::DepthExceeded);
                    }
                    namespaces.open_sliced(&start, work).await?;
                    let element =
                        open_element_sliced(&start, closed, true, &mut budget, &namespaces, work)
                            .await;
                    namespaces.close();
                    place_sliced(element?, &mut stack, &mut root, work).await?;
                }
                Event::End(_) => {
                    let mut done = stack
                        .pop()
                        .ok_or_else(|| ParseError::Malformed("unbalanced end tag".into()))?;
                    namespaces.close();
                    done.content.end = opened.max(done.content.start);
                    place_sliced(done, &mut stack, &mut root, work).await?;
                }
                Event::Text(text) => {
                    let decoded = text.decode().map_err(xml_err)?;
                    push_text_sliced(&decoded, &mut stack, &mut budget, work).await?;
                }
                Event::CData(text) => {
                    let decoded = text.decode().map_err(xml_err)?;
                    push_text_sliced(&decoded, &mut stack, &mut budget, work).await?;
                }
                Event::GeneralRef(entity) => {
                    let name = entity.decode().map_err(xml_err)?;
                    let resolved = resolve_entity(&name)?;
                    push_text_sliced(&resolved, &mut stack, &mut budget, work).await?;
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
        root.take()
            .ok_or_else(|| ParseError::Malformed("no root element".into()))
    }
    .await;
    while let Some(element) = stack.pop() {
        retire_tree(element, work).await;
    }
    if let Some(root) = root {
        retire_tree(root, work).await;
    }
    result.map(|root| (root, nodes - budget.nodes))
}

async fn place_sliced(
    element: Element,
    stack: &mut [Element],
    root: &mut Option<Element>,
    work: &ooxml_opc::WorkBudget,
) -> Result<(), ParseError> {
    if stack.is_empty() && root.is_some() {
        retire_tree(element, work).await;
        return Err(ParseError::Malformed("more than one root element".into()));
    }
    place(element, stack, root)
}

async fn open_element_sliced(
    start: &BytesStart<'_>,
    content_start: usize,
    self_closing: bool,
    budget: &mut Budget,
    namespaces: &Namespaces,
    work: &ooxml_opc::WorkBudget,
) -> Result<Element, ParseError> {
    budget.nodes = budget
        .nodes
        .checked_sub(1)
        .ok_or(ParseError::TreeTooLarge)?;
    let name = decode_name(start.name().as_ref())?;
    let namespace = namespaces.resolve(split_prefix(&name).0);
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        work.step().await;
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
    #[cfg(test)]
    OPENED_ELEMENTS.set(OPENED_ELEMENTS.get() + 1);
    Ok(Element {
        default_cleared: namespaces.default_cleared(),
        name,
        namespace,
        attributes,
        children: Vec::new(),
        content: content_start..content_start,
        self_closing,
    })
}

async fn push_text_sliced(
    text: &str,
    stack: &mut [Element],
    budget: &mut Budget,
    work: &ooxml_opc::WorkBudget,
) -> Result<(), ParseError> {
    let Some(parent) = stack.last_mut() else {
        return Ok(());
    };
    budget.text = budget
        .text
        .checked_sub(text.len())
        .ok_or(ParseError::TreeTooLarge)?;
    match parent.children.last_mut() {
        Some(Node::Text(existing)) => {
            append_text_sliced(existing, text, work).await;
        }
        _ => parent
            .children
            .push(Node::Text(copy_text(text, work).await)),
    }
    Ok(())
}

/// [`parse_tree`] spending at most `nodes` elements and attributes, returning how many it
/// spent. Running out is [`ParseError::TreeTooLarge`].
pub(crate) fn parse_tree_within(data: &[u8], nodes: usize) -> Result<(Element, usize), ParseError> {
    let cap = nodes.min(MAX_TREE_NODES);
    parse_text(&Part::decode(data)?.text, cap)
}

/// Whether a parse failed on a size or depth cap rather than on malformed markup.
pub(crate) fn exceeds_limits(error: &ParseError) -> bool {
    matches!(error, ParseError::TreeTooLarge | ParseError::DepthExceeded)
}

fn parse_text(text: &str, nodes: usize) -> Result<(Element, usize), ParseError> {
    let mut reader = Reader::from_str(text);
    let config = reader.config_mut();
    config.expand_empty_elements = false;
    config.check_end_names = true;

    let mut budget = Budget {
        nodes,
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
    let root = root.ok_or_else(|| ParseError::Malformed("no root element".into()))?;
    Ok((root, nodes - budget.nodes))
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
        default_cleared: namespaces.default_cleared(),
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
