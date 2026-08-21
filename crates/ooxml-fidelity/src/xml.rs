//! Neutral, namespace-resolving XML tree reader for the oracles.
//!
//! Deliberately independent of the engine's parser: an oracle built on the
//! code under test inherits its blind spots. Bounded, resolver-free, strict.

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};

use crate::error::FidelityError;

/// The `xml:` prefix namespace, always in scope.
pub const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";

#[derive(Clone, Debug)]
pub struct XmlLimits {
    pub max_bytes: usize,
    pub max_depth: usize,
    pub max_events: usize,
}

impl Default for XmlLimits {
    fn default() -> Self {
        Self {
            max_bytes: 128 * 1024 * 1024,
            max_depth: 256,
            max_events: 4_000_000,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XmlNode {
    Element(XmlElement),
    Text(String),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlElement {
    /// Resolved namespace URI; empty when the element is in no namespace.
    pub namespace: String,
    pub local: String,
    /// Document order, `xmlns` declarations excluded.
    pub attributes: Vec<XmlAttribute>,
    /// Namespace declarations carried by this element, as `(prefix, uri)`.
    pub bindings: Vec<(String, String)>,
    pub children: Vec<XmlNode>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlAttribute {
    /// Resolved namespace URI; empty for unprefixed attributes.
    pub namespace: String,
    pub local: String,
    pub value: String,
}

impl XmlElement {
    pub fn is(&self, namespace: &str, local: &str) -> bool {
        self.namespace == namespace && self.local == local
    }

    pub fn attribute(&self, namespace: &str, local: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| attribute.namespace == namespace && attribute.local == local)
            .map(|attribute| attribute.value.as_str())
    }

    pub fn element_children(&self) -> impl Iterator<Item = &XmlElement> {
        self.children.iter().filter_map(|child| match child {
            XmlNode::Element(element) => Some(element),
            XmlNode::Text(_) => None,
        })
    }

    pub fn child(&self, namespace: &str, local: &str) -> Option<&XmlElement> {
        self.element_children()
            .find(|element| element.is(namespace, local))
    }
}

/// Parse one XML part into a single-rooted resolved tree.
pub fn parse_part(
    bytes: &[u8],
    part: &str,
    limits: &XmlLimits,
) -> Result<XmlElement, FidelityError> {
    if bytes.len() > limits.max_bytes {
        return Err(limit(part, "xmlBytes"));
    }
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;

    let mut roots: Vec<XmlNode> = Vec::new();
    let mut stack: Vec<XmlElement> = Vec::new();
    let mut scopes: Vec<Vec<(String, String)>> = Vec::new();
    let mut events = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| malformed(part, error))?;
        events += 1;
        if events > limits.max_events {
            return Err(limit(part, "xmlEvents"));
        }
        match event {
            Event::Start(start) => {
                if stack.len() + 1 > limits.max_depth {
                    return Err(limit(part, "xmlDepth"));
                }
                let element = decode_element(&reader, &start, part, &mut scopes)?;
                stack.push(element);
            }
            Event::Empty(start) => {
                let element = decode_element(&reader, &start, part, &mut scopes)?;
                scopes.pop();
                append_element(element, &mut stack, &mut roots);
            }
            Event::End(_) => {
                let element = stack.pop().ok_or_else(|| FidelityError::MalformedXml {
                    part: part.to_owned(),
                    message: "unexpected closing element".to_owned(),
                })?;
                scopes.pop();
                append_element(element, &mut stack, &mut roots);
            }
            Event::Text(text) => {
                let decoded = text.decode().map_err(|error| malformed(part, error))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| malformed(part, error))?;
                append_text(unescaped.into_owned(), &mut stack);
            }
            Event::CData(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| malformed(part, error))?
                    .into_owned();
                append_text(decoded, &mut stack);
            }
            Event::GeneralRef(reference) => {
                let decoded = reference.decode().map_err(|error| malformed(part, error))?;
                let resolved = if reference.is_char_ref() {
                    reference
                        .resolve_char_ref()
                        .map_err(|error| malformed(part, error))?
                        .map(|character| character.to_string())
                } else {
                    quick_xml::escape::resolve_predefined_entity(&decoded).map(str::to_owned)
                };
                let Some(resolved) = resolved else {
                    return Err(FidelityError::UnsafeXml {
                        part: part.to_owned(),
                        kind: "non-predefined entity reference",
                    });
                };
                append_text(resolved, &mut stack);
            }
            Event::DocType(_) => {
                return Err(FidelityError::UnsafeXml {
                    part: part.to_owned(),
                    kind: "DTD/entity declarations are forbidden",
                });
            }
            Event::Decl(_) | Event::PI(_) | Event::Comment(_) => {}
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(FidelityError::MalformedXml {
            part: part.to_owned(),
            message: "unclosed element".to_owned(),
        });
    }
    let mut root_elements = roots.into_iter().filter_map(|node| match node {
        XmlNode::Element(element) => Some(element),
        XmlNode::Text(_) => None,
    });
    let root = root_elements
        .next()
        .ok_or_else(|| FidelityError::MalformedXml {
            part: part.to_owned(),
            message: "no root element".to_owned(),
        })?;
    if root_elements.next().is_some() {
        return Err(FidelityError::MalformedXml {
            part: part.to_owned(),
            message: "more than one root element".to_owned(),
        });
    }
    Ok(root)
}

fn decode_element(
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
    part: &str,
    scopes: &mut Vec<Vec<(String, String)>>,
) -> Result<XmlElement, FidelityError> {
    let name = reader
        .decoder()
        .decode(start.name().as_ref())
        .map_err(|error| malformed(part, error))?
        .into_owned();

    let mut bindings: Vec<(String, String)> = Vec::new();
    let mut raw_attributes: Vec<(String, String)> = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| malformed(part, error))?;
        let key = reader
            .decoder()
            .decode(attribute.key.as_ref())
            .map_err(|error| malformed(part, error))?
            .into_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| malformed(part, error))?
            .into_owned();
        if key == "xmlns" {
            bindings.push((String::new(), value));
        } else if let Some(prefix) = key.strip_prefix("xmlns:") {
            bindings.push((prefix.to_owned(), value));
        } else {
            raw_attributes.push((key, value));
        }
    }
    scopes.push(bindings.clone());

    let (namespace, local) = resolve_name(&name, true, scopes, part)?;
    let mut attributes = Vec::with_capacity(raw_attributes.len());
    for (key, value) in raw_attributes {
        let (namespace, local) = resolve_name(&key, false, scopes, part)?;
        attributes.push(XmlAttribute {
            namespace,
            local,
            value,
        });
    }
    Ok(XmlElement {
        namespace,
        local,
        attributes,
        bindings,
        children: Vec::new(),
    })
}

fn resolve_name(
    name: &str,
    element: bool,
    scopes: &[Vec<(String, String)>],
    part: &str,
) -> Result<(String, String), FidelityError> {
    let (prefix, local) = match name.split_once(':') {
        Some((prefix, local)) => (prefix, local),
        None => ("", name),
    };
    if prefix == "xml" {
        return Ok((XML_NAMESPACE.to_owned(), local.to_owned()));
    }
    if prefix.is_empty() && !element {
        return Ok((String::new(), local.to_owned()));
    }
    for scope in scopes.iter().rev() {
        if let Some((_, uri)) = scope.iter().rev().find(|(bound, _)| bound == prefix) {
            return Ok((uri.clone(), local.to_owned()));
        }
    }
    if prefix.is_empty() {
        return Ok((String::new(), local.to_owned()));
    }
    Err(FidelityError::UnboundPrefix {
        part: part.to_owned(),
        prefix: prefix.to_owned(),
    })
}

fn append_element(element: XmlElement, stack: &mut [XmlElement], roots: &mut Vec<XmlNode>) {
    match stack.last_mut() {
        Some(parent) => parent.children.push(XmlNode::Element(element)),
        None => roots.push(XmlNode::Element(element)),
    }
}

fn append_text(text: String, stack: &mut [XmlElement]) {
    let Some(parent) = stack.last_mut() else {
        return;
    };
    if let Some(XmlNode::Text(existing)) = parent.children.last_mut() {
        existing.push_str(&text);
        return;
    }
    parent.children.push(XmlNode::Text(text));
}

fn malformed(part: &str, error: impl std::fmt::Display) -> FidelityError {
    FidelityError::MalformedXml {
        part: part.to_owned(),
        message: error.to_string(),
    }
}

fn limit(part: &str, kind: &'static str) -> FidelityError {
    FidelityError::ResourceLimit {
        part: part.to_owned(),
        kind,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> XmlElement {
        parse_part(xml.as_bytes(), "test.xml", &XmlLimits::default()).unwrap()
    }

    #[test]
    fn resolves_prefixes_to_uris() {
        let root = parse(r#"<w:document xmlns:w="ns-w"><w:body/></w:document>"#);
        assert!(root.is("ns-w", "document"));
        assert_eq!(root.bindings, vec![("w".to_owned(), "ns-w".to_owned())]);
        assert!(root.element_children().next().unwrap().is("ns-w", "body"));
    }

    #[test]
    fn default_namespace_applies_to_elements_not_attributes() {
        let root = parse(r#"<Types xmlns="ns-ct"><Default Extension="xml"/></Types>"#);
        assert!(root.is("ns-ct", "Types"));
        let child = root.element_children().next().unwrap();
        assert!(child.is("ns-ct", "Default"));
        assert_eq!(child.attributes[0].namespace, "");
        assert_eq!(child.attributes[0].local, "Extension");
    }

    #[test]
    fn xml_prefix_is_always_bound() {
        let root = parse(r#"<t xml:space="preserve"> </t>"#);
        assert_eq!(root.attribute(XML_NAMESPACE, "space"), Some("preserve"));
    }

    #[test]
    fn adjacent_text_and_references_merge() {
        let root = parse("<t>a&amp;b<![CDATA[c]]>&#x64;</t>");
        assert_eq!(root.children, vec![XmlNode::Text("a&bcd".to_owned())]);
    }

    #[test]
    fn inner_binding_shadows_outer() {
        let root = parse(r#"<a xmlns:p="outer"><p:x xmlns:p="inner"/></a>"#);
        assert!(root.element_children().next().unwrap().is("inner", "x"));
    }

    #[test]
    fn unbound_prefix_is_refused() {
        let error = parse_part(b"<w:p/>", "test.xml", &XmlLimits::default()).unwrap_err();
        assert_eq!(
            error,
            FidelityError::UnboundPrefix {
                part: "test.xml".to_owned(),
                prefix: "w".to_owned()
            }
        );
    }

    #[test]
    fn doctype_is_refused() {
        let error = parse_part(
            b"<!DOCTYPE x [<!ENTITY e \"v\">]><x/>",
            "test.xml",
            &XmlLimits::default(),
        )
        .unwrap_err();
        assert!(matches!(error, FidelityError::UnsafeXml { .. }));
    }

    #[test]
    fn depth_limit_is_enforced() {
        let deep = format!("{}{}", "<a>".repeat(300), "</a>".repeat(300));
        let error = parse_part(deep.as_bytes(), "test.xml", &XmlLimits::default()).unwrap_err();
        assert_eq!(
            error,
            FidelityError::ResourceLimit {
                part: "test.xml".to_owned(),
                kind: "xmlDepth"
            }
        );
    }

    #[test]
    fn empty_and_expanded_elements_parse_identically() {
        assert_eq!(parse("<a><b/></a>"), parse("<a><b></b></a>"));
    }
}
