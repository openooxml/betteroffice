use std::collections::BTreeMap;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::PptxError;

#[derive(Clone, Debug)]
pub struct ParseLimits {
    pub max_xml_bytes: usize,
    pub max_xml_events: usize,
    pub max_xml_text_bytes: usize,
    pub max_xml_depth: usize,
    pub max_attributes_per_element: usize,
    pub max_attribute_bytes: usize,
    pub max_relationships: usize,
    pub max_shapes: usize,
    pub max_paragraphs: usize,
    pub max_runs: usize,
}

impl Default for ParseLimits {
    fn default() -> Self {
        Self {
            max_xml_bytes: 128 * 1024 * 1024,
            max_xml_events: 4_000_000,
            max_xml_text_bytes: 128 * 1024 * 1024,
            max_xml_depth: 256,
            max_attributes_per_element: 1_024,
            max_attribute_bytes: 4 * 1024 * 1024,
            max_relationships: 250_000,
            max_shapes: 100_000,
            max_paragraphs: 500_000,
            max_runs: 2_000_000,
        }
    }
}

#[derive(Debug)]
pub(crate) struct ParseBudget<'a> {
    limits: &'a ParseLimits,
    xml_bytes: usize,
    xml_events: usize,
    xml_text_bytes: usize,
    relationships: usize,
    shapes: usize,
    paragraphs: usize,
    runs: usize,
}

impl<'a> ParseBudget<'a> {
    pub fn new(limits: &'a ParseLimits) -> Self {
        Self {
            limits,
            xml_bytes: 0,
            xml_events: 0,
            xml_text_bytes: 0,
            relationships: 0,
            shapes: 0,
            paragraphs: 0,
            runs: 0,
        }
    }

    pub fn charge_relationship(&mut self, part: &str) -> Result<(), PptxError> {
        charge(
            &mut self.relationships,
            1,
            self.limits.max_relationships,
            "relationships",
            part,
        )
    }

    pub fn charge_shape(&mut self, part: &str) -> Result<(), PptxError> {
        charge(&mut self.shapes, 1, self.limits.max_shapes, "shapes", part)
    }

    pub fn charge_paragraph(&mut self, part: &str) -> Result<(), PptxError> {
        charge(
            &mut self.paragraphs,
            1,
            self.limits.max_paragraphs,
            "paragraphs",
            part,
        )
    }

    pub fn charge_run(&mut self, part: &str) -> Result<(), PptxError> {
        charge(&mut self.runs, 1, self.limits.max_runs, "runs", part)
    }
}

fn charge(
    used: &mut usize,
    amount: usize,
    maximum: usize,
    kind: &'static str,
    part: &str,
) -> Result<(), PptxError> {
    *used = used
        .checked_add(amount)
        .ok_or_else(|| PptxError::ResourceLimit {
            part: part.to_owned(),
            kind,
        })?;
    if *used > maximum {
        return Err(PptxError::ResourceLimit {
            part: part.to_owned(),
            kind,
        });
    }
    Ok(())
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct XmlElement {
    pub name: String,
    pub attributes: BTreeMap<String, String>,
    pub children: Vec<XmlNode>,
}

impl XmlElement {
    pub fn local_name(&self) -> &str {
        local_name(&self.name)
    }

    pub fn attribute(&self, name: &str) -> Option<&str> {
        self.attributes.get(name).map(String::as_str)
    }

    pub fn attribute_local(&self, name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|(key, _)| local_name(key) == name)
            .map(|(_, value)| value.as_str())
    }

    pub fn child(&self, name: &str) -> Option<&XmlElement> {
        self.child_elements()
            .find(|child| child.local_name() == name)
    }

    pub fn children_named<'a>(
        &'a self,
        name: &'a str,
    ) -> impl Iterator<Item = &'a XmlElement> + 'a {
        self.child_elements()
            .filter(move |child| child.local_name() == name)
    }

    pub fn child_elements(&self) -> impl Iterator<Item = &XmlElement> {
        self.children.iter().filter_map(|child| match child {
            XmlNode::Element(element) => Some(element),
            XmlNode::Text(_) => None,
        })
    }

    pub fn descendants_named<'a>(&'a self, name: &'a str) -> Vec<&'a XmlElement> {
        let mut output = Vec::new();
        collect_descendants(self, name, &mut output);
        output
    }

    pub fn text_content(&self) -> String {
        let mut output = String::new();
        append_text(self, &mut output);
        output
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum XmlNode {
    Element(XmlElement),
    Text(String),
}

fn collect_descendants<'a>(element: &'a XmlElement, name: &str, output: &mut Vec<&'a XmlElement>) {
    for child in element.child_elements() {
        if child.local_name() == name {
            output.push(child);
        }
        collect_descendants(child, name, output);
    }
}

fn append_text(element: &XmlElement, output: &mut String) {
    for child in &element.children {
        match child {
            XmlNode::Element(element) => append_text(element, output),
            XmlNode::Text(text) => output.push_str(text),
        }
    }
}

pub(crate) fn local_name(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

pub(crate) fn parse_xml(
    xml: &[u8],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<XmlElement, PptxError> {
    charge(
        &mut budget.xml_bytes,
        xml.len(),
        budget.limits.max_xml_bytes,
        "xmlBytes",
        part,
    )?;
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut roots = Vec::new();
    let mut stack = Vec::new();

    loop {
        let event = reader
            .read_event()
            .map_err(|error| malformed(&reader, part, error.to_string()))?;
        charge(
            &mut budget.xml_events,
            1,
            budget.limits.max_xml_events,
            "xmlEvents",
            part,
        )?;
        match event {
            Event::Start(start) => {
                check_depth(stack.len() + 1, budget, part)?;
                stack.push(decode_element(&reader, start, part, budget)?);
            }
            Event::Empty(start) => {
                check_depth(stack.len() + 1, budget, part)?;
                let element = decode_element(&reader, start, part, budget)?;
                append_element(element, &mut stack, &mut roots, part)?;
            }
            Event::End(_) => {
                let element = stack
                    .pop()
                    .ok_or_else(|| malformed(&reader, part, "unexpected closing element"))?;
                append_element(element, &mut stack, &mut roots, part)?;
            }
            Event::Text(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| malformed(&reader, part, error.to_string()))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| malformed(&reader, part, error.to_string()))?;
                charge(
                    &mut budget.xml_text_bytes,
                    unescaped.len(),
                    budget.limits.max_xml_text_bytes,
                    "xmlTextBytes",
                    part,
                )?;
                append_text_node(unescaped.into_owned(), &mut stack, &reader, part)?;
            }
            Event::CData(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| malformed(&reader, part, error.to_string()))?
                    .into_owned();
                charge(
                    &mut budget.xml_text_bytes,
                    decoded.len(),
                    budget.limits.max_xml_text_bytes,
                    "xmlTextBytes",
                    part,
                )?;
                append_text_node(decoded, &mut stack, &reader, part)?;
            }
            Event::DocType(_) => {
                return Err(PptxError::UnsafeXml {
                    part: part.to_owned(),
                    kind: "DTD/entity declarations are forbidden",
                });
            }
            Event::GeneralRef(reference) => {
                let decoded = reference
                    .decode()
                    .map_err(|error| malformed(&reader, part, error.to_string()))?;
                let resolved = if reference.is_char_ref() {
                    reference
                        .resolve_char_ref()
                        .map_err(|error| malformed(&reader, part, error.to_string()))?
                        .filter(|character| is_legal_xml_character(*character))
                        .map(|character| character.to_string())
                } else {
                    quick_xml::escape::resolve_predefined_entity(&decoded).map(str::to_owned)
                };
                let Some(resolved) = resolved else {
                    return Err(PptxError::UnsafeXml {
                        part: part.to_owned(),
                        kind: "non-predefined or illegal entity reference",
                    });
                };
                charge(
                    &mut budget.xml_text_bytes,
                    resolved.len(),
                    budget.limits.max_xml_text_bytes,
                    "xmlTextBytes",
                    part,
                )?;
                append_text_node(resolved, &mut stack, &reader, part)?;
            }
            Event::Decl(_) | Event::PI(_) | Event::Comment(_) => {}
            Event::Eof => break,
        }
    }

    if !stack.is_empty() {
        return Err(malformed(&reader, part, "unclosed element"));
    }
    if roots.len() != 1 {
        return Err(malformed(
            &reader,
            part,
            "XML part must have exactly one root element",
        ));
    }
    Ok(roots.pop().expect("root count checked"))
}

fn check_depth(depth: usize, budget: &ParseBudget<'_>, part: &str) -> Result<(), PptxError> {
    if depth > budget.limits.max_xml_depth {
        return Err(PptxError::ResourceLimit {
            part: part.to_owned(),
            kind: "xmlDepth",
        });
    }
    Ok(())
}

fn decode_element(
    reader: &Reader<&[u8]>,
    start: BytesStart<'_>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<XmlElement, PptxError> {
    let name = reader
        .decoder()
        .decode(start.name().as_ref())
        .map_err(|error| malformed(reader, part, error.to_string()))?
        .into_owned();
    let mut attributes = BTreeMap::new();
    let mut attribute_bytes = 0usize;
    for (index, attribute) in start.attributes().enumerate() {
        if index >= budget.limits.max_attributes_per_element {
            return Err(PptxError::ResourceLimit {
                part: part.to_owned(),
                kind: "attributesPerElement",
            });
        }
        let attribute = attribute.map_err(|error| malformed(reader, part, error.to_string()))?;
        let key = reader
            .decoder()
            .decode(attribute.key.as_ref())
            .map_err(|error| malformed(reader, part, error.to_string()))?
            .into_owned();
        #[allow(deprecated)]
        let value = attribute
            .decode_and_unescape_value(reader.decoder())
            .map_err(|error| malformed(reader, part, error.to_string()))?
            .into_owned();
        attribute_bytes = attribute_bytes
            .checked_add(key.len() + value.len())
            .ok_or_else(|| PptxError::ResourceLimit {
                part: part.to_owned(),
                kind: "attributeBytes",
            })?;
        if attribute_bytes > budget.limits.max_attribute_bytes {
            return Err(PptxError::ResourceLimit {
                part: part.to_owned(),
                kind: "attributeBytes",
            });
        }
        charge(
            &mut budget.xml_text_bytes,
            key.len() + value.len(),
            budget.limits.max_xml_text_bytes,
            "xmlTextBytes",
            part,
        )?;
        if attributes.insert(key.clone(), value).is_some() {
            return Err(malformed(
                reader,
                part,
                format!("duplicate attribute {key}"),
            ));
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
) -> Result<(), PptxError> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlNode::Element(element));
    } else if roots.is_empty() {
        roots.push(element);
    } else {
        return Err(PptxError::MalformedXml {
            part: part.to_owned(),
            offset: 0,
            message: "multiple root elements".to_owned(),
        });
    }
    Ok(())
}

fn append_text_node(
    text: String,
    stack: &mut [XmlElement],
    reader: &Reader<&[u8]>,
    part: &str,
) -> Result<(), PptxError> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlNode::Text(text));
    } else if !text.trim().is_empty() {
        return Err(malformed(reader, part, "text outside root element"));
    }
    Ok(())
}

fn malformed(reader: &Reader<&[u8]>, part: &str, message: impl Into<String>) -> PptxError {
    PptxError::MalformedXml {
        part: part.to_owned(),
        offset: reader.buffer_position(),
        message: message.into(),
    }
}

fn is_legal_xml_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
        || ('\u{20}'..='\u{d7ff}').contains(&character)
        || ('\u{e000}'..='\u{fffd}').contains(&character)
        || ('\u{10000}'..='\u{10ffff}').contains(&character)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_dtd_unknown_entities_and_depth_overflow() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        assert!(matches!(
            parse_xml(b"<!DOCTYPE x><x/>", "x.xml", &mut budget),
            Err(PptxError::UnsafeXml { .. })
        ));

        let mut budget = ParseBudget::new(&limits);
        assert!(matches!(
            parse_xml(b"<x>&file;</x>", "x.xml", &mut budget),
            Err(PptxError::UnsafeXml { .. })
        ));

        let shallow = ParseLimits {
            max_xml_depth: 1,
            ..ParseLimits::default()
        };
        let mut budget = ParseBudget::new(&shallow);
        assert!(matches!(
            parse_xml(b"<x><y/></x>", "x.xml", &mut budget),
            Err(PptxError::ResourceLimit {
                kind: "xmlDepth",
                ..
            })
        ));
    }
}
