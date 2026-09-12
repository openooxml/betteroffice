use std::collections::HashSet;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::VsdxError;

#[derive(Clone, Debug)]
pub struct ParseLimits {
    pub max_xml_bytes: usize,
    pub max_xml_events: usize,
    pub max_xml_text_bytes: usize,
    pub max_xml_depth: usize,
    pub max_attributes_per_element: usize,
    pub max_attribute_bytes: usize,
    pub max_relationships: usize,
    pub max_cells: usize,
    pub max_sections: usize,
    pub max_rows: usize,
    pub max_shapes: usize,
    /// Maximum nesting accepted by the bounded ShapeSheet formula parser.
    pub max_formula_depth: usize,
    /// Maximum AST nodes accepted by the bounded ShapeSheet formula parser.
    pub max_formula_nodes: usize,
    /// Maximum tokens accepted by the bounded ShapeSheet formula tokenizer.
    pub max_formula_tokens: usize,
    /// Maximum evaluator operations for one ShapeSheet formula.
    pub max_formula_steps: usize,
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
            max_cells: 2_000_000,
            max_sections: 500_000,
            max_rows: 2_000_000,
            max_shapes: 100_000,
            max_formula_depth: 256,
            max_formula_nodes: 8_192,
            max_formula_tokens: 16_384,
            max_formula_steps: 32_768,
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
    cells: usize,
    sections: usize,
    rows: usize,
    shapes: usize,
}
impl<'a> ParseBudget<'a> {
    pub fn new(limits: &'a ParseLimits) -> Self {
        Self {
            limits,
            xml_bytes: 0,
            xml_events: 0,
            xml_text_bytes: 0,
            relationships: 0,
            cells: 0,
            sections: 0,
            rows: 0,
            shapes: 0,
        }
    }
    pub fn charge_relationship(&mut self, part: &str) -> Result<(), VsdxError> {
        charge(
            &mut self.relationships,
            1,
            self.limits.max_relationships,
            "relationships",
            part,
        )
    }
    pub fn charge_cell(&mut self, part: &str) -> Result<(), VsdxError> {
        charge(&mut self.cells, 1, self.limits.max_cells, "cells", part)
    }
    pub fn charge_section(&mut self, part: &str) -> Result<(), VsdxError> {
        charge(
            &mut self.sections,
            1,
            self.limits.max_sections,
            "sections",
            part,
        )
    }
    pub fn charge_row(&mut self, part: &str) -> Result<(), VsdxError> {
        charge(&mut self.rows, 1, self.limits.max_rows, "rows", part)
    }
    pub fn charge_shape(&mut self, part: &str) -> Result<(), VsdxError> {
        charge(&mut self.shapes, 1, self.limits.max_shapes, "shapes", part)
    }
}
fn charge(
    used: &mut usize,
    amount: usize,
    maximum: usize,
    kind: &'static str,
    part: &str,
) -> Result<(), VsdxError> {
    *used = used
        .checked_add(amount)
        .ok_or_else(|| VsdxError::ResourceLimit {
            part: part.to_owned(),
            kind,
        })?;
    if *used > maximum {
        return Err(VsdxError::ResourceLimit {
            part: part.to_owned(),
            kind,
        });
    }
    Ok(())
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct XmlElement {
    pub name: String,
    pub attributes: Vec<(String, String)>,
    pub children: Vec<XmlNode>,
}
impl XmlElement {
    pub fn local_name(&self) -> &str {
        local_name(&self.name)
    }
    pub fn attribute(&self, name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.as_str())
    }
    pub fn children_named<'a>(
        &'a self,
        name: &'a str,
    ) -> impl Iterator<Item = &'a XmlElement> + 'a {
        self.children.iter().filter_map(move |node| match node {
            XmlNode::Element(element) if element.local_name() == name => Some(element),
            _ => None,
        })
    }
}
#[allow(dead_code)]
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum XmlNode {
    Element(XmlElement),
    Text(String),
}
fn local_name(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

pub(crate) fn parse_xml(
    xml: &[u8],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<XmlElement, VsdxError> {
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
                append_element(
                    decode_element(&reader, start, part, budget)?,
                    &mut stack,
                    &mut roots,
                    part,
                )?;
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
                let text = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| malformed(&reader, part, error.to_string()))?
                    .into_owned();
                append_text(text, &mut stack, &reader, part, budget)?;
            }
            Event::CData(text) => {
                let text = text
                    .decode()
                    .map_err(|error| malformed(&reader, part, error.to_string()))?
                    .into_owned();
                append_text(text, &mut stack, &reader, part, budget)?;
            }
            Event::DocType(_) => {
                return Err(VsdxError::UnsafeXml {
                    part: part.to_owned(),
                    kind: "DTD/entity declarations are forbidden",
                });
            }
            Event::GeneralRef(reference) => {
                let decoded = reference
                    .decode()
                    .map_err(|error| malformed(&reader, part, error.to_string()))?;
                let text = if reference.is_char_ref() {
                    reference
                        .resolve_char_ref()
                        .map_err(|error| malformed(&reader, part, error.to_string()))?
                        .filter(|character| legal(*character))
                        .map(|character| character.to_string())
                } else {
                    quick_xml::escape::resolve_predefined_entity(&decoded).map(str::to_owned)
                }
                .ok_or_else(|| VsdxError::UnsafeXml {
                    part: part.to_owned(),
                    kind: "non-predefined or illegal entity reference",
                })?;
                append_text(text, &mut stack, &reader, part, budget)?;
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
fn check_depth(depth: usize, budget: &ParseBudget<'_>, part: &str) -> Result<(), VsdxError> {
    if depth > budget.limits.max_xml_depth {
        Err(VsdxError::ResourceLimit {
            part: part.to_owned(),
            kind: "xmlDepth",
        })
    } else {
        Ok(())
    }
}
fn decode_element(
    reader: &Reader<&[u8]>,
    start: BytesStart<'_>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<XmlElement, VsdxError> {
    let name = reader
        .decoder()
        .decode(start.name().as_ref())
        .map_err(|error| malformed(reader, part, error.to_string()))?
        .into_owned();
    let mut attributes = Vec::new();
    let mut attribute_names = HashSet::new();
    let mut bytes = 0_usize;
    for (index, attribute) in start.attributes().enumerate() {
        if index >= budget.limits.max_attributes_per_element {
            return Err(VsdxError::ResourceLimit {
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
        let amount =
            key.len()
                .checked_add(value.len())
                .ok_or_else(|| VsdxError::ResourceLimit {
                    part: part.to_owned(),
                    kind: "attributeBytes",
                })?;
        bytes = bytes
            .checked_add(amount)
            .ok_or_else(|| VsdxError::ResourceLimit {
                part: part.to_owned(),
                kind: "attributeBytes",
            })?;
        if bytes > budget.limits.max_attribute_bytes {
            return Err(VsdxError::ResourceLimit {
                part: part.to_owned(),
                kind: "attributeBytes",
            });
        }
        charge(
            &mut budget.xml_text_bytes,
            amount,
            budget.limits.max_xml_text_bytes,
            "xmlTextBytes",
            part,
        )?;
        if !attribute_names.insert(key.clone()) {
            return Err(malformed(
                reader,
                part,
                format!("duplicate attribute {key}"),
            ));
        }
        attributes.push((key, value));
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
) -> Result<(), VsdxError> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlNode::Element(element));
    } else if roots.is_empty() {
        roots.push(element);
    } else {
        return Err(VsdxError::MalformedXml {
            part: part.to_owned(),
            offset: 0,
            message: "multiple root elements".to_owned(),
        });
    }
    Ok(())
}
fn append_text(
    text: String,
    stack: &mut [XmlElement],
    reader: &Reader<&[u8]>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<(), VsdxError> {
    charge(
        &mut budget.xml_text_bytes,
        text.len(),
        budget.limits.max_xml_text_bytes,
        "xmlTextBytes",
        part,
    )?;
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlNode::Text(text));
    } else if !text.trim().is_empty() {
        return Err(malformed(reader, part, "text outside root element"));
    }
    Ok(())
}
fn malformed(reader: &Reader<&[u8]>, part: &str, message: impl Into<String>) -> VsdxError {
    VsdxError::MalformedXml {
        part: part.to_owned(),
        offset: reader.buffer_position(),
        message: message.into(),
    }
}
fn legal(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
        || ('\u{20}'..='\u{d7ff}').contains(&character)
        || ('\u{e000}'..='\u{fffd}').contains(&character)
        || ('\u{10000}'..='\u{10ffff}').contains(&character)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn attributes_count_toward_xml_text_limit() {
        let limits = ParseLimits {
            max_xml_text_bytes: 3,
            ..ParseLimits::default()
        };
        let mut budget = ParseBudget::new(&limits);
        assert!(matches!(
            parse_xml(b"<Root a='123'/>", "attributes.xml", &mut budget),
            Err(VsdxError::ResourceLimit {
                kind: "xmlTextBytes",
                ..
            })
        ));
    }

    #[test]
    fn rejects_dtd_and_unknown_entities() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        assert!(matches!(
            parse_xml(b"<!DOCTYPE Root><Root/>", "dtd.xml", &mut budget),
            Err(VsdxError::UnsafeXml { .. })
        ));
        let mut budget = ParseBudget::new(&limits);
        assert!(matches!(
            parse_xml(b"<Root>&unknown;</Root>", "entity.xml", &mut budget),
            Err(VsdxError::UnsafeXml { .. })
        ));
    }
}
