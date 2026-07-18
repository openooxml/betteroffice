//! S10 versioned serializer differential wire and canonical XML-event stream.

use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};

use crate::borders::BorderSpec;
use crate::formatting::ConditionalFormatStyle;
use crate::numbering::NumberingDefinitions;
use crate::paragraph::HexIdAllocator;
use crate::section::SectionProperties;
use crate::table::Table;
use crate::vml::Watermark;
use crate::xml::{ParseBudget, ParseError, ParseLimits, XmlElement, XmlNode, parse_xml};

use super::foundation::{
    BorderSide, serialize_border, serialize_conditional_format_style, serialize_table_grid,
};
use super::numbering::serialize_numbering_xml;
use super::raw::CONTENT_FRAGMENT_PREFIX;
use super::section::serialize_section_properties;
use super::watermark::serialize_watermark;

const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";

/// Fixed nondeterminism injected by every serializer differential.
///
/// S10 families do not allocate IDs or timestamps, but freezing the boundary
/// here prevents S11-S13 from adding ambient randomness or wall-clock reads.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SerializerDeterminism {
    pub seed: String,
    pub now: String,
}

impl SerializerDeterminism {
    pub fn validate(&self) -> Result<(), ParseError> {
        // Construct the actual allocator so this remains the same validation
        // and algorithm used by parse and future content serializers.
        let _ids = HexIdAllocator::from_sha256(&self.seed)?;
        if self.now.len() != 24
            || !self.now.ends_with('Z')
            || self.now.as_bytes().get(10) != Some(&b'T')
        {
            return Err(ParseError::Canonical(
                "serializer clock must be a UTC ISO-8601 millisecond timestamp".to_owned(),
            ));
        }
        Ok(())
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "family", rename_all = "camelCase", deny_unknown_fields)]
pub enum S10SerializeRequest {
    Border {
        determinism: SerializerDeterminism,
        #[serde(default)]
        border: Option<BorderSpec>,
        side: String,
    },
    ConditionalFormat {
        determinism: SerializerDeterminism,
        #[serde(default)]
        style: Option<ConditionalFormatStyle>,
    },
    TableGrid {
        determinism: SerializerDeterminism,
        table: Table,
    },
    Section {
        determinism: SerializerDeterminism,
        #[serde(default)]
        properties: Option<SectionProperties>,
    },
    Numbering {
        determinism: SerializerDeterminism,
        numbering: NumberingDefinitions,
    },
    Watermark {
        determinism: SerializerDeterminism,
        watermark: Watermark,
    },
}

impl S10SerializeRequest {
    fn determinism(&self) -> &SerializerDeterminism {
        match self {
            Self::Border { determinism, .. }
            | Self::ConditionalFormat { determinism, .. }
            | Self::TableGrid { determinism, .. }
            | Self::Section { determinism, .. }
            | Self::Numbering { determinism, .. }
            | Self::Watermark { determinism, .. } => determinism,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CanonicalXmlAttribute {
    pub namespace_uri: String,
    pub local_name: String,
    pub value: String,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum CanonicalXmlEvent {
    Start {
        namespace_uri: String,
        local_name: String,
        attributes: Vec<CanonicalXmlAttribute>,
    },
    Text {
        text: String,
    },
    End {
        namespace_uri: String,
        local_name: String,
    },
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S10SerializeResponse {
    pub wire_version: u8,
    pub family: String,
    pub xml: String,
    pub canonical_xml_events: Vec<CanonicalXmlEvent>,
    pub parse_back: serde_json::Value,
}

/// Serialize one S10 family and independently parse it to canonical XML events.
pub fn serialize_s10_wire(
    request: S10SerializeRequest,
) -> Result<S10SerializeResponse, ParseError> {
    request.determinism().validate()?;
    let (family, xml, fragment) = match request {
        S10SerializeRequest::Border { border, side, .. } => (
            "border",
            serialize_border(border.as_ref(), parse_border_side(&side)?),
            true,
        ),
        S10SerializeRequest::ConditionalFormat { style, .. } => (
            "conditionalFormat",
            serialize_conditional_format_style(style.as_ref()),
            true,
        ),
        S10SerializeRequest::TableGrid { table, .. } => {
            ("tableGrid", serialize_table_grid(&table), true)
        }
        S10SerializeRequest::Section { properties, .. } => (
            "section",
            serialize_section_properties(properties.as_ref()),
            true,
        ),
        S10SerializeRequest::Numbering { numbering, .. } => {
            ("numbering", serialize_numbering_xml(&numbering), false)
        }
        S10SerializeRequest::Watermark { watermark, .. } => {
            ("watermark", serialize_watermark(&watermark), true)
        }
    };
    let canonical_xml_events = canonical_xml_events(&xml, fragment)?;
    let parse_back = parse_back(family, &xml)?;
    Ok(S10SerializeResponse {
        wire_version: 1,
        family: family.to_owned(),
        xml,
        canonical_xml_events,
        parse_back,
    })
}

fn parse_back(family: &str, xml: &str) -> Result<serde_json::Value, ParseError> {
    let value = match family {
        "border" => match fragment_child(xml)? {
            Some(element) => {
                serde_json::to_value(crate::borders::parse_border_spec(Some(&element)))
            }
            None => serde_json::to_value(Option::<BorderSpec>::None),
        },
        "conditionalFormat" => match fragment_child(xml)? {
            Some(element) => {
                serde_json::to_value(crate::table::parse_conditional_format_style(Some(&element)))
            }
            None => serde_json::to_value(Option::<ConditionalFormatStyle>::None),
        },
        "tableGrid" => match fragment_child(xml)? {
            Some(element) => serde_json::to_value(crate::table::parse_table_grid(Some(&element))),
            None => serde_json::to_value(Option::<Vec<f64>>::None),
        },
        "section" => match fragment_child(xml)? {
            Some(element) => {
                serde_json::to_value(crate::section::parse_section_properties(Some(&element)))
            }
            None => serde_json::to_value(SectionProperties::default()),
        },
        "numbering" => {
            let limits = ParseLimits::default();
            let definitions = crate::numbering::parse_numbering(
                Some(xml.as_bytes()),
                "word/numbering.xml",
                &mut ParseBudget::new(&limits),
            )?
            .definitions;
            serde_json::to_value(definitions)
        }
        "watermark" => {
            let document = parse_fragment(xml)?;
            serde_json::to_value(crate::vml::extract_watermark(document.root(), None, None))
        }
        _ => unreachable!("all S10 families are matched above"),
    };
    value.map_err(|error| ParseError::Canonical(error.to_string()))
}

fn parse_fragment(xml: &str) -> Result<crate::xml::XmlDocument, ParseError> {
    let source = format!("{CONTENT_FRAGMENT_PREFIX}{xml}</s11:root>");
    let limits = ParseLimits::default();
    parse_xml(
        source.as_bytes(),
        "serializer-parse-back.xml",
        &mut ParseBudget::new(&limits),
    )
}

fn fragment_child(xml: &str) -> Result<Option<XmlElement>, ParseError> {
    if xml.is_empty() {
        return Ok(None);
    }
    let document = parse_fragment(xml)?;
    Ok(document
        .root()
        .and_then(|root| root.child_elements().next())
        .cloned())
}

fn parse_border_side(side: &str) -> Result<BorderSide, ParseError> {
    match side {
        "top" => Ok(BorderSide::Top),
        "bottom" => Ok(BorderSide::Bottom),
        "left" => Ok(BorderSide::Left),
        "right" => Ok(BorderSide::Right),
        "insideH" => Ok(BorderSide::InsideH),
        "insideV" => Ok(BorderSide::InsideV),
        "between" => Ok(BorderSide::Between),
        "bar" => Ok(BorderSide::Bar),
        "start" => Ok(BorderSide::Start),
        "end" => Ok(BorderSide::End),
        "tl2br" => Ok(BorderSide::TopLeftToBottomRight),
        "tr2bl" => Ok(BorderSide::TopRightToBottomLeft),
        _ => Err(ParseError::Canonical(format!(
            "unsupported serializer border side {side:?}"
        ))),
    }
}

/// Canonicalize XML as expanded namespace names, ordered child events, sorted
/// attributes, and exact decoded text. XML namespace declarations are context,
/// not attributes. Fragments are parsed under the complete S10 namespace map.
pub fn canonical_xml_events(
    xml: &str,
    fragment: bool,
) -> Result<Vec<CanonicalXmlEvent>, ParseError> {
    if xml.is_empty() {
        return Ok(Vec::new());
    }
    let owned;
    let source = if fragment {
        owned = format!("{CONTENT_FRAGMENT_PREFIX}{xml}</s11:root>");
        owned.as_str()
    } else {
        xml
    };
    let limits = ParseLimits::default();
    let document = parse_xml(
        source.as_bytes(),
        "serializer-canonical.xml",
        &mut ParseBudget::new(&limits),
    )?;
    let mut events = Vec::new();
    let namespaces = BTreeMap::from([("xml".to_owned(), XML_NAMESPACE.to_owned())]);
    if fragment {
        let root = document.root().ok_or_else(|| {
            ParseError::Canonical("serializer fragment wrapper has no root".to_owned())
        })?;
        let namespaces = namespace_context(root, &namespaces);
        for child in &root.children {
            canonicalize_node(child, &namespaces, &mut events)?;
        }
    } else {
        for root in &document.roots {
            canonicalize_element(root, &namespaces, &mut events)?;
        }
    }
    Ok(events)
}

fn canonicalize_node(
    node: &XmlNode,
    namespaces: &BTreeMap<String, String>,
    events: &mut Vec<CanonicalXmlEvent>,
) -> Result<(), ParseError> {
    match node {
        XmlNode::Element(element) => canonicalize_element(element, namespaces, events),
        XmlNode::Text(text) | XmlNode::CData(text) => {
            events.push(CanonicalXmlEvent::Text { text: text.clone() });
            Ok(())
        }
    }
}

fn canonicalize_element(
    element: &XmlElement,
    inherited_namespaces: &BTreeMap<String, String>,
    events: &mut Vec<CanonicalXmlEvent>,
) -> Result<(), ParseError> {
    let namespaces = namespace_context(element, inherited_namespaces);
    let (namespace_uri, local_name) = expanded_name(&element.name, &namespaces, true)?;
    let mut attributes = Vec::new();
    for (name, value) in &element.attributes {
        if name == "xmlns" || name.starts_with("xmlns:") {
            continue;
        }
        let (namespace_uri, local_name) = expanded_name(name, &namespaces, false)?;
        attributes.push(CanonicalXmlAttribute {
            namespace_uri,
            local_name,
            value: value.clone(),
        });
    }
    attributes.sort_by(|left, right| {
        (&left.namespace_uri, &left.local_name).cmp(&(&right.namespace_uri, &right.local_name))
    });
    events.push(CanonicalXmlEvent::Start {
        namespace_uri: namespace_uri.clone(),
        local_name: local_name.clone(),
        attributes,
    });
    for child in &element.children {
        canonicalize_node(child, &namespaces, events)?;
    }
    events.push(CanonicalXmlEvent::End {
        namespace_uri,
        local_name,
    });
    Ok(())
}

fn namespace_context(
    element: &XmlElement,
    inherited: &BTreeMap<String, String>,
) -> BTreeMap<String, String> {
    let mut namespaces = inherited.clone();
    for (name, value) in &element.attributes {
        if name == "xmlns" {
            namespaces.insert(String::new(), value.clone());
        } else if let Some(prefix) = name.strip_prefix("xmlns:") {
            namespaces.insert(prefix.to_owned(), value.clone());
        }
    }
    namespaces
}

fn expanded_name(
    name: &str,
    namespaces: &BTreeMap<String, String>,
    default_namespace: bool,
) -> Result<(String, String), ParseError> {
    if let Some((prefix, local)) = name.split_once(':') {
        let namespace_uri = namespaces.get(prefix).cloned().ok_or_else(|| {
            ParseError::Canonical(format!("unbound XML namespace prefix {prefix:?}"))
        })?;
        return Ok((namespace_uri, local.to_owned()));
    }
    Ok((
        if default_namespace {
            namespaces.get("").cloned().unwrap_or_default()
        } else {
            String::new()
        },
        name.to_owned(),
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn determinism() -> SerializerDeterminism {
        SerializerDeterminism {
            seed: "0".repeat(64),
            now: "2000-01-01T00:00:00.000Z".to_owned(),
        }
    }

    #[test]
    fn canonical_events_expand_aliases_sort_attributes_and_keep_text() {
        let first = canonical_xml_events(
            "<x:root xmlns:x=\"urn:test\" xmlns:a=\"urn:attr\" plain=\"2\" a:z=\"1\"><x:child> exact </x:child></x:root>",
            false,
        )
        .unwrap();
        let aliased = canonical_xml_events(
            "<q:root xmlns:q=\"urn:test\" xmlns:b=\"urn:attr\" b:z=\"1\" plain=\"2\"><q:child> exact </q:child></q:root>",
            false,
        )
        .unwrap();
        assert_eq!(first, aliased);
        assert_eq!(
            first[0],
            CanonicalXmlEvent::Start {
                namespace_uri: "urn:test".to_owned(),
                local_name: "root".to_owned(),
                attributes: vec![
                    CanonicalXmlAttribute {
                        namespace_uri: String::new(),
                        local_name: "plain".to_owned(),
                        value: "2".to_owned(),
                    },
                    CanonicalXmlAttribute {
                        namespace_uri: "urn:attr".to_owned(),
                        local_name: "z".to_owned(),
                        value: "1".to_owned(),
                    },
                ],
            }
        );
        assert!(first.contains(&CanonicalXmlEvent::Text {
            text: " exact ".to_owned()
        }));

        let escaped = canonical_xml_events("<w:t>safe &lt;&amp;&gt;</w:t>", true).unwrap();
        assert_eq!(
            escaped
                .iter()
                .filter_map(|event| match event {
                    CanonicalXmlEvent::Text { text } => Some(text.as_str()),
                    _ => None,
                })
                .collect::<Vec<_>>(),
            vec!["safe <&>"]
        );
    }

    #[test]
    fn wire_requires_fixed_determinism_and_is_repeatable() {
        let request = S10SerializeRequest::ConditionalFormat {
            determinism: determinism(),
            style: Some(ConditionalFormatStyle {
                first_row: Some(true),
                ..ConditionalFormatStyle::default()
            }),
        };
        let first = serialize_s10_wire(request.clone()).unwrap();
        let second = serialize_s10_wire(request).unwrap();
        assert_eq!(first, second);
        assert_eq!(first.wire_version, 1);
        assert_eq!(first.family, "conditionalFormat");
        assert_eq!(first.canonical_xml_events.len(), 2);
        assert_eq!(
            first.parse_back,
            serde_json::json!({
                "firstRow": true,
                "lastRow": false,
                "firstColumn": false,
                "lastColumn": false,
                "oddHBand": false,
                "evenHBand": false,
                "oddVBand": false,
                "evenVBand": false,
                "nwCell": false,
                "neCell": false,
                "swCell": false,
                "seCell": false,
            })
        );
    }

    #[test]
    fn rejects_invalid_seed_clock_and_dynamic_element_names() {
        let error = serialize_s10_wire(S10SerializeRequest::Border {
            determinism: SerializerDeterminism {
                seed: "random".to_owned(),
                now: "today".to_owned(),
            },
            border: None,
            side: "top\"/><evil".to_owned(),
        })
        .unwrap_err();
        assert!(error.to_string().contains("SHA-256"));

        let error = serialize_s10_wire(S10SerializeRequest::Border {
            determinism: determinism(),
            border: None,
            side: "top\"/><evil".to_owned(),
        })
        .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("unsupported serializer border side")
        );
    }
}
