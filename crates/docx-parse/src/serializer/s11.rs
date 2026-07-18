//! S11 versioned content-serializer wire and parse-back gate.

use serde::{Deserialize, Serialize};

use crate::block::{BlockContent, BlockSdt, StoryParser};
use crate::chart::ChartPartsMap;
use crate::image::Image;
use crate::inline::{InlineNode, InlineSdt, Run};
use crate::media::MediaMap;
use crate::paragraph::{
    DrawingContext, HexIdAllocator, Paragraph, ParagraphContent, parse_run_composed,
};
use crate::shape::Shape;
use crate::smart_art::SmartArtContext;
use crate::table::Table;
use crate::xml::{ParseBudget, ParseError, ParseLimits, XmlElement, parse_xml};

use super::context::SerializerContext;
use super::paragraph::{serialize_inline_sdt, serialize_paragraph};
use super::raw::CONTENT_FRAGMENT_PREFIX;
use super::run::{serialize_drawing_content, serialize_run, serialize_shape_content};
use super::s10::{CanonicalXmlEvent, SerializerDeterminism, canonical_xml_events};
use super::sdt::serialize_block_sdt;
use super::table::serialize_table;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "family", rename_all = "camelCase", deny_unknown_fields)]
pub enum S11SerializeRequest {
    Run {
        determinism: SerializerDeterminism,
        run: Run,
    },
    Drawing {
        determinism: SerializerDeterminism,
        image: Image,
    },
    Shape {
        determinism: SerializerDeterminism,
        shape: Shape,
    },
    Paragraph {
        determinism: SerializerDeterminism,
        paragraph: Paragraph,
    },
    Table {
        determinism: SerializerDeterminism,
        table: Table,
    },
    InlineSdt {
        determinism: SerializerDeterminism,
        sdt: InlineSdt,
    },
    BlockSdt {
        determinism: SerializerDeterminism,
        sdt: BlockSdt,
    },
}

impl S11SerializeRequest {
    fn determinism(&self) -> &SerializerDeterminism {
        match self {
            Self::Run { determinism, .. }
            | Self::Drawing { determinism, .. }
            | Self::Shape { determinism, .. }
            | Self::Paragraph { determinism, .. }
            | Self::Table { determinism, .. }
            | Self::InlineSdt { determinism, .. }
            | Self::BlockSdt { determinism, .. } => determinism,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S11SerializeResponse {
    pub wire_version: u8,
    pub family: String,
    pub xml: String,
    pub canonical_xml_events: Vec<CanonicalXmlEvent>,
    pub parse_back: serde_json::Value,
}

pub fn serialize_s11_wire(
    request: S11SerializeRequest,
) -> Result<S11SerializeResponse, ParseError> {
    let determinism = request.determinism().clone();
    let mut context = SerializerContext::new(&determinism)?;
    let (family, xml) = match request {
        S11SerializeRequest::Run { run, .. } => ("run", serialize_run(&run, &mut context)?),
        S11SerializeRequest::Drawing { image, .. } => {
            ("drawing", serialize_drawing_content(&image, &mut context)?)
        }
        S11SerializeRequest::Shape { shape, .. } => {
            ("shape", serialize_shape_content(&shape, &mut context)?)
        }
        S11SerializeRequest::Paragraph { paragraph, .. } => {
            ("paragraph", serialize_paragraph(&paragraph, &mut context)?)
        }
        S11SerializeRequest::Table { table, .. } => {
            ("table", serialize_table(&table, &mut context)?)
        }
        S11SerializeRequest::InlineSdt { sdt, .. } => {
            ("inlineSdt", serialize_inline_sdt(&sdt, &mut context)?)
        }
        S11SerializeRequest::BlockSdt { sdt, .. } => {
            ("blockSdt", serialize_block_sdt(&sdt, &mut context)?)
        }
    };
    Ok(S11SerializeResponse {
        wire_version: 1,
        family: family.to_owned(),
        canonical_xml_events: canonical_xml_events(&xml, true)?,
        parse_back: parse_back(family, &xml, &determinism.seed)?,
        xml,
    })
}

fn parse_back(family: &str, xml: &str, seed: &str) -> Result<serde_json::Value, ParseError> {
    let value = match family {
        "drawing" => {
            let child = fragment_child(xml)?.ok_or_else(|| {
                ParseError::Canonical("serialized drawing has no root".to_owned())
            })?;
            serde_json::to_value(crate::image::parse_drawing(&child, None, None))
        }
        "shape" => {
            let child = fragment_child(xml)?
                .ok_or_else(|| ParseError::Canonical("serialized shape has no root".to_owned()))?;
            serde_json::to_value(crate::shape::parse_shape_from_drawing(&child))
        }
        "run" => serde_json::to_value(parse_run_back(xml)?),
        "inlineSdt" => serde_json::to_value(find_inline(
            parse_story(&format!("<w:p>{xml}</w:p>"), seed)?,
            "inlineSdt",
        )),
        "paragraph" | "table" | "blockSdt" => {
            let blocks = parse_story(xml, seed)?;
            serde_json::to_value(blocks.into_iter().next())
        }
        _ => unreachable!("all S11 families are matched above"),
    };
    value.map_err(|error| ParseError::Canonical(error.to_string()))
}

fn parse_run_back(xml: &str) -> Result<Run, ParseError> {
    let child = fragment_child(xml)?
        .ok_or_else(|| ParseError::Canonical("serialized run has no root".to_owned()))?;
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let media = MediaMap::new();
    let charts = ChartPartsMap::new();
    let mut smart_art = SmartArtContext::default();
    let mut drawing = DrawingContext {
        media: &media,
        charts: &charts,
        smart_art: &mut smart_art,
    };
    parse_run_composed(
        &child,
        None,
        None,
        None,
        None,
        &mut budget,
        Some(&mut drawing),
    )
}

fn find_inline(blocks: Vec<BlockContent>, node_type: &str) -> Option<InlineNode> {
    let BlockContent::Paragraph(paragraph) = blocks.into_iter().next()? else {
        return None;
    };
    paragraph.content.into_iter().find_map(|content| {
        let ParagraphContent::Inline(node) = content else {
            return None;
        };
        (node.node_type() == node_type).then_some(node)
    })
}

fn parse_story(xml: &str, seed: &str) -> Result<Vec<BlockContent>, ParseError> {
    let document = parse_fragment(xml)?;
    let root = document.root().ok_or_else(|| {
        ParseError::Canonical("serializer parse-back wrapper has no root".to_owned())
    })?;
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let media = MediaMap::new();
    let charts = ChartPartsMap::new();
    let mut smart_art = SmartArtContext::default();
    let mut ids = HexIdAllocator::from_sha256(seed)?;
    StoryParser {
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
        part: "serializer-parse-back.xml",
    }
    .parse_blocks(root, 0, false)
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
    let document = parse_fragment(xml)?;
    Ok(document
        .root()
        .and_then(|root| root.child_elements().next())
        .cloned())
}

#[cfg(test)]
mod tests {
    use crate::inline::{RunContent, RunType};

    use super::*;

    fn determinism() -> SerializerDeterminism {
        SerializerDeterminism {
            seed: "0".repeat(64),
            now: "2000-01-01T00:00:00.000Z".to_owned(),
        }
    }

    #[test]
    fn run_wire_is_repeatable_and_parse_backs_through_the_story_parser() {
        let request = S11SerializeRequest::Run {
            determinism: determinism(),
            run: Run {
                node_type: RunType::Run,
                formatting: None,
                property_changes: None,
                content: vec![RunContent::Text {
                    text: "safe <&>".to_owned(),
                    preserve_space: None,
                }],
            },
        };
        let first = serialize_s11_wire(request.clone()).unwrap();
        assert_eq!(first, serialize_s11_wire(request).unwrap());
        assert_eq!(first.family, "run");
        assert_eq!(first.xml, "<w:r><w:t>safe &lt;&amp;&gt;</w:t></w:r>");
        assert_eq!(first.parse_back["type"], "run");
    }
}
