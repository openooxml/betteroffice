//! Complete S2 package/leaf projection used by the differential corpus gate.

use base64::Engine as _;
use serde::{Deserialize, Serialize};

use crate::borders::{Borders, parse_paragraph_borders, parse_table_borders};
use crate::canonical::{canonical_sha256, from_serializable, to_canonical_bytes};
use crate::fonts::{FontTable, parse_font_table};
use crate::scalars::{
    ColorValue, RunScalarProperties, ShadingProperties, parse_color_value,
    parse_run_scalar_properties, parse_shading_properties,
};
use crate::settings::incumbent_utf8_text_boundary;
use crate::settings::{DocumentSettings, parse_settings};
use crate::tabs::{TabStop, parse_tab_stops};
use crate::theme::{Theme, parse_theme};
use crate::xml::{ParseBudget, ParseError, ParseLimits, XmlElement, XmlNode, parse_xml};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S2Projection {
    pub settings: DocumentSettings,
    pub theme: Theme,
    pub font_table: FontTable,
    pub xml_parts: Vec<S2XmlPart>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S2XmlPart {
    pub path: String,
    pub border_containers: Vec<ElementLeaf<Borders>>,
    pub tab_containers: Vec<ElementLeaf<Vec<TabStop>>>,
    pub shadings: Vec<ElementLeaf<ShadingProperties>>,
    pub colors: Vec<ElementLeaf<ColorValue>>,
    pub run_scalars: Vec<ElementLeaf<RunScalarProperties>>,
}

impl S2XmlPart {
    fn is_empty(&self) -> bool {
        self.border_containers.is_empty()
            && self.tab_containers.is_empty()
            && self.shadings.is_empty()
            && self.colors.is_empty()
            && self.run_scalars.is_empty()
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ElementLeaf<T> {
    pub element: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<T>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S2WireEnvelope {
    pub wire_version: u8,
    pub projection: S2Projection,
    pub canonical_base64: String,
    pub canonical_sha256: String,
}

pub fn parse_docx_s2_projection(data: &[u8]) -> Result<S2Projection, ParseError> {
    let parts = ooxml_opc::unzip_parts(data).map_err(ParseError::Container)?;
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let settings_bytes = find_part(&parts, "word/settings.xml");
    let theme_bytes = find_part(&parts, "word/theme/theme1.xml");
    let font_table_bytes = find_part(&parts, "word/fontTable.xml");
    let settings = parse_settings(settings_bytes, "word/settings.xml", &mut budget)?;
    let theme = parse_theme(theme_bytes, "word/theme/theme1.xml", &mut budget)?;
    let font_table = parse_font_table(font_table_bytes, "word/fontTable.xml", &mut budget)?;

    let mut xml_parts = Vec::new();
    for (path, bytes) in &parts {
        if !path.to_ascii_lowercase().ends_with(".xml") {
            continue;
        }
        if let Some(part) = project_xml_part(bytes, path, &mut budget)? {
            xml_parts.push(part);
        }
    }
    Ok(S2Projection {
        settings,
        theme,
        font_table,
        xml_parts,
    })
}

pub fn s2_wire_envelope(projection: S2Projection) -> Result<S2WireEnvelope, ParseError> {
    let canonical =
        from_serializable(&projection).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let bytes =
        to_canonical_bytes(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let sha =
        canonical_sha256(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    Ok(S2WireEnvelope {
        wire_version: 1,
        projection,
        canonical_base64: base64::engine::general_purpose::STANDARD.encode(bytes),
        canonical_sha256: sha,
    })
}

pub fn parse_docx_s2_wire(data: &[u8]) -> Result<S2WireEnvelope, ParseError> {
    s2_wire_envelope(parse_docx_s2_projection(data)?)
}

pub fn project_xml_part(
    xml: &[u8],
    path: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Option<S2XmlPart>, ParseError> {
    if !incumbent_utf8_text_boundary(xml) {
        return Ok(None);
    }
    let document = parse_xml(xml, path, budget)?;
    let mut projection = S2XmlPart {
        path: path.to_owned(),
        border_containers: Vec::new(),
        tab_containers: Vec::new(),
        shadings: Vec::new(),
        colors: Vec::new(),
        run_scalars: Vec::new(),
    };
    let mut stack: Vec<&XmlElement> = document.roots.iter().rev().collect();
    while let Some(element) = stack.pop() {
        let local_name = element.local_name();
        match local_name {
            "pBdr" => {
                budget.charge_leaf_value(path)?;
                projection.border_containers.push(ElementLeaf {
                    element: local_name.to_owned(),
                    value: parse_paragraph_borders(Some(element)),
                });
            }
            "tblBorders" | "tcBorders" | "pgBorders" => {
                budget.charge_leaf_value(path)?;
                projection.border_containers.push(ElementLeaf {
                    element: local_name.to_owned(),
                    value: parse_table_borders(Some(element)),
                });
            }
            "tabs" => {
                budget.charge_leaf_value(path)?;
                projection.tab_containers.push(ElementLeaf {
                    element: local_name.to_owned(),
                    value: Some(parse_tab_stops(Some(element))),
                });
            }
            "shd" => {
                budget.charge_leaf_value(path)?;
                projection.shadings.push(ElementLeaf {
                    element: local_name.to_owned(),
                    value: parse_shading_properties(Some(element)),
                });
            }
            "color" => {
                budget.charge_leaf_value(path)?;
                projection.colors.push(ElementLeaf {
                    element: local_name.to_owned(),
                    value: Some(parse_color_value(
                        element.attribute(Some("w"), "val"),
                        element.attribute(Some("w"), "themeColor"),
                        element.attribute(Some("w"), "themeTint"),
                        element.attribute(Some("w"), "themeShade"),
                    )),
                });
            }
            "rPr" => {
                budget.charge_leaf_value(path)?;
                projection.run_scalars.push(ElementLeaf {
                    element: local_name.to_owned(),
                    value: parse_run_scalar_properties(Some(element)),
                });
            }
            _ => {}
        }
        for child in element.children.iter().rev() {
            if let XmlNode::Element(child) = child {
                stack.push(child);
            }
        }
    }
    Ok((!projection.is_empty()).then_some(projection))
}

fn find_part<'a>(parts: &'a [(String, Vec<u8>)], path: &str) -> Option<&'a [u8]> {
    parts
        .iter()
        .find(|(candidate, _)| candidate == path)
        .map(|(_, bytes)| bytes.as_slice())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn projects_every_s2_leaf_in_document_order_and_encodes_canonically() {
        let xml = br#"<w:root><w:pBdr><w:top w:val="single"/></w:pBdr><w:tabs><w:tab w:pos="720" w:val="right"/></w:tabs><w:rPr><w:u w:val="double"/><w:color w:val="auto"/><w:highlight w:val="yellow"/><w:shd w:fill="FF0000"/><w:sz w:val="24"/></w:rPr></w:root>"#;
        let limits = ParseLimits::default();
        let part = project_xml_part(xml, "word/document.xml", &mut ParseBudget::new(&limits))
            .unwrap()
            .unwrap();
        assert_eq!(part.border_containers.len(), 1);
        assert_eq!(
            part.tab_containers[0].value.as_ref().unwrap()[0].position,
            720.0
        );
        assert_eq!(part.run_scalars.len(), 1);
        assert_eq!(part.colors.len(), 1);
        assert_eq!(part.shadings.len(), 1);

        let projection = S2Projection {
            settings: DocumentSettings::default(),
            theme: Theme::default(),
            font_table: FontTable::default(),
            xml_parts: vec![part],
        };
        let wire = s2_wire_envelope(projection).unwrap();
        let bytes = base64::engine::general_purpose::STANDARD
            .decode(wire.canonical_base64)
            .unwrap();
        assert!(
            String::from_utf8(bytes)
                .unwrap()
                .starts_with("docx-document-canonical-v1\n{")
        );
        assert_eq!(wire.canonical_sha256.len(), 64);
    }

    #[test]
    fn shared_leaf_budget_is_enforced_before_projection_growth() {
        let limits = ParseLimits {
            max_leaf_values: 1,
            ..ParseLimits::default()
        };
        let error = project_xml_part(
            b"<w:root><w:color/><w:color/></w:root>",
            "word/document.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap_err();
        assert!(matches!(
            error,
            ParseError::ResourceLimit {
                kind: "leafValues",
                ..
            }
        ));
    }
}
