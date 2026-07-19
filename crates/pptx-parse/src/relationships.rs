use serde::{Deserialize, Serialize};

use crate::xml::{ParseBudget, parse_xml};
use crate::{PptxError, xml::XmlElement};

pub mod relationship_types {
    pub const SLIDE: &str = "/slide";
    pub const SLIDE_LAYOUT: &str = "/slideLayout";
    pub const SLIDE_MASTER: &str = "/slideMaster";
    pub const THEME: &str = "/theme";
    pub const IMAGE: &str = "/image";
    pub const CHART: &str = "/chart";
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TargetMode {
    #[default]
    Internal,
    External,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Relationship {
    pub id: String,
    pub relationship_type: String,
    pub target: String,
    pub target_mode: TargetMode,
    pub resolved_target: Option<String>,
}

impl Relationship {
    pub fn has_type(&self, suffix: &str) -> bool {
        self.relationship_type.ends_with(suffix)
    }
}

pub(crate) fn parse_relationships(
    xml: &[u8],
    relationship_part: &str,
    source: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Vec<Relationship>, PptxError> {
    let root = parse_xml(xml, relationship_part, budget)?;
    let mut relationships = Vec::new();
    for element in root.children_named("Relationship") {
        budget.charge_relationship(relationship_part)?;
        relationships.push(parse_relationship(element, source)?);
    }
    Ok(relationships)
}

fn parse_relationship(element: &XmlElement, source: &str) -> Result<Relationship, PptxError> {
    let id = element.attribute("Id").unwrap_or_default().to_owned();
    let relationship_type = element.attribute("Type").unwrap_or_default().to_owned();
    let target = element.attribute("Target").unwrap_or_default().to_owned();
    let target_mode = if element.attribute("TargetMode") == Some("External") {
        TargetMode::External
    } else {
        TargetMode::Internal
    };
    let resolved_target = match target_mode {
        TargetMode::External => None,
        TargetMode::Internal => Some(resolve_target(source, &target)?),
    };
    Ok(Relationship {
        id,
        relationship_type,
        target,
        target_mode,
        resolved_target,
    })
}

pub(crate) fn resolve_target(source: &str, target: &str) -> Result<String, PptxError> {
    if target.contains("://") || target.starts_with("mailto:") || target.starts_with("data:") {
        return Err(PptxError::InvalidRelationship {
            source_part: source.to_owned(),
            target: target.to_owned(),
        });
    }
    let target = target
        .split('#')
        .next()
        .unwrap_or_default()
        .replace('\\', "/");
    let mut segments = Vec::new();
    if !target.starts_with('/')
        && let Some((directory, _)) = source.rsplit_once('/')
    {
        segments.extend(directory.split('/').filter(|segment| !segment.is_empty()));
    }
    for segment in target.trim_start_matches('/').split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                if segments.pop().is_none() {
                    return Err(PptxError::InvalidRelationship {
                        source_part: source.to_owned(),
                        target: target.to_owned(),
                    });
                }
            }
            segment => segments.push(segment),
        }
    }
    if segments.is_empty() {
        return Err(PptxError::InvalidRelationship {
            source_part: source.to_owned(),
            target: target.to_owned(),
        });
    }
    Ok(segments.join("/"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ParseLimits;

    #[test]
    fn resolves_internal_targets_without_fetching_external_targets() {
        assert_eq!(
            resolve_target("ppt/slides/slide1.xml", "../slideLayouts/slideLayout1.xml").unwrap(),
            "ppt/slideLayouts/slideLayout1.xml"
        );
        assert!(resolve_target("ppt/presentation.xml", "../../escape.xml").is_err());

        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let relationships = parse_relationships(
            br#"<Relationships><Relationship Id="rId1" Type="hyperlink" Target="https://example.invalid" TargetMode="External"/></Relationships>"#,
            "ppt/slides/_rels/slide1.xml.rels",
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap();
        assert_eq!(relationships[0].target_mode, TargetMode::External);
        assert_eq!(relationships[0].resolved_target, None);
    }
}
