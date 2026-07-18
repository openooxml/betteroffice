//! Typed, insertion-preserving OPC relationship parsing.

use base64::Engine as _;
use indexmap::IndexMap;
use serde::{Deserialize, Serialize};

use crate::canonical::{CanonicalValue, canonical_sha256, to_canonical_bytes};
use crate::xml::{ParseBudget, ParseError, ParseLimits, parse_xml};

pub type RelationshipMap = IndexMap<String, Relationship>;

pub mod relationship_types {
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const HEADER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    pub const FOOTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    pub const FOOTNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    pub const ENDNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const NUMBERING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    pub const FONT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    pub const THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    pub const SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    pub const WEB_SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
    pub const OLE_OBJECT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    pub const CHART: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    pub const DIAGRAM_DATA: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
    pub const OFFICE_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    pub const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    pub const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    pub const CUSTOM_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    pub const CUSTOM_XML: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
    pub const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub const COMMENTS_EXTENDED: &str =
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
    pub const COMMENTS_IDS: &str =
        "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
    pub const COMMENTS_EXTENSIBLE: &str =
        "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Relationship {
    pub id: String,
    #[serde(rename = "type")]
    pub relationship_type: String,
    pub target: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub target_mode: Option<TargetMode>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum TargetMode {
    External,
    Internal,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum RelationshipTarget {
    Internal(String),
    External(String),
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WireEnvelope {
    pub wire_version: u8,
    pub relationship_parts: Vec<RelationshipPart>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RelationshipPart {
    pub path: String,
    pub relationships: Vec<(String, Relationship)>,
    pub canonical_base64: String,
    pub canonical_sha256: String,
}

impl RelationshipPart {
    pub fn from_map(path: String, relationships: RelationshipMap) -> Result<Self, ParseError> {
        let canonical = relationship_map_canonical_value(&relationships);
        let bytes = to_canonical_bytes(&canonical)
            .map_err(|error| ParseError::Canonical(error.to_string()))?;
        let sha = canonical_sha256(&canonical)
            .map_err(|error| ParseError::Canonical(error.to_string()))?;
        Ok(Self {
            path,
            relationships: relationships.into_iter().collect(),
            canonical_base64: base64::engine::general_purpose::STANDARD.encode(bytes),
            canonical_sha256: sha,
        })
    }
}

pub fn parse_relationships(
    xml: &[u8],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<RelationshipMap, ParseError> {
    let mut relationships = RelationshipMap::new();
    if xml.iter().all(u8::is_ascii_whitespace) {
        return Ok(relationships);
    }
    let document = parse_xml(xml, part, budget)?;
    let Some(root) = document.root() else {
        return Ok(relationships);
    };

    for child in root.child_elements() {
        // Preserve the incumbent's permissive qualified-name check exactly.
        if !child.name.ends_with("Relationship") && !child.name.contains(":Relationship") {
            continue;
        }
        budget.charge_relationship(part)?;
        let id = child.attribute(None, "Id").unwrap_or_default();
        let relationship_type = child.attribute(None, "Type").unwrap_or_default();
        let target = child.attribute(None, "Target").unwrap_or_default();
        if id.is_empty() || relationship_type.is_empty() || target.is_empty() {
            // TypeScript warns and skips incomplete entries; warnings are host
            // diagnostics and are intentionally not part of RelationshipMap.
            continue;
        }
        let target_mode = match child.attribute(None, "TargetMode") {
            Some("External") => Some(TargetMode::External),
            Some("Internal") => Some(TargetMode::Internal),
            _ => None,
        };
        relationships.insert(
            id.to_owned(),
            Relationship {
                id: id.to_owned(),
                relationship_type: relationship_type.to_owned(),
                target: target.to_owned(),
                target_mode,
            },
        );
    }
    Ok(relationships)
}

pub fn get_relationship_type_name(relationship_type: &str) -> &str {
    use relationship_types::*;
    match relationship_type {
        IMAGE => "image",
        HYPERLINK => "hyperlink",
        HEADER => "header",
        FOOTER => "footer",
        FOOTNOTES => "footnotes",
        ENDNOTES => "endnotes",
        STYLES => "styles",
        NUMBERING => "numbering",
        FONT_TABLE => "fontTable",
        THEME => "theme",
        SETTINGS => "settings",
        WEB_SETTINGS => "webSettings",
        OLE_OBJECT => "oleObject",
        CHART => "chart",
        DIAGRAM_DATA => "diagramData",
        OFFICE_DOCUMENT => "officeDocument",
        CORE_PROPERTIES => "coreProperties",
        EXTENDED_PROPERTIES => "extendedProperties",
        CUSTOM_PROPERTIES => "customProperties",
        CUSTOM_XML => "customXml",
        COMMENTS => "comments",
        COMMENTS_EXTENDED => "commentsExtended",
        COMMENTS_IDS => "commentsIds",
        COMMENTS_EXTENSIBLE => "commentsExtensible",
        value => value.rsplit_once('/').map_or("unknown", |(_, name)| name),
    }
}

pub fn is_external_hyperlink(relationship: &Relationship) -> bool {
    relationship.relationship_type == relationship_types::HYPERLINK
        && relationship.target_mode == Some(TargetMode::External)
}

pub fn is_image_relationship(relationship: &Relationship) -> bool {
    relationship.relationship_type == relationship_types::IMAGE
}

pub fn is_header_relationship(relationship: &Relationship) -> bool {
    relationship.relationship_type == relationship_types::HEADER
}

pub fn is_footer_relationship(relationship: &Relationship) -> bool {
    relationship.relationship_type == relationship_types::FOOTER
}

pub fn filter_by_type<'a>(
    map: &'a RelationshipMap,
    relationship_type: &str,
) -> Vec<&'a Relationship> {
    map.values()
        .filter(|relationship| relationship.relationship_type == relationship_type)
        .collect()
}

pub fn images(map: &RelationshipMap) -> Vec<&Relationship> {
    filter_by_type(map, relationship_types::IMAGE)
}

pub fn hyperlinks(map: &RelationshipMap) -> Vec<&Relationship> {
    filter_by_type(map, relationship_types::HYPERLINK)
}

pub fn headers(map: &RelationshipMap) -> Vec<&Relationship> {
    filter_by_type(map, relationship_types::HEADER)
}

pub fn footers(map: &RelationshipMap) -> Vec<&Relationship> {
    filter_by_type(map, relationship_types::FOOTER)
}

pub fn resolve_target<'a>(map: &'a RelationshipMap, id: &str) -> Option<&'a str> {
    map.get(id).map(|relationship| relationship.target.as_str())
}

pub fn resolve_relationship<'a>(map: &'a RelationshipMap, id: &str) -> Option<&'a Relationship> {
    map.get(id)
}

/// Direct package path: `ooxml-opc` owns ZIP budgets/path validation.
pub fn parse_docx_relationship_parts(data: &[u8]) -> Result<WireEnvelope, ParseError> {
    let parts = ooxml_opc::unzip_parts(data).map_err(ParseError::Container)?;
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let mut relationship_parts = Vec::new();
    for (path, bytes) in parts {
        if !path.to_ascii_lowercase().ends_with(".rels") {
            continue;
        }
        let relationships = parse_relationships(&bytes, &path, &mut budget)?;
        relationship_parts.push(RelationshipPart::from_map(path, relationships)?);
    }
    Ok(WireEnvelope {
        wire_version: 1,
        relationship_parts,
    })
}

pub fn relationship_map_canonical_value(map: &RelationshipMap) -> CanonicalValue {
    CanonicalValue::OrderedMap(
        map.iter()
            .map(|(id, relationship)| {
                let mut fields = vec![
                    (
                        "id".to_owned(),
                        CanonicalValue::String(relationship.id.clone()),
                    ),
                    (
                        "type".to_owned(),
                        CanonicalValue::String(relationship.relationship_type.clone()),
                    ),
                    (
                        "target".to_owned(),
                        CanonicalValue::String(relationship.target.clone()),
                    ),
                ];
                if let Some(target_mode) = relationship.target_mode {
                    fields.push((
                        "targetMode".to_owned(),
                        CanonicalValue::String(
                            match target_mode {
                                TargetMode::External => "External",
                                TargetMode::Internal => "Internal",
                            }
                            .to_owned(),
                        ),
                    ));
                }
                (id.clone(), CanonicalValue::Object(fields))
            })
            .collect(),
    )
}

/// Resolve an OPC target while rejecting traversal above the package root.
pub fn resolve_relative_path(base_path: &str, target: &str) -> Result<String, ParseError> {
    if target.contains('\0') || has_drive_prefix(target) {
        return Err(unsafe_relationship(base_path, target));
    }
    let target = target.replace('\\', "/");
    let absolute = target.starts_with('/');
    let mut directory = if absolute {
        Vec::new()
    } else {
        relationship_base_directory(base_path)
    };

    for part in target.trim_start_matches('/').split('/') {
        if part.is_empty() || part == "." {
            continue;
        }
        if part == ".." || is_percent_encoded_parent(part) {
            if directory.pop().is_none() {
                return Err(unsafe_relationship(base_path, &target));
            }
        } else {
            directory.push(part.to_owned());
        }
    }
    Ok(directory.join("/"))
}

pub fn resolve_relationship_target(
    base_path: &str,
    relationship: &Relationship,
) -> Result<RelationshipTarget, ParseError> {
    if relationship.target_mode == Some(TargetMode::External) {
        // External targets are inert data. This crate deliberately has no HTTP
        // client, filesystem resolver, or callback that could dereference one.
        return Ok(RelationshipTarget::External(relationship.target.clone()));
    }
    resolve_relative_path(base_path, &relationship.target).map(RelationshipTarget::Internal)
}

fn relationship_base_directory(base_path: &str) -> Vec<String> {
    let normalized = base_path.replace('\\', "/");
    let directory = normalized
        .rsplit_once('/')
        .map_or("", |(directory, _)| directory);
    let directory = directory
        .strip_suffix("/_rels")
        .or_else(|| (directory == "_rels").then_some(""))
        .unwrap_or(directory);
    directory
        .split('/')
        .filter(|part| !part.is_empty() && *part != ".")
        .map(str::to_owned)
        .collect()
}

fn has_drive_prefix(path: &str) -> bool {
    path.as_bytes().get(1) == Some(&b':')
}

fn is_percent_encoded_parent(part: &str) -> bool {
    matches!(
        part.to_ascii_lowercase().as_str(),
        "%2e%2e" | "%2e." | ".%2e"
    )
}

fn unsafe_relationship(part: &str, target: &str) -> ParseError {
    ParseError::Relationship {
        part: part.to_owned(),
        message: format!("unsafe internal target path {target:?}"),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const RELS: &str = r#"<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId2" Type="image" Target="media/a.png"/>
        <Relationship Id="rId1" Type="hyperlink" Target="https://example.test/x" TargetMode="External"/>
        <Relationship Id="rId2" Type="image-new" Target="media/b.png" TargetMode="Internal"/>
      </Relationships>"#;

    fn parse(input: &str) -> Result<RelationshipMap, ParseError> {
        let limits = ParseLimits::default();
        parse_relationships(
            input.as_bytes(),
            "word/_rels/document.xml.rels",
            &mut ParseBudget::new(&limits),
        )
    }

    #[test]
    fn preserves_map_order_and_duplicate_set_position() {
        let map = parse(RELS).unwrap();
        assert_eq!(
            map.keys().map(String::as_str).collect::<Vec<_>>(),
            ["rId2", "rId1"]
        );
        assert_eq!(map["rId2"].target, "media/b.png");
        assert_eq!(map["rId2"].target_mode, Some(TargetMode::Internal));
    }

    #[test]
    fn canonical_relationship_bytes_match_the_typescript_shape() {
        let part = RelationshipPart::from_map("x.rels".to_owned(), parse(RELS).unwrap()).unwrap();
        let bytes = base64::engine::general_purpose::STANDARD
            .decode(part.canonical_base64)
            .unwrap();
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.starts_with("docx-document-canonical-v1\n{\"$map\":[[\"rId2\","));
        assert!(text.contains("\"targetMode\":\"External\""));
    }

    #[test]
    fn resolves_internal_paths_and_rejects_traversal() {
        assert_eq!(
            resolve_relative_path("word/_rels/document.xml.rels", "media/image.png").unwrap(),
            "word/media/image.png"
        );
        assert_eq!(
            resolve_relative_path("_rels/.rels", "/word/document.xml").unwrap(),
            "word/document.xml"
        );
        for target in [
            "../../evil",
            "..\\..\\evil",
            "%2e%2e/%2e%2e/evil",
            "C:/evil",
        ] {
            assert!(
                resolve_relative_path("_rels/.rels", target).is_err(),
                "{target}"
            );
        }
    }

    #[test]
    fn external_target_is_recorded_but_never_resolved_or_fetched() {
        let relationship = &parse(RELS).unwrap()["rId1"];
        assert_eq!(
            resolve_relationship_target("word/_rels/document.xml.rels", relationship).unwrap(),
            RelationshipTarget::External("https://example.test/x".to_owned())
        );
    }

    #[test]
    fn relationship_type_and_filter_helpers_match_the_incumbent() {
        let mut map = parse(RELS).unwrap();
        map["rId1"].relationship_type = relationship_types::HYPERLINK.to_owned();
        assert_eq!(
            get_relationship_type_name(relationship_types::HYPERLINK),
            "hyperlink"
        );
        assert_eq!(
            get_relationship_type_name("urn:custom/no-slash"),
            "no-slash"
        );
        assert!(is_external_hyperlink(&map["rId1"]));
        assert_eq!(hyperlinks(&map).len(), 1);
        assert_eq!(resolve_target(&map, "rId1"), Some("https://example.test/x"));
        assert!(resolve_relationship(&map, "missing").is_none());
    }

    #[test]
    fn relationship_limit_is_shared_and_checked_before_insert() {
        let limits = ParseLimits {
            max_relationships: 2,
            ..ParseLimits::default()
        };
        let mut budget = ParseBudget::new(&limits);
        let error =
            parse_relationships(RELS.as_bytes(), "word/_rels/document.xml.rels", &mut budget)
                .unwrap_err();
        assert!(matches!(
            error,
            ParseError::ResourceLimit {
                kind: "relationships",
                ..
            }
        ));
    }
}
