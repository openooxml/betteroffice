use std::collections::{BTreeMap, HashMap, HashSet};

use crate::model::{PackagePart, VsdxPackage};
use crate::relationships::{Relationship, parse_relationships, relationship_types};
use crate::sheet::{parse_records, parse_sheet};
use crate::xml::{ParseBudget, XmlElement, parse_xml};
use crate::{ParseLimits, Sheet, VsdxError};

pub fn parse_vsdx(data: &[u8]) -> Result<VsdxPackage, VsdxError> {
    parse_vsdx_with_limits(data, &ParseLimits::default())
}

pub fn parse_vsdx_with_limits(data: &[u8], limits: &ParseLimits) -> Result<VsdxPackage, VsdxError> {
    let source_parts = ooxml_opc::unzip_parts_with_limits(data, limits.max_expanded_bytes)
        .map_err(VsdxError::Container)?;
    match ooxml_opc::detect_package_kind(&source_parts) {
        Ok(ooxml_opc::DocumentKind::Vsdx) => {}
        Ok(kind) => return Err(VsdxError::UnsupportedDocumentKind(kind)),
        Err(ooxml_opc::DocumentKindError::ConflictingMainDocumentRelationships(targets)) => {
            return Err(VsdxError::ConflictingMainDocumentRelationships(targets));
        }
        Err(ooxml_opc::DocumentKindError::MissingMainDocumentPart(part)) => {
            return Err(VsdxError::MissingPart(part));
        }
        Err(error) => return Err(VsdxError::Container(error.to_string())),
    }
    let parts: HashMap<&str, &[u8]> = source_parts
        .iter()
        .map(|(path, bytes)| (path.as_str(), bytes.as_slice()))
        .collect();
    let mut budget = ParseBudget::new(limits);
    let mut xml_parts = HashMap::new();
    let mut relationships = BTreeMap::new();
    let root_relationships = load_relationships("", &parts, &mut relationships, &mut budget)?;
    let document_path = root_relationships
        .iter()
        .find(|relationship| relationship.relationship_type == relationship_types::DOCUMENT)
        .and_then(|relationship| relationship.resolved_target.clone())
        .ok_or_else(|| VsdxError::MissingPart("root Visio document relationship".to_owned()))?;
    require_part(&parts, &document_path)?;
    parse_part(&parts, &document_path, &mut xml_parts, &mut budget)?;
    let document_relationships =
        load_relationships(&document_path, &parts, &mut relationships, &mut budget)?;
    let pages_part_path = target_by_type(document_relationships, relationship_types::PAGES);
    let masters_part_path = target_by_type(document_relationships, relationship_types::MASTERS);
    let theme_part_paths = targets_by_type(document_relationships, relationship_types::THEME);
    let windows_part_path = target_by_type(document_relationships, relationship_types::WINDOWS);
    let mut page_part_paths = Vec::new();
    if let Some(path) = &pages_part_path {
        require_part(&parts, path)?;
        parse_part(&parts, path, &mut xml_parts, &mut budget)?;
        let rels = load_relationships(path, &parts, &mut relationships, &mut budget)?;
        page_part_paths = targets_by_type(rels, relationship_types::PAGE);
        for path in &page_part_paths {
            require_part(&parts, path)?;
            parse_part(&parts, path, &mut xml_parts, &mut budget)?;
        }
    }
    let mut master_part_paths = Vec::new();
    if let Some(path) = &masters_part_path {
        require_part(&parts, path)?;
        parse_part(&parts, path, &mut xml_parts, &mut budget)?;
        let rels = load_relationships(path, &parts, &mut relationships, &mut budget)?;
        master_part_paths = targets_by_type(rels, relationship_types::MASTER);
        for path in &master_part_paths {
            require_part(&parts, path)?;
            parse_part(&parts, path, &mut xml_parts, &mut budget)?;
        }
    }
    for path in theme_part_paths.iter().chain(windows_part_path.iter()) {
        require_part(&parts, path)?;
        parse_part(&parts, path, &mut xml_parts, &mut budget)?;
    }
    load_reachable_relationships(&parts, &mut relationships, &mut budget)?;
    let document = xml_parts.remove(&document_path).expect("document parsed");
    let document_sheet = document
        .children_named("DocumentSheet")
        .next()
        .map(|sheet| parse_sheet(sheet, &document_path, &mut budget))
        .transpose()?;
    let style_sheets = document
        .children_named("StyleSheets")
        .flat_map(|styles| styles.children_named("StyleSheet"))
        .map(|style| parse_sheet(style, &document_path, &mut budget))
        .collect::<Result<Vec<_>, _>>()?;
    let colors = parse_records(document.children_named("Colors").next());
    let face_names = parse_records(document.children_named("FaceNames").next());
    let page_sheets = parse_catalog_sheets(
        xml_parts.get(pages_part_path.as_deref().unwrap_or("")),
        "Page",
        &document_path,
        &mut budget,
    )?;
    let master_sheets = parse_catalog_sheets(
        xml_parts.get(masters_part_path.as_deref().unwrap_or("")),
        "Master",
        &document_path,
        &mut budget,
    )?;
    let page_part_ids = catalog_part_ids(
        xml_parts.get(pages_part_path.as_deref().unwrap_or("")),
        "Page",
        pages_part_path.as_deref().unwrap_or(""),
        relationships
            .get(pages_part_path.as_deref().unwrap_or(""))
            .map(Vec::as_slice),
    )?;
    let master_part_ids = catalog_part_ids(
        xml_parts.get(masters_part_path.as_deref().unwrap_or("")),
        "Master",
        masters_part_path.as_deref().unwrap_or(""),
        relationships
            .get(masters_part_path.as_deref().unwrap_or(""))
            .map(Vec::as_slice),
    )?;
    let page_contents = parse_part_sheets(&page_part_paths, &mut xml_parts, &mut budget)?;
    let master_contents = parse_part_sheets(&master_part_paths, &mut xml_parts, &mut budget)?;
    Ok(VsdxPackage {
        document_part_path: document_path,
        pages_part_path,
        masters_part_path,
        page_part_paths,
        master_part_paths,
        theme_part_paths,
        windows_part_path,
        relationships,
        document_sheet,
        style_sheets,
        colors,
        face_names,
        page_sheets,
        master_sheets,
        page_part_ids,
        master_part_ids,
        page_contents,
        master_contents,
        parts: source_parts
            .into_iter()
            .map(|(path, bytes)| PackagePart { path, bytes })
            .collect(),
    })
}

fn catalog_part_ids(
    root: Option<&XmlElement>,
    item: &str,
    part: &str,
    relationships: Option<&[Relationship]>,
) -> Result<BTreeMap<String, u32>, VsdxError> {
    let mut ids = BTreeMap::new();
    for element in root.into_iter().flat_map(|root| root.children_named(item)) {
        let Some(id) = element.attribute("ID").and_then(|value| value.parse().ok()) else {
            continue;
        };
        let Some(relationship_id) = element
            .attributes
            .iter()
            .find(|(name, _)| name == "r:id" || name == "id")
            .map(|(_, value)| value.as_str())
        else {
            continue;
        };
        let resolved = relationships
            .into_iter()
            .flatten()
            .find(|relationship| relationship.id == relationship_id)
            .and_then(|relationship| relationship.resolved_target.clone());
        match resolved {
            Some(path) => {
                ids.insert(path, id);
            }
            None => {
                return Err(VsdxError::InvalidRelationship {
                    source_part: part.to_owned(),
                    target: relationship_id.to_owned(),
                });
            }
        }
    }
    Ok(ids)
}

fn parse_part_sheets(
    paths: &[String],
    xml_parts: &mut HashMap<String, XmlElement>,
    budget: &mut ParseBudget<'_>,
) -> Result<BTreeMap<String, Sheet>, VsdxError> {
    paths
        .iter()
        .map(|path| {
            let root = xml_parts
                .remove(path)
                .ok_or_else(|| VsdxError::MissingPart(path.clone()))?;
            Ok((path.clone(), parse_sheet(&root, path, budget)?))
        })
        .collect()
}

fn parse_catalog_sheets(
    root: Option<&XmlElement>,
    item: &str,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<BTreeMap<u32, Sheet>, VsdxError> {
    root.into_iter()
        .flat_map(|root| root.children_named(item))
        .map(|element| {
            let id = element
                .attribute("ID")
                .ok_or_else(|| VsdxError::MalformedXml {
                    part: part.to_owned(),
                    offset: 0,
                    message: format!("missing {item} ID"),
                })?
                .parse()
                .map_err(|_| VsdxError::MalformedXml {
                    part: part.to_owned(),
                    offset: 0,
                    message: format!("invalid {item} ID"),
                })?;
            let sheet = element
                .children_named("PageSheet")
                .next()
                .map(|sheet| parse_sheet(sheet, part, budget))
                .transpose()?
                .unwrap_or_default();
            Ok((id, sheet))
        })
        .collect()
}

fn parse_part(
    parts: &HashMap<&str, &[u8]>,
    path: &str,
    xml_parts: &mut HashMap<String, XmlElement>,
    budget: &mut ParseBudget<'_>,
) -> Result<(), VsdxError> {
    if !xml_parts.contains_key(path) {
        let bytes = parts
            .get(path)
            .ok_or_else(|| VsdxError::MissingPart(path.to_owned()))?;
        xml_parts.insert(path.to_owned(), parse_xml(bytes, path, budget)?);
    }
    Ok(())
}

fn load_reachable_relationships(
    parts: &HashMap<&str, &[u8]>,
    relationships: &mut BTreeMap<String, Vec<Relationship>>,
    budget: &mut ParseBudget<'_>,
) -> Result<(), VsdxError> {
    let mut pending: Vec<String> = relationships.keys().cloned().collect();
    let mut index = 0;
    while let Some(source) = pending.get(index).cloned() {
        index += 1;
        let targets: Vec<String> = relationships[&source]
            .iter()
            .filter_map(|relationship| relationship.resolved_target.clone())
            .collect();
        for target in targets {
            let path = relationship_path(&target);
            if parts.contains_key(path.as_str()) && !relationships.contains_key(&target) {
                load_relationships(&target, parts, relationships, budget)?;
                pending.push(target);
            }
        }
    }
    Ok(())
}

pub fn write_vsdx(package: &VsdxPackage) -> Result<Vec<u8>, VsdxError> {
    ooxml_opc::rezip_parts(
        &package
            .parts
            .iter()
            .map(|part| (part.path.clone(), part.bytes.clone()))
            .collect::<Vec<_>>(),
    )
    .map_err(VsdxError::Container)
}

fn load_relationships<'a>(
    source: &str,
    parts: &HashMap<&str, &[u8]>,
    relationships: &'a mut BTreeMap<String, Vec<Relationship>>,
    budget: &mut ParseBudget<'_>,
) -> Result<&'a [Relationship], VsdxError> {
    if !relationships.contains_key(source) {
        let path = relationship_path(source);
        let parsed = match parts.get(path.as_str()) {
            Some(bytes) => parse_relationships(bytes, &path, source, budget)?,
            None => Vec::new(),
        };
        relationships.insert(source.to_owned(), parsed);
    }
    Ok(relationships
        .get(source)
        .map(Vec::as_slice)
        .expect("inserted relationship entry"))
}
fn relationship_path(source: &str) -> String {
    if source.is_empty() {
        "_rels/.rels".to_owned()
    } else if let Some((directory, name)) = source.rsplit_once('/') {
        format!("{directory}/_rels/{name}.rels")
    } else {
        format!("_rels/{source}.rels")
    }
}
fn require_part(parts: &HashMap<&str, &[u8]>, path: &str) -> Result<(), VsdxError> {
    if parts.contains_key(path) {
        Ok(())
    } else {
        Err(VsdxError::MissingPart(path.to_owned()))
    }
}
fn target_by_type(relationships: &[Relationship], kind: &str) -> Option<String> {
    relationships
        .iter()
        .find(|relationship| relationship.has_type(kind))
        .and_then(|relationship| relationship.resolved_target.clone())
}
fn targets_by_type(relationships: &[Relationship], kind: &str) -> Vec<String> {
    let mut seen = HashSet::new();
    relationships
        .iter()
        .filter(|relationship| relationship.has_type(kind))
        .filter_map(|relationship| relationship.resolved_target.clone())
        .filter(|path| seen.insert(path.clone()))
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::serialize_sheet;
    use crate::xml::{XmlNode, parse_xml};
    use ooxml_opc::{rezip_parts, unzip_parts};
    use std::path::PathBuf;

    #[test]
    fn round_trips_committed_fixture_parts_in_order_and_byte_for_byte() {
        let source = include_bytes!("../tests/fixtures/foundation.vsdx");
        let written = write_vsdx(&parse_vsdx(source).unwrap()).unwrap();
        assert_eq!(unzip_parts(&written).unwrap(), unzip_parts(source).unwrap());
    }

    #[test]
    fn preserves_external_corpus_parts_when_available() {
        let Some(directory) = std::env::var_os("VSDX_CORPUS_DIR") else {
            eprintln!("WARNING: VSDX corpus test skipped; VSDX_CORPUS_DIR is unset");
            return;
        };
        for file in ["lichtsysteme.vsdx", "soundplan.vsdx"] {
            let path = PathBuf::from(&directory).join(file);
            assert!(path.is_file(), "missing corpus file: {}", path.display());
            let source = std::fs::read(&path).unwrap();
            let written = write_vsdx(&parse_vsdx(&source).unwrap()).unwrap();
            assert_eq!(
                parse_vsdx(&written).unwrap().page_contents,
                parse_vsdx(&source).unwrap().page_contents,
                "{}",
                path.display()
            );
            assert_eq!(
                unzip_parts(&written).unwrap(),
                unzip_parts(&source).unwrap(),
                "{}",
                path.display()
            );
        }
    }

    #[test]
    fn models_lossless_shapesheet_features() {
        let package = parse_vsdx(include_bytes!("../tests/fixtures/foundation.vsdx")).unwrap();
        let sheet = &package.page_contents["visio/pages/page1.xml"];
        let shape = sheet.shapes().next().unwrap();
        assert!(
            shape
                .cells()
                .any(|cell| cell.name == "LineWeight" && cell.del)
        );
        assert!(
            shape
                .sections()
                .any(|section| section.name == "Scratch" && section.del)
        );
        let geometry = shape
            .sections()
            .find(|section| section.name == "Geometry")
            .unwrap();
        let geometry_rows: Vec<_> = geometry.rows().collect();
        assert_eq!(geometry_rows[1].name.as_deref(), Some("LineTo"));
        assert_eq!(geometry_rows[2].index, Some(2));
        assert!(geometry_rows[2].del);
        assert_eq!(
            shape.text(),
            Some(
                [
                    crate::TextToken::Literal(" A".to_owned()),
                    crate::TextToken::CharacterRun(1),
                    crate::TextToken::Literal("B".to_owned()),
                    crate::TextToken::ParagraphRun(2),
                    crate::TextToken::Tab(3),
                    crate::TextToken::Field(0),
                    crate::TextToken::Literal(" C ".to_owned())
                ]
                .as_slice()
            )
        );
        assert_eq!(sheet.connects().next().unwrap().from_part, Some(9));
        assert_eq!(sheet.connects().next().unwrap().to_part, Some(3));
        assert_eq!(
            package.page_sheets[&1].cells().next().unwrap().name,
            "PageWidth"
        );
        assert_eq!(
            package.master_sheets[&1].cells().next().unwrap().name,
            "PageHeight"
        );
        assert_eq!(geometry_rows[0].local_name.as_deref(), Some("Start"));
        assert!(
            shape
                .cells()
                .any(|cell| cell.name == "FOnly" && cell.formula.is_some() && cell.value.is_none())
        );
        assert!(
            shape
                .cells()
                .any(|cell| cell.name == "VOnly" && cell.formula.is_none() && cell.value.is_some())
        );
        assert!(
            shape
                .cells()
                .any(|cell| cell.name == "Both" && cell.formula.is_some() && cell.value.is_some())
        );
        let user = shape
            .sections()
            .find(|section| section.name == "User")
            .unwrap();
        let user_rows: Vec<_> = user.rows().collect();
        assert_eq!(user_rows[0].name.as_deref(), Some("visVersion"));
        assert_eq!(user_rows[1].index, Some(3));
        assert_eq!(user_rows[1].row_type.as_deref(), Some("UnknownRow"));
        assert_eq!(
            user_rows[1].cells().next().unwrap().other_attrs,
            [("UnknownAttr".to_owned(), "kept".to_owned())]
        );
        let written = write_vsdx(&package).unwrap();
        assert_eq!(
            parse_vsdx(&written).unwrap().page_contents,
            package.page_contents
        );
    }

    #[test]
    fn model_serializer_matches_original_parts_and_reparse() {
        assert_model_serialization(
            include_bytes!("../tests/fixtures/foundation.vsdx"),
            "foundation.vsdx",
        );
        let Some(directory) = std::env::var_os("VSDX_CORPUS_DIR") else {
            eprintln!("WARNING: VSDX corpus comparator skipped; VSDX_CORPUS_DIR is unset");
            return;
        };
        for file in ["lichtsysteme.vsdx", "soundplan.vsdx"] {
            let path = PathBuf::from(&directory).join(file);
            assert!(path.is_file(), "missing corpus file: {}", path.display());
            assert_model_serialization(&std::fs::read(&path).unwrap(), &path.display().to_string());
        }
    }

    fn assert_model_serialization(source: &[u8], label: &str) {
        let package = parse_vsdx(source).unwrap();
        let original = unzip_parts(source).unwrap();
        let document_bytes = original
            .iter()
            .find(|(path, _)| path == &package.document_part_path)
            .unwrap()
            .1
            .as_slice();
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(document_bytes, &package.document_part_path, &mut budget).unwrap();
        if let Some(sheet) = &package.document_sheet {
            assert_sheet_matches_element(
                sheet,
                "DocumentSheet",
                document.children_named("DocumentSheet").next().unwrap(),
                label,
            );
        }
        for (sheet, original) in package.style_sheets.iter().zip(
            document
                .children_named("StyleSheets")
                .flat_map(|styles| styles.children_named("StyleSheet")),
        ) {
            assert_sheet_matches_element(sheet, "StyleSheet", original, label);
        }
        if let Some(path) = &package.pages_part_path {
            assert_catalog_sheets(&original, path, "Page", &package.page_sheets, label);
        }
        if let Some(path) = &package.masters_part_path {
            assert_catalog_sheets(&original, path, "Master", &package.master_sheets, label);
        }
        for (path, sheet) in &package.page_contents {
            let original = original
                .iter()
                .find(|(candidate, _)| candidate == path)
                .unwrap()
                .1
                .as_slice();
            let serialized = serialize_sheet("PageContents", sheet);
            assert_canonical_xml_eq(original, serialized.as_bytes(), &format!("{label}:{path}"));
            let reparsed = parse_sheet_for_test(serialized.as_bytes(), path);
            assert_eq!(&reparsed, sheet, "{label}:{path}");
        }
        for (path, sheet) in &package.master_contents {
            let original = original
                .iter()
                .find(|(candidate, _)| candidate == path)
                .unwrap()
                .1
                .as_slice();
            let serialized = serialize_sheet("MasterContents", sheet);
            assert_canonical_xml_eq(original, serialized.as_bytes(), &format!("{label}:{path}"));
            let reparsed = parse_sheet_for_test(serialized.as_bytes(), path);
            assert_eq!(&reparsed, sheet, "{label}:{path}");
        }
    }

    fn assert_catalog_sheets(
        original: &[(String, Vec<u8>)],
        path: &str,
        item: &str,
        sheets: &BTreeMap<u32, Sheet>,
        label: &str,
    ) {
        let bytes = original
            .iter()
            .find(|(candidate, _)| candidate == path)
            .unwrap()
            .1
            .as_slice();
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(bytes, path, &mut budget).unwrap();
        for item_element in root.children_named(item) {
            if let Some(original) = item_element.children_named("PageSheet").next() {
                let id = item_element.attribute("ID").unwrap().parse().unwrap();
                assert_sheet_matches_element(&sheets[&id], "PageSheet", original, label);
            }
        }
    }

    fn assert_sheet_matches_element(sheet: &Sheet, root: &str, original: &XmlElement, label: &str) {
        let serialized = serialize_sheet(root, sheet);
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let serialized = parse_xml(serialized.as_bytes(), label, &mut budget).unwrap();
        assert_eq!(canonical(original), canonical(&serialized), "{label}");
        assert_eq!(
            parse_sheet_for_test(serialize_sheet(root, sheet).as_bytes(), label),
            *sheet,
            "{label}"
        );
    }

    fn parse_sheet_for_test(bytes: &[u8], part: &str) -> Sheet {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(bytes, part, &mut budget).unwrap();
        parse_sheet(&root, part, &mut budget).unwrap()
    }

    fn assert_canonical_xml_eq(left: &[u8], right: &[u8], label: &str) {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let left = parse_xml(left, label, &mut budget).unwrap();
        let mut budget = ParseBudget::new(&limits);
        let right = parse_xml(right, label, &mut budget).unwrap();
        assert_eq!(canonical(&left), canonical(&right), "{label}");
    }

    #[derive(Debug, PartialEq, Eq)]
    enum CanonicalNode {
        Element(String, Vec<(String, String)>, Vec<CanonicalNode>),
        Text(String),
    }
    fn canonical(element: &XmlElement) -> CanonicalNode {
        let mut attributes = canonical_attributes(element);
        attributes.sort();
        let mut children = Vec::new();
        for child in &element.children {
            match child {
                XmlNode::Element(element) => children.push(CanonicalNode::Element(
                    element.local_name().to_owned(),
                    canonical_attributes(element),
                    canonical_children(element),
                )),
                XmlNode::Text(text) if !text.is_empty() => {
                    if let Some(CanonicalNode::Text(previous)) = children.last_mut() {
                        previous.push_str(text);
                    } else {
                        children.push(CanonicalNode::Text(text.clone()));
                    }
                }
                XmlNode::Text(_) => {}
            }
        }
        CanonicalNode::Element(element.local_name().to_owned(), attributes, children)
    }
    fn canonical_attributes(element: &XmlElement) -> Vec<(String, String)> {
        let mut attributes = element
            .attributes
            .iter()
            .filter(|(name, _)| name != "xmlns" && !name.starts_with("xmlns:"))
            .map(|(name, value)| {
                (
                    name.rsplit_once(':')
                        .map_or(name.as_str(), |(_, local)| local)
                        .to_owned(),
                    value.clone(),
                )
            })
            .collect::<Vec<_>>();
        attributes.sort();
        attributes
    }
    fn canonical_children(element: &XmlElement) -> Vec<CanonicalNode> {
        match canonical(element) {
            CanonicalNode::Element(_, _, children) => children,
            CanonicalNode::Text(_) => unreachable!(),
        }
    }

    #[test]
    fn enforces_sheet_resource_boundaries() {
        let source = include_bytes!("../tests/fixtures/foundation.vsdx");
        for limits in [
            ParseLimits {
                max_cells: 0,
                ..ParseLimits::default()
            },
            ParseLimits {
                max_sections: 0,
                ..ParseLimits::default()
            },
            ParseLimits {
                max_rows: 0,
                ..ParseLimits::default()
            },
            ParseLimits {
                max_shapes: 0,
                ..ParseLimits::default()
            },
        ] {
            assert!(matches!(
                parse_vsdx_with_limits(source, &limits),
                Err(VsdxError::ResourceLimit { .. })
            ));
        }
    }

    #[test]
    fn threads_the_caller_expanded_data_ceiling_into_extraction() {
        let mut parts = unzip_parts(include_bytes!("../tests/fixtures/foundation.vsdx")).unwrap();
        parts.push((
            "visio/media/filler.bin".to_owned(),
            vec![0; 8 * 1024 * 1024],
        ));
        let source = rezip_parts(&parts).unwrap();
        assert!(source.len() < 1024 * 1024, "fixture must stay compressible");
        assert!(matches!(
            parse_vsdx_with_limits(
                &source,
                &ParseLimits {
                    max_expanded_bytes: 1024 * 1024,
                    ..ParseLimits::default()
                }
            ),
            Err(VsdxError::Container(message)) if message.contains("inflated size exceeds")
        ));
        assert!(parse_vsdx_with_limits(&source, &ParseLimits::default()).is_ok());
    }

    #[test]
    fn enforces_exact_package_wide_sheet_budgets() {
        let source = rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>"#.to_vec()),
            ("visio/document.xml".to_owned(), br#"<VisioDocument><DocumentSheet><Cell N='DocumentCell'/><Section N='DocumentSection'><Row IX='0'><Cell N='DocumentRowCell'/></Row></Section></DocumentSheet><StyleSheets><StyleSheet ID='1'><Cell N='StyleCell'/><Section N='StyleSection'><Row IX='0'><Cell N='StyleRowCell'/></Row></Section></StyleSheet></StyleSheets></VisioDocument>"#.to_vec()),
            ("visio/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/pages' Target='pages/pages.xml'/></Relationships>"#.to_vec()),
            ("visio/pages/pages.xml".to_owned(), br#"<Pages><Page ID='1'/></Pages>"#.to_vec()),
            ("visio/pages/_rels/pages.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/page' Target='page1.xml'/></Relationships>"#.to_vec()),
            ("visio/pages/page1.xml".to_owned(), br#"<PageContents><Cell N='PageCell'/><Section N='PageSection'><Row IX='0'><Cell N='PageRowCell'/></Row></Section><Shapes><Shape ID='1'><Cell N='ShapeCell'/><Section N='ShapeSection'><Row IX='0'><Cell N='ShapeRowCell0'/></Row><Row IX='1'><Cell N='ShapeRowCell1'/></Row></Section><Shapes><Shape ID='2'><Cell N='NestedShapeCell'/></Shape></Shapes></Shape></Shapes></PageContents>"#.to_vec()),
        ]).unwrap();
        let limits_for = |kind: &str, value: usize| match kind {
            "cells" => ParseLimits {
                max_cells: value,
                ..ParseLimits::default()
            },
            "sections" => ParseLimits {
                max_sections: value,
                ..ParseLimits::default()
            },
            "rows" => ParseLimits {
                max_rows: value,
                ..ParseLimits::default()
            },
            "shapes" => ParseLimits {
                max_shapes: value,
                ..ParseLimits::default()
            },
            _ => unreachable!(),
        };
        for (kind, minimum) in [("cells", 10), ("sections", 4), ("rows", 5), ("shapes", 2)] {
            assert!(parse_vsdx_with_limits(&source, &limits_for(kind, minimum)).is_ok());
            assert!(matches!(
                parse_vsdx_with_limits(&source, &limits_for(kind, minimum - 1)),
                Err(VsdxError::ResourceLimit { kind: actual, .. }) if actual == kind
            ));
        }
    }

    #[test]
    fn models_namespaced_deleted_sheet_content() {
        let source = include_bytes!("../tests/fixtures/foundation.vsdx");
        let mut parts = unzip_parts(source).unwrap();
        let page = parts
            .iter_mut()
            .find(|(path, _)| path == "visio/pages/page1.xml")
            .unwrap();
        page.1 = br#"<v:PageContents xmlns:v='urn:visio'><v:Shapes><v:Shape ID='1' Del='1'><v:Cell N='PinX' Del='1'/><v:Section N='Geometry' Del='1'><v:Row IX='0' Del='1'/></v:Section></v:Shape></v:Shapes></v:PageContents>"#.to_vec();
        let package = parse_vsdx(&rezip_parts(&parts).unwrap()).unwrap();
        let shape = package.page_contents["visio/pages/page1.xml"]
            .shapes()
            .next()
            .unwrap();
        assert!(shape.del);
        assert!(shape.cells().next().unwrap().del);
        assert!(shape.sections().next().unwrap().del);
        assert!(shape.sections().next().unwrap().rows().next().unwrap().del);
    }

    #[test]
    fn discovers_parts_only_through_relationships() {
        let package = rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>"#.to_vec()),
            ("visio/document.xml".to_owned(), b"<VisioDocument/>".to_vec()),
            ("visio/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/pages' Target='pages/pages.xml'/></Relationships>"#.to_vec()),
            ("visio/pages/pages.xml".to_owned(), b"<Pages/>".to_vec()),
            ("visio/pages/_rels/pages.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/page' Target='page9.xml'/></Relationships>"#.to_vec()),
            ("visio/pages/page9.xml".to_owned(), b"<PageContents/>".to_vec()),
        ]).unwrap();
        let parsed = parse_vsdx(&package).unwrap();
        assert_eq!(parsed.page_part_paths, ["visio/pages/page9.xml"]);
    }

    #[test]
    fn opens_parts_that_carry_no_relationship_part() {
        let content_types = ("[Content_Types].xml".to_owned(), br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec());
        let root_rels = ("_rels/.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>"#.to_vec());
        let document = (
            "visio/document.xml".to_owned(),
            b"<VisioDocument/>".to_vec(),
        );
        let bare =
            rezip_parts(&[content_types.clone(), root_rels.clone(), document.clone()]).unwrap();
        let parsed = parse_vsdx(&bare).unwrap();
        assert_eq!(parsed.pages_part_path, None);
        assert!(parsed.page_part_paths.is_empty());

        let catalog = rezip_parts(&[
            content_types,
            root_rels,
            document,
            ("visio/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/pages' Target='pages/pages.xml'/></Relationships>"#.to_vec()),
            ("visio/pages/pages.xml".to_owned(), b"<Pages/>".to_vec()),
        ]).unwrap();
        let parsed = parse_vsdx(&catalog).unwrap();
        assert_eq!(
            parsed.pages_part_path.as_deref(),
            Some("visio/pages/pages.xml")
        );
        assert!(parsed.page_part_paths.is_empty());
    }

    #[test]
    fn rejects_page_catalog_entries_that_cannot_resolve() {
        let package = rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>"#.to_vec()),
            ("visio/document.xml".to_owned(), b"<VisioDocument/>".to_vec()),
            ("visio/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/pages' Target='pages/pages.xml'/></Relationships>"#.to_vec()),
            ("visio/pages/pages.xml".to_owned(), br#"<Pages><Page ID='1' r:id='rId1'/></Pages>"#.to_vec()),
        ]).unwrap();
        assert!(matches!(
            parse_vsdx(&package),
            Err(VsdxError::InvalidRelationship { .. })
        ));
    }

    #[test]
    fn rejects_master_catalog_entries_that_cannot_resolve() {
        let package = rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>"#.to_vec()),
            ("visio/document.xml".to_owned(), b"<VisioDocument/>".to_vec()),
            ("visio/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/masters' Target='masters/masters.xml'/></Relationships>"#.to_vec()),
            ("visio/masters/masters.xml".to_owned(), br#"<Masters><Master ID='1' r:id='rId1'/></Masters>"#.to_vec()),
        ]).unwrap();
        assert!(matches!(
            parse_vsdx(&package),
            Err(VsdxError::InvalidRelationship { .. })
        ));
    }

    #[test]
    fn rejects_dangling_relationship_targets() {
        let package = rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec()),
            (
                "_rels/.rels".to_owned(),
                br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>"#.to_vec(),
            ),
        ])
        .unwrap();
        assert!(matches!(
            parse_vsdx(&package),
            Err(VsdxError::MissingPart(_))
        ));
    }

    #[test]
    fn validates_reachable_page_relationship_parts() {
        let package = rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>"#.to_vec()),
            ("visio/document.xml".to_owned(), b"<VisioDocument/>".to_vec()),
            ("visio/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/pages' Target='pages/pages.xml'/></Relationships>"#.to_vec()),
            ("visio/pages/pages.xml".to_owned(), b"<Pages/>".to_vec()),
            ("visio/pages/_rels/pages.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/page' Target='page1.xml'/></Relationships>"#.to_vec()),
            ("visio/pages/page1.xml".to_owned(), b"<PageContents/>".to_vec()),
            ("visio/pages/_rels/page1.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='x/image' Target='javascript:x'/></Relationships>"#.to_vec()),
        ]).unwrap();
        assert!(matches!(
            parse_vsdx(&package),
            Err(VsdxError::InvalidRelationship { .. })
        ));
    }

    #[test]
    fn rejects_unsupported_or_conflicting_document_kinds() {
        for content_type in [
            "application/vnd.ms-visio.drawing.macroEnabled.main+xml",
            "application/vnd.ms-visio.stencil.main+xml",
            "application/vnd.ms-visio.template.main+xml",
        ] {
            let package = rezip_parts(&[
                ("_rels/.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>"#.to_vec()),
                ("[Content_Types].xml".to_owned(), format!("<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='{content_type}'/></Types>").into_bytes()),
                ("visio/document.xml".to_owned(), b"<VisioDocument/>".to_vec()),
            ]).unwrap();
            assert!(matches!(
                parse_vsdx(&package),
                Err(VsdxError::UnsupportedDocumentKind(_))
            ));
        }
        let package = rezip_parts(&[
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/><Relationship Id='r2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>"#.to_vec()),
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>"#.to_vec()),
            ("visio/document.xml".to_owned(), b"<VisioDocument/>".to_vec()),
            ("word/document.xml".to_owned(), b"<document/>".to_vec()),
        ]).unwrap();
        assert!(matches!(
            parse_vsdx(&package),
            Err(VsdxError::ConflictingMainDocumentRelationships(_))
        ));
    }

    #[test]
    fn rejects_missing_or_wrong_content_types() {
        let missing = rezip_parts(&[(
            "visio/document.xml".to_owned(),
            b"<VisioDocument/>".to_vec(),
        )])
        .unwrap();
        assert!(parse_vsdx(&missing).is_err());
        let wrong = rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>"#.to_vec()),
            ("word/document.xml".to_owned(), b"<document/>".to_vec()),
        ]).unwrap();
        assert!(matches!(
            parse_vsdx(&wrong),
            Err(VsdxError::UnsupportedDocumentKind(_))
        ));
    }
}
