use std::collections::{BTreeMap, HashMap, HashSet};

use crate::model::{PackagePart, VsdxPackage};
use crate::patch::{
    AttributeSpan, CellEdit, MAX_PATCH_BYTES, MAX_PATCH_EDITS, SpanEdit, apply_span_edits,
    escape_attribute_value, scan_element_spans,
};
use crate::relationships::{Relationship, parse_relationships, relationship_types};
use crate::sheet::{parse_records, parse_sheet};
use crate::xml::{ParseBudget, XmlElement, XmlNode, parse_xml};
use crate::{
    CellAttribute, ConnectsChild, ParseLimits, Shape, Sheet, SheetChild, ThemeEffectColor,
    ThemeEffectStyle, ThemeEffects, ThemeOuterShadow, ThemeVariationScheme, VsdxError,
};
use ooxml_drawingml::Theme;

/// Identifies the ShapeSheet containing a semantic cell.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum CellSheet {
    Document,
    Page(u32),
    Master(u32),
}

/// Identifies a row within a ShapeSheet section.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum CellRow {
    Index(u32),
    Name(String),
}

/// Stable semantic identity for a ShapeSheet cell.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellLocator {
    pub sheet: CellSheet,
    pub shape_id: Option<u32>,
    pub section: Option<String>,
    pub section_index: Option<u32>,
    pub row: Option<CellRow>,
    pub cell_name: String,
}

/// A formula edit with a fresh cache; None removes the old cache.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SemanticCellEdit {
    pub locator: CellLocator,
    pub gesture: MutationGesture,
    pub formula: Option<String>,
    pub value: Option<String>,
    /// Row type for cells that create their container row.
    pub row_type: Option<String>,
}

/// Plain-text content of a shape's `Text` element.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SemanticTextEdit {
    pub locator: CellLocator,
    pub text: String,
}

/// The user action that requested a ShapeSheet mutation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MutationGesture {
    CellEdit,
    MoveX,
    MoveY,
    ResizeWidth,
    ResizeHeight,
    ResizeAspect,
    Rotate,
    TextEdit,
    Format,
    Delete,
}

/// A structural change applied atomically with any cell edits.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum StructuralEdit {
    /// Reorders the complete page catalog and its page relationships.
    ReorderPages {
        page_ids: Vec<u32>,
    },
    /// Inserts a complete Shape fragment using the next unused page-local ID.
    AddShape {
        page_id: u32,
        shape_xml: Vec<u8>,
    },
    DeleteShape {
        page_id: u32,
        shape_id: u32,
    },
    AddConnect {
        page_id: u32,
        from_sheet: u32,
        from_cell: String,
        to_sheet: u32,
        to_cell: String,
    },
    /// Moves a shape before `before_shape_id`, or to the end when it is absent.
    ReorderShape {
        page_id: u32,
        shape_id: u32,
        before_shape_id: Option<u32>,
    },
}

type PendingInsertion = (crate::SourceSpan, u8, Vec<(CellAttribute, String)>);

struct CellRemoval {
    part_path: String,
    cell_span: crate::SourceSpan,
    attribute: CellAttribute,
}

struct NewCell {
    part_path: String,
    owner_span: crate::SourceSpan,
    name: String,
    formula: String,
    value: Option<String>,
}

struct NewContainerCell {
    part_path: String,
    owner_span: crate::SourceSpan,
    section: Option<String>,
    section_index: Option<u32>,
    row: Option<CellRow>,
    row_type: Option<String>,
    name: String,
    formula: String,
    value: Option<String>,
}

pub fn parse_vsdx(data: &[u8]) -> Result<VsdxPackage, VsdxError> {
    parse_vsdx_with_limits(data, &ParseLimits::default())
}

pub fn parse_vsdx_with_limits(data: &[u8], limits: &ParseLimits) -> Result<VsdxPackage, VsdxError> {
    let source_parts = ooxml_opc::unzip_parts_with_limits(data, limits.max_expanded_bytes)
        .map_err(VsdxError::Container)?;
    match ooxml_opc::detect_package_kind(&source_parts) {
        Ok(ooxml_opc::DocumentKind::Vsdx) | Ok(ooxml_opc::DocumentKind::Vstx) => {}
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
    let mut theme_part_paths = targets_by_type(document_relationships, relationship_types::THEME);
    for path in targets_by_type(document_relationships, relationship_types::THEME_OOX) {
        if !theme_part_paths.contains(&path) {
            theme_part_paths.push(path);
        }
    }
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
    let catalog = xml_parts.get(pages_part_path.as_deref().unwrap_or(""));
    let mut page_positions = HashMap::new();
    let mut page_names = BTreeMap::new();
    for (position, page) in catalog
        .into_iter()
        .flat_map(|root| root.children_named("Page"))
        .enumerate()
    {
        let Some(id) = page.attribute("ID").and_then(|id| id.parse::<u32>().ok()) else {
            continue;
        };
        page_positions.insert(id, position);
        if let Some(name) = page.attribute("Name").or_else(|| page.attribute("NameU")) {
            page_names.insert(id, name.to_owned());
        }
    }
    page_part_paths.sort_by_key(|path| {
        page_part_ids
            .get(path)
            .and_then(|id| page_positions.get(id))
            .copied()
            .unwrap_or(usize::MAX)
    });
    let mut master_names = BTreeMap::new();
    for master in xml_parts
        .get(masters_part_path.as_deref().unwrap_or(""))
        .into_iter()
        .flat_map(|root| root.children_named("Master"))
    {
        let Some(id) = master.attribute("ID").and_then(|id| id.parse::<u32>().ok()) else {
            continue;
        };
        if let Some(name) = master
            .attribute("Name")
            .or_else(|| master.attribute("NameU"))
        {
            master_names.insert(id, name.to_owned());
        }
    }
    let page_contents = parse_part_sheets(&page_part_paths, &mut xml_parts, &mut budget)?;
    let master_contents = parse_part_sheets(&master_part_paths, &mut xml_parts, &mut budget)?;
    let mut themes: BTreeMap<u32, Theme> = BTreeMap::new();
    let mut theme_effects = BTreeMap::new();
    for (index, path) in theme_part_paths.iter().enumerate() {
        let root = xml_parts
            .get(path)
            .ok_or_else(|| VsdxError::MissingPart(path.clone()))?;
        let theme = parse_theme(root, path)?;
        let key = (index + 1) as u32;
        themes.insert(key, theme.clone());
        theme_effects.insert(key, parse_theme_effects(root));
        if let Some(id) = theme_scheme_enum(root)
            && !themes.contains_key(&id)
        {
            themes.insert(id, theme);
        }
    }
    let sheet_part_paths: HashSet<&str> = std::iter::once(document_path.as_str())
        .chain(pages_part_path.iter().map(String::as_str))
        .chain(masters_part_path.iter().map(String::as_str))
        .chain(page_part_paths.iter().map(String::as_str))
        .chain(master_part_paths.iter().map(String::as_str))
        .collect();
    let package_parts = source_parts
        .into_iter()
        .map(|(path, bytes)| {
            let spans = if sheet_part_paths.contains(path.as_str()) {
                scan_element_spans(&bytes).map_err(|_| VsdxError::MalformedXml {
                    part: path.clone(),
                    offset: 0,
                    message: "invalid lexical XML structure".to_owned(),
                })?
            } else {
                Vec::new()
            };
            Ok(PackagePart { path, bytes, spans })
        })
        .collect::<Result<Vec<_>, VsdxError>>()?;
    Ok(VsdxPackage {
        document_part_path: document_path,
        pages_part_path,
        masters_part_path,
        page_part_paths,
        master_part_paths,
        theme_part_paths,
        themes,
        theme_effects,
        windows_part_path,
        relationships,
        document_sheet,
        style_sheets,
        colors,
        face_names,
        page_sheets,
        master_sheets,
        page_part_ids,
        page_names,
        master_part_ids,
        master_names,
        page_contents,
        master_contents,
        parts: package_parts,
        source_container: Some(ooxml_opc::SourceContainer::new(data.to_vec())),
    })
}

fn parse_theme(root: &XmlElement, part: &str) -> Result<Theme, VsdxError> {
    let mut theme = Theme {
        name: root.attribute("name").unwrap_or("Office Theme").to_owned(),
        ..Theme::default()
    };
    let Some(theme_elements) = root.children_named("themeElements").next() else {
        return Ok(theme);
    };
    let scheme = theme_elements
        .children_named("clrScheme")
        .next()
        .ok_or_else(|| VsdxError::MalformedXml {
            part: part.to_owned(),
            offset: 0,
            message: "theme is missing themeElements/clrScheme".to_owned(),
        })?;
    for slot in [
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ] {
        let Some(value) = scheme.children_named(slot).next().and_then(|slot| {
            slot.children.iter().find_map(|child| match child {
                XmlNode::Element(value) => value
                    .attribute("lastClr")
                    .or_else(|| value.attribute("val")),
                XmlNode::Text(_) => None,
            })
        }) else {
            continue;
        };
        theme.color_scheme.set(slot, value.to_owned());
    }
    Ok(theme)
}

fn theme_scheme_enum(root: &XmlElement) -> Option<u32> {
    let elements = root.children_named("themeElements").next()?;
    let scheme = elements.children_named("clrScheme").next()?;
    let list = scheme.children_named("extLst").next()?;
    for ext in list.children_named("ext") {
        for id in ext.children_named("schemeID") {
            if let Some(value) = id
                .attribute("schemeEnum")
                .and_then(|value| value.parse().ok())
            {
                return Some(value);
            }
        }
    }
    None
}

fn parse_theme_effects(root: &XmlElement) -> ThemeEffects {
    let mut effects = ThemeEffects::default();
    let Some(elements) = root.children_named("themeElements").next() else {
        return effects;
    };
    for ext in elements
        .children_named("extLst")
        .flat_map(|list| list.children_named("ext"))
    {
        if let Some(scheme) = ext.children_named("themeScheme").next()
            && let Some(id) = scheme
                .children_named("schemeID")
                .find_map(|id| id.attribute("schemeEnum"))
                .and_then(|id| id.parse().ok())
        {
            effects.scheme_id = Some(id);
        }
        for scheme in ext.children_named("variationStyleSchemeLst") {
            for variant in scheme.children_named("variationStyleScheme") {
                effects.variation_schemes.push(ThemeVariationScheme {
                    embellishment: variant
                        .attribute("embellishment")
                        .and_then(|value| value.parse().ok())
                        .unwrap_or(0),
                    effect_indexes: variant
                        .children_named("varStyle")
                        .filter_map(|style| style.attribute("effectIdx"))
                        .filter_map(|index| index.parse().ok())
                        .collect(),
                });
            }
        }
        for schemes in ext.children_named("variationClrSchemeLst") {
            for scheme in schemes.children_named("variationClrScheme") {
                let mut colors = Vec::new();
                for position in 1..=7 {
                    let name = format!("varColor{position}");
                    colors.push(scheme.children_named(&name).next().and_then(|slot| {
                        slot.children.iter().find_map(|child| match child {
                            XmlNode::Element(value) => {
                                let hex = value
                                    .attribute("lastClr")
                                    .or_else(|| value.attribute("val"))?;
                                (value.local_name() == "srgbClr").then(|| hex.to_ascii_uppercase())
                            }
                            XmlNode::Text(_) => None,
                        })
                    }));
                }
                effects.variation_colors.push(colors);
            }
        }
    }
    if let Some(styles) = elements
        .children_named("fmtScheme")
        .next()
        .and_then(|scheme| scheme.children_named("effectStyleLst").next())
    {
        for style in styles.children_named("effectStyle") {
            let list = style.children_named("effectLst").next();
            effects.effect_styles.push(ThemeEffectStyle {
                outer_shadow: list.and_then(|list| {
                    list.children_named("outerShdw")
                        .next()
                        .map(parse_outer_shadow)
                }),
                has_bevel: style.children_named("sp3d").next().is_some(),
            });
        }
    }
    effects
}

fn parse_outer_shadow(style: &XmlElement) -> ThemeOuterShadow {
    let number = |name: &str| {
        style
            .attribute(name)
            .and_then(|value| value.parse().ok())
            .unwrap_or(0)
    };
    let mut shadow = ThemeOuterShadow {
        blur_emu: number("blurRad"),
        dist_emu: number("dist"),
        direction_60k: number("dir"),
        color: ThemeEffectColor::Placeholder,
        alpha_1000pct: None,
    };
    let Some(color) = style.children.iter().find_map(|child| match child {
        XmlNode::Element(value) if matches!(value.local_name(), "srgbClr" | "schemeClr") => {
            Some(value)
        }
        _ => None,
    }) else {
        return shadow;
    };
    shadow.color = match color.local_name() {
        "srgbClr" => color
            .attribute("val")
            .map(|hex| ThemeEffectColor::Srgb(hex.to_ascii_uppercase()))
            .unwrap_or_default(),
        _ => color
            .attribute("val")
            .map(|slot| ThemeEffectColor::Scheme(slot.to_owned()))
            .unwrap_or_default(),
    };
    shadow.alpha_1000pct = color
        .children_named("alpha")
        .find_map(|alpha| alpha.attribute("val"))
        .and_then(|value| value.parse().ok());
    shadow
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
            .children_named("Rel")
            .find_map(|rel| rel.attribute("r:id").or_else(|| rel.attribute("id")))
            .or_else(|| {
                element
                    .attribute("r:id")
                    .or_else(|| element.attribute("id"))
            })
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
    let parts = package
        .parts
        .iter()
        .map(|part| (part.path.clone(), part.bytes.clone()))
        .collect::<Vec<_>>();
    match &package.source_container {
        Some(source) => ooxml_opc::rezip_parts_preserving(&parts, source.as_bytes()),
        None => ooxml_opc::rezip_parts(&parts),
    }
    .map_err(VsdxError::Container)
}

/// Applies Cell@V and Cell@F attribute edits without mutating `package`.
#[cfg(test)]
pub(crate) fn save_cell_edits(
    package: &VsdxPackage,
    edits: &[CellEdit],
) -> Result<Vec<u8>, VsdxError> {
    save_cell_edits_with_new_cells(package, edits, &[], &[], &[])
}

struct ContainerCell {
    name: String,
    formula: String,
    value: Option<String>,
}

struct ContainerRow {
    row: Option<CellRow>,
    row_type: Option<String>,
    cells: Vec<ContainerCell>,
}

struct ContainerGroup {
    part_path: String,
    owner_span: crate::SourceSpan,
    section: Option<String>,
    section_index: Option<u32>,
    rows: Vec<ContainerRow>,
}

/// Merges container cells that share an insertion point into one section block.
/// Rows join one group only when the section itself is new; an existing section
/// keeps a group per row so each row resolves its own insertion point.
fn group_container_cells(cells: &[NewContainerCell]) -> Vec<ContainerGroup> {
    let mut groups: Vec<ContainerGroup> = Vec::new();
    for cell in cells {
        let group = match groups.iter_mut().find(|group| {
            group.part_path == cell.part_path
                && group.owner_span == cell.owner_span
                && group.section == cell.section
                && group.section_index == cell.section_index
                && (cell.section.is_some()
                    || group
                        .rows
                        .iter()
                        .all(|row| row.row == cell.row && row.row_type == cell.row_type))
        }) {
            Some(group) => group,
            None => {
                groups.push(ContainerGroup {
                    part_path: cell.part_path.clone(),
                    owner_span: cell.owner_span,
                    section: cell.section.clone(),
                    section_index: cell.section_index,
                    rows: Vec::new(),
                });
                groups.last_mut().expect("group was just pushed")
            }
        };
        match group
            .rows
            .iter_mut()
            .find(|row| row.row == cell.row && row.row_type == cell.row_type)
        {
            Some(row) => row.cells.push(ContainerCell {
                name: cell.name.clone(),
                formula: cell.formula.clone(),
                value: cell.value.clone(),
            }),
            None => group.rows.push(ContainerRow {
                row: cell.row.clone(),
                row_type: cell.row_type.clone(),
                cells: vec![ContainerCell {
                    name: cell.name.clone(),
                    formula: cell.formula.clone(),
                    value: cell.value.clone(),
                }],
            }),
        }
    }
    for group in &mut groups {
        for row in &mut group.rows {
            row.cells.sort_by(|left, right| left.name.cmp(&right.name));
        }
        group
            .rows
            .sort_by(|left, right| row_sort_key(&left.row).cmp(&row_sort_key(&right.row)));
    }
    groups
}

fn row_sort_key(row: &Option<CellRow>) -> (u8, u32, &str) {
    match row {
        Some(CellRow::Index(index)) => (0, *index, ""),
        Some(CellRow::Name(name)) => (1, 0, name),
        None => (2, 0, ""),
    }
}

fn save_cell_edits_with_new_cells(
    package: &VsdxPackage,
    edits: &[CellEdit],
    removals: &[CellRemoval],
    new_cells: &[NewCell],
    new_container_cells: &[NewContainerCell],
) -> Result<Vec<u8>, VsdxError> {
    if edits.len() + removals.len() > MAX_PATCH_EDITS {
        return Err(VsdxError::PatchLimit { kind: "editCount" });
    }
    let mut replacement_bytes = 0_usize;
    let mut validated = Vec::with_capacity(edits.len());
    let mut insertions: BTreeMap<(&str, crate::SourceSpan), PendingInsertion> = BTreeMap::new();
    for edit in edits {
        if !is_shapesheet_part(package, &edit.part_path) {
            return Err(VsdxError::InvalidCellEdit {
                part: edit.part_path.clone(),
                message: "part is not an authoritative ShapeSheet part".to_owned(),
            });
        }
        let part = package
            .parts
            .iter()
            .find(|part| part.path == edit.part_path)
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: edit.part_path.clone(),
                message: "part does not exist".to_owned(),
            })?;
        let cell = part
            .spans
            .iter()
            .find(|span| {
                span.name
                    .rsplit_once(':')
                    .map_or(span.name.as_str(), |(_, name)| name)
                    == "Cell"
                    && span.span == edit.cell_span
            })
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: edit.part_path.clone(),
                message: "span is not an existing Cell".to_owned(),
            })?;
        if let Some(attribute) = cell.attributes.get(edit.attribute.name()) {
            let replacement = escape_attribute_value(&edit.value, attribute.quote)?;
            replacement_bytes = replacement_bytes
                .checked_add(replacement.len())
                .ok_or(VsdxError::PatchLimit { kind: "editBytes" })?;
            validated.push((
                part.path.as_str(),
                SpanEdit {
                    span: attribute.value,
                    replacement,
                },
            ));
        } else {
            let quote = cell
                .attributes
                .values()
                .next()
                .map(|attribute| attribute.quote)
                .ok_or_else(|| VsdxError::InvalidCellEdit {
                    part: edit.part_path.clone(),
                    message: "Cell has no attribute quote style".to_owned(),
                })?;
            let entry = insertions
                .entry((part.path.as_str(), cell.start_tag))
                .or_insert_with(|| {
                    let end = cell.start_tag.end().expect("scanner span cannot overflow");
                    let offset = if part.bytes[end - 2] == b'/' {
                        end - 2
                    } else {
                        end - 1
                    };
                    (crate::SourceSpan { offset, length: 0 }, quote, Vec::new())
                });
            if entry
                .2
                .iter()
                .any(|(attribute, _)| *attribute == edit.attribute)
            {
                return Err(VsdxError::InvalidCellEdit {
                    part: edit.part_path.clone(),
                    message: format!("duplicate Cell@{} edit", edit.attribute.name()),
                });
            }
            entry.2.push((edit.attribute, edit.value.clone()));
        }
    }
    let mut removal_spans: BTreeMap<(&str, crate::SourceSpan), SpanEdit> = BTreeMap::new();
    for removal in removals {
        if !is_shapesheet_part(package, &removal.part_path) {
            return Err(VsdxError::InvalidCellEdit {
                part: removal.part_path.clone(),
                message: "part is not an authoritative ShapeSheet part".to_owned(),
            });
        }
        let part = package
            .parts
            .iter()
            .find(|part| part.path == removal.part_path)
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: removal.part_path.clone(),
                message: "part does not exist".to_owned(),
            })?;
        let cell = part
            .spans
            .iter()
            .find(|span| {
                span.name
                    .rsplit_once(':')
                    .map_or(span.name.as_str(), |(_, name)| name)
                    == "Cell"
                    && span.span == removal.cell_span
            })
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: removal.part_path.clone(),
                message: "span is not an existing Cell".to_owned(),
            })?;
        let Some(attribute) = cell.attributes.get(removal.attribute.name()) else {
            continue;
        };
        let mut start = attribute.name.offset;
        while start > 0 && part.bytes[start - 1].is_ascii_whitespace() {
            start -= 1;
        }
        let end = attribute
            .value
            .end()
            .and_then(|end| end.checked_add(1))
            .ok_or(VsdxError::InvalidSpan)?;
        let length = end.checked_sub(start).ok_or(VsdxError::InvalidSpan)?;
        removal_spans.insert(
            (part.path.as_str(), cell.start_tag),
            SpanEdit {
                span: crate::SourceSpan {
                    offset: start,
                    length,
                },
                replacement: Vec::new(),
            },
        );
    }
    for ((path, start_tag), (span, quote, attributes)) in insertions {
        let mut replacement = Vec::new();
        for (attribute, value) in attributes {
            replacement.push(b' ');
            replacement.extend_from_slice(attribute.name().as_bytes());
            replacement.push(b'=');
            replacement.push(quote);
            replacement.extend_from_slice(&escape_attribute_value(&value, quote)?);
            replacement.push(quote);
        }
        replacement_bytes = replacement_bytes
            .checked_add(replacement.len())
            .ok_or(VsdxError::PatchLimit { kind: "editBytes" })?;
        if let Some(removal) = removal_spans.get_mut(&(path, start_tag))
            && removal.span.end() == Some(span.offset)
        {
            removal.replacement.extend(replacement);
            continue;
        }
        validated.push((path, SpanEdit { span, replacement }));
    }
    for ((path, _), removal) in removal_spans {
        validated.push((path, removal));
    }
    for new_cell in new_cells {
        let part = package
            .parts
            .iter()
            .find(|part| part.path == new_cell.part_path)
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: new_cell.part_path.clone(),
                message: "part does not exist".to_owned(),
            })?;
        let owner = part
            .spans
            .iter()
            .find(|span| span.span == new_cell.owner_span)
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: new_cell.part_path.clone(),
                message: "local cell owner does not exist".to_owned(),
            })?;
        let quote = owner
            .attributes
            .values()
            .next()
            .map(|attribute| attribute.quote)
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: new_cell.part_path.clone(),
                message: "local cell owner has no attribute quote style".to_owned(),
            })?;
        let end = owner.span.end().ok_or(VsdxError::InvalidSpan)?;
        let (span, closes_owner) = if end >= 2 && part.bytes[end - 2..end] == *b"/>" {
            (
                crate::SourceSpan {
                    offset: end - 2,
                    length: 2,
                },
                true,
            )
        } else {
            (
                crate::SourceSpan {
                    offset: part.bytes[..end]
                        .iter()
                        .rposition(|byte| *byte == b'<')
                        .ok_or(VsdxError::InvalidSpan)?,
                    length: 0,
                },
                false,
            )
        };
        let mut replacement = b"<Cell N=".to_vec();
        if closes_owner {
            replacement.insert(0, b'>');
        }
        replacement.push(quote);
        replacement.extend_from_slice(&escape_attribute_value(&new_cell.name, quote)?);
        replacement.push(quote);
        let mut attributes: Vec<(&str, &str)> = vec![("F", &new_cell.formula)];
        if let Some(value) = &new_cell.value {
            attributes.push(("V", value));
        }
        for (name, value) in attributes {
            replacement.push(b' ');
            replacement.extend_from_slice(name.as_bytes());
            replacement.push(b'=');
            replacement.push(quote);
            replacement.extend_from_slice(&escape_attribute_value(value, quote)?);
            replacement.push(quote);
        }
        replacement.extend_from_slice(b"/>");
        if closes_owner {
            replacement.extend_from_slice(format!("</{}>", local_name(&owner.name)).as_bytes());
        }
        replacement_bytes = replacement_bytes
            .checked_add(replacement.len())
            .ok_or(VsdxError::PatchLimit { kind: "editBytes" })?;
        validated.push((part.path.as_str(), SpanEdit { span, replacement }));
    }
    for group in group_container_cells(new_container_cells) {
        let part = package
            .parts
            .iter()
            .find(|part| part.path == group.part_path)
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: group.part_path.clone(),
                message: "part does not exist".to_owned(),
            })?;
        let owner = part
            .spans
            .iter()
            .find(|span| span.span == group.owner_span)
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: group.part_path.clone(),
                message: "local container owner does not exist".to_owned(),
            })?;
        let quote = owner
            .attributes
            .values()
            .next()
            .map(|attribute| attribute.quote)
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: group.part_path.clone(),
                message: "local container owner has no attribute quote style".to_owned(),
            })?;
        let (span, closes_owner) = container_insertion_point(
            part,
            owner,
            group.section.as_deref(),
            group.rows.first().and_then(|row| row.row.as_ref()),
        )?;
        let mut replacement = Vec::new();
        if closes_owner {
            replacement.push(b'>');
        }
        if let Some(section) = &group.section {
            replacement.extend_from_slice(b"<Section N=");
            push_quoted(&mut replacement, section, quote)?;
            if let Some(index) = group.section_index {
                replacement.extend_from_slice(b" IX=");
                push_quoted(&mut replacement, &index.to_string(), quote)?;
            }
            replacement.push(b'>');
        }
        for row in &group.rows {
            if let Some(row_key) = &row.row {
                replacement.extend_from_slice(b"<Row ");
                match row_key {
                    CellRow::Index(_) => replacement.extend_from_slice(b"IX="),
                    CellRow::Name(_) => replacement.extend_from_slice(b"N="),
                }
                let row_value = match row_key {
                    CellRow::Index(index) => index.to_string(),
                    CellRow::Name(name) => name.clone(),
                };
                push_quoted(&mut replacement, &row_value, quote)?;
                if let Some(row_type) = &row.row_type {
                    replacement.extend_from_slice(b" T=");
                    push_quoted(&mut replacement, row_type, quote)?;
                }
                replacement.push(b'>');
            }
            for cell in &row.cells {
                replacement.extend_from_slice(b"<Cell N=");
                push_quoted(&mut replacement, &cell.name, quote)?;
                replacement.extend_from_slice(b" F=");
                push_quoted(&mut replacement, &cell.formula, quote)?;
                if let Some(value) = &cell.value {
                    replacement.extend_from_slice(b" V=");
                    push_quoted(&mut replacement, value, quote)?;
                }
                replacement.extend_from_slice(b"/>");
            }
            if row.row.is_some() {
                replacement.extend_from_slice(b"</Row>");
            }
        }
        if group.section.is_some() {
            replacement.extend_from_slice(b"</Section>");
        }
        if closes_owner {
            replacement.extend_from_slice(format!("</{}>", local_name(&owner.name)).as_bytes());
        }
        replacement_bytes = replacement_bytes
            .checked_add(replacement.len())
            .ok_or(VsdxError::PatchLimit { kind: "editBytes" })?;
        validated.push((part.path.as_str(), SpanEdit { span, replacement }));
    }
    if replacement_bytes > MAX_PATCH_BYTES {
        return Err(VsdxError::PatchLimit { kind: "editBytes" });
    }
    let mut part_edits: BTreeMap<&str, Vec<SpanEdit>> = BTreeMap::new();
    for (path, edit) in validated {
        part_edits.entry(path).or_default().push(edit);
    }
    let mut output = package.clone();
    for (path, mut edits) in part_edits {
        edits.sort_by_key(|edit| edit.span.offset);
        let mut merged: Vec<SpanEdit> = Vec::with_capacity(edits.len());
        for edit in edits {
            if let Some(previous) = merged.last_mut()
                && previous.span.offset == edit.span.offset
                && previous.span.length == 0
                && edit.span.length == 0
            {
                previous.replacement.extend(edit.replacement);
                continue;
            }
            merged.push(edit);
        }
        let part = output
            .parts
            .iter_mut()
            .find(|part| part.path == path)
            .expect("validated source part");
        part.bytes = apply_span_edits(&part.bytes, &merged)?;
    }
    let bytes = write_vsdx(&output)?;
    parse_vsdx(&bytes)?;
    Ok(bytes)
}

/// Saves semantic formula edits, replacing or removing their cached values.
pub fn save_semantic_cell_edits(
    package: &VsdxPackage,
    edits: &[SemanticCellEdit],
) -> Result<Vec<u8>, VsdxError> {
    let mut lexical = Vec::new();
    let mut removals = Vec::new();
    let mut new_cells = Vec::new();
    let mut new_container_cells = Vec::new();
    for edit in edits {
        let Some(formula) = &edit.formula else {
            return Err(VsdxError::InvalidCellEdit {
                part: format!("{:?}", edit.locator.sheet),
                message: "semantic edits require a formula".to_owned(),
            });
        };
        match resolve_cell_locator(package, &edit.locator)? {
            LocalCell::Existing(part_path, cell_span) => {
                lexical.push(CellEdit {
                    part_path: part_path.clone(),
                    cell_span,
                    attribute: CellAttribute::Formula,
                    value: formula.clone(),
                });
                match &edit.value {
                    Some(value) => lexical.push(CellEdit {
                        part_path,
                        cell_span,
                        attribute: CellAttribute::Value,
                        value: value.clone(),
                    }),
                    None => removals.push(CellRemoval {
                        part_path,
                        cell_span,
                        attribute: CellAttribute::Value,
                    }),
                }
            }
            LocalCell::New(part_path, owner_span) => new_cells.push(NewCell {
                part_path,
                owner_span,
                name: edit.locator.cell_name.clone(),
                formula: formula.clone(),
                value: edit.value.clone(),
            }),
            LocalCell::NewContainer(part_path, owner_span, section, row) => new_container_cells
                .push(NewContainerCell {
                    part_path,
                    owner_span,
                    section,
                    section_index: edit.locator.section_index,
                    row,
                    row_type: edit.row_type.clone(),
                    name: edit.locator.cell_name.clone(),
                    formula: formula.clone(),
                    value: edit.value.clone(),
                }),
        }
    }
    if new_cells.is_empty() && new_container_cells.is_empty() {
        save_cell_edits_with_new_cells(package, &lexical, &removals, &[], &[])
    } else {
        save_cell_edits_with_new_cells(
            package,
            &lexical,
            &removals,
            &new_cells,
            &new_container_cells,
        )
    }
}

/// Patches only the inner XML of each edited shape's `Text` element.
pub fn save_semantic_text_edits(
    package: &VsdxPackage,
    edits: &[SemanticTextEdit],
) -> Result<Vec<u8>, VsdxError> {
    let limits = ParseLimits::default();
    let mut part_edits: BTreeMap<String, Vec<crate::SpanEdit>> = BTreeMap::new();
    for edit in edits {
        validate_text_content(&edit.text, &limits)?;
        let (page_id, shape_id) = match (&edit.locator.sheet, edit.locator.shape_id) {
            (CellSheet::Page(page_id), Some(shape_id)) => (*page_id, shape_id),
            _ => {
                return Err(VsdxError::InvalidCellEdit {
                    part: format!("{:?}", edit.locator.sheet),
                    message: "text edits require a page shape".to_owned(),
                });
            }
        };
        let path = package
            .page_part_ids
            .iter()
            .find_map(|(path, id)| (*id == page_id).then(|| path.clone()))
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: page_id.to_string(),
                message: "page does not exist".to_owned(),
            })?;
        let part = package
            .parts
            .iter()
            .find(|part| part.path == path)
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: path.clone(),
                message: "part does not exist".to_owned(),
            })?;
        let shape = part
            .spans
            .iter()
            .find(|span| {
                local_name(&span.name) == "Shape"
                    && attribute_equals(&part.bytes, span, "ID", &shape_id.to_string())
            })
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: path.clone(),
                message: format!("shape {shape_id} does not exist"),
            })?;
        let text = part.spans.iter().find(|span| {
            local_name(&span.name) == "Text"
                && nearest_parent(part, span.span, "Shape")
                    .is_some_and(|parent| parent.span == shape.span)
        });
        let escaped = escape_text_content(&edit.text);
        match text {
            Some(text) => {
                let outer_end = text.span.end().ok_or(VsdxError::InvalidSpan)?;
                let open_end = tag_close(&part.bytes, text.span.offset, outer_end)
                    .ok_or(VsdxError::InvalidSpan)?;
                if part.bytes[open_end.saturating_sub(1)] == b'/' {
                    part_edits.entry(path).or_default().push(crate::SpanEdit {
                        span: text.span,
                        replacement: format!("<{}>{escaped}</{}>", text.name, text.name)
                            .into_bytes(),
                    });
                } else {
                    let inner_start = open_end + 1;
                    let closing = part.bytes[inner_start..outer_end]
                        .iter()
                        .rposition(|byte| *byte == b'<')
                        .map(|relative| relative + inner_start)
                        .ok_or(VsdxError::InvalidSpan)?;
                    part_edits.entry(path).or_default().push(crate::SpanEdit {
                        span: crate::SourceSpan {
                            offset: inner_start,
                            length: closing - inner_start,
                        },
                        replacement: escaped.into_bytes(),
                    });
                }
            }
            None => {
                let outer_end = shape.span.end().ok_or(VsdxError::InvalidSpan)?;
                if outer_end >= 2 && part.bytes[outer_end - 2..outer_end] == *b"/>" {
                    part_edits.entry(path).or_default().push(crate::SpanEdit {
                        span: crate::SourceSpan {
                            offset: outer_end - 2,
                            length: 2,
                        },
                        replacement: format!("><Text>{escaped}</Text></{}>", shape.name)
                            .into_bytes(),
                    });
                } else {
                    let closing = part.bytes[..outer_end]
                        .iter()
                        .rposition(|byte| *byte == b'<')
                        .ok_or(VsdxError::InvalidSpan)?;
                    part_edits.entry(path).or_default().push(crate::SpanEdit {
                        span: crate::SourceSpan {
                            offset: closing,
                            length: 0,
                        },
                        replacement: format!("<Text>{escaped}</Text>").into_bytes(),
                    });
                }
            }
        }
    }
    if part_edits.values().all(|edits| edits.is_empty()) {
        return write_vsdx(package);
    }
    let mut output = package.clone();
    for (path, edits) in &part_edits {
        let part = package
            .parts
            .iter()
            .find(|part| &part.path == path)
            .expect("validated source part");
        let bytes = apply_span_edits(&part.bytes, edits)?;
        let target = output
            .parts
            .iter_mut()
            .find(|part| &part.path == path)
            .expect("validated source part");
        target.bytes = bytes;
    }
    let bytes = write_vsdx(&output)?;
    parse_vsdx(&bytes)?;
    Ok(bytes)
}

fn validate_text_content(value: &str, limits: &ParseLimits) -> Result<(), VsdxError> {
    if value.len() > limits.max_xml_text_bytes {
        return Err(VsdxError::PatchLimit { kind: "editBytes" });
    }
    if value.chars().any(|c| !matches!(c, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')) {
        return Err(VsdxError::InvalidXmlCharacter);
    }
    Ok(())
}

fn escape_text_content(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(character),
        }
    }
    output
}

/// Applies structural edits and validates references; deletion removes incident Connects.
pub fn save_structural_edits(
    package: &VsdxPackage,
    edits: &[StructuralEdit],
) -> Result<Vec<u8>, VsdxError> {
    let page_reorders = edits
        .iter()
        .filter_map(|edit| match edit {
            StructuralEdit::ReorderPages { page_ids } => Some(page_ids.as_slice()),
            _ => None,
        })
        .collect::<Vec<_>>();
    if page_reorders.len() > 1 {
        return Err(VsdxError::InvalidCellEdit {
            part: "pages".to_owned(),
            message: "multiple page reorder edits are not supported".to_owned(),
        });
    }
    let mut by_part: BTreeMap<String, Vec<&StructuralEdit>> = BTreeMap::new();
    for edit in edits {
        let page_id = match edit {
            StructuralEdit::AddShape { page_id, .. }
            | StructuralEdit::DeleteShape { page_id, .. }
            | StructuralEdit::ReorderShape { page_id, .. }
            | StructuralEdit::AddConnect { page_id, .. } => page_id,
            StructuralEdit::ReorderPages { .. } => continue,
        };
        let path = package
            .page_part_ids
            .iter()
            .find_map(|(path, id)| (*id == *page_id).then(|| path.clone()))
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: page_id.to_string(),
                message: "page does not exist".to_owned(),
            })?;
        by_part.entry(path).or_default().push(edit);
    }
    let mut output = package.clone();
    for (path, part_edits) in by_part {
        let part = package
            .parts
            .iter()
            .find(|part| part.path == path)
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: path.clone(),
                message: "part does not exist".to_owned(),
            })?;
        page_shapes(part).ok_or_else(|| VsdxError::InvalidCellEdit {
            part: path.clone(),
            message: "page has no Shapes container".to_owned(),
        })?;
        let mut bytes = part.bytes.clone();
        for edit in part_edits {
            let current = PackagePart {
                path: path.clone(),
                bytes: bytes.clone(),
                spans: scan_element_spans(&bytes)?,
            };
            let shapes = page_shapes(&current).ok_or_else(|| VsdxError::InvalidCellEdit {
                part: path.clone(),
                message: "page has no Shapes container after structural edit".to_owned(),
            })?;
            let replacement = match edit {
                StructuralEdit::AddShape { shape_xml, .. } => {
                    add_shape(&current, shapes, shape_xml)?
                }
                StructuralEdit::DeleteShape { shape_id, .. } => {
                    delete_shape(&current, shapes, *shape_id)?
                }
                StructuralEdit::ReorderShape {
                    shape_id,
                    before_shape_id,
                    ..
                } => reorder_shape(&current, shapes, *shape_id, *before_shape_id)?,
                StructuralEdit::AddConnect {
                    from_sheet,
                    from_cell,
                    to_sheet,
                    to_cell,
                    ..
                } => add_connect(&current, *from_sheet, from_cell, *to_sheet, to_cell)?,
                StructuralEdit::ReorderPages { .. } => unreachable!(),
            };
            bytes = apply_span_edits(&bytes, &replacement)?;
        }
        let target = output
            .parts
            .iter_mut()
            .find(|part| part.path == path)
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: path.clone(),
                message: "output part does not exist".to_owned(),
            })?;
        target.bytes = bytes;
    }
    if let Some(page_ids) = page_reorders.first() {
        reorder_pages(&mut output, package, page_ids)?;
    }
    let bytes = write_vsdx(&output)?;
    let reparsed = parse_vsdx(&bytes)?;
    validate_structure(&reparsed)?;
    Ok(bytes)
}

fn reorder_pages(
    output: &mut VsdxPackage,
    package: &VsdxPackage,
    page_ids: &[u32],
) -> Result<(), VsdxError> {
    let expected = package
        .page_part_paths
        .iter()
        .map(|path| package.page_part_ids.get(path).copied())
        .collect::<Option<Vec<_>>>()
        .ok_or_else(|| VsdxError::InvalidCellEdit {
            part: "pages".to_owned(),
            message: "page catalog is missing a page ID".to_owned(),
        })?;
    let expected_ids = expected.iter().copied().collect::<HashSet<_>>();
    if page_ids.len() != expected.len()
        || page_ids.iter().copied().collect::<HashSet<_>>() != expected_ids
    {
        return Err(VsdxError::InvalidCellEdit {
            part: "pages".to_owned(),
            message: "page reorder must contain every page exactly once".to_owned(),
        });
    }
    let pages_path =
        package
            .pages_part_path
            .as_deref()
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: "pages".to_owned(),
                message: "package has no page catalog".to_owned(),
            })?;
    let catalog = package
        .parts
        .iter()
        .find(|part| part.path == pages_path)
        .ok_or_else(|| VsdxError::InvalidCellEdit {
            part: pages_path.to_owned(),
            message: "page catalog part does not exist".to_owned(),
        })?;
    let pages = catalog
        .spans
        .iter()
        .find(|span| local_name(&span.name) == "Pages")
        .ok_or_else(|| VsdxError::InvalidCellEdit {
            part: pages_path.to_owned(),
            message: "page catalog has no Pages container".to_owned(),
        })?;
    let page_entries = catalog
        .spans
        .iter()
        .filter(|span| {
            local_name(&span.name) == "Page"
                && immediate_parent(catalog, span.span)
                    .is_some_and(|parent| parent.span == pages.span)
        })
        .collect::<Vec<_>>();
    let mut ordered_pages = Vec::with_capacity(page_ids.len());
    let mut relationship_ids = Vec::with_capacity(page_ids.len());
    for id in page_ids {
        let page = page_entries
            .iter()
            .copied()
            .find(|page| attribute_equals(&catalog.bytes, page, "ID", &id.to_string()))
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: pages_path.to_owned(),
                message: format!("page {id} does not exist in the catalog"),
            })?;
        let relationship_id = catalog
            .spans
            .iter()
            .find(|span| {
                local_name(&span.name) == "Rel"
                    && immediate_parent(catalog, span.span)
                        .is_some_and(|parent| parent.span == page.span)
            })
            .and_then(|rel| {
                rel.attributes
                    .get("r:id")
                    .or_else(|| rel.attributes.get("id"))
                    .and_then(|attribute| attribute_value(&catalog.bytes, attribute))
            })
            .or_else(|| {
                page.attributes
                    .get("r:id")
                    .or_else(|| page.attributes.get("id"))
                    .and_then(|attribute| attribute_value(&catalog.bytes, attribute))
            })
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: pages_path.to_owned(),
                message: format!("page {id} has no relationship"),
            })?
            .into_owned();
        ordered_pages.push(page);
        relationship_ids.push(relationship_id);
    }
    let mut catalog_output = catalog.clone();
    catalog_output.bytes = apply_span_edits(
        &catalog.bytes,
        &[reorder_selected_children(
            catalog,
            pages,
            &page_entries,
            &ordered_pages,
        )?],
    )?;
    replace_part(output, catalog_output)?;

    let rels_path = relationship_path(pages_path);
    let mut rels = package
        .parts
        .iter()
        .find(|part| part.path == rels_path)
        .cloned()
        .ok_or_else(|| VsdxError::InvalidCellEdit {
            part: rels_path.clone(),
            message: "page relationships part does not exist".to_owned(),
        })?;
    rels.spans = scan_element_spans(&rels.bytes)?;
    let relationships = rels
        .spans
        .iter()
        .find(|span| local_name(&span.name) == "Relationships")
        .ok_or_else(|| VsdxError::InvalidCellEdit {
            part: rels_path.clone(),
            message: "page relationships have no container".to_owned(),
        })?;
    let entries = rels
        .spans
        .iter()
        .filter(|span| {
            local_name(&span.name) == "Relationship"
                && immediate_parent(&rels, span.span)
                    .is_some_and(|parent| parent.span == relationships.span)
        })
        .collect::<Vec<_>>();
    let page_entries = entries
        .iter()
        .copied()
        .filter(|entry| {
            entry
                .attributes
                .get("Id")
                .and_then(|attribute| attribute_value(&rels.bytes, attribute))
                .is_some_and(|id| relationship_ids.iter().any(|expected| expected == &id))
        })
        .collect::<Vec<_>>();
    let ordered_relationships = relationship_ids
        .iter()
        .map(|id| {
            page_entries
                .iter()
                .copied()
                .find(|entry| attribute_equals(&rels.bytes, entry, "Id", id))
                .ok_or_else(|| VsdxError::InvalidCellEdit {
                    part: rels_path.clone(),
                    message: format!("page relationship {id} does not exist"),
                })
        })
        .collect::<Result<Vec<_>, _>>()?;
    rels.bytes = apply_span_edits(
        &rels.bytes,
        &[reorder_selected_children(
            &rels,
            relationships,
            &page_entries,
            &ordered_relationships,
        )?],
    )?;
    replace_part(output, rels)
}

fn reorder_selected_children(
    part: &PackagePart,
    container: &crate::ElementSpan,
    children: &[&crate::ElementSpan],
    ordered: &[&crate::ElementSpan],
) -> Result<SpanEdit, VsdxError> {
    if children.len() != ordered.len() {
        return Err(VsdxError::InvalidSpan);
    }
    let mut replacement = Vec::new();
    let mut cursor = container.span.offset;
    for (child, desired) in children.iter().zip(ordered) {
        replacement.extend_from_slice(&part.bytes[cursor..child.span.offset]);
        replacement.extend_from_slice(
            &part.bytes[desired.span.offset..desired.span.end().ok_or(VsdxError::InvalidSpan)?],
        );
        cursor = child.span.end().ok_or(VsdxError::InvalidSpan)?;
    }
    replacement.extend_from_slice(
        &part.bytes[cursor..container.span.end().ok_or(VsdxError::InvalidSpan)?],
    );
    Ok(SpanEdit {
        span: container.span,
        replacement,
    })
}

fn replace_part(output: &mut VsdxPackage, replacement: PackagePart) -> Result<(), VsdxError> {
    let part = output
        .parts
        .iter_mut()
        .find(|part| part.path == replacement.path)
        .ok_or_else(|| VsdxError::InvalidCellEdit {
            part: replacement.path.clone(),
            message: "output part does not exist".to_owned(),
        })?;
    *part = replacement;
    Ok(())
}

fn shape_children<'a>(
    part: &'a PackagePart,
    shapes: &crate::ElementSpan,
) -> Vec<&'a crate::ElementSpan> {
    let mut children: Vec<_> = part
        .spans
        .iter()
        .filter(|candidate| {
            local_name(&candidate.name) == "Shape"
                && nearest_parent(part, candidate.span, &shapes.name)
                    .is_some_and(|parent| parent.span == shapes.span)
        })
        .collect();
    children.sort_by_key(|shape| shape.span.offset);
    children
}

fn shape_id(part: &PackagePart, shape: &crate::ElementSpan) -> Option<u32> {
    shape
        .attributes
        .get("ID")
        .and_then(|attribute| attribute_value(&part.bytes, attribute))
        .and_then(|value| value.parse().ok())
}

fn add_shape(
    part: &PackagePart,
    shapes: &crate::ElementSpan,
    fragment: &[u8],
) -> Result<Vec<SpanEdit>, VsdxError> {
    let fragment_spans = scan_element_spans(fragment)?;
    let root = fragment_spans
        .iter()
        .find(|span| span.span.offset == 0 && local_name(&span.name) == "Shape")
        .ok_or_else(|| VsdxError::InvalidCellEdit {
            part: part.path.clone(),
            message: "new shape must be one complete Shape element".to_owned(),
        })?;
    if root.span.end() != Some(fragment.len()) {
        return Err(VsdxError::InvalidCellEdit {
            part: part.path.clone(),
            message: "new shape fragment has trailing bytes".to_owned(),
        });
    }
    if fragment_spans
        .iter()
        .any(|span| local_name(&span.name) == "Rel")
    {
        return Err(VsdxError::InvalidCellEdit {
            part: part.path.clone(),
            message: "new shape fragments cannot contain relationship references".to_owned(),
        });
    }
    let requested_shape_id = root
        .attributes
        .get("ID")
        .and_then(|attribute| attribute_value(fragment, attribute))
        .and_then(|id| id.parse().ok());
    let new_shape_id = requested_shape_id.unwrap_or(
        all_shape_spans(part, shapes)
            .iter()
            .filter_map(|shape| shape_id(part, shape))
            .max()
            .map_or(Ok(1), |id| {
                id.checked_add(1).ok_or_else(|| VsdxError::InvalidCellEdit {
                    part: part.path.clone(),
                    message: "shape ID space is exhausted".to_owned(),
                })
            })?,
    );
    if all_shape_spans(part, shapes)
        .iter()
        .filter_map(|shape| shape_id(part, shape))
        .any(|shape_id| shape_id == new_shape_id)
    {
        return Err(VsdxError::InvalidCellEdit {
            part: part.path.clone(),
            message: "shape ID is already in use".to_owned(),
        });
    }
    let mut new_shape = fragment.to_vec();
    let id = new_shape_id.to_string().into_bytes();
    if let Some(attribute) = root.attributes.get("ID") {
        new_shape = apply_span_edits(
            &new_shape,
            &[SpanEdit {
                span: attribute.value,
                replacement: id,
            }],
        )?;
    } else {
        let end = fragment[root.span.offset..root.span.end().ok_or(VsdxError::InvalidSpan)?]
            .iter()
            .position(|byte| *byte == b'>')
            .map(|offset| root.span.offset + offset + 1)
            .ok_or(VsdxError::InvalidSpan)?;
        let offset = if new_shape.get(end - 2) == Some(&b'/') {
            end - 2
        } else {
            end - 1
        };
        new_shape = apply_span_edits(
            &new_shape,
            &[SpanEdit {
                span: crate::SourceSpan { offset, length: 0 },
                replacement: format!(" ID='{new_shape_id}'").into_bytes(),
            }],
        )?;
    }
    let end = shapes.span.end().ok_or(VsdxError::InvalidSpan)?;
    let (span, replacement) = if part.bytes.get(end - 2..end) == Some(b"/>") {
        (
            crate::SourceSpan {
                offset: end - 2,
                length: 2,
            },
            [
                b">".as_slice(),
                new_shape.as_slice(),
                format!("</{}>", shapes.name).as_bytes(),
            ]
            .concat(),
        )
    } else {
        let close = part.bytes[..end]
            .iter()
            .rposition(|byte| *byte == b'<')
            .ok_or(VsdxError::InvalidSpan)?;
        (
            crate::SourceSpan {
                offset: close,
                length: 0,
            },
            new_shape,
        )
    };
    Ok(vec![SpanEdit { span, replacement }])
}

fn delete_shape(
    part: &PackagePart,
    shapes: &crate::ElementSpan,
    shape_id_to_delete: u32,
) -> Result<Vec<SpanEdit>, VsdxError> {
    let shapes = shape_container(part, shapes, shape_id_to_delete).ok_or_else(|| {
        VsdxError::InvalidCellEdit {
            part: part.path.clone(),
            message: format!("shape {shape_id_to_delete} does not exist"),
        }
    })?;
    let children = shape_children(part, shapes);
    let deleted_shape = *children
        .iter()
        .find(|shape| shape_id(part, shape) == Some(shape_id_to_delete))
        .ok_or_else(|| VsdxError::InvalidCellEdit {
            part: part.path.clone(),
            message: format!("shape {shape_id_to_delete} does not exist"),
        })?;
    let deleted: HashSet<_> = std::iter::once(deleted_shape)
        .chain(all_shape_spans(part, deleted_shape))
        .filter_map(|shape| shape_id(part, shape))
        .collect();
    let mut replacements = vec![container_without(part, shapes, "Shape", |shape| {
        shape_id(part, shape).is_some_and(|id| deleted.contains(&id))
    })?];
    if let Some(connects) = direct_child(part, "Connects", None) {
        replacements.push(container_without(part, connects, "Connect", |connect| {
            references_deleted_shape(
                ["FromSheet", "ToSheet"].iter().filter_map(|name| {
                    connect
                        .attributes
                        .get(*name)
                        .and_then(|attribute| attribute_value(&part.bytes, attribute))
                        .and_then(|value| value.parse::<u32>().ok())
                }),
                &deleted,
            )
        })?);
    }
    Ok(replacements)
}

fn add_connect(
    part: &PackagePart,
    from_sheet: u32,
    from_cell: &str,
    to_sheet: u32,
    to_cell: &str,
) -> Result<Vec<SpanEdit>, VsdxError> {
    if !matches!(from_cell, "BeginX" | "BeginY" | "EndX" | "EndY") {
        return Err(VsdxError::InvalidCellEdit {
            part: part.path.clone(),
            message: "Connect FromCell must be a connector endpoint".to_owned(),
        });
    }
    if !valid_glue_target(to_cell) {
        return Err(VsdxError::InvalidCellEdit {
            part: part.path.clone(),
            message: "Connect ToCell must be PinX, PinY, or Connections.XN".to_owned(),
        });
    }
    let connects = direct_child(part, "Connects", None);
    let contents = part
        .spans
        .iter()
        .find(|span| local_name(&span.name) == "PageContents")
        .ok_or_else(|| VsdxError::InvalidCellEdit {
            part: part.path.clone(),
            message: "page has no PageContents container".to_owned(),
        })?;
    let owner = connects.unwrap_or(contents);
    let prefix = owner.name.rsplit_once(':').map_or("", |(prefix, _)| prefix);
    let qualify = |name: &str| {
        if prefix.is_empty() {
            name.to_owned()
        } else {
            format!("{prefix}:{name}")
        }
    };
    let mut fragment = format!("<{} FromSheet=\"", qualify("Connect")).into_bytes();
    fragment.extend_from_slice(from_sheet.to_string().as_bytes());
    fragment.extend_from_slice(b"\" FromCell=\"");
    push_quoted_fragment(&mut fragment, from_cell)?;
    fragment.extend_from_slice(b"\" ToSheet=\"");
    fragment.extend_from_slice(to_sheet.to_string().as_bytes());
    fragment.extend_from_slice(b"\" ToCell=\"");
    push_quoted_fragment(&mut fragment, to_cell)?;
    fragment.extend_from_slice(b"\"/>");
    if let Some(connects) = connects {
        let end = connects.span.end().ok_or(VsdxError::InvalidSpan)?;
        if end >= 2 && part.bytes[end - 2..end] == *b"/>" {
            let mut replacement = Vec::from(b">".as_slice());
            replacement.extend_from_slice(&fragment);
            replacement.extend_from_slice(format!("</{}>", connects.name).as_bytes());
            return Ok(vec![SpanEdit {
                span: crate::SourceSpan {
                    offset: end - 2,
                    length: 2,
                },
                replacement,
            }]);
        }
        let close = part.bytes[..end]
            .iter()
            .rposition(|byte| *byte == b'<')
            .ok_or(VsdxError::InvalidSpan)?;
        return Ok(vec![SpanEdit {
            span: crate::SourceSpan {
                offset: close,
                length: 0,
            },
            replacement: fragment,
        }]);
    }
    let end = contents.span.end().ok_or(VsdxError::InvalidSpan)?;
    let close = part.bytes[..end]
        .iter()
        .rposition(|byte| *byte == b'<')
        .ok_or(VsdxError::InvalidSpan)?;
    let container = qualify("Connects");
    let mut replacement = format!("<{container}>").into_bytes();
    replacement.extend_from_slice(&fragment);
    replacement.extend_from_slice(format!("</{container}>").as_bytes());
    Ok(vec![SpanEdit {
        span: crate::SourceSpan {
            offset: close,
            length: 0,
        },
        replacement,
    }])
}

fn valid_glue_target(cell: &str) -> bool {
    if matches!(cell, "PinX" | "PinY") {
        return true;
    }
    cell.strip_prefix("Connections.X")
        .and_then(|ordinal| ordinal.parse::<u32>().ok())
        .is_some_and(|ordinal| ordinal >= 1)
}

fn push_quoted_fragment(output: &mut Vec<u8>, value: &str) -> Result<(), VsdxError> {
    output.extend_from_slice(&escape_attribute_value(value, b'"')?);
    Ok(())
}

pub fn remove_connects_referencing_shapes(sheet: &mut Sheet, deleted: &HashSet<u32>) {
    for child in &mut sheet.children {
        let SheetChild::Connects(connects) = child else {
            continue;
        };
        connects.retain(|connect| match connect {
            ConnectsChild::Connect(connect) => {
                !references_deleted_shape([connect.from_sheet, connect.to_sheet], deleted)
            }
            ConnectsChild::Unknown(_) => true,
        });
    }
}

fn references_deleted_shape(
    referenced_shape_ids: impl IntoIterator<Item = u32>,
    deleted: &HashSet<u32>,
) -> bool {
    referenced_shape_ids
        .into_iter()
        .any(|shape_id| deleted.contains(&shape_id))
}

fn reorder_shape(
    part: &PackagePart,
    shapes: &crate::ElementSpan,
    moving_id: u32,
    before_id: Option<u32>,
) -> Result<Vec<SpanEdit>, VsdxError> {
    let shapes =
        shape_container(part, shapes, moving_id).ok_or_else(|| VsdxError::InvalidCellEdit {
            part: part.path.clone(),
            message: format!("shape {moving_id} does not exist"),
        })?;
    let children = shape_children(part, shapes);
    let moving = children
        .iter()
        .find(|shape| shape_id(part, shape) == Some(moving_id))
        .ok_or_else(|| VsdxError::InvalidCellEdit {
            part: part.path.clone(),
            message: format!("shape {moving_id} does not exist"),
        })?;
    let mut ordered: Vec<_> = children
        .iter()
        .copied()
        .filter(|shape| shape.span != moving.span)
        .collect();
    let position = match before_id {
        Some(id) => ordered
            .iter()
            .position(|shape| shape_id(part, shape) == Some(id))
            .ok_or_else(|| VsdxError::InvalidCellEdit {
                part: part.path.clone(),
                message: format!("shape {id} does not exist"),
            })?,
        None => ordered.len(),
    };
    ordered.insert(position, moving);
    let mut replacement = Vec::new();
    let mut cursor = shapes.span.offset;
    for shape in children {
        replacement.extend_from_slice(&part.bytes[cursor..shape.span.offset]);
        let desired = ordered.remove(0);
        replacement.extend_from_slice(
            &part.bytes[desired.span.offset..desired.span.end().ok_or(VsdxError::InvalidSpan)?],
        );
        cursor = shape.span.end().ok_or(VsdxError::InvalidSpan)?;
    }
    replacement
        .extend_from_slice(&part.bytes[cursor..shapes.span.end().ok_or(VsdxError::InvalidSpan)?]);
    Ok(vec![SpanEdit {
        span: shapes.span,
        replacement,
    }])
}

fn shape_container<'a>(
    part: &'a PackagePart,
    page_shapes: &'a crate::ElementSpan,
    id: u32,
) -> Option<&'a crate::ElementSpan> {
    std::iter::once(page_shapes)
        .chain(part.spans.iter().filter(|span| {
            local_name(&span.name) == "Shapes" && contains(page_shapes.span, span.span)
        }))
        .find(|shapes| {
            shape_children(part, shapes)
                .iter()
                .any(|shape| shape_id(part, shape) == Some(id))
        })
}

fn direct_child<'a>(
    part: &'a PackagePart,
    name: &str,
    parent: Option<crate::SourceSpan>,
) -> Option<&'a crate::ElementSpan> {
    part.spans.iter().find(|candidate| {
        local_name(&candidate.name) == name
            && match parent {
                Some(parent) => nearest_parent(part, candidate.span, "Shapes")
                    .is_some_and(|owner| owner.span == parent),
                None => immediate_parent(part, candidate.span)
                    .is_some_and(|owner| local_name(&owner.name) == "PageContents"),
            }
    })
}

/// Inserts containers in schema order and indexed rows in numeric order.
fn container_insertion_point(
    part: &PackagePart,
    owner: &crate::ElementSpan,
    section: Option<&str>,
    row: Option<&CellRow>,
) -> Result<(crate::SourceSpan, bool), VsdxError> {
    let anchor = part
        .spans
        .iter()
        .filter(|candidate| {
            immediate_parent(part, candidate.span).is_some_and(|parent| parent.span == owner.span)
                && match (section, row) {
                    (Some(_), _) => matches!(
                        local_name(&candidate.name),
                        "Text" | "ForeignData" | "Shapes"
                    ),
                    (None, Some(CellRow::Index(index))) => {
                        local_name(&candidate.name) == "Row"
                            && candidate
                                .attributes
                                .get("IX")
                                .and_then(|attribute| attribute_value(&part.bytes, attribute))
                                .and_then(|value| value.parse::<u32>().ok())
                                .is_some_and(|existing| existing > *index)
                    }
                    _ => false,
                }
        })
        .min_by_key(|candidate| candidate.span.offset);
    if let Some(anchor) = anchor {
        return Ok((
            crate::SourceSpan {
                offset: anchor.span.offset,
                length: 0,
            },
            false,
        ));
    }
    let end = owner.span.end().ok_or(VsdxError::InvalidSpan)?;
    if end >= 2 && part.bytes[end - 2..end] == *b"/>" {
        Ok((
            crate::SourceSpan {
                offset: end - 2,
                length: 2,
            },
            true,
        ))
    } else {
        Ok((
            crate::SourceSpan {
                offset: part.bytes[..end]
                    .iter()
                    .rposition(|byte| *byte == b'<')
                    .ok_or(VsdxError::InvalidSpan)?,
                length: 0,
            },
            false,
        ))
    }
}

fn page_shapes(part: &PackagePart) -> Option<&crate::ElementSpan> {
    direct_child(part, "Shapes", None)
}

fn all_shape_spans<'a>(
    part: &'a PackagePart,
    container: &crate::ElementSpan,
) -> Vec<&'a crate::ElementSpan> {
    part.spans
        .iter()
        .filter(|candidate| {
            local_name(&candidate.name) == "Shape" && contains(container.span, candidate.span)
        })
        .collect()
}

fn container_without(
    part: &PackagePart,
    container: &crate::ElementSpan,
    child_name: &str,
    remove: impl Fn(&crate::ElementSpan) -> bool,
) -> Result<SpanEdit, VsdxError> {
    let mut children: Vec<_> = part
        .spans
        .iter()
        .filter(|candidate| {
            local_name(&candidate.name) == child_name
                && nearest_parent(part, candidate.span, &container.name)
                    .is_some_and(|parent| parent.span == container.span)
        })
        .collect();
    children.sort_by_key(|child| child.span.offset);
    let mut replacement = Vec::new();
    let mut cursor = container.span.offset;
    for child in children {
        if remove(child) {
            replacement.extend_from_slice(&part.bytes[cursor..child.span.offset]);
            cursor = child.span.end().ok_or(VsdxError::InvalidSpan)?;
        }
    }
    replacement.extend_from_slice(
        &part.bytes[cursor..container.span.end().ok_or(VsdxError::InvalidSpan)?],
    );
    Ok(SpanEdit {
        span: container.span,
        replacement,
    })
}

/// Checks page-local shape IDs and connector endpoints after structural edits.
pub fn validate_structure(package: &VsdxPackage) -> Result<(), VsdxError> {
    for (path, sheet) in &package.page_contents {
        let mut ids = HashSet::new();
        for shape in all_shapes(sheet) {
            if !ids.insert(shape.id) {
                return Err(VsdxError::InvalidCellEdit {
                    part: path.clone(),
                    message: format!("duplicate shape ID {}", shape.id),
                });
            }
        }
        for connect in sheet.connects() {
            if !ids.contains(&connect.from_sheet) || !ids.contains(&connect.to_sheet) {
                return Err(VsdxError::InvalidCellEdit {
                    part: path.clone(),
                    message: "Connect references a missing shape".to_owned(),
                });
            }
        }
    }
    Ok(())
}

fn all_shapes(sheet: &Sheet) -> Vec<&Shape> {
    fn visit<'a>(shape: &'a Shape, shapes: &mut Vec<&'a Shape>) {
        shapes.push(shape);
        for child in shape.shapes() {
            visit(child, shapes);
        }
    }

    let mut shapes = Vec::new();
    for shape in sheet.shapes() {
        visit(shape, &mut shapes);
    }
    shapes
}

enum LocalCell {
    Existing(String, crate::SourceSpan),
    New(String, crate::SourceSpan),
    NewContainer(String, crate::SourceSpan, Option<String>, Option<CellRow>),
}

fn section_matches_locator(
    part: &PackagePart,
    section: &crate::ElementSpan,
    name: &str,
    index: Option<u32>,
) -> bool {
    attribute_equals(&part.bytes, section, "N", name)
        && section
            .attributes
            .get("IX")
            .and_then(|attribute| attribute_value(&part.bytes, attribute))
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(0)
            == index.unwrap_or(0)
}

fn resolve_cell_locator(
    package: &VsdxPackage,
    locator: &CellLocator,
) -> Result<LocalCell, VsdxError> {
    let path = match locator.sheet {
        CellSheet::Document => Some(package.document_part_path.clone()),
        CellSheet::Page(id) if locator.shape_id.is_some() => package
            .page_part_ids
            .iter()
            .find_map(|(path, candidate)| (*candidate == id).then(|| path.clone())),
        CellSheet::Page(id) => package.parts.iter().find_map(|part| {
            part.spans.iter().find_map(|page| {
                (local_name(&page.name) == "Page"
                    && attribute_equals(&part.bytes, page, "ID", &id.to_string()))
                .then(|| {
                    part.spans
                        .iter()
                        .any(|sheet| {
                            local_name(&sheet.name) == "PageSheet"
                                && contains(page.span, sheet.span)
                        })
                        .then(|| part.path.clone())
                })
                .flatten()
            })
        }),
        CellSheet::Master(id) => package
            .master_part_ids
            .iter()
            .find_map(|(path, candidate)| (*candidate == id).then(|| path.clone())),
    }
    .ok_or_else(|| VsdxError::InvalidCellEdit {
        part: format!("{:?}", locator.sheet),
        message: "sheet does not exist".to_owned(),
    })?;
    let part = package
        .parts
        .iter()
        .find(|part| part.path == path)
        .expect("catalogued part exists");
    let cell = part
        .spans
        .iter()
        .filter(|span| {
            local_name(&span.name) == "Cell"
                && attribute_equals(&part.bytes, span, "N", &locator.cell_name)
        })
        .find(|cell| {
            let nearest = |name: &str| {
                part.spans
                    .iter()
                    .filter(|parent| {
                        local_name(&parent.name) == name && contains(parent.span, cell.span)
                    })
                    .min_by_key(|parent| parent.span.length)
            };
            let shape_matches = match (locator.shape_id, nearest("Shape")) {
                (None, None) => true,
                (Some(id), Some(shape)) => {
                    attribute_equals(&part.bytes, shape, "ID", &id.to_string())
                }
                _ => false,
            };
            let section_matches = match (&locator.section, nearest("Section")) {
                (None, None) => true,
                (Some(name), Some(section)) => {
                    section_matches_locator(part, section, name, locator.section_index)
                }
                _ => false,
            };
            let row_matches = match (&locator.row, nearest("Row")) {
                (None, None) => true,
                (Some(CellRow::Index(index)), Some(row)) => {
                    attribute_equals(&part.bytes, row, "IX", &index.to_string())
                }
                (Some(CellRow::Name(name)), Some(row)) => {
                    attribute_equals(&part.bytes, row, "N", name)
                }
                _ => false,
            };
            shape_matches && section_matches && row_matches
        });
    if let Some(cell) = cell {
        return Ok(LocalCell::Existing(path, cell.span));
    }
    if let (Some(section_name), Some(row)) = (&locator.section, &locator.row) {
        let shape_matches = |candidate: &crate::ElementSpan| match locator.shape_id {
            Some(id) => nearest_parent(part, candidate.span, "Shape")
                .is_some_and(|shape| attribute_equals(&part.bytes, shape, "ID", &id.to_string())),
            None => nearest_parent(part, candidate.span, "Shape").is_none(),
        };
        if let Some(existing_row) = part.spans.iter().find(|candidate| {
            local_name(&candidate.name) == "Row"
                && nearest_parent(part, candidate.span, "Section").is_some_and(|section| {
                    section_matches_locator(part, section, section_name, locator.section_index)
                })
                && shape_matches(candidate)
                && match row {
                    CellRow::Index(index) => {
                        attribute_equals(&part.bytes, candidate, "IX", &index.to_string())
                    }
                    CellRow::Name(name) => attribute_equals(&part.bytes, candidate, "N", name),
                }
        }) {
            return Ok(LocalCell::New(path, existing_row.span));
        }
        if let Some(section) = part.spans.iter().find(|candidate| {
            local_name(&candidate.name) == "Section"
                && section_matches_locator(part, candidate, section_name, locator.section_index)
                && shape_matches(candidate)
        }) {
            return Ok(LocalCell::NewContainer(
                path,
                section.span,
                None,
                Some(row.clone()),
            ));
        }
        if let Some(shape) = part.spans.iter().find(|candidate| {
            local_name(&candidate.name) == "Shape"
                && locator.shape_id.is_some_and(|id| {
                    attribute_equals(&part.bytes, candidate, "ID", &id.to_string())
                })
        }) {
            return Ok(LocalCell::NewContainer(
                path,
                shape.span,
                Some(section_name.clone()),
                Some(row.clone()),
            ));
        }
    }
    let owner_name = if locator.section.is_some() {
        "Row"
    } else if locator.shape_id.is_some() {
        "Shape"
    } else if matches!(locator.sheet, CellSheet::Page(_)) {
        "PageSheet"
    } else {
        "DocumentSheet"
    };
    let owner = part
        .spans
        .iter()
        .filter(|span| local_name(&span.name) == owner_name)
        .find(|owner| {
            let owner = *owner;
            let shape = (local_name(&owner.name) == "Shape")
                .then_some(owner)
                .or_else(|| nearest_parent(part, owner.span, "Shape"));
            let shape_matches = match (locator.shape_id, shape) {
                (None, None) => true,
                (Some(id), Some(shape)) => {
                    attribute_equals(&part.bytes, shape, "ID", &id.to_string())
                }
                _ => false,
            };
            let section_matches = match (
                &locator.section,
                nearest_parent(part, owner.span, "Section"),
            ) {
                (None, None) => true,
                (Some(name), Some(section)) => {
                    section_matches_locator(part, section, name, locator.section_index)
                }
                _ => false,
            };
            let row_matches = match (&locator.row, Some(owner)) {
                (None, _) => true,
                (Some(CellRow::Index(index)), Some(row)) => {
                    attribute_equals(&part.bytes, row, "IX", &index.to_string())
                }
                (Some(CellRow::Name(name)), Some(row)) => {
                    attribute_equals(&part.bytes, row, "N", name)
                }
                _ => false,
            };
            shape_matches && section_matches && row_matches
        });
    owner.map(|owner| LocalCell::New(path.clone(), owner.span)).ok_or_else(|| VsdxError::InvalidCellEdit {
        part: path,
        message: "semantic cell does not exist locally; creating inherited Section or Row is unsupported".to_owned(),
    })
}

fn push_quoted(output: &mut Vec<u8>, value: &str, quote: u8) -> Result<(), VsdxError> {
    output.push(quote);
    output.extend_from_slice(&escape_attribute_value(value, quote)?);
    output.push(quote);
    Ok(())
}

fn nearest_parent<'a>(
    part: &'a PackagePart,
    child: crate::SourceSpan,
    name: &str,
) -> Option<&'a crate::ElementSpan> {
    part.spans
        .iter()
        .filter(|parent| local_name(&parent.name) == name && contains(parent.span, child))
        .min_by_key(|parent| parent.span.length)
}

fn immediate_parent(part: &PackagePart, child: crate::SourceSpan) -> Option<&crate::ElementSpan> {
    part.spans
        .iter()
        .filter(|parent| contains(parent.span, child))
        .min_by_key(|parent| parent.span.length)
}

fn local_name(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, name)| name)
}

/// Offset of the `>` closing the tag opened at `start`, skipping quoted values.
fn tag_close(source: &[u8], start: usize, end: usize) -> Option<usize> {
    let mut index = start;
    let mut quote = None;
    while index < end {
        let byte = source[index];
        if let Some(active) = quote {
            if byte == active {
                quote = None;
            }
        } else if byte == b'\'' || byte == b'"' {
            quote = Some(byte);
        } else if byte == b'>' {
            return Some(index);
        }
        index += 1;
    }
    None
}

fn contains(parent: crate::SourceSpan, child: crate::SourceSpan) -> bool {
    parent.offset < child.offset
        && parent
            .end()
            .is_some_and(|end| child.end().is_some_and(|child_end| child_end <= end))
}

fn attribute_equals(source: &[u8], span: &crate::ElementSpan, name: &str, expected: &str) -> bool {
    span.attributes.get(name).is_some_and(|attribute| {
        attribute_value(source, attribute).is_some_and(|value| value == expected)
    })
}

fn attribute_value<'a>(
    source: &'a [u8],
    attribute: &AttributeSpan,
) -> Option<std::borrow::Cow<'a, str>> {
    let value = source.get(attribute.value.offset..attribute.value.end()?)?;
    quick_xml::escape::unescape(std::str::from_utf8(value).ok()?).ok()
}

fn is_shapesheet_part(package: &VsdxPackage, path: &str) -> bool {
    path == package.document_part_path
        || package.pages_part_path.as_deref() == Some(path)
        || package.masters_part_path.as_deref() == Some(path)
        || package
            .page_part_paths
            .iter()
            .any(|candidate| candidate == path)
        || package
            .master_part_paths
            .iter()
            .any(|candidate| candidate == path)
}

fn load_relationships<'a>(
    source: &str,
    parts: &HashMap<&str, &[u8]>,
    relationships: &'a mut BTreeMap<String, Vec<Relationship>>,
    budget: &mut ParseBudget<'_>,
) -> Result<&'a [Relationship], VsdxError> {
    if !relationships.contains_key(source) {
        let path = relationship_path(source);
        let parsed = parts
            .get(path.as_str())
            .map(|bytes| parse_relationships(bytes, &path, source, budget))
            .transpose()?
            .unwrap_or_default();
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
    #[test]
    fn collects_themes_from_both_relationship_types_and_scheme_ids() {
        fn theme(name: &str, accent: &str, scheme: Option<&str>) -> Vec<u8> {
            let ext = scheme.map_or(String::new(), |value| {
                format!(
                    "<a:extLst><a:ext uri='{{x}}'><vt:schemeID xmlns:vt='http://schemas.microsoft.com/office/visio/2012/main' schemeEnum='{value}'/></a:ext></a:extLst>"
                )
            });
            format!(
                "<a:theme xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' name='{name}'><a:themeElements><a:clrScheme name='{name}'><a:accent1><a:srgbClr val='{accent}'/></a:accent1>{ext}</a:clrScheme></a:themeElements></a:theme>"
            )
            .into_bytes()
        }
        let package = rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>"#.to_vec()),
            ("visio/document.xml".to_owned(), b"<VisioDocument/>".to_vec()),
            ("visio/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/theme' Target='theme/theme1.xml'/><Relationship Id='r2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' Target='theme/theme2.xml'/></Relationships>"#.to_vec()),
            ("visio/theme/theme1.xml".to_owned(), theme("First", "AAAAAA", Some("5"))),
            ("visio/theme/theme2.xml".to_owned(), theme("Second", "BBBBBB", None)),
        ])
        .unwrap();
        let parsed = parse_vsdx(&package).unwrap();
        let accent = |index: u32| {
            parsed
                .themes
                .get(&index)
                .map(|theme| theme.color_scheme.accent1.clone())
        };
        assert_eq!(accent(1).as_deref(), Some("AAAAAA"));
        assert_eq!(accent(2).as_deref(), Some("BBBBBB"));
        assert_eq!(accent(5).as_deref(), Some("AAAAAA"));
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

    use super::*;
    use crate::sheet::serialize_sheet;
    use crate::xml::{XmlNode, parse_xml};
    use crate::{CellAttribute, CellEdit, Shape, SourceSpan};
    use ooxml_opc::{rezip_parts, unzip_parts};

    fn first_value_edit(package: &VsdxPackage, part_path: &str, value: &str) -> CellEdit {
        let cell = package
            .element_spans(part_path)
            .unwrap()
            .iter()
            .find(|span| {
                span.name
                    .rsplit_once(':')
                    .map_or(span.name.as_str(), |(_, name)| name)
                    == "Cell"
                    && span.attributes.contains_key("V")
            })
            .unwrap();
        CellEdit {
            part_path: part_path.to_owned(),
            cell_span: cell.span,
            attribute: CellAttribute::Value,
            value: value.to_owned(),
        }
    }

    fn package_with_page_xml(xml: &[u8]) -> (VsdxPackage, String) {
        let source = include_bytes!("../tests/fixtures/foundation.vsdx");
        let mut parts = unzip_parts(source).unwrap();
        let path = parse_vsdx(source).unwrap().page_part_paths.remove(0);
        parts
            .iter_mut()
            .find(|(candidate, _)| candidate == &path)
            .unwrap()
            .1 = xml.to_vec();
        (parse_vsdx(&rezip_parts(&parts).unwrap()).unwrap(), path)
    }

    fn nested_groups_package() -> (VsdxPackage, String) {
        let package = parse_vsdx(include_bytes!("../tests/fixtures/nested-groups.vsdx")).unwrap();
        let path = package.page_part_paths[0].clone();
        (package, path)
    }

    fn shape_bytes_by_id(package: &VsdxPackage, path: &str) -> BTreeMap<u32, Vec<u8>> {
        let bytes = package.part_bytes(path).unwrap();
        let part = package.parts.iter().find(|part| part.path == path).unwrap();
        package
            .element_spans(path)
            .unwrap()
            .iter()
            .filter(|span| local_name(&span.name) == "Shape")
            .map(|span| {
                (
                    shape_id(part, span).unwrap(),
                    bytes[span.span.offset..span.span.end().unwrap()].to_vec(),
                )
            })
            .collect()
    }

    fn cell_span(package: &VsdxPackage, part_path: &str) -> SourceSpan {
        package
            .element_spans(part_path)
            .unwrap()
            .iter()
            .find(|span| span.name == "Cell")
            .unwrap()
            .span
    }

    #[test]
    fn inserts_a_missing_local_singleton_cell() {
        let source = b"<PageContents><Shapes><Shape ID='1'><Cell N='Other' V='1'/></Shape></Shapes></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let saved = save_semantic_cell_edits(
            &package,
            &[SemanticCellEdit {
                locator: CellLocator {
                    sheet: CellSheet::Page(page_id),
                    shape_id: Some(1),
                    section: None,
                    section_index: None,
                    row: None,
                    cell_name: "Width".to_owned(),
                },
                gesture: MutationGesture::CellEdit,
                row_type: None,
                formula: Some("4".to_owned()),
                value: Some("4".to_owned()),
            }],
        )
        .unwrap();
        let saved = parse_vsdx(&saved).unwrap();
        let after = saved.part_bytes(&path).unwrap();
        assert_eq!(after, b"<PageContents><Shapes><Shape ID='1'><Cell N='Other' V='1'/><Cell N='Width' F='4' V='4'/></Shape></Shapes></PageContents>");
    }

    #[test]
    fn formula_only_edits_drop_existing_caches_and_save_uncached_cells() {
        let source = b"<PageContents><Shapes><Shape ID='1'><Cell N='FOnly' F='a'/><Cell N='Both' F='b' V='2'/><Cell N='VOnly' V='3'/></Shape></Shapes></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let edit = |name: &str| SemanticCellEdit {
            locator: CellLocator {
                sheet: CellSheet::Page(page_id),
                shape_id: Some(1),
                section: None,
                section_index: None,
                row: None,
                cell_name: name.to_owned(),
            },
            gesture: MutationGesture::CellEdit,
            row_type: None,
            formula: Some("9".to_owned()),
            value: None,
        };
        let saved =
            save_semantic_cell_edits(&package, &[edit("FOnly"), edit("Both"), edit("VOnly")])
                .unwrap();
        assert_eq!(
            parse_vsdx(&saved).unwrap().part_bytes(&path).unwrap(),
            b"<PageContents><Shapes><Shape ID='1'><Cell N='FOnly' F='9'/><Cell N='Both' F='9'/><Cell N='VOnly' F='9'/></Shape></Shapes></PageContents>"
        );
    }

    #[test]
    fn inserts_formula_only_cells_without_a_cache() {
        let source = b"<PageContents><Shapes><Shape ID='1'><Cell N='Other' V='1'/></Shape></Shapes></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let saved = save_semantic_cell_edits(
            &package,
            &[SemanticCellEdit {
                locator: CellLocator {
                    sheet: CellSheet::Page(page_id),
                    shape_id: Some(1),
                    section: None,
                    section_index: None,
                    row: None,
                    cell_name: "Width".to_owned(),
                },
                gesture: MutationGesture::CellEdit,
                row_type: None,
                formula: Some("4".to_owned()),
                value: None,
            }],
        )
        .unwrap();
        assert_eq!(
            parse_vsdx(&saved).unwrap().part_bytes(&path).unwrap(),
            b"<PageContents><Shapes><Shape ID='1'><Cell N='Other' V='1'/><Cell N='Width' F='4'/></Shape></Shapes></PageContents>"
        );
        let saved = save_semantic_cell_edits(
            &package,
            &[SemanticCellEdit {
                locator: CellLocator {
                    sheet: CellSheet::Page(page_id),
                    shape_id: Some(1),
                    section: Some("User".to_owned()),
                    section_index: None,
                    row: Some(CellRow::Name("New".to_owned())),
                    cell_name: "Value".to_owned(),
                },
                gesture: MutationGesture::CellEdit,
                row_type: None,
                formula: Some("4".to_owned()),
                value: None,
            }],
        )
        .unwrap();
        assert_eq!(
            parse_vsdx(&saved).unwrap().part_bytes(&path).unwrap(),
            b"<PageContents><Shapes><Shape ID='1'><Cell N='Other' V='1'/><Section N='User'><Row N='New'><Cell N='Value' F='4'/></Row></Section></Shape></Shapes></PageContents>"
        );
    }

    #[test]
    fn semantic_edits_require_a_formula() {
        let source = b"<PageContents><Shapes><Shape ID='1'><Cell N='Width' V='1'/></Shape></Shapes></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let error = save_semantic_cell_edits(
            &package,
            &[SemanticCellEdit {
                locator: CellLocator {
                    sheet: CellSheet::Page(page_id),
                    shape_id: Some(1),
                    section: None,
                    section_index: None,
                    row: None,
                    cell_name: "Width".to_owned(),
                },
                gesture: MutationGesture::CellEdit,
                row_type: None,
                formula: None,
                value: Some("4".to_owned()),
            }],
        )
        .unwrap_err();
        assert!(
            matches!(error, VsdxError::InvalidCellEdit { ref message, .. } if message == "semantic edits require a formula")
        );
    }

    #[test]
    fn inserts_cell_in_missing_row_without_rewriting_section_descendants() {
        let source = b"<PageContents><Shapes><Shape ID='1'><Section N='User'><Row N='Keep'><Cell N='Value' V='old'/></Row></Section></Shape></Shapes></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let saved = save_semantic_cell_edits(
            &package,
            &[semantic_edit(
                page_id,
                "User",
                CellRow::Name("New".to_owned()),
            )],
        )
        .unwrap();
        let after = parse_vsdx(&saved)
            .unwrap()
            .part_bytes(&path)
            .unwrap()
            .to_vec();
        assert_eq!(after, b"<PageContents><Shapes><Shape ID='1'><Section N='User'><Row N='Keep'><Cell N='Value' V='old'/></Row><Row N='New'><Cell N='Value' F='4' V='4'/></Row></Section></Shape></Shapes></PageContents>");
    }

    #[test]
    fn inserts_cell_with_missing_section_and_row() {
        let source = b"<PageContents><Shapes><Shape ID='1'><Cell N='Keep' V='old'/></Shape></Shapes></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let saved = save_semantic_cell_edits(
            &package,
            &[semantic_edit(
                page_id,
                "User",
                CellRow::Name("New".to_owned()),
            )],
        )
        .unwrap();
        let after = parse_vsdx(&saved)
            .unwrap()
            .part_bytes(&path)
            .unwrap()
            .to_vec();
        assert_eq!(after, b"<PageContents><Shapes><Shape ID='1'><Cell N='Keep' V='old'/><Section N='User'><Row N='New'><Cell N='Value' F='4' V='4'/></Row></Section></Shape></Shapes></PageContents>");
    }

    #[test]
    fn inserts_sections_and_indexed_rows_at_shapesheet_anchors() {
        let source = b"<PageContents><Shapes><Shape ID=\"1\"><Cell N=\"Keep\"/><Text>text</Text><ForeignData/><Shapes><Shape ID=\"2\"/></Shapes></Shape></Shapes></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let saved = save_semantic_cell_edits(
            &package,
            &[semantic_edit(page_id, "User", CellRow::Index(0))],
        )
        .unwrap();
        let after = parse_vsdx(&saved)
            .unwrap()
            .part_bytes(&path)
            .unwrap()
            .to_vec();
        assert_eq!(after, b"<PageContents><Shapes><Shape ID=\"1\"><Cell N=\"Keep\"/><Section N=\"User\"><Row IX=\"0\"><Cell N=\"Value\" F=\"4\" V=\"4\"/></Row></Section><Text>text</Text><ForeignData/><Shapes><Shape ID=\"2\"/></Shapes></Shape></Shapes></PageContents>");

        let source = b"<PageContents><Shapes><Shape ID='1'><Section N='User'><Row IX='1'/><Row IX='3'/></Section></Shape></Shapes></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let saved = save_semantic_cell_edits(
            &package,
            &[semantic_edit(page_id, "User", CellRow::Index(2))],
        )
        .unwrap();
        assert_eq!(parse_vsdx(&saved).unwrap().part_bytes(&path).unwrap(), b"<PageContents><Shapes><Shape ID='1'><Section N='User'><Row IX='1'/><Row IX='2'><Cell N='Value' F='4' V='4'/></Row><Row IX='3'/></Section></Shape></Shapes></PageContents>");
    }

    #[test]
    fn groups_typed_container_rows_into_one_section() {
        let source =
            b"<PageContents><Shapes><Shape ID='1'><Cell N='Keep'/></Shape></Shapes></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let edit = |row: u32, name: &str, row_type: &str| SemanticCellEdit {
            locator: CellLocator {
                sheet: CellSheet::Page(page_id),
                shape_id: Some(1),
                section: Some("Geometry".to_owned()),
                section_index: None,
                row: Some(CellRow::Index(row)),
                cell_name: name.to_owned(),
            },
            gesture: MutationGesture::CellEdit,
            row_type: Some(row_type.to_owned()),
            formula: Some("1".to_owned()),
            value: None,
        };
        let saved = save_semantic_cell_edits(
            &package,
            &[
                edit(0, "X", "MoveTo"),
                edit(0, "Y", "MoveTo"),
                edit(1, "X", "LineTo"),
                edit(1, "Y", "LineTo"),
            ],
        )
        .unwrap();
        assert_eq!(
            parse_vsdx(&saved).unwrap().part_bytes(&path).unwrap(),
            b"<PageContents><Shapes><Shape ID='1'><Cell N='Keep'/><Section N='Geometry'><Row IX='0' T='MoveTo'><Cell N='X' F='1'/><Cell N='Y' F='1'/></Row><Row IX='1' T='LineTo'><Cell N='X' F='1'/><Cell N='Y' F='1'/></Row></Section></Shape></Shapes></PageContents>"
        );
    }

    #[test]
    fn inserts_cell_into_self_closing_local_row() {
        let source = b"<PageContents><Shapes><Shape ID='1'><Section N='User'><Row N='New'/></Section></Shape></Shapes></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let saved = save_semantic_cell_edits(
            &package,
            &[semantic_edit(
                page_id,
                "User",
                CellRow::Name("New".to_owned()),
            )],
        )
        .unwrap();
        let after = parse_vsdx(&saved)
            .unwrap()
            .part_bytes(&path)
            .unwrap()
            .to_vec();
        assert_eq!(after, b"<PageContents><Shapes><Shape ID='1'><Section N='User'><Row N='New'><Cell N='Value' F='4' V='4'/></Row></Section></Shape></Shapes></PageContents>");
    }

    #[test]
    fn deletes_a_shape_and_its_connects_without_rewriting_siblings() {
        let source = b"<PageContents><Shapes><Shape ID='1'><Cell N='Keep' V='one'/></Shape><Shape ID='2'><Cell N='Keep' V='two'/></Shape><Shape ID='3'><Cell N='Keep' V='three'/></Shape></Shapes><Connects><Connect FromSheet='1' ToSheet='2'/><Connect FromSheet='2' ToSheet='3'/></Connects></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let saved = save_structural_edits(
            &package,
            &[StructuralEdit::DeleteShape {
                page_id,
                shape_id: 2,
            }],
        )
        .unwrap();
        let after = parse_vsdx(&saved)
            .unwrap()
            .part_bytes(&path)
            .unwrap()
            .to_vec();
        assert_eq!(after, b"<PageContents><Shapes><Shape ID='1'><Cell N='Keep' V='one'/></Shape><Shape ID='3'><Cell N='Keep' V='three'/></Shape></Shapes><Connects></Connects></PageContents>");
    }

    #[test]
    fn adds_shapes_with_page_local_ids_in_normal_and_self_closing_containers() {
        for (source, expected) in [
            (
                b"<PageContents><Shapes><Shape ID='4'>one</Shape><Shape ID='7'>two</Shape></Shapes></PageContents>".as_slice(),
                b"<PageContents><Shapes><Shape ID='4'>one</Shape><Shape ID='7'>two</Shape><Shape ID='8'><Cell N='Width' V='1'/></Shape></Shapes></PageContents>".as_slice(),
            ),
            (
                b"<PageContents><Shapes/></PageContents>".as_slice(),
                b"<PageContents><Shapes><Shape ID='1'><Cell N='Width' V='1'/></Shape></Shapes></PageContents>".as_slice(),
            ),
        ] {
            let (package, path) = package_with_page_xml(source);
            let page_id = package.page_part_ids[&path];
            let saved = save_structural_edits(&package, &[StructuralEdit::AddShape { page_id, shape_xml: b"<Shape><Cell N='Width' V='1'/></Shape>".to_vec() }]).unwrap();
            let saved = parse_vsdx(&saved).unwrap();
            assert_eq!(saved.part_bytes(&path).unwrap(), expected);
            validate_structure(&saved).unwrap();
        }
    }

    #[test]
    fn adding_a_relationship_bearing_shape_is_rejected_without_changing_the_package() {
        let (package, path) =
            package_with_page_xml(b"<PageContents><Shapes><Shape ID='1'/></Shapes></PageContents>");
        let page_id = package.page_part_ids[&path];
        let before: BTreeMap<_, _> = package
            .parts
            .iter()
            .map(|part| (part.path.clone(), part.bytes.clone()))
            .collect();
        assert!(matches!(
            save_structural_edits(
                &package,
                &[StructuralEdit::AddShape {
                    page_id,
                    shape_xml: b"<Shape><ForeignData><Rel r:id='missing' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'/></ForeignData></Shape>".to_vec(),
                }],
            ),
            Err(VsdxError::InvalidCellEdit { message, .. }) if message == "new shape fragments cannot contain relationship references"
        ));
        for part in &package.parts {
            assert_eq!(part.bytes, before[&part.path], "{}", part.path);
        }
    }

    #[test]
    fn adding_a_shape_rejects_id_exhaustion() {
        let (package, path) = package_with_page_xml(
            b"<PageContents><Shapes><Shape ID='4294967295'/></Shapes></PageContents>",
        );
        let page_id = package.page_part_ids[&path];
        assert!(matches!(
            save_structural_edits(
                &package,
                &[StructuralEdit::AddShape {
                    page_id,
                    shape_xml: b"<Shape/>".to_vec()
                }]
            ),
            Err(VsdxError::InvalidCellEdit { message, .. }) if message == "shape ID space is exhausted"
        ));
    }

    #[test]
    fn structural_edits_preserve_group_boundaries_in_nested_groups_fixture() {
        let (package, path) = nested_groups_package();
        let page_id = package.page_part_ids[&path];
        let original = package.part_bytes(&path).unwrap();
        let group = package
            .element_spans(&path)
            .unwrap()
            .iter()
            .find(|span| span.name == "Shape" && attribute_equals(original, span, "ID", "1"))
            .unwrap();
        let group_bytes = &original[group.span.offset..group.span.end().unwrap()];

        let added = save_structural_edits(
            &package,
            &[StructuralEdit::AddShape {
                page_id,
                shape_xml: b"<Shape Type='Shape'/>".to_vec(),
            }],
        )
        .unwrap();
        let added = parse_vsdx(&added).unwrap();
        let bytes = added.part_bytes(&path).unwrap();
        assert!(
            bytes
                .windows(b"<Shape Type='Shape' ID='6'/>".len())
                .any(|window| window == b"<Shape Type='Shape' ID='6'/>")
        );
        assert!(
            bytes
                .windows(group_bytes.len())
                .any(|window| window == group_bytes)
        );

        let deleted = save_structural_edits(
            &package,
            &[StructuralEdit::DeleteShape {
                page_id,
                shape_id: 1,
            }],
        )
        .unwrap();
        let deleted = parse_vsdx(&deleted).unwrap();
        let deleted_bytes = deleted.part_bytes(&path).unwrap();
        for id in [1, 2, 3, 4, 5] {
            assert!(!shape_bytes_by_id(&deleted, &path).contains_key(&id));
        }
        assert!(
            !deleted_bytes
                .windows(b"<Connect ".len())
                .any(|window| window == b"<Connect ")
        );
        validate_structure(&deleted).unwrap();

        let reordered = save_structural_edits(
            &package,
            &[StructuralEdit::ReorderShape {
                page_id,
                shape_id: 0,
                before_shape_id: Some(1),
            }],
        )
        .unwrap();
        let reordered = parse_vsdx(&reordered).unwrap();
        let reordered_bytes = reordered.part_bytes(&path).unwrap();
        assert!(
            reordered_bytes
                .windows(b"<Shape ID='0' Type='Shape'><Text>page sibling</Text></Shape><Shape ID='1'".len())
                .any(|window| window == b"<Shape ID='0' Type='Shape'><Text>page sibling</Text></Shape><Shape ID='1'")
        );
        assert_eq!(shape_bytes_by_id(&reordered, &path)[&1], group_bytes);
    }

    #[test]
    fn allocation_and_connect_validation_include_group_members() {
        let source = b"<PageContents><Shapes><Shape ID='1' Type='Group'><Shapes><Shape ID='2'/></Shapes></Shape></Shapes><Connects><Connect FromSheet='1' ToSheet='2'/></Connects></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let saved = save_structural_edits(
            &package,
            &[StructuralEdit::AddShape {
                page_id,
                shape_xml: b"<Shape/>".to_vec(),
            }],
        )
        .unwrap();
        assert!(
            parse_vsdx(&saved)
                .unwrap()
                .part_bytes(&path)
                .unwrap()
                .windows(b"<Shape ID='3'/>".len())
                .any(|window| window == b"<Shape ID='3'/>")
        );

        let saved = save_structural_edits(
            &package,
            &[StructuralEdit::DeleteShape {
                page_id,
                shape_id: 1,
            }],
        )
        .unwrap();
        let saved = parse_vsdx(&saved).unwrap();
        assert_eq!(
            saved.part_bytes(&path).unwrap(),
            b"<PageContents><Shapes></Shapes><Connects></Connects></PageContents>"
        );
        validate_structure(&saved).unwrap();
    }

    #[test]
    fn add_connect_appends_glue_without_touching_existing_spans() {
        let source = b"<PageContents><Shapes><Shape ID='1'/><Shape ID='2'/></Shapes><Connects><Connect FromSheet='1' FromCell='BeginX' ToSheet='1' ToCell='PinX'/></Connects></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let saved = save_structural_edits(
            &package,
            &[StructuralEdit::AddConnect {
                page_id,
                from_sheet: 2,
                from_cell: "EndX".to_owned(),
                to_sheet: 1,
                to_cell: "Connections.X1".to_owned(),
            }],
        )
        .unwrap();
        let reparsed = parse_vsdx(&saved).unwrap();
        let after = reparsed.part_bytes(&path).unwrap();
        assert!(
            after
                .windows(
                    b"<Connect FromSheet='1' FromCell='BeginX' ToSheet='1' ToCell='PinX'/>".len()
                )
                .any(|window| window
                    == b"<Connect FromSheet='1' FromCell='BeginX' ToSheet='1' ToCell='PinX'/>")
        );
        assert!(after
            .windows(b"<Connect FromSheet=\"2\" FromCell=\"EndX\" ToSheet=\"1\" ToCell=\"Connections.X1\"/>".len())
            .any(|window| window
                == b"<Connect FromSheet=\"2\" FromCell=\"EndX\" ToSheet=\"1\" ToCell=\"Connections.X1\"/>"));
        assert!(
            after
                .windows(b"<Shape ID='1'/>".len())
                .any(|window| window == b"<Shape ID='1'/>")
        );
        let connects = reparsed.page_contents[&path].connects().collect::<Vec<_>>();
        assert_eq!(connects.len(), 2);
        assert_eq!(connects[1].from_sheet, 2);
        assert_eq!(connects[1].from_cell.as_deref(), Some("EndX"));
        assert_eq!(connects[1].to_sheet, 1);
        assert_eq!(connects[1].to_cell.as_deref(), Some("Connections.X1"));
        validate_structure(&reparsed).unwrap();
    }

    #[test]
    fn add_connect_creates_a_connects_container_when_missing() {
        let source =
            b"<PageContents><Shapes><Shape ID='1'/><Shape ID='2'/></Shapes></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        let saved = save_structural_edits(
            &package,
            &[StructuralEdit::AddConnect {
                page_id,
                from_sheet: 2,
                from_cell: "BeginX".to_owned(),
                to_sheet: 1,
                to_cell: "PinX".to_owned(),
            }],
        )
        .unwrap();
        let reparsed = parse_vsdx(&saved).unwrap();
        assert_eq!(
            reparsed.part_bytes(&path).unwrap(),
            b"<PageContents><Shapes><Shape ID='1'/><Shape ID='2'/></Shapes><Connects><Connect FromSheet=\"2\" FromCell=\"BeginX\" ToSheet=\"1\" ToCell=\"PinX\"/></Connects></PageContents>"
        );
        validate_structure(&reparsed).unwrap();
    }

    #[test]
    fn add_shape_preserves_a_prefixed_empty_shapes_container() {
        let source = b"<v:PageContents xmlns:v='http://schemas.microsoft.com/office/visio/2012/main'><v:Shapes/></v:PageContents>";
        let (package, path) = package_with_page_xml(source);
        let saved = save_structural_edits(
            &package,
            &[StructuralEdit::AddShape {
                page_id: package.page_part_ids[&path],
                shape_xml: b"<v:Shape><v:Cell N='Width' V='1'/></v:Shape>".to_vec(),
            }],
        )
        .unwrap();
        let reparsed = parse_vsdx(&saved).unwrap();
        let after = std::str::from_utf8(reparsed.part_bytes(&path).unwrap()).unwrap();
        assert!(after.contains("</v:Shapes>"), "{after}");
        assert!(!after.contains("</Shapes>"), "{after}");
        assert_eq!(reparsed.page_contents[&path].shapes().count(), 1);
        validate_structure(&reparsed).unwrap();
    }

    #[test]
    fn add_connect_preserves_prefixed_containers() {
        for container in ["<v:Connects/>", "<v:Connects></v:Connects>", ""] {
            let source = format!(
                "<v:PageContents xmlns:v='http://schemas.microsoft.com/office/visio/2012/main'><v:Shapes><v:Shape ID='1'/><v:Shape ID='2'/></v:Shapes>{container}</v:PageContents>"
            );
            let (package, path) = package_with_page_xml(source.as_bytes());
            let saved = save_structural_edits(
                &package,
                &[StructuralEdit::AddConnect {
                    page_id: package.page_part_ids[&path],
                    from_sheet: 2,
                    from_cell: "BeginX".to_owned(),
                    to_sheet: 1,
                    to_cell: "PinX".to_owned(),
                }],
            )
            .unwrap();
            let reparsed = parse_vsdx(&saved).unwrap();
            let after = std::str::from_utf8(reparsed.part_bytes(&path).unwrap()).unwrap();
            assert!(after.contains("<v:Connect "));
            assert!(after.contains("</v:Connects>"));
            assert_eq!(reparsed.page_contents[&path].connects().count(), 1);
            validate_structure(&reparsed).unwrap();
        }
    }

    #[test]
    fn add_connect_refuses_cells_outside_the_glue_contract() {
        let source =
            b"<PageContents><Shapes><Shape ID='1'/><Shape ID='2'/></Shapes></PageContents>";
        let (package, path) = package_with_page_xml(source);
        let page_id = package.page_part_ids[&path];
        for (from_cell, to_cell) in [
            ("PinX", "PinX"),
            ("BeginX", "Width"),
            ("BeginX", "Connections.X0"),
            ("BeginX", "Connections.X"),
        ] {
            assert!(
                save_structural_edits(
                    &package,
                    &[StructuralEdit::AddConnect {
                        page_id,
                        from_sheet: 2,
                        from_cell: from_cell.to_owned(),
                        to_sheet: 1,
                        to_cell: to_cell.to_owned(),
                    }],
                )
                .is_err(),
                "{from_cell} -> {to_cell}"
            );
        }
        assert!(
            save_structural_edits(
                &package,
                &[StructuralEdit::AddConnect {
                    page_id,
                    from_sheet: 2,
                    from_cell: "BeginX".to_owned(),
                    to_sheet: 99,
                    to_cell: "PinX".to_owned(),
                }],
            )
            .is_err()
        );
    }

    #[test]
    fn add_connect_leaves_other_parts_byte_identical() {
        let package = parse_vsdx(include_bytes!("../tests/fixtures/foundation.vsdx")).unwrap();
        let path = package.page_part_paths[0].clone();
        let page_id = package.page_part_ids[&path];
        let saved = save_structural_edits(
            &package,
            &[StructuralEdit::AddConnect {
                page_id,
                from_sheet: 1,
                from_cell: "EndX".to_owned(),
                to_sheet: 1,
                to_cell: "PinY".to_owned(),
            }],
        )
        .unwrap();
        let reparsed = parse_vsdx(&saved).unwrap();
        for part in &package.parts {
            if part.path == path {
                continue;
            }
            assert_eq!(
                reparsed.part_bytes(&part.path),
                Some(part.bytes.as_slice()),
                "untouched part changed: {}",
                part.path
            );
        }
        let after = reparsed.part_bytes(&path).unwrap();
        assert!(
            after
                .windows(b"Mystery='yes'".len())
                .any(|window| window == b"Mystery='yes'")
        );
        assert!(
            after
                .windows(b"<UnknownConnect Flag='yes'>".len())
                .any(|window| window == b"<UnknownConnect Flag='yes'>")
        );
    }

    #[test]
    fn deleting_a_missing_shape_fails() {
        let (package, path) =
            package_with_page_xml(b"<PageContents><Shapes><Shape ID='1'/></Shapes></PageContents>");
        let page_id = package.page_part_ids[&path];
        assert!(matches!(
            save_structural_edits(&package, &[StructuralEdit::DeleteShape { page_id, shape_id: 2 }]),
            Err(VsdxError::InvalidCellEdit { message, .. }) if message == "shape 2 does not exist"
        ));
    }

    #[test]
    fn reorders_shapes_without_rewriting_any_shape_bytes() {
        for (shape_id, before, expected) in [
            (
                1,
                None,
                b"<Shape ID='2'>two</Shape><Shape ID='3'>three</Shape><Shape ID='1'>one</Shape>"
                    .as_slice(),
            ),
            (
                3,
                Some(1),
                b"<Shape ID='3'>three</Shape><Shape ID='1'>one</Shape><Shape ID='2'>two</Shape>"
                    .as_slice(),
            ),
            (
                3,
                Some(2),
                b"<Shape ID='1'>one</Shape><Shape ID='3'>three</Shape><Shape ID='2'>two</Shape>"
                    .as_slice(),
            ),
        ] {
            let (package, path) = package_with_page_xml(b"<PageContents><Shapes><Shape ID='1'>one</Shape><Shape ID='2'>two</Shape><Shape ID='3'>three</Shape></Shapes></PageContents>");
            let page_id = package.page_part_ids[&path];
            let saved = save_structural_edits(
                &package,
                &[StructuralEdit::ReorderShape {
                    page_id,
                    shape_id,
                    before_shape_id: before,
                }],
            )
            .unwrap();
            let saved = parse_vsdx(&saved).unwrap();
            let after = saved.part_bytes(&path).unwrap();
            assert_eq!(
                after,
                [
                    b"<PageContents><Shapes>".as_slice(),
                    expected,
                    b"</Shapes></PageContents>"
                ]
                .concat()
            );
            for shape in [
                b"<Shape ID='1'>one</Shape>".as_slice(),
                b"<Shape ID='2'>two</Shape>",
                b"<Shape ID='3'>three</Shape>",
            ] {
                assert!(after.windows(shape.len()).any(|window| window == shape));
            }
        }
    }

    fn semantic_edit(page_id: u32, section: &str, row: CellRow) -> SemanticCellEdit {
        SemanticCellEdit {
            locator: CellLocator {
                sheet: CellSheet::Page(page_id),
                shape_id: Some(1),
                section: Some(section.to_owned()),
                section_index: None,
                row: Some(row),
                cell_name: "Value".to_owned(),
            },
            gesture: MutationGesture::CellEdit,
            row_type: None,
            formula: Some("4".to_owned()),
            value: Some("4".to_owned()),
        }
    }

    fn assert_only_span_changed(
        before: &[u8],
        after: &[u8],
        span: SourceSpan,
        replacement_length: usize,
    ) {
        assert_eq!(&before[..span.offset], &after[..span.offset]);
        assert_eq!(
            &before[span.offset + span.length..],
            &after[span.offset + replacement_length..]
        );
    }

    fn sheet_has_value(sheet: &Sheet, value: &str) -> bool {
        sheet
            .cells()
            .any(|cell| cell.value.as_deref() == Some(value))
            || sheet.shapes().any(|shape| shape_has_value(shape, value))
    }

    fn shape_has_value(shape: &Shape, value: &str) -> bool {
        shape
            .cells()
            .any(|cell| cell.value.as_deref() == Some(value))
            || shape.sections().any(|section| {
                section
                    .rows()
                    .flat_map(crate::Row::cells)
                    .any(|cell| cell.value.as_deref() == Some(value))
            })
            || shape.shapes().any(|shape| shape_has_value(shape, value))
    }

    #[test]
    fn records_source_spans_for_shapesheet_entities() {
        let package = parse_vsdx(include_bytes!("../tests/fixtures/foundation.vsdx")).unwrap();
        let part_path = package.page_part_paths.first().unwrap();
        let bytes = package.part_bytes(part_path).unwrap();
        let spans = package.element_spans(part_path).unwrap();
        for name in ["Shape", "Section", "Row", "Cell", "Text"] {
            assert!(spans.iter().any(|span| span.name == name), "{name}");
        }
        for span in spans {
            assert_eq!(bytes[span.span.offset], b'<');
            assert_eq!(bytes[span.span.offset + span.span.length - 1], b'>');
        }
    }

    #[test]
    fn records_exact_spans_for_shapesheet_entities_and_mixed_text() {
        let source = b"<PageContents><Shapes><Shape ID='7'><Section N='Geometry'><Row IX='2'><Cell N='X'/></Row></Section><Text>before<cp IX='0'/>after</Text></Shape></Shapes></PageContents>";
        let spans = scan_element_spans(source).unwrap();
        for (name, expected) in [
            (
                "Shape",
                "<Shape ID='7'><Section N='Geometry'><Row IX='2'><Cell N='X'/></Row></Section><Text>before<cp IX='0'/>after</Text></Shape>",
            ),
            (
                "Section",
                "<Section N='Geometry'><Row IX='2'><Cell N='X'/></Row></Section>",
            ),
            ("Row", "<Row IX='2'><Cell N='X'/></Row>"),
            ("Cell", "<Cell N='X'/>"),
            ("Text", "<Text>before<cp IX='0'/>after</Text>"),
        ] {
            let span = spans.iter().find(|span| span.name == name).unwrap().span;
            assert_eq!(
                &source[span.offset..span.end().unwrap()],
                expected.as_bytes(),
                "{name}"
            );
        }
    }

    #[test]
    fn inserts_missing_cell_attributes_with_existing_quote_style() {
        for (xml, edits, expected) in [
            (b"<PageContents><Shapes><Shape ID='1'><Cell N='X' F='a'/></Shape></Shapes></PageContents>".as_slice(), vec![(CellAttribute::Value, "v")], "<Cell N='X' F='a' V='v'/>"),
            (b"<PageContents><Shapes><Shape ID='1'><Cell N='X' V='v'/></Shape></Shapes></PageContents>".as_slice(), vec![(CellAttribute::Formula, "a")], "<Cell N='X' V='v' F='a'/>"),
            (b"<PageContents><Shapes><Shape ID='1'><Cell N='X'/></Shape></Shapes></PageContents>".as_slice(), vec![(CellAttribute::Formula, "a"), (CellAttribute::Value, "v")], "<Cell N='X' F='a' V='v'/>"),
        ] {
            let (package, path) = package_with_page_xml(xml);
            let span = cell_span(&package, &path);
            let edits = edits.into_iter().map(|(attribute, value)| CellEdit { part_path: path.clone(), cell_span: span, attribute, value: value.to_owned() }).collect::<Vec<_>>();
            let saved = save_cell_edits(&package, &edits).unwrap();
            let part = unzip_parts(&saved).unwrap().into_iter().find(|(candidate, _)| candidate == &path).unwrap().1;
            assert!(std::str::from_utf8(&part).unwrap().contains(expected));
        }
    }

    #[test]
    fn saves_cell_value_without_rewriting_other_parts_or_lexical_spans() {
        let source = include_bytes!("../tests/fixtures/foundation.vsdx");
        let package = parse_vsdx(source).unwrap();
        let part_path = package.page_part_paths.first().unwrap();
        let edit = first_value_edit(&package, part_path, "patched");
        let before: BTreeMap<_, _> = unzip_parts(source).unwrap().into_iter().collect();
        let saved = save_cell_edits(&package, std::slice::from_ref(&edit)).unwrap();
        let after: BTreeMap<_, _> = unzip_parts(&saved).unwrap().into_iter().collect();
        for (path, bytes) in &before {
            if path == &edit.part_path {
                let original = package.part_bytes(path).unwrap();
                let cell = package
                    .element_spans(path)
                    .unwrap()
                    .iter()
                    .find(|span| span.span == edit.cell_span)
                    .unwrap();
                assert_only_span_changed(
                    original,
                    &after[path],
                    cell.attributes["V"].value,
                    edit.value.len(),
                );
            } else {
                assert_eq!(&after[path], bytes, "{path}");
            }
        }
        let reparsed = parse_vsdx(&saved).unwrap();
        assert!(
            reparsed
                .page_contents
                .values()
                .any(|sheet| sheet_has_value(sheet, "patched"))
        );
    }

    #[test]
    fn saves_shape_text_without_rewriting_other_parts_or_lexical_spans() {
        let source = include_bytes!("../tests/fixtures/foundation.vsdx");
        let package = parse_vsdx(source).unwrap();
        let part_path = package.page_part_paths.first().unwrap().clone();
        let page_id = package.page_part_ids.get(&part_path).copied().unwrap();
        let edit = SemanticTextEdit {
            locator: CellLocator {
                sheet: CellSheet::Page(page_id),
                shape_id: Some(1),
                section: None,
                section_index: None,
                row: None,
                cell_name: "Text".to_owned(),
            },
            text: "Hello & <World>\nNew line".to_owned(),
        };
        let before: BTreeMap<_, _> = unzip_parts(source).unwrap().into_iter().collect();
        let saved = save_semantic_text_edits(&package, std::slice::from_ref(&edit)).unwrap();
        let after: BTreeMap<_, _> = unzip_parts(&saved).unwrap().into_iter().collect();
        for (path, bytes) in &before {
            if path == &part_path {
                let original = package.part_bytes(path).unwrap();
                let text = package
                    .element_spans(path)
                    .unwrap()
                    .iter()
                    .find(|span| {
                        span.name == "Text"
                            && original[span.span.offset..span.span.end().unwrap()]
                                .windows(b"<fld".len())
                                .any(|window| window == b"<fld")
                    })
                    .unwrap();
                let outer_end = text.span.end().unwrap();
                let inner_start = original[text.span.offset..outer_end]
                    .iter()
                    .position(|byte| *byte == b'>')
                    .unwrap()
                    + text.span.offset
                    + 1;
                let closing = original[inner_start..outer_end]
                    .iter()
                    .rposition(|byte| *byte == b'<')
                    .unwrap()
                    + inner_start;
                assert_only_span_changed(
                    original,
                    &after[path],
                    crate::SourceSpan {
                        offset: inner_start,
                        length: closing - inner_start,
                    },
                    "Hello &amp; &lt;World&gt;\nNew line".len(),
                );
            } else {
                assert_eq!(&after[path], bytes, "{path}");
            }
        }
        let reparsed = parse_vsdx(&saved).unwrap();
        let sheet = reparsed.page_contents.get(&part_path).unwrap();
        let shape = sheet.shapes().find(|shape| shape.id == 1).unwrap();
        let round_tripped = shape
            .text()
            .unwrap_or_default()
            .iter()
            .filter_map(|token| match token {
                crate::TextToken::Literal(value) => Some(value.clone()),
                _ => None,
            })
            .collect::<String>();
        assert_eq!(round_tripped, "Hello & <World>\nNew line");
    }

    #[test]
    fn saving_shape_text_rejects_non_shape_targets_and_forbidden_characters() {
        let package = parse_vsdx(include_bytes!("../tests/fixtures/foundation.vsdx")).unwrap();
        let part_path = package.page_part_paths.first().unwrap().clone();
        let page_id = package.page_part_ids.get(&part_path).copied().unwrap();
        let locator = |shape_id: Option<u32>| CellLocator {
            sheet: CellSheet::Page(page_id),
            shape_id,
            section: None,
            section_index: None,
            row: None,
            cell_name: "Text".to_owned(),
        };
        assert!(matches!(
            save_semantic_text_edits(
                &package,
                &[SemanticTextEdit {
                    locator: locator(None),
                    text: "x".to_owned()
                }]
            ),
            Err(VsdxError::InvalidCellEdit { .. })
        ));
        assert!(matches!(
            save_semantic_text_edits(
                &package,
                &[SemanticTextEdit {
                    locator: locator(Some(999)),
                    text: "x".to_owned()
                }]
            ),
            Err(VsdxError::InvalidCellEdit { .. })
        ));
        assert!(matches!(
            save_semantic_text_edits(
                &package,
                &[SemanticTextEdit {
                    locator: locator(Some(1)),
                    text: "bad\0".to_owned()
                }]
            ),
            Err(VsdxError::InvalidXmlCharacter)
        ));
    }

    #[test]
    fn resolves_semantic_locators_with_escaped_attributes() {
        let (package, path) = package_with_page_xml(
            b"<PageContents><Shapes><Shape ID='&#49;'><Cell N='A&amp;B' V='direct'/><Section N='A&#x26;B'><Row IX='0'><Cell N='SectionCell' V='section'/></Row></Section><Cell N='IdCell' V='id'/></Shape></Shapes></PageContents>",
        );
        let cases = [
            (
                CellLocator {
                    sheet: CellSheet::Page(*package.page_part_ids.get(&path).unwrap()),
                    shape_id: Some(1),
                    section: None,
                    section_index: None,
                    row: None,
                    cell_name: "A&B".to_owned(),
                },
                "A&B",
                "patched-direct",
            ),
            (
                CellLocator {
                    sheet: CellSheet::Page(*package.page_part_ids.get(&path).unwrap()),
                    shape_id: Some(1),
                    section: Some("A&B".to_owned()),
                    section_index: None,
                    row: Some(CellRow::Index(0)),
                    cell_name: "SectionCell".to_owned(),
                },
                "SectionCell",
                "patched-section",
            ),
            (
                CellLocator {
                    sheet: CellSheet::Page(*package.page_part_ids.get(&path).unwrap()),
                    shape_id: Some(1),
                    section: None,
                    section_index: None,
                    row: None,
                    cell_name: "IdCell".to_owned(),
                },
                "IdCell",
                "patched-id",
            ),
        ];
        for (locator, _, new_value) in cases {
            let saved = save_semantic_cell_edits(
                &package,
                &[SemanticCellEdit {
                    locator,
                    gesture: MutationGesture::CellEdit,
                    row_type: None,
                    formula: Some("2".to_owned()),
                    value: Some(new_value.to_owned()),
                }],
            )
            .unwrap();
            let reparsed = parse_vsdx(&saved).unwrap();
            assert!(
                reparsed
                    .page_contents
                    .values()
                    .any(|sheet| sheet_has_value(sheet, new_value))
            );
        }
    }

    #[test]
    fn saves_existing_cell_formula() {
        let source = include_bytes!("../tests/fixtures/foundation.vsdx");
        let package = parse_vsdx(source).unwrap();
        let part_path = package.page_part_paths.first().unwrap();
        let cell = package
            .element_spans(part_path)
            .unwrap()
            .iter()
            .find(|span| {
                span.name == "Cell"
                    && span.attributes.contains_key("F")
                    && span.attributes.contains_key("V")
            })
            .unwrap();
        let saved = save_cell_edits(
            &package,
            &[CellEdit {
                part_path: part_path.to_owned(),
                cell_span: cell.span,
                attribute: CellAttribute::Formula,
                value: "Width*3".to_owned(),
            }],
        )
        .unwrap();
        let reparsed = parse_vsdx(&saved).unwrap();
        assert!(
            reparsed
                .page_contents
                .values()
                .flat_map(Sheet::shapes)
                .any(|shape| shape
                    .cells()
                    .any(|cell| cell.formula.as_deref() == Some("Width*3")))
        );
    }

    #[test]
    fn saves_real_corpus_cells_without_rewriting_other_parts() {
        let Some(directory) = std::env::var_os("VSDX_CORPUS_DIR") else {
            eprintln!("SKIPPED REVIEW CORPUS TEST: VSDX_CORPUS_DIR is unset");
            return;
        };
        for name in ["lichtsysteme.vsdx", "soundplan.vsdx"] {
            let source = std::fs::read(std::path::Path::new(&directory).join(name)).unwrap();
            let package = parse_vsdx(&source).unwrap();
            let part_path = package.page_part_paths.first().unwrap();
            let edit = first_value_edit(&package, part_path, "patched");
            let before: BTreeMap<_, _> = unzip_parts(&source).unwrap().into_iter().collect();
            let saved = save_cell_edits(&package, std::slice::from_ref(&edit)).unwrap();
            let after: BTreeMap<_, _> = unzip_parts(&saved).unwrap().into_iter().collect();
            for (path, bytes) in before {
                if path == part_path.as_str() {
                    let cell = package
                        .element_spans(&path)
                        .unwrap()
                        .iter()
                        .find(|span| span.span == edit.cell_span)
                        .unwrap();
                    assert_only_span_changed(
                        &bytes,
                        &after[&path],
                        cell.attributes["V"].value,
                        edit.value.len(),
                    );
                } else {
                    assert_eq!(after[&path], bytes, "{name}: {path}");
                }
            }
        }
    }

    #[test]
    fn cell_edits_preserve_quotes_escape_values_and_are_transactional() {
        let source = include_bytes!("../tests/fixtures/foundation.vsdx");
        let package = parse_vsdx(source).unwrap();
        let part_path = package.page_part_paths.first().unwrap();
        let mut edit = first_value_edit(&package, part_path, "<&\"'");
        let cell = package
            .element_spans(part_path)
            .unwrap()
            .iter()
            .find(|span| span.span == edit.cell_span)
            .unwrap();
        assert_eq!(cell.attributes["V"].quote, b'\'');
        let saved = save_cell_edits(&package, &[edit.clone()]).unwrap();
        let mut parts: BTreeMap<_, _> = unzip_parts(&saved).unwrap().into_iter().collect();
        let part = parts.remove(part_path).unwrap();
        assert!(
            std::str::from_utf8(&part)
                .unwrap()
                .contains("V='&lt;&amp;&quot;&apos;'")
        );
        let reparsed = parse_vsdx(&saved).unwrap();
        assert!(
            reparsed
                .page_contents
                .values()
                .any(|sheet| sheet_has_value(sheet, "<&\"'"))
        );

        edit.cell_span.offset = usize::MAX;
        let before = package.part_bytes(part_path).unwrap().to_vec();
        assert!(matches!(
            save_cell_edits(
                &package,
                &[first_value_edit(&package, part_path, "ok"), edit]
            ),
            Err(VsdxError::InvalidCellEdit { .. })
        ));
        assert_eq!(package.part_bytes(part_path).unwrap(), before);
    }

    #[test]
    fn only_authoritative_shapesheet_parts_receive_spans_or_edits() {
        let source = include_bytes!("../tests/fixtures/foundation.vsdx");
        let mut parts = unzip_parts(source).unwrap();
        parts.push((
            "visio/media/forged.png".to_owned(),
            b"not PNG <Cell V='1'/>".to_vec(),
        ));
        parts.push((
            "customXml/item1.xml".to_owned(),
            b"<root><Cell V='1'/></root>".to_vec(),
        ));
        let package = parse_vsdx(&rezip_parts(&parts).unwrap()).unwrap();
        for path in ["visio/media/forged.png", "customXml/item1.xml"] {
            assert_eq!(package.element_spans(path), Some([].as_slice()));
            assert!(matches!(
                save_cell_edits(
                    &package,
                    &[CellEdit {
                        part_path: path.to_owned(),
                        cell_span: SourceSpan {
                            offset: 0,
                            length: 13
                        },
                        attribute: CellAttribute::Value,
                        value: "changed".to_owned(),
                    }]
                ),
                Err(VsdxError::InvalidCellEdit { .. })
            ));
        }
    }

    #[test]
    fn replacing_a_part_invalidates_its_edit_provenance() {
        let source = include_bytes!("../tests/fixtures/foundation.vsdx");
        let mut package = parse_vsdx(source).unwrap();
        let part_path = package.page_part_paths.first().unwrap().clone();
        let edit = first_value_edit(&package, &part_path, "changed");
        let replacement = package.part_bytes(&part_path).unwrap().to_vec();
        assert!(package.replace_part(&part_path, replacement.clone()));
        assert_eq!(package.element_spans(&part_path), Some([].as_slice()));
        assert!(matches!(
            save_cell_edits(&package, &[edit]),
            Err(VsdxError::InvalidCellEdit { .. })
        ));
        assert_eq!(package.part_bytes(&part_path), Some(replacement.as_slice()));
    }

    #[test]
    fn rejects_an_aggregate_over_limit_before_mutating_the_package() {
        let source = include_bytes!("../tests/fixtures/foundation.vsdx");
        let package = parse_vsdx(source).unwrap();
        let part_path = package.page_part_paths.first().unwrap();
        let edit = first_value_edit(&package, part_path, "changed");
        let edits = vec![edit; MAX_PATCH_EDITS + 1];
        let before = package.part_bytes(part_path).unwrap().to_vec();
        assert!(matches!(
            save_cell_edits(&package, &edits),
            Err(VsdxError::PatchLimit { kind: "editCount" })
        ));
        assert_eq!(package.part_bytes(part_path), Some(before.as_slice()));

        let oversized = "x".repeat(MAX_PATCH_BYTES / 2 + 1);
        let edits = [
            first_value_edit(&package, part_path, &oversized),
            first_value_edit(&package, part_path, &oversized),
        ];
        assert!(matches!(
            save_cell_edits(&package, &edits),
            Err(VsdxError::PatchLimit { kind: "editBytes" })
        ));
        assert_eq!(package.part_bytes(part_path), Some(before.as_slice()));
    }

    #[test]
    fn failed_edits_after_a_valid_edit_leave_the_package_unchanged() {
        let (mut package, part_path) = package_with_page_xml(
            b"<PageContents><Shapes><Shape ID='1'><Cell N='Valid' V='old'/><Cell N='Invalid'/></Shape></Shapes></PageContents>",
        );
        let invalid_span = package
            .element_spans(&part_path)
            .unwrap()
            .iter()
            .find(|span| {
                span.name == "Cell"
                    && span.attributes.contains_key("N")
                    && !span.attributes.contains_key("V")
            })
            .unwrap()
            .span;
        package
            .parts
            .iter_mut()
            .find(|part| part.path == part_path)
            .unwrap()
            .spans
            .iter_mut()
            .find(|span| span.span == invalid_span)
            .unwrap()
            .attributes
            .clear();
        let valid = first_value_edit(&package, &part_path, "valid");
        let before = package.part_bytes(&part_path).unwrap().to_vec();
        let stale = CellEdit {
            cell_span: SourceSpan {
                offset: usize::MAX,
                length: 0,
            },
            ..valid.clone()
        };
        let invalid_insertion = CellEdit {
            part_path: part_path.clone(),
            cell_span: invalid_span,
            attribute: CellAttribute::Formula,
            value: "x".to_owned(),
        };
        let cases = vec![
            vec![valid.clone(), valid.clone()],
            vec![
                valid.clone(),
                CellEdit {
                    value: "x".repeat(MAX_PATCH_BYTES),
                    ..valid.clone()
                },
            ],
            vec![
                valid.clone(),
                CellEdit {
                    value: "bad\0".to_owned(),
                    ..valid.clone()
                },
            ],
            vec![valid.clone(), stale],
            vec![valid.clone(), invalid_insertion],
        ];
        for edits in cases {
            assert!(save_cell_edits(&package, &edits).is_err());
            assert_eq!(package.part_bytes(&part_path), Some(before.as_slice()));
        }
    }

    #[test]
    fn saves_xml_whitespace_character_references_exactly() {
        let source = include_bytes!("../tests/fixtures/foundation.vsdx");
        let package = parse_vsdx(source).unwrap();
        let part_path = package.page_part_paths.first().unwrap();
        let value = "a\tb\nc\rd";
        let saved =
            save_cell_edits(&package, &[first_value_edit(&package, part_path, value)]).unwrap();
        let reparsed = parse_vsdx(&saved).unwrap();
        assert!(
            reparsed
                .page_contents
                .values()
                .any(|sheet| sheet_has_value(sheet, value))
        );
    }

    #[test]
    fn rejects_incomplete_theme_parts() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let part = "visio/theme/theme1.xml";
        let root = parse_xml(
            br#"<a:theme xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><a:themeElements/></a:theme>"#,
            part,
            &mut budget,
        )
        .unwrap();
        assert!(matches!(
            parse_theme(&root, part),
            Err(VsdxError::MalformedXml { message, .. }) if message == "theme is missing themeElements/clrScheme"
        ));
    }

    #[test]
    fn parses_theme_shadow_and_variant_effects() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<a:theme xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:vt='http://schemas.microsoft.com/office/visio/2012/theme'><a:themeElements><a:fmtScheme><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="38100" dist="25420" dir="5400000"><a:srgbClr val="A5A5A5"><a:alpha val="60000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst/><a:sp3d/></a:effectStyle></a:effectStyleLst></a:fmtScheme><a:extLst><a:ext><vt:themeScheme><vt:schemeID schemeEnum="34"/></vt:themeScheme></a:ext><a:ext><vt:variationStyleSchemeLst><vt:variationStyleScheme embellishment="2"><vt:varStyle fillIdx="2" lineIdx="2" effectIdx="2" fontIdx="2"/></vt:variationStyleScheme></vt:variationStyleSchemeLst></a:ext><a:ext><vt:variationClrSchemeLst><vt:variationClrScheme><vt:varColor1><a:srgbClr val="268FEE"/></vt:varColor1></vt:variationClrScheme></vt:variationClrSchemeLst></a:ext></a:extLst></a:themeElements></a:theme>"#,
            "visio/theme/theme1.xml",
            &mut budget,
        )
        .unwrap();
        let effects = parse_theme_effects(&root);
        assert_eq!(effects.scheme_id, Some(34));
        assert_eq!(effects.effect_styles.len(), 3);
        assert_eq!(effects.effect_styles[0].outer_shadow, None);
        let shadow = effects.effect_styles[1].outer_shadow.as_ref().unwrap();
        assert_eq!(shadow.blur_emu, 38100);
        assert_eq!(shadow.dist_emu, 25420);
        assert_eq!(shadow.direction_60k, 5400000);
        assert_eq!(shadow.color, ThemeEffectColor::Srgb("A5A5A5".to_owned()));
        assert_eq!(shadow.alpha_1000pct, Some(60000));
        assert!(effects.effect_styles[2].has_bevel);
        assert_eq!(effects.variation_schemes.len(), 1);
        assert_eq!(effects.variation_schemes[0].embellishment, 2);
        assert_eq!(effects.variation_schemes[0].effect_indexes, vec![2]);
        assert_eq!(
            effects.variation_colors,
            vec![vec![
                Some("268FEE".to_owned()),
                None,
                None,
                None,
                None,
                None,
                None
            ]]
        );
    }

    #[test]
    fn maps_catalog_ids_through_child_relationships() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<Masters><Master ID='15'><Rel r:id='rId2'/></Master></Masters>"#,
            "visio/masters/masters.xml",
            &mut budget,
        )
        .unwrap();
        let relationships = [Relationship {
            id: "rId2".into(),
            relationship_type: relationship_types::MASTER.into(),
            target: "master2.xml".into(),
            target_mode: crate::TargetMode::Internal,
            resolved_target: Some("visio/masters/master2.xml".into()),
        }];

        assert_eq!(
            catalog_part_ids(
                Some(&root),
                "Master",
                "visio/masters/masters.xml",
                Some(&relationships)
            )
            .unwrap(),
            [("visio/masters/master2.xml".into(), 15)].into()
        );
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
    fn structural_edits_preserve_external_corpus_parts_when_available() {
        let Some(directory) = std::env::var_os("VSDX_CORPUS_DIR") else {
            eprintln!("SKIPPED REVIEW CORPUS TEST: VSDX_CORPUS_DIR is unset");
            return;
        };
        for file in ["lichtsysteme.vsdx", "soundplan.vsdx"] {
            let path = PathBuf::from(&directory).join(file);
            assert!(path.is_file(), "missing corpus file: {}", path.display());
            let source = std::fs::read(&path).unwrap();
            let package = parse_vsdx(&source).unwrap();
            let (page_path, page) = package
                .page_contents
                .iter()
                .find(|(_, page)| page.shapes().count() >= 2)
                .unwrap_or_else(|| panic!("{file}: no page with two shapes"));
            let page_id = package.page_part_ids[page_path];
            let shapes: Vec<_> = page.shapes().collect();
            let shape_ids: Vec<_> = shapes.iter().map(|shape| shape.id).collect();
            let shape_bytes = shape_bytes_by_id(&package, page_path);
            let operations = [
                (
                    "add",
                    StructuralEdit::AddShape {
                        page_id,
                        shape_xml: b"<Shape/>".to_vec(),
                    },
                ),
                (
                    "delete",
                    StructuralEdit::DeleteShape {
                        page_id,
                        shape_id: shapes[0].id,
                    },
                ),
                (
                    "reorder",
                    StructuralEdit::ReorderShape {
                        page_id,
                        shape_id: shapes[0].id,
                        before_shape_id: None,
                    },
                ),
            ];
            for (operation, edit) in operations {
                let saved = save_structural_edits(&package, &[edit])
                    .unwrap_or_else(|error| panic!("{file} {operation}: {error}"));
                let reparsed = parse_vsdx(&saved)
                    .unwrap_or_else(|error| panic!("{file} {operation}: {error}"));
                validate_structure(&reparsed)
                    .unwrap_or_else(|error| panic!("{file} {operation}: {error}"));
                let after_page = &reparsed.page_contents[page_path];
                let after_ids: Vec<_> = after_page.shapes().map(|shape| shape.id).collect();
                match operation {
                    "add" => {
                        assert_eq!(after_ids.len(), shape_ids.len() + 1, "{file} add");
                        assert_eq!(
                            after_ids.last(),
                            Some(&(shape_bytes.keys().max().unwrap() + 1))
                        );
                    }
                    "delete" => {
                        assert_eq!(after_ids.len(), shape_ids.len() - 1, "{file} delete");
                        assert!(!after_ids.contains(&shape_ids[0]), "{file} delete");
                        assert!(after_page.connects().all(|connect| {
                            connect.from_sheet != shape_ids[0] && connect.to_sheet != shape_ids[0]
                        }));
                    }
                    "reorder" => {
                        let mut expected = shape_ids[1..].to_vec();
                        expected.push(shape_ids[0]);
                        assert_eq!(after_ids, expected, "{file} reorder");
                    }
                    _ => unreachable!(),
                }
                let after_shape_bytes = shape_bytes_by_id(&reparsed, page_path);
                for (id, bytes) in &shape_bytes {
                    if operation != "delete" || *id != shape_ids[0] {
                        assert_eq!(
                            after_shape_bytes[id], *bytes,
                            "{file} {operation} shape {id}"
                        );
                    }
                }
                let before = unzip_parts(&source).unwrap();
                let after = unzip_parts(&saved).unwrap();
                for (part, bytes) in before {
                    if &part != page_path {
                        assert_eq!(
                            after
                                .iter()
                                .find(|(candidate, _)| candidate == &part)
                                .unwrap()
                                .1,
                            bytes,
                            "{file} {operation}: {part}"
                        );
                    }
                }
                eprintln!("CORPUS {file} {operation}: PASS");
            }
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
        // Fixture totals: 10 cells, 4 sections, 5 rows, 2 shapes.
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
    fn a_missing_companion_rels_file_yields_an_empty_relationship_set() {
        let package = rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/></Relationships>"#.to_vec()),
            ("visio/document.xml".to_owned(), b"<VisioDocument/>".to_vec()),
            ("visio/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/pages' Target='pages/pages.xml'/><Relationship Id='r2' Type='http://schemas.microsoft.com/visio/2010/relationships/masters' Target='masters/masters.xml'/></Relationships>"#.to_vec()),
            ("visio/pages/pages.xml".to_owned(), br#"<Pages><Page ID='1'/></Pages>"#.to_vec()),
            ("visio/pages/_rels/pages.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Type='http://schemas.microsoft.com/visio/2010/relationships/page' Target='page1.xml'/></Relationships>"#.to_vec()),
            ("visio/pages/page1.xml".to_owned(), b"<PageContents/>".to_vec()),
            // No `visio/masters/_rels/masters.xml.rels` companion part: a Masters part with no
            // masters legitimately has no relationships file at all.
            ("visio/masters/masters.xml".to_owned(), b"<Masters/>".to_vec()),
        ])
        .unwrap();
        let parsed = parse_vsdx(&package).unwrap();
        assert_eq!(parsed.master_part_paths, Vec::<String>::new());
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
            "application/vnd.ms-visio.template.macroEnabled.main+xml",
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
    fn opens_template_packages_and_round_trips_byte_for_byte() {
        let source = include_bytes!("../tests/fixtures/template.vstx");
        let package = parse_vsdx(source).unwrap();
        let content_types = package.part_bytes("[Content_Types].xml").unwrap();
        assert!(
            std::str::from_utf8(content_types)
                .unwrap()
                .contains("application/vnd.ms-visio.template.main+xml")
        );
        let written = write_vsdx(&package).unwrap();
        assert_eq!(unzip_parts(&written).unwrap(), unzip_parts(source).unwrap());
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
