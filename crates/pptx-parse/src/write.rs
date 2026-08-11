//! Edit-driven package writes. Parts the deck did not change are copied
//! through byte for byte; edited slides are patched at the XML level so
//! unmodeled markup survives.

use std::collections::{BTreeMap, HashMap, HashSet};

use ooxml_drawingml::{ColorValue, ShapeFill, ShapeOutline};

use crate::PptxError;
use crate::model::{Bullet, PptxPackage, RunProperties, SlideReference};
use crate::xml::{ParseBudget, ParseLimits, XmlElement, XmlNode, parse_xml, serialize_xml};

const DRAWINGML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PRESENTATIONML_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const OFFICE_RELATIONSHIPS_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SLIDE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_LAYOUT_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const MIN_SLIDE_ID: u32 = 256;

/// The desired final deck, expressed against the parsed source package.
pub struct DeckWrite {
    pub slides: Vec<SlideWrite>,
}

pub enum SlideWrite {
    /// Copy the source part through untouched.
    Keep { part_path: String },
    /// Patch the source part; `shapes` is the full top-level list in final order.
    Patch {
        part_path: String,
        shapes: Vec<ShapeWrite>,
    },
    /// Mint a new slide part.
    Add {
        name: Option<String>,
        layout_part_path: Option<String>,
        shapes: Vec<ShapeAdd>,
    },
}

pub enum ShapeWrite {
    /// Keep the source shape element untouched. `source_index` counts shape
    /// elements at the same nesting level, in document order.
    Keep {
        source_index: usize,
    },
    Patch {
        source_index: usize,
        patch: Box<ShapePatch>,
    },
    Add(Box<ShapeAdd>),
}

#[derive(Default)]
pub struct ShapePatch {
    pub offset: Option<(i64, i64)>,
    pub extent: Option<(i64, i64)>,
    pub fill: Option<ShapeFill>,
    /// A default outline clears the stroke.
    pub outline: Option<ShapeOutline>,
    pub adjust_values: Option<BTreeMap<String, f64>>,
    pub texts: Vec<TextWrite>,
    /// Non-empty only when a group's child list must be rebuilt.
    pub children: Vec<ShapeWrite>,
}

pub struct TextWrite {
    pub target: TextTarget,
    pub paragraphs: Vec<ParagraphWrite>,
}

pub enum TextTarget {
    Body,
    TableCell { row: usize, cell: usize },
}

pub struct ParagraphWrite {
    /// Index of the paragraph in the source text body, when it survives.
    pub source_index: Option<usize>,
    /// `false` keeps the source paragraph verbatim.
    pub rebuild: bool,
    pub properties_changed: bool,
    pub alignment: Option<String>,
    pub level: u32,
    pub bullet: Option<Bullet>,
    pub runs: Vec<RunWrite>,
}

pub struct RunWrite {
    pub text: String,
    pub properties: RunProperties,
}

pub struct ShapeAdd {
    pub name: String,
    pub geometry: String,
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
    pub adjust_values: BTreeMap<String, f64>,
    pub fill: Option<ShapeFill>,
    pub outline: Option<ShapeOutline>,
    pub paragraphs: Option<Vec<ParagraphWrite>>,
}

/// Writes the package with `deck` applied. Slides marked [`SlideWrite::Keep`]
/// and every non-slide part keep their exact source bytes.
pub fn write_pptx_with_edits(
    package: &PptxPackage,
    deck: &DeckWrite,
) -> Result<Vec<u8>, PptxError> {
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);

    let final_paths: HashSet<&str> = deck
        .slides
        .iter()
        .filter_map(|slide| match slide {
            SlideWrite::Keep { part_path } | SlideWrite::Patch { part_path, .. } => {
                Some(part_path.as_str())
            }
            SlideWrite::Add { .. } => None,
        })
        .collect();
    let removed: Vec<&SlideReference> = package
        .presentation
        .slides
        .iter()
        .filter(|reference| !final_paths.contains(reference.part_path.as_str()))
        .collect();

    let mut replacements: HashMap<String, Vec<u8>> = HashMap::new();
    let mut new_parts: Vec<(String, Vec<u8>)> = Vec::new();

    for slide in &deck.slides {
        if let SlideWrite::Patch { part_path, shapes } = slide {
            let bytes = package
                .part_bytes(part_path)
                .ok_or_else(|| PptxError::MissingPart(part_path.clone()))?;
            let mut root = parse_xml(bytes, part_path, &mut budget)?;
            patch_slide(&mut root, shapes, part_path)?;
            replacements.insert(part_path.clone(), serialize_xml(&root));
        }
    }

    let mut minted = Vec::new();
    let mut next_slide_number = next_slide_number(package);
    let mut next_slide_id = next_slide_id(package);
    let mut next_relationship = next_relationship_number(package);
    let mut structural = false;
    let mut final_order: Vec<FinalSlide<'_>> = Vec::new();
    for slide in &deck.slides {
        match slide {
            SlideWrite::Keep { part_path } | SlideWrite::Patch { part_path, .. } => {
                let reference = package
                    .presentation
                    .slides
                    .iter()
                    .find(|reference| &reference.part_path == part_path)
                    .ok_or_else(|| write_error(part_path, "not a slide of this deck"))?;
                final_order.push(FinalSlide::Existing(reference));
            }
            SlideWrite::Add {
                name,
                layout_part_path,
                shapes,
            } => {
                structural = true;
                let part_path = format!("ppt/slides/slide{next_slide_number}.xml");
                next_slide_number += 1;
                let relationship_id = format!("rId{next_relationship}");
                next_relationship += 1;
                let slide_id = next_slide_id;
                next_slide_id += 1;
                let layout = layout_part_path.clone().or_else(|| {
                    package
                        .layouts
                        .first()
                        .map(|layout| layout.part_path.clone())
                });
                new_parts.push((
                    part_path.clone(),
                    slide_xml(name.as_deref(), shapes, &part_path)?,
                ));
                if let Some(layout) = &layout {
                    new_parts.push((
                        slide_relationships_path(&part_path),
                        slide_relationships_xml(&part_path, layout),
                    ));
                }
                minted.push(MintedSlide {
                    part_path,
                    relationship_id,
                    slide_id,
                });
                final_order.push(FinalSlide::Added(minted.len() - 1));
            }
        }
    }
    let original_paths: Vec<&str> = package
        .presentation
        .slides
        .iter()
        .map(|reference| reference.part_path.as_str())
        .collect();
    let kept_paths: Vec<&str> = final_order
        .iter()
        .filter_map(|slide| match slide {
            FinalSlide::Existing(reference) => Some(reference.part_path.as_str()),
            FinalSlide::Added(_) => None,
        })
        .collect();
    structural = structural || kept_paths != original_paths;

    if structural {
        patch_structure(
            package,
            &final_order,
            &minted,
            &removed,
            &mut replacements,
            &mut budget,
        )?;
    }

    let mut removed_paths = HashSet::new();
    for reference in &removed {
        removed_paths.insert(reference.part_path.clone());
        removed_paths.insert(slide_relationships_path(&reference.part_path));
    }
    let mut parts = Vec::with_capacity(package.parts.len() + new_parts.len());
    for part in &package.parts {
        if removed_paths.contains(&part.path) {
            continue;
        }
        match replacements.remove(&part.path) {
            Some(bytes) => parts.push((part.path.clone(), bytes)),
            None => parts.push((part.path.clone(), part.bytes.clone())),
        }
    }
    parts.extend(new_parts);
    ooxml_opc::rezip_parts(&parts).map_err(PptxError::Container)
}

struct MintedSlide {
    part_path: String,
    relationship_id: String,
    slide_id: u32,
}

enum FinalSlide<'a> {
    Existing(&'a SlideReference),
    Added(usize),
}

fn write_error(part: &str, message: impl Into<String>) -> PptxError {
    PptxError::Write {
        part: part.to_owned(),
        message: message.into(),
    }
}

fn next_slide_number(package: &PptxPackage) -> usize {
    let paths: HashSet<&str> = package
        .parts
        .iter()
        .map(|part| part.path.as_str())
        .collect();
    let mut number = package
        .parts
        .iter()
        .filter_map(|part| {
            part.path
                .strip_prefix("ppt/slides/slide")?
                .strip_suffix(".xml")?
                .parse::<usize>()
                .ok()
        })
        .max()
        .unwrap_or(0)
        + 1;
    while paths.contains(format!("ppt/slides/slide{number}.xml").as_str())
        || paths.contains(format!("ppt/slides/_rels/slide{number}.xml.rels").as_str())
    {
        number += 1;
    }
    number
}

fn next_slide_id(package: &PptxPackage) -> u32 {
    package
        .presentation
        .slides
        .iter()
        .map(|reference| reference.id)
        .max()
        .unwrap_or(0)
        .max(MIN_SLIDE_ID - 1)
        + 1
}

fn next_relationship_number(package: &PptxPackage) -> usize {
    package
        .relationships
        .get(&package.presentation.part_path)
        .into_iter()
        .flatten()
        .filter_map(|relationship| relationship.id.strip_prefix("rId")?.parse::<usize>().ok())
        .max()
        .unwrap_or(0)
        + 1
}

fn slide_relationships_path(part_path: &str) -> String {
    match part_path.rsplit_once('/') {
        Some((directory, name)) => format!("{directory}/_rels/{name}.rels"),
        None => format!("_rels/{part_path}.rels"),
    }
}

fn relative_target(from_part: &str, to_part: &str) -> String {
    let from: Vec<&str> = from_part.split('/').collect();
    let to: Vec<&str> = to_part.split('/').collect();
    let from_directory = &from[..from.len().saturating_sub(1)];
    let to_directory = &to[..to.len().saturating_sub(1)];
    let shared = from_directory
        .iter()
        .zip(to_directory)
        .take_while(|(a, b)| a == b)
        .count();
    let mut segments: Vec<String> = vec!["..".to_owned(); from_directory.len() - shared];
    segments.extend(to[shared..].iter().map(|segment| (*segment).to_owned()));
    segments.join("/")
}

// --- structural parts -------------------------------------------------------

fn patch_structure(
    package: &PptxPackage,
    final_order: &[FinalSlide<'_>],
    minted: &[MintedSlide],
    removed: &[&SlideReference],
    replacements: &mut HashMap<String, Vec<u8>>,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    let presentation_path = package.presentation.part_path.as_str();
    let bytes = package
        .part_bytes(presentation_path)
        .ok_or_else(|| PptxError::MissingPart(presentation_path.to_owned()))?;
    let mut root = parse_xml(bytes, presentation_path, budget)?;
    patch_slide_id_list(&mut root, final_order, minted, presentation_path)?;
    replacements.insert(presentation_path.to_owned(), serialize_xml(&root));

    let relationships_path = slide_relationships_path(presentation_path);
    let bytes = package
        .part_bytes(&relationships_path)
        .ok_or_else(|| PptxError::MissingPart(relationships_path.clone()))?;
    let mut root = parse_xml(bytes, &relationships_path, budget)?;
    let removed_ids: HashSet<&str> = removed
        .iter()
        .map(|reference| reference.relationship_id.as_str())
        .collect();
    root.children.retain(|child| match child {
        XmlNode::Element(element) if element.local_name() == "Relationship" => element
            .attribute("Id")
            .is_none_or(|id| !removed_ids.contains(id)),
        _ => true,
    });
    for slide in minted {
        root.children.push(XmlNode::Element(
            XmlElement::new("Relationship")
                .with_attribute("Id", slide.relationship_id.clone())
                .with_attribute("Type", SLIDE_RELATIONSHIP_TYPE)
                .with_attribute(
                    "Target",
                    relative_target(presentation_path, &slide.part_path),
                ),
        ));
    }
    replacements.insert(relationships_path, serialize_xml(&root));

    let content_types_path = "[Content_Types].xml";
    let bytes = package
        .part_bytes(content_types_path)
        .ok_or_else(|| PptxError::MissingPart(content_types_path.to_owned()))?;
    let mut root = parse_xml(bytes, content_types_path, budget)?;
    let removed_names: HashSet<String> = removed
        .iter()
        .map(|reference| format!("/{}", reference.part_path))
        .collect();
    root.children.retain(|child| match child {
        XmlNode::Element(element) if element.local_name() == "Override" => element
            .attribute("PartName")
            .is_none_or(|name| !removed_names.contains(name)),
        _ => true,
    });
    for slide in minted {
        root.children.push(XmlNode::Element(
            XmlElement::new("Override")
                .with_attribute("PartName", format!("/{}", slide.part_path))
                .with_attribute("ContentType", SLIDE_CONTENT_TYPE),
        ));
    }
    replacements.insert(content_types_path.to_owned(), serialize_xml(&root));
    Ok(())
}

fn patch_slide_id_list(
    root: &mut XmlElement,
    final_order: &[FinalSlide<'_>],
    minted: &[MintedSlide],
    part: &str,
) -> Result<(), PptxError> {
    let presentation_prefix = resolve_prefix(root, PRESENTATIONML_NS, "p");
    let relationship_prefix = resolve_prefix(root, OFFICE_RELATIONSHIPS_NS, "r");
    let list_name = qualified(&presentation_prefix, "sldIdLst");
    if root.child_mut("sldIdLst").is_none() {
        let position = root
            .children
            .iter()
            .position(|child| {
                matches!(child, XmlNode::Element(element) if element.local_name() == "sldMasterIdLst")
            })
            .map(|index| index + 1)
            .unwrap_or(0);
        root.children.insert(
            position,
            XmlNode::Element(XmlElement::new(list_name.clone())),
        );
    }
    let relationship_attribute = qualified(&relationship_prefix, "id");
    let list = root
        .child_mut("sldIdLst")
        .ok_or_else(|| write_error(part, "missing slide id list"))?;
    let mut existing: HashMap<String, XmlElement> = HashMap::new();
    for child in std::mem::take(&mut list.children) {
        if let XmlNode::Element(element) = child
            && element.local_name() == "sldId"
        {
            let relationship_id = element.attribute(&relationship_attribute).or_else(|| {
                element
                    .attributes
                    .iter()
                    .find(|(key, _)| key.ends_with(":id"))
                    .map(|(_, value)| value.as_str())
            });
            if let Some(id) = relationship_id {
                existing.insert(id.to_owned(), element);
            }
        }
    }
    let entry_name = qualified(&presentation_prefix, "sldId");
    for slide in final_order {
        let element = match slide {
            FinalSlide::Existing(reference) => existing
                .remove(&reference.relationship_id)
                .unwrap_or_else(|| {
                    XmlElement::new(entry_name.clone())
                        .with_attribute("id", reference.id.to_string())
                        .with_attribute(
                            relationship_attribute.clone(),
                            reference.relationship_id.clone(),
                        )
                }),
            FinalSlide::Added(index) => {
                let slide = &minted[*index];
                XmlElement::new(entry_name.clone())
                    .with_attribute("id", slide.slide_id.to_string())
                    .with_attribute(
                        relationship_attribute.clone(),
                        slide.relationship_id.clone(),
                    )
            }
        };
        list.children.push(XmlNode::Element(element));
    }
    Ok(())
}

// --- namespace prefixes -----------------------------------------------------

fn find_prefix(root: &XmlElement, namespace: &str) -> Option<String> {
    root.attributes.iter().find_map(|(key, value)| {
        if value != namespace {
            return None;
        }
        if key == "xmlns" {
            Some(String::new())
        } else {
            key.strip_prefix("xmlns:").map(str::to_owned)
        }
    })
}

fn resolve_prefix(root: &mut XmlElement, namespace: &str, fallback: &str) -> String {
    if let Some(prefix) = find_prefix(root, namespace) {
        return prefix;
    }
    root.set_attribute(format!("xmlns:{fallback}"), namespace);
    fallback.to_owned()
}

fn qualified(prefix: &str, local: &str) -> String {
    if prefix.is_empty() {
        local.to_owned()
    } else {
        format!("{prefix}:{local}")
    }
}

struct Prefixes {
    drawing: String,
    presentation: String,
}

impl Prefixes {
    fn from_root(root: &mut XmlElement) -> Self {
        Self {
            drawing: resolve_prefix(root, DRAWINGML_NS, "a"),
            presentation: resolve_prefix(root, PRESENTATIONML_NS, "p"),
        }
    }

    fn drawing(&self, local: &str) -> String {
        qualified(&self.drawing, local)
    }

    fn presentation(&self, local: &str) -> String {
        qualified(&self.presentation, local)
    }
}

// --- slide patching ---------------------------------------------------------

fn is_shape_element(local: &str) -> bool {
    matches!(local, "sp" | "pic" | "graphicFrame" | "grpSp")
}

fn patch_slide(root: &mut XmlElement, shapes: &[ShapeWrite], part: &str) -> Result<(), PptxError> {
    let prefixes = Prefixes::from_root(root);
    let mut next_shape_id = max_shape_id(root) + 1;
    let tree = root
        .child_mut("cSld")
        .and_then(|common| common.child_mut("spTree"))
        .ok_or_else(|| write_error(part, "slide has no shape tree"))?;
    patch_shape_children(tree, shapes, &mut next_shape_id, &prefixes, part)
}

fn max_shape_id(root: &XmlElement) -> u32 {
    root.descendants_named("cNvPr")
        .iter()
        .filter_map(|element| element.attribute("id")?.parse::<u32>().ok())
        .max()
        .unwrap_or(1)
}

fn patch_shape_children(
    parent: &mut XmlElement,
    writes: &[ShapeWrite],
    next_shape_id: &mut u32,
    prefixes: &Prefixes,
    part: &str,
) -> Result<(), PptxError> {
    let mut leading = Vec::new();
    let mut originals: Vec<Option<XmlElement>> = Vec::new();
    let mut trailing = Vec::new();
    for child in std::mem::take(&mut parent.children) {
        match child {
            XmlNode::Element(element) if is_shape_element(element.local_name()) => {
                originals.push(Some(element));
            }
            XmlNode::Text(text) if text.trim().is_empty() => {}
            other => {
                if originals.is_empty() {
                    leading.push(other);
                } else {
                    trailing.push(other);
                }
            }
        }
    }
    let mut children = leading;
    for write in writes {
        let node = match write {
            ShapeWrite::Keep { source_index } => {
                take_original(&mut originals, *source_index, part)?
            }
            ShapeWrite::Patch {
                source_index,
                patch,
            } => {
                let mut element = take_original(&mut originals, *source_index, part)?;
                patch_shape(&mut element, patch, next_shape_id, prefixes, part)?;
                element
            }
            ShapeWrite::Add(add) => shape_element(add, next_shape_id, prefixes, part)?,
        };
        children.push(XmlNode::Element(node));
    }
    children.extend(trailing);
    parent.children = children;
    Ok(())
}

fn take_original(
    originals: &mut [Option<XmlElement>],
    index: usize,
    part: &str,
) -> Result<XmlElement, PptxError> {
    originals
        .get_mut(index)
        .and_then(Option::take)
        .ok_or_else(|| write_error(part, format!("shape index {index} is not available")))
}

fn patch_shape(
    element: &mut XmlElement,
    patch: &ShapePatch,
    next_shape_id: &mut u32,
    prefixes: &Prefixes,
    part: &str,
) -> Result<(), PptxError> {
    if patch.offset.is_some() || patch.extent.is_some() {
        patch_transform(element, patch.offset, patch.extent, prefixes, part)?;
    }
    if let Some(fill) = &patch.fill {
        let properties = shape_properties_mut(element, part)?;
        set_fill(properties, fill, prefixes, part)?;
    }
    if let Some(outline) = &patch.outline {
        let properties = shape_properties_mut(element, part)?;
        set_outline(properties, outline, prefixes);
    }
    if let Some(adjust_values) = &patch.adjust_values {
        let properties = shape_properties_mut(element, part)?;
        set_adjust_values(properties, adjust_values, prefixes);
    }
    for text in &patch.texts {
        patch_text(element, text, prefixes, part)?;
    }
    if !patch.children.is_empty() {
        patch_shape_children(element, &patch.children, next_shape_id, prefixes, part)?;
    }
    Ok(())
}

fn shape_properties_mut<'a>(
    element: &'a mut XmlElement,
    part: &str,
) -> Result<&'a mut XmlElement, PptxError> {
    let container = match element.local_name() {
        "grpSp" => "grpSpPr",
        _ => "spPr",
    };
    element
        .child_mut(container)
        .ok_or_else(|| write_error(part, "shape has no properties element"))
}

fn patch_transform(
    element: &mut XmlElement,
    offset: Option<(i64, i64)>,
    extent: Option<(i64, i64)>,
    prefixes: &Prefixes,
    part: &str,
) -> Result<(), PptxError> {
    let (container, transform_name) = match element.local_name() {
        "graphicFrame" => (None, prefixes.presentation("xfrm")),
        "grpSp" => (Some("grpSpPr"), prefixes.drawing("xfrm")),
        _ => (Some("spPr"), prefixes.drawing("xfrm")),
    };
    let parent = match container {
        Some(name) => element
            .child_mut(name)
            .ok_or_else(|| write_error(part, "shape has no properties element"))?,
        None => element,
    };
    if parent.child_mut("xfrm").is_none() {
        let position = parent
            .children
            .iter()
            .position(|child| {
                matches!(child, XmlNode::Element(element) if element.local_name() != "nvGraphicFramePr")
            })
            .unwrap_or(parent.children.len());
        parent
            .children
            .insert(position, XmlNode::Element(XmlElement::new(transform_name)));
    }
    let transform = parent.child_mut("xfrm").expect("transform ensured above");
    if transform.child_mut("off").is_none() {
        transform.children.insert(
            0,
            XmlNode::Element(
                XmlElement::new(prefixes.drawing("off"))
                    .with_attribute("x", "0")
                    .with_attribute("y", "0"),
            ),
        );
    }
    if transform.child_mut("ext").is_none() {
        let position = transform
            .children
            .iter()
            .position(
                |child| matches!(child, XmlNode::Element(element) if element.local_name() == "off"),
            )
            .map(|index| index + 1)
            .unwrap_or(transform.children.len());
        transform.children.insert(
            position,
            XmlNode::Element(
                XmlElement::new(prefixes.drawing("ext"))
                    .with_attribute("cx", "0")
                    .with_attribute("cy", "0"),
            ),
        );
    }
    if let Some((x, y)) = offset {
        let element = transform.child_mut("off").expect("offset ensured above");
        element.set_attribute("x", x.to_string());
        element.set_attribute("y", y.to_string());
    }
    if let Some((width, height)) = extent {
        let element = transform.child_mut("ext").expect("extent ensured above");
        element.set_attribute("cx", width.to_string());
        element.set_attribute("cy", height.to_string());
    }
    Ok(())
}

const FILL_ELEMENTS: [&str; 6] = [
    "noFill",
    "solidFill",
    "gradFill",
    "blipFill",
    "pattFill",
    "grpFill",
];
const POST_FILL_ELEMENTS: [&str; 6] = ["ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"];

fn set_fill(
    properties: &mut XmlElement,
    fill: &ShapeFill,
    prefixes: &Prefixes,
    part: &str,
) -> Result<(), PptxError> {
    let element = fill_element(fill, prefixes, part)?;
    let existing = properties.children.iter().position(|child| {
        matches!(child, XmlNode::Element(element) if FILL_ELEMENTS.contains(&element.local_name()))
    });
    match existing {
        Some(index) => properties.children[index] = XmlNode::Element(element),
        None => {
            let position = properties
                .children
                .iter()
                .position(|child| {
                    matches!(
                        child,
                        XmlNode::Element(element)
                            if POST_FILL_ELEMENTS.contains(&element.local_name())
                    )
                })
                .unwrap_or(properties.children.len());
            properties
                .children
                .insert(position, XmlNode::Element(element));
        }
    }
    Ok(())
}

fn fill_element(
    fill: &ShapeFill,
    prefixes: &Prefixes,
    part: &str,
) -> Result<XmlElement, PptxError> {
    match fill.fill_type.as_str() {
        "none" => Ok(XmlElement::new(prefixes.drawing("noFill"))),
        "solid" => {
            let color = fill
                .color
                .as_ref()
                .and_then(|color| color_element(color, prefixes))
                .ok_or_else(|| write_error(part, "solid fill without a serializable color"))?;
            Ok(XmlElement::new(prefixes.drawing("solidFill")).with_child(color))
        }
        "gradient" => {
            let gradient = fill
                .gradient
                .as_ref()
                .ok_or_else(|| write_error(part, "gradient fill without stops"))?;
            let mut stops = XmlElement::new(prefixes.drawing("gsLst"));
            for stop in &gradient.stops {
                let color = color_element(&stop.color, prefixes)
                    .ok_or_else(|| write_error(part, "gradient stop without a color"))?;
                stops = stops.with_child(
                    XmlElement::new(prefixes.drawing("gs"))
                        .with_attribute("pos", format_fixed(stop.position))
                        .with_child(color),
                );
            }
            let mut element = XmlElement::new(prefixes.drawing("gradFill")).with_child(stops);
            match gradient.gradient_type.as_str() {
                "linear" => {
                    element = element.with_child(
                        XmlElement::new(prefixes.drawing("lin"))
                            .with_attribute(
                                "ang",
                                format_fixed(gradient.angle.unwrap_or_default() * 60_000.0),
                            )
                            .with_attribute("scaled", "1"),
                    );
                }
                kind => {
                    let path = match kind {
                        "radial" => "circle",
                        "rectangular" => "rect",
                        _ => "shape",
                    };
                    element = element.with_child(
                        XmlElement::new(prefixes.drawing("path")).with_attribute("path", path),
                    );
                }
            }
            Ok(element)
        }
        other => Err(write_error(
            part,
            format!("unsupported fill type {other:?}"),
        )),
    }
}

fn color_element(color: &ColorValue, prefixes: &Prefixes) -> Option<XmlElement> {
    let mut element = if let Some(rgb) = &color.rgb {
        XmlElement::new(prefixes.drawing("srgbClr")).with_attribute("val", rgb.clone())
    } else if let Some(theme) = &color.theme_color {
        XmlElement::new(prefixes.drawing("schemeClr"))
            .with_attribute("val", denormalize_scheme_color(theme))
    } else {
        return None;
    };
    let fractions = [
        ("alpha", color.alpha),
        ("lumMod", color.luminance_modulation),
        ("lumOff", color.luminance_offset),
        ("satMod", color.saturation_modulation),
    ];
    for (name, value) in fractions {
        if let Some(value) = value {
            element = element.with_child(
                XmlElement::new(prefixes.drawing(name))
                    .with_attribute("val", format_fixed(value * 100_000.0)),
            );
        }
    }
    let modifiers = [("tint", &color.theme_tint), ("shade", &color.theme_shade)];
    for (name, value) in modifiers {
        if let Some(byte) = value
            .as_deref()
            .and_then(|value| u8::from_str_radix(value, 16).ok())
        {
            element = element.with_child(
                XmlElement::new(prefixes.drawing(name))
                    .with_attribute("val", format_fixed(f64::from(byte) / 255.0 * 100_000.0)),
            );
        }
    }
    Some(element)
}

fn denormalize_scheme_color(value: &str) -> String {
    match value {
        "text1" => "tx1",
        "text2" => "tx2",
        "background1" => "bg1",
        "background2" => "bg2",
        value => value,
    }
    .to_owned()
}

fn format_fixed(value: f64) -> String {
    (value.round() as i64).to_string()
}

fn set_outline(properties: &mut XmlElement, outline: &ShapeOutline, prefixes: &Prefixes) {
    let element = outline_element(outline, prefixes);
    let existing = properties.children.iter().position(
        |child| matches!(child, XmlNode::Element(element) if element.local_name() == "ln"),
    );
    match existing {
        Some(index) => properties.children[index] = XmlNode::Element(element),
        None => {
            let position = properties
                .children
                .iter()
                .position(|child| {
                    matches!(
                        child,
                        XmlNode::Element(element)
                            if POST_FILL_ELEMENTS[1..].contains(&element.local_name())
                    )
                })
                .unwrap_or(properties.children.len());
            properties
                .children
                .insert(position, XmlNode::Element(element));
        }
    }
}

fn outline_element(outline: &ShapeOutline, prefixes: &Prefixes) -> XmlElement {
    let mut element = XmlElement::new(prefixes.drawing("ln"));
    if *outline == ShapeOutline::default() {
        return element.with_child(XmlElement::new(prefixes.drawing("noFill")));
    }
    if let Some(width) = outline.width {
        element.set_attribute("w", format_fixed(width));
    }
    if let Some(cap) = &outline.cap {
        element.set_attribute("cap", cap.clone());
    }
    if let Some(color) = outline
        .color
        .as_ref()
        .and_then(|color| color_element(color, prefixes))
    {
        element =
            element.with_child(XmlElement::new(prefixes.drawing("solidFill")).with_child(color));
    }
    if let Some(style) = &outline.style {
        element = element.with_child(
            XmlElement::new(prefixes.drawing("prstDash")).with_attribute("val", style.clone()),
        );
    }
    if let Some(join) = &outline.join {
        element = element.with_child(XmlElement::new(prefixes.drawing(join)));
    }
    let ends = [
        ("headEnd", &outline.head_end),
        ("tailEnd", &outline.tail_end),
    ];
    for (name, end) in ends {
        if let Some(end) = end {
            let mut end_element = XmlElement::new(prefixes.drawing(name))
                .with_attribute("type", end.end_type.clone());
            if let Some(width) = &end.width {
                end_element.set_attribute("w", width.clone());
            }
            if let Some(length) = &end.length {
                end_element.set_attribute("len", length.clone());
            }
            element = element.with_child(end_element);
        }
    }
    element
}

fn set_adjust_values(
    properties: &mut XmlElement,
    adjust_values: &BTreeMap<String, f64>,
    prefixes: &Prefixes,
) {
    // Only preset geometries carry an editable adjustment list.
    let Some(geometry) = properties.child_mut("prstGeom") else {
        return;
    };
    let list_name = prefixes.drawing("avLst");
    if geometry.child_mut("avLst").is_none() {
        geometry
            .children
            .insert(0, XmlNode::Element(XmlElement::new(list_name)));
    }
    let list = geometry.child_mut("avLst").expect("list ensured above");
    list.children.clear();
    for (name, value) in adjust_values {
        list.children.push(XmlNode::Element(
            XmlElement::new(prefixes.drawing("gd"))
                .with_attribute("name", name.clone())
                .with_attribute("fmla", format!("val {}", format_fixed(value * 100_000.0))),
        ));
    }
}

// --- text -------------------------------------------------------------------

fn patch_text(
    element: &mut XmlElement,
    text: &TextWrite,
    prefixes: &Prefixes,
    part: &str,
) -> Result<(), PptxError> {
    let body = match &text.target {
        TextTarget::Body => element.child_mut("txBody"),
        TextTarget::TableCell { row, cell } => element
            .child_mut("graphic")
            .and_then(|graphic| graphic.child_mut("graphicData"))
            .and_then(|data| data.child_mut("tbl"))
            .and_then(|table| nth_child_mut(table, "tr", *row))
            .and_then(|table_row| nth_child_mut(table_row, "tc", *cell))
            .and_then(|table_cell| table_cell.child_mut("txBody")),
    };
    let body = body.ok_or_else(|| write_error(part, "text target has no body"))?;
    rebuild_paragraphs(body, &text.paragraphs, prefixes);
    Ok(())
}

fn nth_child_mut<'a>(
    parent: &'a mut XmlElement,
    local: &str,
    index: usize,
) -> Option<&'a mut XmlElement> {
    parent
        .children
        .iter_mut()
        .filter_map(|child| match child {
            XmlNode::Element(element) if element.local_name() == local => Some(element),
            _ => None,
        })
        .nth(index)
}

fn rebuild_paragraphs(body: &mut XmlElement, paragraphs: &[ParagraphWrite], prefixes: &Prefixes) {
    let mut preamble = Vec::new();
    let mut originals: Vec<Option<XmlElement>> = Vec::new();
    for child in std::mem::take(&mut body.children) {
        match child {
            XmlNode::Element(element) if element.local_name() == "p" => {
                originals.push(Some(element));
            }
            XmlNode::Text(text) if text.trim().is_empty() => {}
            other => preamble.push(other),
        }
    }
    let mut children = preamble;
    for paragraph in paragraphs {
        let source = paragraph
            .source_index
            .and_then(|index| originals.get_mut(index).and_then(Option::take));
        let element = match source {
            Some(element) if !paragraph.rebuild => element,
            source => build_paragraph(paragraph, source, prefixes),
        };
        children.push(XmlNode::Element(element));
    }
    if children
        .iter()
        .all(|child| !matches!(child, XmlNode::Element(element) if element.local_name() == "p"))
    {
        children.push(XmlNode::Element(XmlElement::new(prefixes.drawing("p"))));
    }
    body.children = children;
}

fn build_paragraph(
    write: &ParagraphWrite,
    source: Option<XmlElement>,
    prefixes: &Prefixes,
) -> XmlElement {
    let mut paragraph = match source {
        Some(mut element) => {
            element.children.retain(|child| {
                !matches!(
                    child,
                    XmlNode::Element(element)
                        if matches!(element.local_name(), "r" | "br" | "fld")
                ) && !matches!(child, XmlNode::Text(text) if text.trim().is_empty())
            });
            element
        }
        None => XmlElement::new(prefixes.drawing("p")),
    };
    let fresh = write.source_index.is_none();
    if write.properties_changed || fresh {
        apply_paragraph_properties(&mut paragraph, write, prefixes);
    }
    let mut runs = Vec::new();
    for run in &write.runs {
        runs.extend(run_nodes(run, prefixes));
    }
    let position = paragraph
        .children
        .iter()
        .position(|child| {
            matches!(child, XmlNode::Element(element) if element.local_name() == "endParaRPr")
        })
        .unwrap_or(paragraph.children.len());
    paragraph.children.splice(position..position, runs);
    paragraph
}

fn apply_paragraph_properties(
    paragraph: &mut XmlElement,
    write: &ParagraphWrite,
    prefixes: &Prefixes,
) {
    let needs_properties = write.alignment.is_some() || write.level > 0 || write.bullet.is_some();
    if paragraph.child_mut("pPr").is_none() {
        if !needs_properties {
            return;
        }
        paragraph.children.insert(
            0,
            XmlNode::Element(XmlElement::new(prefixes.drawing("pPr"))),
        );
    }
    let properties = paragraph.child_mut("pPr").expect("ensured above");
    match &write.alignment {
        Some(alignment) => properties.set_attribute("algn", alignment.clone()),
        None => {
            properties.attributes.remove("algn");
        }
    }
    if write.level > 0 {
        properties.set_attribute("lvl", write.level.to_string());
    } else {
        properties.attributes.remove("lvl");
    }
    properties.children.retain(|child| {
        !matches!(
            child,
            XmlNode::Element(element)
                if matches!(element.local_name(), "buNone" | "buChar" | "buAutoNum")
        )
    });
    let bullet = match &write.bullet {
        Some(Bullet::Character { value }) => {
            Some(XmlElement::new(prefixes.drawing("buChar")).with_attribute("char", value.clone()))
        }
        Some(Bullet::AutoNumber { scheme, start_at }) => {
            let mut element = XmlElement::new(prefixes.drawing("buAutoNum"))
                .with_attribute("type", scheme.clone());
            if *start_at != 1 {
                element.set_attribute("startAt", start_at.to_string());
            }
            Some(element)
        }
        Some(Bullet::None) => Some(XmlElement::new(prefixes.drawing("buNone"))),
        None => None,
    };
    if let Some(bullet) = bullet {
        let position = properties
            .children
            .iter()
            .position(|child| {
                matches!(
                    child,
                    XmlNode::Element(element)
                        if matches!(element.local_name(), "tabLst" | "defRPr" | "extLst")
                )
            })
            .unwrap_or(properties.children.len());
        properties
            .children
            .insert(position, XmlNode::Element(bullet));
    }
}

fn run_nodes(run: &RunWrite, prefixes: &Prefixes) -> Vec<XmlNode> {
    let mut nodes = Vec::new();
    for (index, segment) in run.text.split('\n').enumerate() {
        if index > 0 {
            let mut line_break = XmlElement::new(prefixes.drawing("br"));
            if let Some(properties) = run_properties_element(&run.properties, prefixes) {
                line_break = line_break.with_child(properties);
            }
            nodes.push(XmlNode::Element(line_break));
        }
        if segment.is_empty() {
            continue;
        }
        let mut element = XmlElement::new(prefixes.drawing("r"));
        if let Some(properties) = run_properties_element(&run.properties, prefixes) {
            element = element.with_child(properties);
        }
        element = element.with_child(XmlElement::new(prefixes.drawing("t")).with_text(segment));
        nodes.push(XmlNode::Element(element));
    }
    nodes
}

fn run_properties_element(properties: &RunProperties, prefixes: &Prefixes) -> Option<XmlElement> {
    let mut element = XmlElement::new(prefixes.drawing("rPr"));
    let mut present = false;
    if let Some(language) = &properties.language {
        element.set_attribute("lang", language.clone());
        present = true;
    }
    if let Some(size) = properties.font_size_pt {
        element.set_attribute("sz", format_fixed(size * 100.0));
        present = true;
    }
    if let Some(bold) = properties.bold {
        element.set_attribute("b", if bold { "1" } else { "0" });
        present = true;
    }
    if let Some(italic) = properties.italic {
        element.set_attribute("i", if italic { "1" } else { "0" });
        present = true;
    }
    if let Some(underline) = &properties.underline {
        element.set_attribute("u", underline.clone());
        present = true;
    }
    if let Some(color) = properties
        .color
        .as_ref()
        .and_then(|color| color_element(color, prefixes))
    {
        element =
            element.with_child(XmlElement::new(prefixes.drawing("solidFill")).with_child(color));
        present = true;
    }
    if let Some(family) = &properties.font_family {
        element = element.with_child(
            XmlElement::new(prefixes.drawing("latin")).with_attribute("typeface", family.clone()),
        );
        present = true;
    }
    present.then_some(element)
}

// --- new shapes and slides --------------------------------------------------

fn shape_element(
    add: &ShapeAdd,
    next_shape_id: &mut u32,
    prefixes: &Prefixes,
    part: &str,
) -> Result<XmlElement, PptxError> {
    if matches!(add.geometry.as_str(), "custom" | "group") {
        return Err(write_error(
            part,
            format!("unsupported geometry {:?} for a new shape", add.geometry),
        ));
    }
    let shape_id = *next_shape_id;
    *next_shape_id += 1;
    let non_visual = XmlElement::new(prefixes.presentation("nvSpPr"))
        .with_child(
            XmlElement::new(prefixes.presentation("cNvPr"))
                .with_attribute("id", shape_id.to_string())
                .with_attribute("name", add.name.clone()),
        )
        .with_child(XmlElement::new(prefixes.presentation("cNvSpPr")))
        .with_child(XmlElement::new(prefixes.presentation("nvPr")));
    let transform = XmlElement::new(prefixes.drawing("xfrm"))
        .with_child(
            XmlElement::new(prefixes.drawing("off"))
                .with_attribute("x", add.x.to_string())
                .with_attribute("y", add.y.to_string()),
        )
        .with_child(
            XmlElement::new(prefixes.drawing("ext"))
                .with_attribute("cx", add.width.to_string())
                .with_attribute("cy", add.height.to_string()),
        );
    let mut adjust_list = XmlElement::new(prefixes.drawing("avLst"));
    for (name, value) in &add.adjust_values {
        adjust_list = adjust_list.with_child(
            XmlElement::new(prefixes.drawing("gd"))
                .with_attribute("name", name.clone())
                .with_attribute("fmla", format!("val {}", format_fixed(value * 100_000.0))),
        );
    }
    let geometry = XmlElement::new(prefixes.drawing("prstGeom"))
        .with_attribute("prst", add.geometry.clone())
        .with_child(adjust_list);
    let mut properties = XmlElement::new(prefixes.presentation("spPr"))
        .with_child(transform)
        .with_child(geometry);
    if let Some(fill) = &add.fill {
        properties = properties.with_child(fill_element(fill, prefixes, part)?);
    }
    if let Some(outline) = &add.outline {
        properties = properties.with_child(outline_element(outline, prefixes));
    }
    let mut shape = XmlElement::new(prefixes.presentation("sp"))
        .with_child(non_visual)
        .with_child(properties);
    if let Some(paragraphs) = &add.paragraphs {
        let mut body = XmlElement::new(prefixes.presentation("txBody"))
            .with_child(XmlElement::new(prefixes.drawing("bodyPr")))
            .with_child(XmlElement::new(prefixes.drawing("lstStyle")));
        for paragraph in paragraphs {
            body = body.with_child(build_paragraph(paragraph, None, prefixes));
        }
        if paragraphs.is_empty() {
            body = body.with_child(XmlElement::new(prefixes.drawing("p")));
        }
        shape = shape.with_child(body);
    }
    Ok(shape)
}

fn slide_xml(name: Option<&str>, shapes: &[ShapeAdd], part: &str) -> Result<Vec<u8>, PptxError> {
    let mut root = XmlElement::new("p:sld")
        .with_attribute("xmlns:a", DRAWINGML_NS)
        .with_attribute("xmlns:r", OFFICE_RELATIONSHIPS_NS)
        .with_attribute("xmlns:p", PRESENTATIONML_NS);
    let prefixes = Prefixes {
        drawing: "a".to_owned(),
        presentation: "p".to_owned(),
    };
    let group_transform = XmlElement::new(prefixes.drawing("xfrm"))
        .with_child(
            XmlElement::new(prefixes.drawing("off"))
                .with_attribute("x", "0")
                .with_attribute("y", "0"),
        )
        .with_child(
            XmlElement::new(prefixes.drawing("ext"))
                .with_attribute("cx", "0")
                .with_attribute("cy", "0"),
        )
        .with_child(
            XmlElement::new(prefixes.drawing("chOff"))
                .with_attribute("x", "0")
                .with_attribute("y", "0"),
        )
        .with_child(
            XmlElement::new(prefixes.drawing("chExt"))
                .with_attribute("cx", "0")
                .with_attribute("cy", "0"),
        );
    let mut tree = XmlElement::new(prefixes.presentation("spTree"))
        .with_child(
            XmlElement::new(prefixes.presentation("nvGrpSpPr"))
                .with_child(
                    XmlElement::new(prefixes.presentation("cNvPr"))
                        .with_attribute("id", "1")
                        .with_attribute("name", ""),
                )
                .with_child(XmlElement::new(prefixes.presentation("cNvGrpSpPr")))
                .with_child(XmlElement::new(prefixes.presentation("nvPr"))),
        )
        .with_child(XmlElement::new(prefixes.presentation("grpSpPr")).with_child(group_transform));
    let mut next_shape_id = 2;
    for shape in shapes {
        tree = tree.with_child(shape_element(shape, &mut next_shape_id, &prefixes, part)?);
    }
    let mut common = XmlElement::new(prefixes.presentation("cSld"));
    if let Some(name) = name {
        common.set_attribute("name", name);
    }
    root = root.with_child(common.with_child(tree)).with_child(
        XmlElement::new(prefixes.presentation("clrMapOvr"))
            .with_child(XmlElement::new(prefixes.drawing("masterClrMapping"))),
    );
    Ok(serialize_xml(&root))
}

fn slide_relationships_xml(part_path: &str, layout_part_path: &str) -> Vec<u8> {
    let root = XmlElement::new("Relationships")
        .with_attribute(
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        )
        .with_child(
            XmlElement::new("Relationship")
                .with_attribute("Id", "rId1")
                .with_attribute("Type", SLIDE_LAYOUT_RELATIONSHIP_TYPE)
                .with_attribute("Target", relative_target(part_path, layout_part_path)),
        );
    serialize_xml(&root)
}
