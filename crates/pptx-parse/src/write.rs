//! Edit-driven package writes. Parts the deck did not change are copied
//! through byte for byte; edited slides are patched at the XML level so
//! unmodeled markup survives.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::ops::Range;

use ooxml_drawingml::{
    ColorValue, GradientFill, ShapeFill, ShapeOutline, Theme, resolve_color_value_to_hex_with_theme,
};

use crate::PptxError;
use crate::comments::{CommentFlavor, CommentSlide, CommentsWrite, authors_xml, comments_xml};
use crate::drawing::parse_run_properties;
use crate::model::{Bullet, PptxPackage, RunProperties, ShapeElements, SlideReference};
use crate::xml::{
    DRAWINGML_NS, PRESENTATIONML_NS, ParseBudget, ParseLimits, XmlElement, XmlNode,
    alternate_content_branch_index, parse_xml, serialize_xml,
};

const OFFICE_RELATIONSHIPS_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SLIDE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_LAYOUT_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const IMAGE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const MIN_SLIDE_ID: u32 = 256;
const MAX_SLIDE_ID: u32 = 2_147_483_647;

/// The desired final deck, expressed against the parsed source package.
pub struct DeckWrite {
    pub slides: Vec<SlideWrite>,
    pub comments: Option<CommentsWrite>,
    pub notes: Option<NotesWrite>,
}

/// Changed speaker notes; empty text clears the body placeholder.
pub struct NotesWrite {
    pub per_slide: Vec<(CommentSlide, String)>,
}

const NOTES_SLIDE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";

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
    /// The halves the edit changed. Transform pieces the source already
    /// spells out and the edit did not touch are left as they are.
    pub offset: Option<(i64, i64)>,
    pub extent: Option<(i64, i64)>,
    /// Fills in transform pieces the source never spelled out.
    pub inherited: Option<InheritedTransform>,
    pub fill: Option<ShapeFill>,
    /// A default outline clears the stroke.
    pub outline: Option<ShapeOutline>,
    pub adjust_values: Option<BTreeMap<String, f64>>,
    pub texts: Vec<TextWrite>,
    /// Non-empty only when a group's child list must be rebuilt.
    pub children: Vec<ShapeWrite>,
}

/// The transform a shape inherits from its layout or master placeholder.
#[derive(Clone, Copy)]
pub struct InheritedTransform {
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
    pub rotation_deg: f64,
    pub flip_horizontal: bool,
    pub flip_vertical: bool,
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
    /// A picture instead of an autoshape; mints its own media part and relationship.
    pub picture: Option<PictureAdd>,
}

pub struct PictureAdd {
    pub media_bytes: Vec<u8>,
    pub content_type: String,
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

    let mut minted = Vec::new();
    let mut next_slide_number = next_slide_number(package);
    let mut minted_slide_id: Option<u32> = None;
    let mut next_relationship = next_relationship_number(package);
    let mut structural = false;
    let mut final_order: Vec<FinalSlide<'_>> = Vec::new();
    let mut minted_by_slide: HashMap<usize, usize> = HashMap::new();
    {
        let mut sink = PartSink {
            package,
            replacements: &mut replacements,
            new_parts: &mut new_parts,
        };

        for slide in &deck.slides {
            if let SlideWrite::Patch { part_path, shapes } = slide {
                let bytes = package
                    .part_bytes(part_path)
                    .ok_or_else(|| PptxError::MissingPart(part_path.clone()))?;
                let mut root = parse_xml(bytes, part_path, &mut budget)?;
                let original_root = root.clone();
                let theme = crate::slide_theme(package, Some(part_path), None);
                patch_slide(
                    &mut root,
                    shapes,
                    Some(&theme),
                    part_path,
                    package.shape_elements,
                    &mut sink,
                    &mut budget,
                )?;
                if root != original_root {
                    sink.store(part_path, serialize_xml(&root));
                }
            }
        }

        for (slide_index, slide) in deck.slides.iter().enumerate() {
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
                    let slide_id = match minted_slide_id {
                        None => next_slide_id(package)?,
                        Some(previous) => previous
                            .checked_add(1)
                            .filter(|id| *id <= MAX_SLIDE_ID)
                            .ok_or_else(|| {
                                write_error(
                                    &package.presentation.part_path,
                                    "the slide id space is exhausted",
                                )
                            })?,
                    };
                    minted_slide_id = Some(slide_id);
                    let layout = layout_part_path
                        .clone()
                        .filter(|path| {
                            package
                                .layouts
                                .iter()
                                .any(|layout| &layout.part_path == path)
                        })
                        .or_else(|| {
                            package
                                .layouts
                                .first()
                                .map(|layout| layout.part_path.clone())
                        });
                    if let Some(layout) = &layout {
                        sink.store(
                            &slide_relationships_path(&part_path),
                            slide_relationships_xml(&part_path, layout),
                        );
                    }
                    let xml =
                        slide_xml(name.as_deref(), shapes, &part_path, &mut sink, &mut budget)?;
                    sink.store(&part_path, xml);
                    minted.push(MintedSlide {
                        part_path,
                        relationship_id,
                        slide_id,
                    });
                    minted_by_slide.insert(slide_index, minted.len() - 1);
                    final_order.push(FinalSlide::Added(minted.len() - 1));
                }
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

    let mut removed_paths = HashSet::new();
    for reference in &removed {
        removed_paths.insert(reference.part_path.clone());
        removed_paths.insert(slide_relationships_path(&reference.part_path));
    }
    let orphans = orphaned_parts(package, &removed);
    for orphan in &orphans {
        removed_paths.insert(orphan.clone());
        removed_paths.insert(slide_relationships_path(orphan));
    }

    if structural {
        patch_structure(
            package,
            &final_order,
            &minted,
            &removed,
            &orphans,
            &mut replacements,
            &mut budget,
        )?;
    }
    if let Some(comments) = &deck.comments {
        patch_comment_parts(
            package,
            comments,
            &MintedSlides {
                slides: &minted,
                by_slide_index: &minted_by_slide,
            },
            &mut replacements,
            &mut new_parts,
            &mut removed_paths,
            &mut budget,
        )?;
    }
    if let Some(notes) = &deck.notes {
        patch_notes_parts(
            package,
            notes,
            &MintedSlides {
                slides: &minted,
                by_slide_index: &minted_by_slide,
            },
            &mut replacements,
            &mut new_parts,
            &mut removed_paths,
            &mut budget,
        )?;
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

fn next_slide_id(package: &PptxPackage) -> Result<u32, PptxError> {
    let next = package
        .presentation
        .slides
        .iter()
        .map(|reference| reference.id)
        .max()
        .unwrap_or(0)
        .max(MIN_SLIDE_ID - 1)
        .checked_add(1)
        .filter(|id| *id <= MAX_SLIDE_ID);
    next.ok_or_else(|| {
        write_error(
            &package.presentation.part_path,
            "the slide id space is exhausted",
        )
    })
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

/// Parts that were reachable from the package root only through the removed
/// slides (notes slides, media used nowhere else, ...), transitively.
fn orphaned_parts(package: &PptxPackage, removed: &[&SlideReference]) -> Vec<String> {
    if removed.is_empty() {
        return Vec::new();
    }
    let excluded: HashSet<&str> = removed
        .iter()
        .map(|reference| reference.part_path.as_str())
        .collect();
    let before = reachable_parts(package, &HashSet::new());
    let after = reachable_parts(package, &excluded);
    before
        .into_iter()
        .filter(|part| !after.contains(part) && !excluded.contains(part.as_str()))
        .collect()
}

fn reachable_parts(package: &PptxPackage, excluded: &HashSet<&str>) -> HashSet<String> {
    let mut reachable = HashSet::new();
    let mut queue = vec![String::new()];
    while let Some(source) = queue.pop() {
        for relationship in package.relationships.get(&source).into_iter().flatten() {
            let Some(target) = &relationship.resolved_target else {
                continue;
            };
            if excluded.contains(target.as_str()) {
                continue;
            }
            if reachable.insert(target.clone()) {
                queue.push(target.clone());
            }
        }
    }
    reachable
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
    orphans: &[String],
    replacements: &mut HashMap<String, Vec<u8>>,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    let presentation_path = package.presentation.part_path.as_str();
    let bytes = package
        .part_bytes(presentation_path)
        .ok_or_else(|| PptxError::MissingPart(presentation_path.to_owned()))?;
    let mut root = parse_xml(bytes, presentation_path, budget)?;
    patch_slide_id_list(&mut root, final_order, minted, presentation_path)?;
    prune_slide_references(&mut root, removed);
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
    let bytes = match replacements.get(content_types_path) {
        Some(bytes) => bytes.clone(),
        None => package
            .part_bytes(content_types_path)
            .ok_or_else(|| PptxError::MissingPart(content_types_path.to_owned()))?
            .to_vec(),
    };
    let mut root = parse_xml(&bytes, content_types_path, budget)?;
    let removed_names: HashSet<String> = removed
        .iter()
        .map(|reference| format!("/{}", reference.part_path))
        .chain(orphans.iter().map(|orphan| format!("/{orphan}")))
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

const PACKAGE_RELATIONSHIPS_NS: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";

struct MintedSlides<'a> {
    slides: &'a [MintedSlide],
    by_slide_index: &'a HashMap<usize, usize>,
}

impl MintedSlides<'_> {
    fn get(&self, slide_index: usize) -> Option<&MintedSlide> {
        self.by_slide_index
            .get(&slide_index)
            .and_then(|at| self.slides.get(*at))
    }
}

fn patch_comment_parts(
    package: &PptxPackage,
    write: &CommentsWrite,
    minted: &MintedSlides<'_>,
    replacements: &mut HashMap<String, Vec<u8>>,
    new_parts: &mut Vec<(String, Vec<u8>)>,
    removed_paths: &mut HashSet<String>,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    let mut taken: HashSet<String> = package
        .parts
        .iter()
        .map(|part| part.path.clone())
        .chain(
            package
                .relationships
                .values()
                .flatten()
                .filter(|relationship| relationship.is_type(write.comments_relationship_type()))
                .filter_map(|relationship| relationship.resolved_target.clone()),
        )
        .collect();

    let mut sink = PartSink {
        package,
        replacements,
        new_parts,
    };
    for (slide, comments) in &write.per_slide {
        let (slide_part_path, slide_id) = match slide {
            CommentSlide::Existing(part_path) => {
                let reference = package
                    .presentation
                    .slides
                    .iter()
                    .find(|reference| &reference.part_path == part_path)
                    .ok_or_else(|| write_error(part_path, "not a slide of this deck"))?;
                (part_path.clone(), reference.id)
            }
            CommentSlide::Added(index) => {
                let Some(slide) = minted.get(*index) else {
                    continue;
                };
                (slide.part_path.clone(), slide.slide_id)
            }
        };
        let relationships_path = slide_relationships_path(&slide_part_path);
        let existing = existing_comment_part(package, &slide_part_path, write);

        if comments.is_empty() {
            if let Some(part_path) = existing {
                removed_paths.insert(part_path.clone());
                sink.forget(&part_path);
                remove_content_type_override(&mut sink, &part_path, budget)?;
                remove_relationship(
                    &mut sink,
                    &relationships_path,
                    write.comments_relationship_type(),
                    budget,
                )?;
            }
            continue;
        }

        let is_existing = existing.is_some();
        let part_path = match existing {
            Some(part_path) => part_path,
            None => mint_comment_part_path(write, &slide_part_path, &mut taken),
        };
        removed_paths.remove(&part_path);
        let bytes = match sink.current(&part_path) {
            Some(bytes) => crate::comment_patch::patch_comments_xml(
                &bytes, &part_path, write, comments, slide_id, budget,
            )?,
            None => comments_xml(write, comments, slide_id),
        };
        sink.store(&part_path, bytes);
        if !is_existing {
            set_content_type_override(
                &mut sink,
                &part_path,
                write.comments_content_type(),
                budget,
            )?;
            set_relationship(
                &mut sink,
                &relationships_path,
                write.comments_relationship_type(),
                &relative_target(&slide_part_path, &part_path),
                budget,
            )?;
        }
    }

    let presentation_relationships = slide_relationships_path(&package.presentation.part_path);
    let authors_path = existing_authors_part(package, write)
        .unwrap_or_else(|| write.authors_part_path().to_owned());
    if write.authors.is_empty() {
        removed_paths.insert(authors_path.clone());
        sink.forget(&authors_path);
        remove_content_type_override(&mut sink, &authors_path, budget)?;
        remove_relationship(
            &mut sink,
            &presentation_relationships,
            write.authors_relationship_type(),
            budget,
        )?;
    } else {
        removed_paths.remove(&authors_path);
        let existing = sink.current(&authors_path);
        let bytes = match &existing {
            Some(bytes) => {
                crate::comment_patch::patch_authors_xml(bytes, &authors_path, write, budget)?
            }
            None => authors_xml(write),
        };
        sink.store(&authors_path, bytes);
        if existing.is_none() {
            set_content_type_override(
                &mut sink,
                &authors_path,
                write.authors_content_type(),
                budget,
            )?;
            set_relationship(
                &mut sink,
                &presentation_relationships,
                write.authors_relationship_type(),
                &relative_target(&package.presentation.part_path, &authors_path),
                budget,
            )?;
        }
    }

    if package.comment_flavor != Some(write.flavor) {
        drop_other_flavor(package, write, &mut sink, removed_paths, budget)?;
    }
    Ok(())
}

fn patch_notes_parts(
    package: &PptxPackage,
    write: &NotesWrite,
    minted: &MintedSlides<'_>,
    replacements: &mut HashMap<String, Vec<u8>>,
    new_parts: &mut Vec<(String, Vec<u8>)>,
    removed_paths: &mut HashSet<String>,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    let mut taken: HashSet<String> = package
        .parts
        .iter()
        .map(|part| part.path.clone())
        .chain(
            package
                .relationships
                .values()
                .flatten()
                .filter(|relationship| relationship.is_type(NOTES_SLIDE_RELATIONSHIP_TYPE))
                .filter_map(|relationship| relationship.resolved_target.clone()),
        )
        .collect();

    let mut sink = PartSink {
        package,
        replacements,
        new_parts,
    };
    for (slide, text) in &write.per_slide {
        let slide_part_path = match slide {
            CommentSlide::Existing(part_path) => part_path.clone(),
            CommentSlide::Added(index) => {
                let Some(slide) = minted.get(*index) else {
                    continue;
                };
                slide.part_path.clone()
            }
        };
        let relationships_path = slide_relationships_path(&slide_part_path);
        let existing = existing_notes_part(package, &slide_part_path);

        if text.is_empty() && existing.is_none() {
            continue;
        }

        let is_existing = existing.is_some();
        let part_path = match existing {
            Some(part_path) => part_path,
            None => mint_notes_part_path(&slide_part_path, &mut taken),
        };
        removed_paths.remove(&part_path);
        let bytes = match sink.current(&part_path) {
            Some(bytes) => crate::notes::patch_notes_xml(&bytes, &part_path, text, budget)?,
            None => crate::notes::notes_slide_xml(text),
        };
        sink.store(&part_path, bytes);
        if !is_existing {
            set_content_type_override(&mut sink, &part_path, crate::notes::CT_NOTES_SLIDE, budget)?;
            set_relationship(
                &mut sink,
                &relationships_path,
                NOTES_SLIDE_RELATIONSHIP_TYPE,
                &relative_target(&slide_part_path, &part_path),
                budget,
            )?;
            let notes_relationships = slide_relationships_path(&part_path);
            set_relationship(
                &mut sink,
                &notes_relationships,
                SLIDE_RELATIONSHIP_TYPE,
                &relative_target(&part_path, &slide_part_path),
                budget,
            )?;
            if let Some(master) = package
                .relationships
                .get(&package.presentation.part_path)
                .into_iter()
                .flatten()
                .find(|relationship| relationship.has_type("/notesMaster"))
                .and_then(|relationship| relationship.resolved_target.as_deref())
            {
                set_relationship(
                    &mut sink,
                    &notes_relationships,
                    &format!("{OFFICE_RELATIONSHIPS_NS}/notesMaster"),
                    &relative_target(&part_path, master),
                    budget,
                )?;
            }
        }
    }
    Ok(())
}

fn existing_notes_part(package: &PptxPackage, slide_part_path: &str) -> Option<String> {
    package
        .relationships
        .get(slide_part_path)?
        .iter()
        .find(|relationship| relationship.is_type(NOTES_SLIDE_RELATIONSHIP_TYPE))
        .and_then(|relationship| relationship.resolved_target.clone())
}

fn mint_notes_part_path(slide_part_path: &str, taken: &mut HashSet<String>) -> String {
    let preferred = slide_number(slide_part_path);
    let mut number = preferred;
    loop {
        let candidate = format!("ppt/notesSlides/notesSlide{number}.xml");
        if taken.insert(candidate.clone()) {
            return candidate;
        }
        number += 1;
    }
}

fn mint_comment_part_path(
    write: &CommentsWrite,
    slide_part_path: &str,
    taken: &mut HashSet<String>,
) -> String {
    let preferred = slide_number(slide_part_path);
    let mut number = preferred;
    loop {
        let candidate = write.comments_part_path(number);
        if taken.insert(candidate.clone()) {
            return candidate;
        }
        number += 1;
    }
}

fn drop_other_flavor(
    package: &PptxPackage,
    write: &CommentsWrite,
    sink: &mut PartSink<'_>,
    removed_paths: &mut HashSet<String>,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    let other = CommentsWrite {
        flavor: match write.flavor {
            CommentFlavor::Legacy => CommentFlavor::Modern,
            CommentFlavor::Modern => CommentFlavor::Legacy,
        },
        authors: Vec::new(),
        per_slide: Vec::new(),
    };
    for reference in &package.presentation.slides {
        let Some(part_path) = existing_comment_part(package, &reference.part_path, &other) else {
            continue;
        };
        removed_paths.insert(part_path.clone());
        sink.forget(&part_path);
        remove_content_type_override(sink, &part_path, budget)?;
        remove_relationship(
            sink,
            &slide_relationships_path(&reference.part_path),
            other.comments_relationship_type(),
            budget,
        )?;
    }
    let authors_path = existing_authors_part(package, &other)
        .unwrap_or_else(|| other.authors_part_path().to_owned());
    if package.part_bytes(&authors_path).is_some() {
        removed_paths.insert(authors_path.clone());
        sink.forget(&authors_path);
        remove_content_type_override(sink, &authors_path, budget)?;
        remove_relationship(
            sink,
            &slide_relationships_path(&package.presentation.part_path),
            other.authors_relationship_type(),
            budget,
        )?;
    }
    Ok(())
}

fn existing_authors_part(package: &PptxPackage, write: &CommentsWrite) -> Option<String> {
    package
        .relationships
        .get(&package.presentation.part_path)?
        .iter()
        .find(|relationship| relationship.is_type(write.authors_relationship_type()))
        .and_then(|relationship| relationship.resolved_target.clone())
}

fn existing_comment_part(
    package: &PptxPackage,
    slide_part_path: &str,
    write: &CommentsWrite,
) -> Option<String> {
    package
        .relationships
        .get(slide_part_path)?
        .iter()
        .find(|relationship| relationship.is_type(write.comments_relationship_type()))
        .and_then(|relationship| relationship.resolved_target.clone())
}

fn slide_number(slide_part_path: &str) -> usize {
    slide_part_path
        .rsplit_once("/slide")
        .and_then(|(_, tail)| tail.strip_suffix(".xml"))
        .and_then(|digits| digits.parse().ok())
        .unwrap_or(1)
}

struct PartSink<'a> {
    package: &'a PptxPackage,
    replacements: &'a mut HashMap<String, Vec<u8>>,
    new_parts: &'a mut Vec<(String, Vec<u8>)>,
}

impl PartSink<'_> {
    fn current(&self, path: &str) -> Option<Vec<u8>> {
        if let Some(bytes) = self.replacements.get(path) {
            return Some(bytes.clone());
        }
        if let Some((_, bytes)) = self.new_parts.iter().find(|(name, _)| name == path) {
            return Some(bytes.clone());
        }
        self.package.part_bytes(path).map(<[u8]>::to_vec)
    }

    fn store(&mut self, path: &str, bytes: Vec<u8>) {
        if self.package.part_bytes(path).is_some() {
            self.replacements.insert(path.to_owned(), bytes);
            return;
        }
        match self.new_parts.iter_mut().find(|(name, _)| name == path) {
            Some((_, existing)) => *existing = bytes,
            None => self.new_parts.push((path.to_owned(), bytes)),
        }
    }

    fn forget(&mut self, path: &str) {
        self.replacements.remove(path);
        self.new_parts.retain(|(name, _)| name != path);
    }
}

fn set_relationship(
    sink: &mut PartSink<'_>,
    relationships_path: &str,
    relationship_type: &str,
    target: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    let mut root = match sink.current(relationships_path) {
        Some(bytes) => parse_xml(&bytes, relationships_path, budget)?,
        None => XmlElement::new("Relationships").with_attribute("xmlns", PACKAGE_RELATIONSHIPS_NS),
    };
    let existing = root.children.iter_mut().find_map(|child| match child {
        XmlNode::Element(element)
            if element.local_name() == "Relationship"
                && element.attribute("Type") == Some(relationship_type) =>
        {
            Some(element)
        }
        _ => None,
    });
    match existing {
        Some(element) => element.set_attribute("Target", target),
        None => {
            let id = format!("rId{}", max_relationship_id(&root) + 1);
            root.children.push(XmlNode::Element(
                XmlElement::new("Relationship")
                    .with_attribute("Id", id)
                    .with_attribute("Type", relationship_type)
                    .with_attribute("Target", target),
            ));
        }
    }
    sink.store(relationships_path, serialize_xml(&root));
    Ok(())
}

fn remove_relationship(
    sink: &mut PartSink<'_>,
    relationships_path: &str,
    relationship_type: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    let Some(bytes) = sink.current(relationships_path) else {
        return Ok(());
    };
    let mut root = parse_xml(&bytes, relationships_path, budget)?;
    root.children.retain(|child| match child {
        XmlNode::Element(element) if element.local_name() == "Relationship" => {
            element.attribute("Type") != Some(relationship_type)
        }
        _ => true,
    });
    sink.store(relationships_path, serialize_xml(&root));
    Ok(())
}

fn max_relationship_id(root: &XmlElement) -> usize {
    root.children_named("Relationship")
        .filter_map(|element| element.attribute("Id"))
        .filter_map(|id| id.strip_prefix("rId"))
        .filter_map(|number| number.parse::<usize>().ok())
        .max()
        .unwrap_or(0)
}

fn set_content_type_override(
    sink: &mut PartSink<'_>,
    part_path: &str,
    content_type: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    let path = "[Content_Types].xml";
    let bytes = sink
        .current(path)
        .ok_or_else(|| PptxError::MissingPart(path.to_owned()))?;
    let mut root = parse_xml(&bytes, path, budget)?;
    let prefix = root.name.rsplit_once(':').map_or("", |(prefix, _)| prefix);
    let override_name = qualified(prefix, "Override");
    let name = format!("/{part_path}");
    let existing = root.children.iter_mut().find_map(|child| match child {
        XmlNode::Element(element)
            if element.local_name() == "Override"
                && element.attribute("PartName") == Some(name.as_str()) =>
        {
            Some(element)
        }
        _ => None,
    });
    match existing {
        Some(element) => element.set_attribute("ContentType", content_type),
        None => root.children.push(XmlNode::Element(
            XmlElement::new(override_name)
                .with_attribute("PartName", name)
                .with_attribute("ContentType", content_type),
        )),
    }
    sink.store(path, serialize_xml(&root));
    Ok(())
}

fn remove_content_type_override(
    sink: &mut PartSink<'_>,
    part_path: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    let path = "[Content_Types].xml";
    let Some(bytes) = sink.current(path) else {
        return Ok(());
    };
    let mut root = parse_xml(&bytes, path, budget)?;
    let name = format!("/{part_path}");
    root.children.retain(|child| match child {
        XmlNode::Element(element) if element.local_name() == "Override" => {
            element.attribute("PartName") != Some(name.as_str())
        }
        _ => true,
    });
    sink.store(path, serialize_xml(&root));
    Ok(())
}

/// Drops references to removed slides from custom shows (`r:id`) and
/// section lists (`p14:sldId` inside `extLst`).
fn prune_slide_references(root: &mut XmlElement, removed: &[&SlideReference]) {
    if removed.is_empty() {
        return;
    }
    let removed_relationship_ids: HashSet<&str> = removed
        .iter()
        .map(|reference| reference.relationship_id.as_str())
        .collect();
    let removed_slide_ids: HashSet<String> = removed
        .iter()
        .map(|reference| reference.id.to_string())
        .collect();
    for child in root.children.iter_mut() {
        let XmlNode::Element(element) = child else {
            continue;
        };
        match element.local_name() {
            "custShowLst" => remove_matching_descendants(element, &|candidate| {
                candidate.local_name() == "sld"
                    && candidate.attributes.iter().any(|(key, value)| {
                        key.ends_with(":id") && removed_relationship_ids.contains(value.as_str())
                    })
            }),
            "extLst" => remove_matching_descendants(element, &|candidate| {
                candidate.local_name() == "sldId"
                    && candidate
                        .attribute("id")
                        .is_some_and(|id| removed_slide_ids.contains(id))
            }),
            _ => {}
        }
    }
}

fn remove_matching_descendants(element: &mut XmlElement, matches: &dyn Fn(&XmlElement) -> bool) {
    element.children.retain(|child| match child {
        XmlNode::Element(candidate) => !matches(candidate),
        _ => true,
    });
    for child in element.children.iter_mut() {
        if let XmlNode::Element(child) = child {
            remove_matching_descendants(child, matches);
        }
    }
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
    relationship: String,
}

impl Prefixes {
    fn from_root(root: &mut XmlElement) -> Self {
        Self {
            drawing: resolve_prefix(root, DRAWINGML_NS, "a"),
            presentation: resolve_prefix(root, PRESENTATIONML_NS, "p"),
            relationship: resolve_prefix(root, OFFICE_RELATIONSHIPS_NS, "r"),
        }
    }

    fn drawing(&self, local: &str) -> String {
        qualified(&self.drawing, local)
    }

    fn presentation(&self, local: &str) -> String {
        qualified(&self.presentation, local)
    }

    fn relationship(&self, local: &str) -> String {
        qualified(&self.relationship, local)
    }
}

// --- slide patching ---------------------------------------------------------

/// Resolves source ordinals with the original parser filter.
fn patch_slide(
    root: &mut XmlElement,
    shapes: &[ShapeWrite],
    theme: Option<&Theme>,
    part: &str,
    elements: ShapeElements,
    sink: &mut PartSink<'_>,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    let prefixes = Prefixes::from_root(root);
    let mut next_shape_id = max_shape_id(root).checked_add(1);
    let tree = root
        .child_mut("cSld")
        .and_then(|common| common.child_mut("spTree"))
        .ok_or_else(|| write_error(part, "slide has no shape tree"))?;
    patch_shape_children(
        tree,
        shapes,
        &mut next_shape_id,
        theme,
        &prefixes,
        part,
        elements,
        sink,
        budget,
    )
}

fn max_shape_id(root: &XmlElement) -> u32 {
    root.descendants_named("cNvPr")
        .iter()
        .filter_map(|element| element.attribute("id")?.parse::<u32>().ok())
        .max()
        .unwrap_or(1)
}

fn alloc_shape_id(next_shape_id: &mut Option<u32>, part: &str) -> Result<u32, PptxError> {
    let shape_id =
        next_shape_id.ok_or_else(|| write_error(part, "the shape id space is exhausted"))?;
    if shape_id == 0 {
        return Err(write_error(part, "the shape id space is exhausted"));
    }
    *next_shape_id = shape_id.checked_add(1);
    Ok(shape_id)
}

/// A parsed shape's source element: a direct child of the tree, or one the
/// `mc:AlternateContent` at `position` contributes through `path`.
struct ShapeSlot {
    position: usize,
    path: Vec<usize>,
}

/// Rebuilds the shape run in place: non-shape siblings keep their slots
/// relative to the shape that follows them, and elements trailing the last
/// shape (`p:extLst` in particular) stay last.
#[allow(clippy::too_many_arguments)]
fn patch_shape_children(
    parent: &mut XmlElement,
    writes: &[ShapeWrite],
    next_shape_id: &mut Option<u32>,
    theme: Option<&Theme>,
    prefixes: &Prefixes,
    part: &str,
    elements: ShapeElements,
    sink: &mut PartSink<'_>,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    let mut slots: Vec<Option<XmlNode>> = std::mem::take(&mut parent.children)
        .into_iter()
        .map(Some)
        .collect();
    let mut shape_slots = shape_slots(&slots, elements);
    let mut kept = vec![false; shape_slots.len()];
    for write in writes {
        if let ShapeWrite::Keep { source_index } | ShapeWrite::Patch { source_index, .. } = write
            && let Some(flag) = kept.get_mut(*source_index)
        {
            *flag = true;
        }
    }
    // A wrapper is not a shape element, so the sibling flush carries it through whatever became of
    // the shapes inside it: a deleted one leaves the branch, an emptied wrapper leaves the file.
    let mut nested: BTreeMap<usize, (usize, Vec<Vec<usize>>)> = BTreeMap::new();
    for (slot, kept) in shape_slots.iter().zip(&kept) {
        if slot.path.is_empty() {
            continue;
        }
        let entry = nested.entry(slot.position).or_default();
        entry.0 += 1;
        if !kept {
            entry.1.push(slot.path.clone());
        }
    }
    for (position, (total, mut paths)) in nested {
        if paths.len() == total && branch_holds_only(&slots[position], &paths) {
            slots[position] = None;
            continue;
        }
        paths.sort();
        for path in paths.iter().rev() {
            let Some(node) = slots[position].as_mut() else {
                continue;
            };
            remove_at_path(node, path);
            for slot in shape_slots
                .iter_mut()
                .filter(|slot| slot.position == position)
            {
                shift_path(&mut slot.path, path);
            }
        }
    }
    let mut children = Vec::with_capacity(slots.len());
    let mut hosts: BTreeMap<usize, usize> = BTreeMap::new();
    let first_ext_list = |slots: &[Option<XmlNode>]| {
        slots
            .iter()
            .position(|slot| {
                matches!(slot, Some(XmlNode::Element(element)) if element.local_name() == "extLst")
            })
            .unwrap_or(slots.len())
    };
    let prologue_end = shape_slots
        .first()
        .map(|slot| slot.position)
        .unwrap_or_else(|| first_ext_list(&slots));
    emit_sibling_slots(&mut slots, &mut children, prologue_end, elements);
    let last_kept = writes
        .iter()
        .rposition(|write| !matches!(write, ShapeWrite::Add(_)));
    for (index, write) in writes.iter().enumerate() {
        match write {
            ShapeWrite::Keep { source_index } | ShapeWrite::Patch { source_index, .. } => {
                let slot = shape_slots.get(*source_index).ok_or_else(|| {
                    write_error(part, format!("shape index {source_index} is not available"))
                })?;
                emit_sibling_slots(&mut slots, &mut children, slot.position, elements);
                let target = emit_source_element(&mut slots, &mut children, &mut hosts, slot)
                    .ok_or_else(|| {
                        write_error(
                            part,
                            format!("shape index {source_index} was already written"),
                        )
                    })?;
                if let ShapeWrite::Patch { patch, .. } = write {
                    let element =
                        element_at_path(&mut children[target], &slot.path).ok_or_else(|| {
                            write_error(
                                part,
                                format!("shape index {source_index} is not available"),
                            )
                        })?;
                    patch_shape(
                        element,
                        patch,
                        next_shape_id,
                        theme,
                        prefixes,
                        part,
                        elements,
                        sink,
                        budget,
                    )?;
                }
            }
            ShapeWrite::Add(add) => {
                // A trailing add goes on top: after every remaining sibling
                // except a final extLst.
                if last_kept.is_none_or(|kept| index > kept) {
                    let flush_end = first_ext_list(&slots);
                    emit_sibling_slots(&mut slots, &mut children, flush_end, elements);
                }
                children.push(XmlNode::Element(add_shape_element(
                    add,
                    next_shape_id,
                    prefixes,
                    part,
                    sink,
                    budget,
                )?));
            }
        }
    }
    for slot in slots.into_iter().flatten() {
        if !matches!(&slot, XmlNode::Element(element) if elements.contains(element.local_name())) {
            children.push(slot);
        }
    }
    parent.children = children;
    Ok(())
}

/// Source ordinals in the order [`common_slide_data`] parses them.
fn shape_slots(slots: &[Option<XmlNode>], elements: ShapeElements) -> Vec<ShapeSlot> {
    let mut found = Vec::new();
    for (position, slot) in slots.iter().enumerate() {
        let Some(XmlNode::Element(element)) = slot else {
            continue;
        };
        if elements.contains(element.local_name()) {
            found.push(ShapeSlot {
                position,
                path: Vec::new(),
            });
        } else if element.local_name() == "AlternateContent" {
            collect_branch_slots(element, position, &mut Vec::new(), elements, &mut found);
        }
    }
    found
}

fn collect_branch_slots(
    alternate: &XmlElement,
    position: usize,
    path: &mut Vec<usize>,
    elements: ShapeElements,
    found: &mut Vec<ShapeSlot>,
) {
    let Some(index) = alternate_content_branch_index(alternate) else {
        return;
    };
    let XmlNode::Element(branch) = &alternate.children[index] else {
        return;
    };
    path.push(index);
    for (index, child) in branch.children.iter().enumerate() {
        let XmlNode::Element(child) = child else {
            continue;
        };
        path.push(index);
        if elements.contains(child.local_name()) {
            found.push(ShapeSlot {
                position,
                path: path.clone(),
            });
        } else if child.local_name() == "AlternateContent" {
            collect_branch_slots(child, position, path, elements, found);
        }
        path.pop();
    }
    path.pop();
}

/// Moves a shape's source element into `children`, or reuses the
/// `mc:AlternateContent` already moved there for an earlier nested shape.
fn emit_source_element(
    slots: &mut [Option<XmlNode>],
    children: &mut Vec<XmlNode>,
    hosts: &mut BTreeMap<usize, usize>,
    slot: &ShapeSlot,
) -> Option<usize> {
    if let Some(host) = hosts.get(&slot.position) {
        return Some(*host);
    }
    let XmlNode::Element(element) = slots[slot.position].take()? else {
        return None;
    };
    children.push(XmlNode::Element(element));
    if !slot.path.is_empty() {
        hosts.insert(slot.position, children.len() - 1);
    }
    Some(children.len() - 1)
}

/// Whether the branch `slot` reads holds nothing but the elements at `paths`.
fn branch_holds_only(slot: &Option<XmlNode>, paths: &[Vec<usize>]) -> bool {
    let Some(XmlNode::Element(alternate)) = slot else {
        return true;
    };
    let Some(index) = alternate_content_branch_index(alternate) else {
        return true;
    };
    let XmlNode::Element(branch) = &alternate.children[index] else {
        return true;
    };
    branch.children.iter().enumerate().all(|(child, node)| {
        !matches!(node, XmlNode::Element(_)) || paths.iter().any(|path| path == &[index, child])
    })
}

fn remove_at_path(node: &mut XmlNode, path: &[usize]) {
    let Some((last, parents)) = path.split_last() else {
        return;
    };
    if let Some(parent) = element_at_path(node, parents)
        && *last < parent.children.len()
    {
        parent.children.remove(*last);
    }
}

/// Closes the gap [`remove_at_path`] leaves in the paths of its later siblings.
fn shift_path(path: &mut [usize], removed: &[usize]) {
    let Some((last, parents)) = removed.split_last() else {
        return;
    };
    if path.len() > parents.len() && path.starts_with(parents) && path[parents.len()] > *last {
        path[parents.len()] -= 1;
    }
}

fn element_at_path<'a>(node: &'a mut XmlNode, path: &[usize]) -> Option<&'a mut XmlElement> {
    let XmlNode::Element(element) = node else {
        return None;
    };
    let mut current = element;
    for index in path {
        match current.children.get_mut(*index)? {
            XmlNode::Element(child) => current = child,
            XmlNode::Text(_) => return None,
        }
    }
    Some(current)
}

/// Emits the not-yet-written non-shape nodes that precede `position`.
fn emit_sibling_slots(
    slots: &mut [Option<XmlNode>],
    children: &mut Vec<XmlNode>,
    position: usize,
    elements: ShapeElements,
) {
    for slot in slots.iter_mut().take(position) {
        let is_other = !matches!(
            slot,
            Some(XmlNode::Element(element)) if elements.contains(element.local_name())
        );
        if is_other && let Some(node) = slot.take() {
            children.push(node);
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn patch_shape(
    element: &mut XmlElement,
    patch: &ShapePatch,
    next_shape_id: &mut Option<u32>,
    theme: Option<&Theme>,
    prefixes: &Prefixes,
    part: &str,
    elements: ShapeElements,
    sink: &mut PartSink<'_>,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    if patch.offset.is_some() || patch.extent.is_some() {
        patch_transform(element, patch, prefixes, part)?;
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
        set_adjust_values(properties, adjust_values, prefixes, part)?;
    }
    for text in &patch.texts {
        patch_text(element, text, theme, prefixes, part)?;
    }
    if !patch.children.is_empty() {
        patch_shape_children(
            element,
            &patch.children,
            next_shape_id,
            theme,
            prefixes,
            part,
            elements,
            sink,
            budget,
        )?;
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

/// Writes the edited transform halves. Pieces the source spells out and the
/// edit did not change are left untouched; pieces the source lacks are
/// materialized from the inherited transform, rotation and flips included.
fn patch_transform(
    element: &mut XmlElement,
    patch: &ShapePatch,
    prefixes: &Prefixes,
    part: &str,
) -> Result<(), PptxError> {
    let offset = patch
        .offset
        .or(patch.inherited.map(|inherited| (inherited.x, inherited.y)));
    let extent = patch.extent.or(patch
        .inherited
        .map(|inherited| (inherited.width, inherited.height)));
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
        if offset.is_none() || extent.is_none() {
            return Err(write_error(
                part,
                "cannot write half a transform onto a shape without one",
            ));
        }
        let position = parent
            .children
            .iter()
            .position(|child| {
                matches!(child, XmlNode::Element(element) if element.local_name() != "nvGraphicFramePr")
            })
            .unwrap_or(parent.children.len());
        let mut created = XmlElement::new(transform_name);
        if let Some(inherited) = patch.inherited {
            if inherited.rotation_deg != 0.0 {
                created.set_attribute("rot", format_fixed(inherited.rotation_deg * 60_000.0));
            }
            if inherited.flip_horizontal {
                created.set_attribute("flipH", "1");
            }
            if inherited.flip_vertical {
                created.set_attribute("flipV", "1");
            }
        }
        parent.children.insert(position, XmlNode::Element(created));
    }
    let transform = parent.child_mut("xfrm").expect("transform ensured above");
    if let Some((x, y)) = offset
        && (patch.offset.is_some() || transform.child_mut("off").is_none())
    {
        if transform.child_mut("off").is_none() {
            transform.children.insert(
                0,
                XmlNode::Element(XmlElement::new(prefixes.drawing("off"))),
            );
        }
        let element = transform.child_mut("off").expect("offset ensured above");
        element.set_attribute("x", x.to_string());
        element.set_attribute("y", y.to_string());
    }
    if let Some((width, height)) = extent
        && (patch.extent.is_some() || transform.child_mut("ext").is_none())
    {
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
                XmlNode::Element(XmlElement::new(prefixes.drawing("ext"))),
            );
        }
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
            gradient_element(gradient, prefixes)
                .ok_or_else(|| write_error(part, "gradient stop without a color"))
        }
        other => Err(write_error(
            part,
            format!("unsupported fill type {other:?}"),
        )),
    }
}

fn gradient_element(gradient: &GradientFill, prefixes: &Prefixes) -> Option<XmlElement> {
    let mut stops = XmlElement::new(prefixes.drawing("gsLst"));
    for stop in &gradient.stops {
        stops = stops.with_child(
            XmlElement::new(prefixes.drawing("gs"))
                .with_attribute("pos", format_fixed(stop.position))
                .with_child(color_element(&stop.color, prefixes)?),
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
            element = element
                .with_child(XmlElement::new(prefixes.drawing("path")).with_attribute("path", path));
        }
    }
    Some(element)
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
    if let Some(line) = properties.child_mut("ln")
        && let Some(mut original) = crate::drawing::parse_outline_element(line)
    {
        original.width = outline.width;
        if original == *outline {
            if let Some(width) = outline.width {
                line.set_attribute("w", format_fixed(width));
            } else {
                line.attributes.remove("w");
            }
            return;
        }
    }
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
    } else if let Some(gradient) = outline
        .gradient
        .as_ref()
        .and_then(|gradient| gradient_element(gradient, prefixes))
    {
        element = element.with_child(gradient);
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
    part: &str,
) -> Result<(), PptxError> {
    let Some(geometry) = properties.child_mut("prstGeom") else {
        return Err(write_error(
            part,
            "cannot write adjustments onto a shape without preset geometry",
        ));
    };
    let list_name = prefixes.drawing("avLst");
    if geometry.child_mut("avLst").is_none() {
        geometry
            .children
            .insert(0, XmlNode::Element(XmlElement::new(list_name)));
    }
    let list = geometry.child_mut("avLst").expect("list ensured above");
    // Guides absent from the model (unevaluated formulas) are kept as-is.
    for (name, value) in adjust_values {
        let formula = format!("val {}", format_fixed(value * 100_000.0));
        let existing = list.children.iter_mut().find_map(|child| match child {
            XmlNode::Element(element)
                if element.local_name() == "gd" && element.attribute("name") == Some(name) =>
            {
                Some(element)
            }
            _ => None,
        });
        match existing {
            Some(guide) => guide.set_attribute("fmla", formula),
            None => list.children.push(XmlNode::Element(
                XmlElement::new(prefixes.drawing("gd"))
                    .with_attribute("name", name.clone())
                    .with_attribute("fmla", formula),
            )),
        }
    }
    Ok(())
}

// --- text -------------------------------------------------------------------

fn patch_text(
    element: &mut XmlElement,
    text: &TextWrite,
    theme: Option<&Theme>,
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
    rebuild_paragraphs(body, &text.paragraphs, theme, prefixes);
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

fn rebuild_paragraphs(
    body: &mut XmlElement,
    paragraphs: &[ParagraphWrite],
    theme: Option<&Theme>,
    prefixes: &Prefixes,
) {
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
            source => build_paragraph(paragraph, source, theme, prefixes),
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

fn is_run_element(local: &str) -> bool {
    matches!(local, "r" | "br" | "fld")
}

/// One piece of the paragraph's target text: a line break or a styled span.
struct RunSegment<'a> {
    line_break: bool,
    text: &'a str,
    properties: &'a RunProperties,
}

fn run_segments(runs: &[RunWrite]) -> Vec<RunSegment<'_>> {
    let mut segments = Vec::new();
    for run in runs {
        for (index, piece) in run.text.split('\n').enumerate() {
            if index > 0 {
                segments.push(RunSegment {
                    line_break: true,
                    text: "",
                    properties: &run.properties,
                });
            }
            if !piece.is_empty() {
                segments.push(RunSegment {
                    line_break: false,
                    text: piece,
                    properties: &run.properties,
                });
            }
        }
    }
    segments
}

/// Matches text and modeled styling against a source run.
fn segment_matches(element: &XmlElement, segment: &RunSegment<'_>, theme: Option<&Theme>) -> bool {
    if element.local_name() == "br" {
        return segment.line_break;
    }
    if segment.line_break {
        return false;
    }
    let text = element
        .child("t")
        .map(XmlElement::text_content)
        .unwrap_or_default();
    if text != segment.text {
        return false;
    }
    let source = parse_run_properties(element.child("rPr"));
    let target = segment.properties;
    source.baseline_pct == target.baseline_pct
        && source.bold == target.bold
        && source.italic == target.italic
        && source.font_size_pt == target.font_size_pt
        && source.spacing_pt == target.spacing_pt
        && source.underline == target.underline
        && source.caps == target.caps
        && source.font_family == target.font_family
        && resolve_color_value_to_hex_with_theme(source.color.as_ref(), theme)
            == resolve_color_value_to_hex_with_theme(target.color.as_ref(), theme)
}

fn build_paragraph(
    write: &ParagraphWrite,
    source: Option<XmlElement>,
    theme: Option<&Theme>,
    prefixes: &Prefixes,
) -> XmlElement {
    let (mut paragraph, source_runs) = match source {
        Some(mut element) => {
            let mut runs = Vec::new();
            let mut kept = Vec::new();
            for child in std::mem::take(&mut element.children) {
                match child {
                    XmlNode::Element(child) if is_run_element(child.local_name()) => {
                        runs.push(child);
                    }
                    XmlNode::Text(text) if text.trim().is_empty() => {}
                    other => kept.push(other),
                }
            }
            element.children = kept;
            (element, runs)
        }
        None => (XmlElement::new(prefixes.drawing("p")), Vec::new()),
    };
    let fresh = write.source_index.is_none();
    if write.properties_changed || fresh {
        apply_paragraph_properties(&mut paragraph, write, prefixes);
    }
    let segments = run_segments(&write.runs);
    let mut front = 0;
    let mut back = 0;
    while front < source_runs.len()
        && front < segments.len()
        && segment_matches(&source_runs[front], &segments[front], theme)
    {
        let steals = front + 1 < source_runs.len() - back
            && segment_matches(&source_runs[front + 1], &segments[front], theme)
            && source_span_len(&source_runs, front + 1..source_runs.len() - back)
                > target_span_len(&segments, front + 1..segments.len() - back);
        if steals {
            break;
        }
        front += 1;
    }
    while back < source_runs.len() - front
        && back < segments.len() - front
        && segment_matches(
            &source_runs[source_runs.len() - 1 - back],
            &segments[segments.len() - 1 - back],
            theme,
        )
    {
        let src = source_runs.len() - 1 - back;
        let seg = segments.len() - 1 - back;
        let steals = src > front
            && segment_matches(&source_runs[src - 1], &segments[seg], theme)
            && source_span_len(&source_runs, front..src) > target_span_len(&segments, front..seg);
        if steals {
            break;
        }
        back += 1;
    }
    let tail_start = source_runs.len() - back;
    let rebuilt = span_elements(
        &segments[front..segments.len() - back],
        &source_runs,
        front..tail_start,
        theme,
        prefixes,
    );
    let mut runs: Vec<XmlNode> = Vec::with_capacity(segments.len());
    let mut source_runs = source_runs.into_iter();
    for element in source_runs.by_ref().take(front) {
        runs.push(XmlNode::Element(element));
    }
    runs.extend(rebuilt);
    for element in source_runs.skip(tail_start - front) {
        runs.push(XmlNode::Element(element));
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

/// Target slice rebuilt onto one source run.
struct TargetRange {
    end: usize,
    source: usize,
    /// Source text kept unchanged, in order.
    verbatim: bool,
    /// Intact `a:fld` kept as a field.
    field: bool,
}

/// Rebuilds the span onto source runs, preserving unmodeled markup.
fn span_elements(
    segments: &[RunSegment<'_>],
    source_runs: &[XmlElement],
    span: Range<usize>,
    theme: Option<&Theme>,
    prefixes: &Prefixes,
) -> Vec<XmlNode> {
    if source_runs.is_empty() {
        return segments
            .iter()
            .map(|segment| XmlNode::Element(segment_element(segment, None, theme, prefixes)))
            .collect();
    }
    let ranges = align_span(segments, source_runs, span);
    let mut output = Vec::new();
    let mut offset = 0;
    let mut range_index = 0;
    let mut range_start = 0;
    for segment in segments {
        let mut remaining = segment.text;
        loop {
            while ranges[range_index].end <= offset {
                range_start = ranges[range_index].end;
                range_index += 1;
            }
            let range = &ranges[range_index];
            let source = &source_runs[range.source];
            if segment.line_break {
                output.push(XmlNode::Element(segment_element(
                    segment,
                    source.child("rPr"),
                    theme,
                    prefixes,
                )));
                offset += 1;
                break;
            }
            let length = remaining.len().min(range.end - offset);
            let piece = RunSegment {
                text: &remaining[..length],
                ..*segment
            };
            let element = if range.field && offset == range_start && offset + length == range.end {
                field_element(source, &piece, theme, prefixes)
            } else {
                segment_element(&piece, source.child("rPr"), theme, prefixes)
            };
            output.push(XmlNode::Element(element));
            offset += length;
            remaining = &remaining[length..];
            if remaining.is_empty() {
                break;
            }
        }
    }
    output
}

fn run_text(run: &XmlElement) -> String {
    if run.local_name() == "br" {
        return "\n".to_owned();
    }
    run.child("t")
        .map(XmlElement::text_content)
        .unwrap_or_default()
}

fn segment_text<'a>(segment: &RunSegment<'a>) -> &'a str {
    if segment.line_break {
        "\n"
    } else {
        segment.text
    }
}

fn source_span_len(runs: &[XmlElement], span: Range<usize>) -> usize {
    runs[span].iter().map(|run| run_text(run).len()).sum()
}

fn target_span_len(segments: &[RunSegment<'_>], span: Range<usize>) -> usize {
    segments[span]
        .iter()
        .map(|segment| segment_text(segment).len())
        .sum()
}

fn push_range(ranges: &mut Vec<TargetRange>, end: usize, source: usize, verbatim: bool) {
    if end <= ranges.last().map_or(0, |last| last.end) {
        return;
    }
    match ranges.last_mut() {
        Some(last) if last.source == source && last.verbatim == verbatim => last.end = end,
        _ => ranges.push(TargetRange {
            end,
            source,
            verbatim,
            field: false,
        }),
    }
}

fn align_span(
    segments: &[RunSegment<'_>],
    source_runs: &[XmlElement],
    span: Range<usize>,
) -> Vec<TargetRange> {
    let runs = &source_runs[span.clone()];
    let mut source_text = String::new();
    let mut source_ends = Vec::with_capacity(runs.len());
    for run in runs {
        source_text.push_str(&run_text(run));
        source_ends.push(source_text.len());
    }
    let target_text: String = segments.iter().map(segment_text).collect();
    let run_at = |offset: usize| {
        span.start
            + source_ends
                .iter()
                .position(|&end| end > offset)
                .unwrap_or(source_ends.len().saturating_sub(1))
    };
    let prefix: usize = source_text
        .chars()
        .zip(target_text.chars())
        .take_while(|(left, right)| left == right)
        .map(|(value, _)| value.len_utf8())
        .sum();
    let suffix: usize = source_text[prefix..]
        .chars()
        .rev()
        .zip(target_text[prefix..].chars().rev())
        .take_while(|(left, right)| left == right)
        .map(|(value, _)| value.len_utf8())
        .sum();
    let source_end = source_text.len() - suffix;
    let target_end = target_text.len() - suffix;

    let mut ranges: Vec<TargetRange> = Vec::new();
    for (index, &end) in source_ends.iter().enumerate() {
        push_range(&mut ranges, end.min(prefix), span.start + index, true);
    }
    let seed = span.start.saturating_sub(1);
    let (mut source_offset, mut target_offset) = (prefix, prefix);
    let mut replaced: Option<usize> = None;
    for op in char_diff(
        &source_text[prefix..source_end],
        &target_text[prefix..target_end],
    ) {
        match op {
            DiffOp::Match(length) => {
                let mut done = 0;
                while done < length {
                    let run = run_at(source_offset + done);
                    let piece =
                        (source_ends[run - span.start] - (source_offset + done)).min(length - done);
                    push_range(&mut ranges, target_offset + done + piece, run, true);
                    done += piece;
                }
                source_offset += length;
                target_offset += length;
                replaced = None;
            }
            DiffOp::Delete(length) => {
                if replaced.is_none() {
                    replaced = Some(run_at(source_offset));
                }
                source_offset += length;
            }
            DiffOp::Insert(length) => {
                let source = replaced
                    .take()
                    .or_else(|| ranges.last().map(|range| range.source))
                    .unwrap_or(seed);
                push_range(&mut ranges, target_offset + length, source, false);
                target_offset += length;
            }
        }
    }
    for (index, &end) in source_ends.iter().enumerate() {
        if end > source_end {
            push_range(
                &mut ranges,
                target_end + end - source_end,
                span.start + index,
                true,
            );
        }
    }
    let mut kept = vec![false; runs.len()];
    for range in &ranges {
        if range.verbatim {
            kept[range.source - span.start] = true;
        }
    }
    for field_idx in 0..runs.len() {
        if runs[field_idx].local_name() != "fld" || kept[field_idx] {
            continue;
        }
        let field_text = run_text(&runs[field_idx]);
        if field_text.is_empty() {
            continue;
        }
        let mut donor: Option<usize> = None;
        for candidate in [
            field_idx.checked_sub(1),
            field_idx.checked_add(1).filter(|next| *next < runs.len()),
        ]
        .into_iter()
        .flatten()
        {
            if runs[candidate].local_name() == "r"
                && kept[candidate]
                && run_text(&runs[candidate]) == field_text
            {
                donor = Some(candidate);
                break;
            }
        }
        if let Some(donor) = donor {
            for range in &mut ranges {
                if range.verbatim && range.source == span.start + donor {
                    range.source = span.start + field_idx;
                }
            }
            kept[donor] = false;
            kept[field_idx] = true;
        }
    }

    let mut verbatim_ranges = vec![0usize; runs.len()];
    for range in ranges.iter().filter(|range| range.verbatim) {
        verbatim_ranges[range.source - span.start] += 1;
    }
    let mut merged: Vec<TargetRange> = Vec::with_capacity(ranges.len());
    let mut start = 0;
    for mut range in ranges {
        range.field = range.verbatim && {
            let relative = range.source - span.start;
            let run_start = relative
                .checked_sub(1)
                .map_or(0, |previous| source_ends[previous]);
            runs[relative].local_name() == "fld"
                && verbatim_ranges[relative] == 1
                && range.end - start == source_ends[relative] - run_start
        };
        start = range.end;
        match merged.last_mut() {
            Some(last) if last.source == range.source && !last.field && !range.field => {
                last.end = range.end
            }
            _ => merged.push(range),
        }
    }
    merged
}

/// Diff step in bytes.
enum DiffOp {
    Match(usize),
    Delete(usize),
    Insert(usize),
}

const DIFF_CELL_LIMIT: usize = 1 << 20;

fn push_op(ops: &mut Vec<DiffOp>, op: DiffOp) {
    match (ops.last_mut(), op) {
        (Some(DiffOp::Match(last)), DiffOp::Match(length))
        | (Some(DiffOp::Delete(last)), DiffOp::Delete(length))
        | (Some(DiffOp::Insert(last)), DiffOp::Insert(length)) => *last += length,
        (_, op) => ops.push(op),
    }
}

/// LCS alignment; ties delete first, oversized counts as replacement.
fn char_diff(source: &str, target: &str) -> Vec<DiffOp> {
    let source_chars: Vec<char> = source.chars().collect();
    let target_chars: Vec<char> = target.chars().collect();
    let (rows, columns) = (source_chars.len(), target_chars.len());
    let mut ops = Vec::new();
    if rows == 0 || columns == 0 || rows.saturating_mul(columns) > DIFF_CELL_LIMIT {
        if rows > 0 {
            push_op(&mut ops, DiffOp::Delete(source.len()));
        }
        if columns > 0 {
            push_op(&mut ops, DiffOp::Insert(target.len()));
        }
        return ops;
    }
    let width = columns + 1;
    let mut table = vec![0u16; (rows + 1) * width];
    for row in (0..rows).rev() {
        for column in (0..columns).rev() {
            table[row * width + column] = if source_chars[row] == target_chars[column] {
                table[(row + 1) * width + column + 1] + 1
            } else {
                table[(row + 1) * width + column].max(table[row * width + column + 1])
            };
        }
    }
    let (mut row, mut column) = (0, 0);
    while row < rows && column < columns {
        if source_chars[row] == target_chars[column] {
            push_op(&mut ops, DiffOp::Match(source_chars[row].len_utf8()));
            row += 1;
            column += 1;
        } else if table[(row + 1) * width + column] >= table[row * width + column + 1] {
            push_op(&mut ops, DiffOp::Delete(source_chars[row].len_utf8()));
            row += 1;
        } else {
            push_op(&mut ops, DiffOp::Insert(target_chars[column].len_utf8()));
            column += 1;
        }
    }
    for value in &source_chars[row..] {
        push_op(&mut ops, DiffOp::Delete(value.len_utf8()));
    }
    for value in &target_chars[column..] {
        push_op(&mut ops, DiffOp::Insert(value.len_utf8()));
    }
    ops
}

fn segment_properties(
    segment: &RunSegment<'_>,
    template: Option<&XmlElement>,
    theme: Option<&Theme>,
    prefixes: &Prefixes,
) -> Option<XmlElement> {
    match template.filter(|template| !segment.line_break || template.child("gradFill").is_some()) {
        Some(template) => {
            let mut base = template.clone();
            apply_run_properties(&mut base, segment.properties, theme, prefixes);
            Some(base)
        }
        None => run_properties_element(segment.properties, prefixes),
    }
}

fn segment_element(
    segment: &RunSegment<'_>,
    template: Option<&XmlElement>,
    theme: Option<&Theme>,
    prefixes: &Prefixes,
) -> XmlElement {
    let mut element =
        XmlElement::new(prefixes.drawing(if segment.line_break { "br" } else { "r" }));
    if let Some(properties) = segment_properties(segment, template, theme, prefixes) {
        element = element.with_child(properties);
    }
    if segment.line_break {
        element
    } else {
        element.with_child(XmlElement::new(prefixes.drawing("t")).with_text(segment.text))
    }
}

/// Intact `a:fld` keeps its binding.
fn field_element(
    source: &XmlElement,
    segment: &RunSegment<'_>,
    theme: Option<&Theme>,
    prefixes: &Prefixes,
) -> XmlElement {
    let mut element = XmlElement::new(source.name.clone());
    element.attributes = source.attributes.clone();
    if let Some(properties) = segment_properties(segment, source.child("rPr"), theme, prefixes) {
        element.children.push(XmlNode::Element(properties));
    }
    element.children.extend(
        source
            .child_elements()
            .filter(|child| !matches!(child.local_name(), "rPr" | "t"))
            .cloned()
            .map(XmlNode::Element),
    );
    element.children.push(XmlNode::Element(
        XmlElement::new(prefixes.drawing("t")).with_text(segment.text),
    ));
    element
}

const POST_LATIN_ELEMENTS: [&str; 7] = [
    "ea",
    "cs",
    "sym",
    "hlinkClick",
    "hlinkMouseOver",
    "rtl",
    "extLst",
];

/// Updates modeled styling while preserving unmodeled XML.
fn apply_run_properties(
    base: &mut XmlElement,
    properties: &RunProperties,
    theme: Option<&Theme>,
    prefixes: &Prefixes,
) {
    match properties.font_size_pt {
        Some(size) => base.set_attribute("sz", format_fixed(size * 100.0)),
        None => {
            base.attributes.remove("sz");
        }
    }
    match properties.spacing_pt {
        Some(spacing) => base.set_attribute("spc", format_fixed(spacing * 100.0)),
        None => {
            base.attributes.remove("spc");
        }
    }
    match properties.baseline_pct {
        Some(baseline) => base.set_attribute("baseline", format_fixed(baseline * 1000.0)),
        None => {
            base.attributes.remove("baseline");
        }
    }
    match properties.caps {
        Some(caps) => base.set_attribute("cap", caps.as_attribute()),
        None => {
            base.attributes.remove("cap");
        }
    }
    let toggles = [("b", properties.bold), ("i", properties.italic)];
    for (name, value) in toggles {
        match value {
            Some(value) => base.set_attribute(name, if value { "1" } else { "0" }),
            None => {
                base.attributes.remove(name);
            }
        }
    }
    match &properties.underline {
        Some(underline) => base.set_attribute("u", underline.clone()),
        None => {
            base.attributes.remove("u");
        }
    }
    let keeps_gradient = properties.color.as_ref().is_some_and(|color| {
        base.child("gradFill")
            .and_then(crate::drawing::run_gradient_color)
            .is_some_and(|authored| {
                &authored == color
                    || resolve_color_value_to_hex_with_theme(Some(&authored), theme).is_some_and(
                        |authored| {
                            Some(authored)
                                == resolve_color_value_to_hex_with_theme(Some(color), theme)
                        },
                    )
            })
    });
    // A colour write replaces the whole fill choice; clearing the colour only
    // drops an explicit solidFill so an unmodeled noFill/gradFill survives.
    let removed_fills: &[&str] = if keeps_gradient {
        &[]
    } else if properties.color.is_some() {
        &FILL_ELEMENTS
    } else {
        &["solidFill"]
    };
    base.children.retain(|child| {
        !matches!(
            child,
            XmlNode::Element(element) if removed_fills.contains(&element.local_name())
        )
    });
    if let Some(color) = properties
        .color
        .as_ref()
        .filter(|_| !keeps_gradient)
        .and_then(|color| color_element(color, prefixes))
    {
        let position = base
            .children
            .iter()
            .position(
                |child| !matches!(child, XmlNode::Element(element) if element.local_name() == "ln"),
            )
            .unwrap_or(base.children.len());
        base.children.insert(
            position,
            XmlNode::Element(XmlElement::new(prefixes.drawing("solidFill")).with_child(color)),
        );
    }
    base.children.retain(|child| {
        !matches!(
            child,
            XmlNode::Element(element) if element.local_name() == "latin"
        )
    });
    if let Some(family) = &properties.font_family {
        let position = base
            .children
            .iter()
            .position(|child| {
                matches!(
                    child,
                    XmlNode::Element(element)
                        if POST_LATIN_ELEMENTS.contains(&element.local_name())
                )
            })
            .unwrap_or(base.children.len());
        base.children.insert(
            position,
            XmlNode::Element(
                XmlElement::new(prefixes.drawing("latin"))
                    .with_attribute("typeface", family.clone()),
            ),
        );
    }
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
        Some(Bullet::AutoNumber {
            scheme,
            start_at,
            restart,
        }) => {
            let mut element = XmlElement::new(prefixes.drawing("buAutoNum"))
                .with_attribute("type", scheme.clone());
            if *restart || *start_at != 1 {
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
    if let Some(spacing) = properties.spacing_pt {
        element.set_attribute("spc", format_fixed(spacing * 100.0));
        present = true;
    }
    if let Some(baseline) = properties.baseline_pct {
        element.set_attribute("baseline", format_fixed(baseline * 1000.0));
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
    if let Some(caps) = properties.caps {
        element.set_attribute("cap", caps.as_attribute());
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
    next_shape_id: &mut Option<u32>,
    prefixes: &Prefixes,
    part: &str,
) -> Result<XmlElement, PptxError> {
    if matches!(add.geometry.as_str(), "custom" | "group") {
        return Err(write_error(
            part,
            format!("unsupported geometry {:?} for a new shape", add.geometry),
        ));
    }
    let shape_id = alloc_shape_id(next_shape_id, part)?;
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
            body = body.with_child(build_paragraph(paragraph, None, None, prefixes));
        }
        if paragraphs.is_empty() {
            body = body.with_child(XmlElement::new(prefixes.drawing("p")));
        }
        shape = shape.with_child(body);
    }
    Ok(shape)
}

/// A new shape, or a picture that also needs a media part and relationship.
fn add_shape_element(
    add: &ShapeAdd,
    next_shape_id: &mut Option<u32>,
    prefixes: &Prefixes,
    part: &str,
    sink: &mut PartSink<'_>,
    budget: &mut ParseBudget<'_>,
) -> Result<XmlElement, PptxError> {
    match &add.picture {
        Some(picture) => {
            picture_shape_element(add, picture, next_shape_id, prefixes, part, sink, budget)
        }
        None => shape_element(add, next_shape_id, prefixes, part),
    }
}

fn image_content_type_extension(content_type: &str) -> Option<&'static str> {
    match content_type {
        "image/png" => Some("png"),
        "image/jpeg" | "image/jpg" => Some("jpeg"),
        "image/gif" => Some("gif"),
        "image/bmp" => Some("bmp"),
        "image/tiff" => Some("tiff"),
        "image/webp" => Some("webp"),
        "image/svg+xml" => Some("svg"),
        _ => None,
    }
}

fn image_extension(content_type: &str) -> Result<&'static str, PptxError> {
    image_content_type_extension(content_type)
        .ok_or_else(|| write_error("media", format!("unsupported image type {content_type:?}")))
}

/// Whether [`write_pptx_with_edits`] can mint a media part for this MIME type.
pub fn is_supported_image_content_type(content_type: &str) -> bool {
    image_content_type_extension(content_type).is_some()
}

fn next_media_part_path(
    package: &PptxPackage,
    new_parts: &[(String, Vec<u8>)],
    extension: &str,
) -> String {
    let existing: HashSet<&str> = package
        .parts
        .iter()
        .map(|part| part.path.as_str())
        .chain(new_parts.iter().map(|(path, _)| path.as_str()))
        .collect();
    let mut number = 1_u64;
    while existing.contains(format!("ppt/media/image{number}.{extension}").as_str()) {
        number += 1;
    }
    format!("ppt/media/image{number}.{extension}")
}

/// Registers the extension as a package-wide default, unless already declared.
fn ensure_media_content_type(
    sink: &mut PartSink<'_>,
    extension: &str,
    part_path: &str,
    content_type: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<(), PptxError> {
    let path = "[Content_Types].xml";
    let bytes = sink
        .current(path)
        .ok_or_else(|| PptxError::MissingPart(path.to_owned()))?;
    let mut root = parse_xml(&bytes, path, budget)?;
    let default = root.children.iter().find_map(|child| match child {
        XmlNode::Element(element)
            if element.local_name() == "Default"
                && element
                    .attribute("Extension")
                    .is_some_and(|value| value.eq_ignore_ascii_case(extension)) =>
        {
            Some(element)
        }
        _ => None,
    });
    if let Some(default) = default {
        if default.attribute("ContentType") != Some(content_type) {
            set_content_type_override(sink, part_path, content_type, budget)?;
        }
    } else {
        let prefix = root.name.rsplit_once(':').map_or("", |(prefix, _)| prefix);
        root.children.push(XmlNode::Element(
            XmlElement::new(qualified(prefix, "Default"))
                .with_attribute("Extension", extension)
                .with_attribute("ContentType", content_type),
        ));
        sink.store(path, serialize_xml(&root));
    }
    Ok(())
}

/// Unlike [`set_relationship`], always mints a fresh one rather than
/// replacing an existing match, since a slide can carry many images.
fn add_relationship(
    sink: &mut PartSink<'_>,
    relationships_path: &str,
    relationship_type: &str,
    target: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<String, PptxError> {
    let mut root = match sink.current(relationships_path) {
        Some(bytes) => parse_xml(&bytes, relationships_path, budget)?,
        None => XmlElement::new("Relationships").with_attribute("xmlns", PACKAGE_RELATIONSHIPS_NS),
    };
    let used: HashSet<&str> = root
        .children_named("Relationship")
        .filter_map(|element| element.attribute("Id"))
        .collect();
    let mut number = 1_u64;
    while used.contains(format!("rId{number}").as_str()) {
        number += 1;
    }
    let id = format!("rId{number}");
    let prefix = root.name.rsplit_once(':').map_or("", |(prefix, _)| prefix);
    root.children.push(XmlNode::Element(
        XmlElement::new(qualified(prefix, "Relationship"))
            .with_attribute("Id", id.clone())
            .with_attribute("Type", relationship_type)
            .with_attribute("Target", target),
    ));
    sink.store(relationships_path, serialize_xml(&root));
    Ok(id)
}

fn picture_shape_element(
    add: &ShapeAdd,
    picture: &PictureAdd,
    next_shape_id: &mut Option<u32>,
    prefixes: &Prefixes,
    slide_part_path: &str,
    sink: &mut PartSink<'_>,
    budget: &mut ParseBudget<'_>,
) -> Result<XmlElement, PptxError> {
    let extension = image_extension(&picture.content_type)?;
    let media_part_path = next_media_part_path(sink.package, sink.new_parts.as_slice(), extension);
    sink.store(&media_part_path, picture.media_bytes.clone());
    ensure_media_content_type(
        sink,
        extension,
        &media_part_path,
        &picture.content_type,
        budget,
    )?;
    let relationships_path = slide_relationships_path(slide_part_path);
    let target = relative_target(slide_part_path, &media_part_path);
    let relationship_id = add_relationship(
        sink,
        &relationships_path,
        IMAGE_RELATIONSHIP_TYPE,
        &target,
        budget,
    )?;

    let shape_id = alloc_shape_id(next_shape_id, slide_part_path)?;
    let non_visual = XmlElement::new(prefixes.presentation("nvPicPr"))
        .with_child(
            XmlElement::new(prefixes.presentation("cNvPr"))
                .with_attribute("id", shape_id.to_string())
                .with_attribute("name", add.name.clone()),
        )
        .with_child(
            XmlElement::new(prefixes.presentation("cNvPicPr")).with_child(
                XmlElement::new(prefixes.drawing("picLocks")).with_attribute("noChangeAspect", "1"),
            ),
        )
        .with_child(XmlElement::new(prefixes.presentation("nvPr")));
    let blip_fill = XmlElement::new(prefixes.presentation("blipFill"))
        .with_child(
            XmlElement::new(prefixes.drawing("blip"))
                .with_attribute(prefixes.relationship("embed"), relationship_id),
        )
        .with_child(
            XmlElement::new(prefixes.drawing("stretch"))
                .with_child(XmlElement::new(prefixes.drawing("fillRect"))),
        );
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
    let geometry = XmlElement::new(prefixes.drawing("prstGeom"))
        .with_attribute("prst", "rect")
        .with_child(XmlElement::new(prefixes.drawing("avLst")));
    let properties = XmlElement::new(prefixes.presentation("spPr"))
        .with_child(transform)
        .with_child(geometry);
    Ok(XmlElement::new(prefixes.presentation("pic"))
        .with_child(non_visual)
        .with_child(blip_fill)
        .with_child(properties))
}

fn slide_xml(
    name: Option<&str>,
    shapes: &[ShapeAdd],
    part: &str,
    sink: &mut PartSink<'_>,
    budget: &mut ParseBudget<'_>,
) -> Result<Vec<u8>, PptxError> {
    let mut root = XmlElement::new("p:sld")
        .with_attribute("xmlns:a", DRAWINGML_NS)
        .with_attribute("xmlns:r", OFFICE_RELATIONSHIPS_NS)
        .with_attribute("xmlns:p", PRESENTATIONML_NS);
    let prefixes = Prefixes {
        drawing: "a".to_owned(),
        presentation: "p".to_owned(),
        relationship: "r".to_owned(),
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
    let mut next_shape_id = Some(2);
    for shape in shapes {
        tree = tree.with_child(add_shape_element(
            shape,
            &mut next_shape_id,
            &prefixes,
            part,
            sink,
            budget,
        )?);
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ParseLimits;
    use crate::drawing::parse_run_properties;
    use crate::model::ShapeElements;
    use crate::xml::{ParseBudget, parse_xml};

    #[test]
    fn a_new_gradient_outline_writes_its_stops_and_angle() {
        let package =
            crate::parse_pptx(include_bytes!("../tests/fixtures/gradient-outline.pptx")).unwrap();
        let crate::ShapeNode::Shape(shape) = &package.slides[0].shapes[0] else {
            panic!("shape")
        };
        let outline = shape.outline.as_ref().unwrap();
        let mut properties = XmlElement::new("p:spPr");
        let prefixes = Prefixes::from_root(&mut properties);
        set_outline(&mut properties, outline, &prefixes);
        let line = properties.child("ln").unwrap();
        assert_eq!(line.attribute("w"), Some("76200"));
        let gradient = line.child("gradFill").unwrap();
        assert_eq!(gradient.child("gsLst").unwrap().child_elements().count(), 3);
        assert_eq!(gradient.child("lin").unwrap().attribute("ang"), Some("0"));
        assert!(line.child("solidFill").is_none());
        assert_eq!(
            crate::drawing::parse_outline_element(line).as_ref(),
            Some(outline)
        );
    }

    #[test]
    fn source_ordinals_cover_the_shapes_an_alternate_content_contributes() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let part = "ppt/slides/slide1.xml";
        let mut root = parse_xml(
            br#"<p:sld><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="one"/></p:nvSpPr></p:sp><mc:AlternateContent><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:sp><p:nvSpPr><p:cNvPr id="3" name="choice"/></p:nvSpPr></p:sp></mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="4" name="two"/></p:nvSpPr></p:sp><p:pic><p:nvPicPr><p:cNvPr id="5" name="three"/></p:nvPicPr></p:pic></mc:Fallback></mc:AlternateContent><p:sp><p:nvSpPr><p:cNvPr id="6" name="four"/></p:nvSpPr></p:sp></p:spTree></p:cSld></p:sld>"#,
            part,
            &mut budget,
        )
        .unwrap();
        let parsed = crate::drawing::common_slide_data(
            &root,
            &[],
            part,
            &mut budget,
            ShapeElements::WithConnectors,
        )
        .unwrap();
        let tree = root
            .child_mut("cSld")
            .and_then(|common| common.child_mut("spTree"))
            .unwrap();
        let slots: Vec<Option<XmlNode>> = std::mem::take(&mut tree.children)
            .into_iter()
            .map(Some)
            .collect();

        let found = shape_slots(&slots, ShapeElements::WithConnectors);

        assert_eq!(found.len(), parsed.shapes.len());
        let names: Vec<String> = found
            .iter()
            .map(|slot| {
                let mut node = slots[slot.position].clone().unwrap();
                let element = element_at_path(&mut node, &slot.path).unwrap();
                element.descendants_named("cNvPr")[0]
                    .attribute("name")
                    .unwrap()
                    .to_owned()
            })
            .collect();
        assert_eq!(names, ["one", "two", "three", "four"]);
    }

    #[test]
    fn deleting_one_of_the_shapes_a_branch_holds_leaves_the_others_in_place() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let part = "ppt/slides/slide1.xml";
        let mut root = parse_xml(
            br#"<p:sld><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="one"/></p:nvSpPr><p:spPr/></p:sp><mc:AlternateContent><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="3" name="two"/></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="three"/></p:nvSpPr><p:spPr/></p:sp></mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#,
            part,
            &mut budget,
        )
        .unwrap();

        let package = PptxPackage::default();
        let mut replacements = HashMap::new();
        let mut new_parts = Vec::new();
        let mut sink = PartSink {
            package: &package,
            replacements: &mut replacements,
            new_parts: &mut new_parts,
        };
        patch_slide(
            &mut root,
            &[
                ShapeWrite::Keep { source_index: 0 },
                ShapeWrite::Patch {
                    source_index: 2,
                    patch: Box::new(ShapePatch {
                        fill: Some(ShapeFill {
                            color: Some(ColorValue::from_attribute("DC2626")),
                            ..ShapeFill::named("solid")
                        }),
                        ..ShapePatch::default()
                    }),
                },
            ],
            None,
            part,
            ShapeElements::WithConnectors,
            &mut sink,
            &mut budget,
        )
        .unwrap();

        let xml = String::from_utf8(serialize_xml(&root)).unwrap();
        assert!(xml.contains(r#"name="one""#));
        assert!(!xml.contains(r#"name="two""#), "{xml}");
        let survivor = xml.split(r#"name="three""#).nth(1).unwrap();
        assert!(survivor.contains(r#"<a:srgbClr val="DC2626"/>"#), "{xml}");
    }

    #[test]
    fn a_branch_that_still_holds_unmodeled_markup_outlives_its_only_shape() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let part = "ppt/slides/slide1.xml";
        let mut root = parse_xml(
            br#"<p:sld><p:cSld><p:spTree><mc:AlternateContent><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="2" name="one"/></p:nvSpPr></p:sp><p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="two"/></p:nvCxnSpPr></p:cxnSp></mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#,
            part,
            &mut budget,
        )
        .unwrap();

        let package = PptxPackage::default();
        let mut replacements = HashMap::new();
        let mut new_parts = Vec::new();
        let mut sink = PartSink {
            package: &package,
            replacements: &mut replacements,
            new_parts: &mut new_parts,
        };
        patch_slide(
            &mut root,
            &[],
            None,
            part,
            ShapeElements::WithoutConnectors,
            &mut sink,
            &mut budget,
        )
        .unwrap();

        let xml = String::from_utf8(serialize_xml(&root)).unwrap();
        assert!(!xml.contains(r#"name="one""#), "{xml}");
        assert!(xml.contains(r#"name="two""#), "{xml}");
    }

    fn run_properties(xml: &[u8]) -> (XmlElement, Prefixes, RunProperties) {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let mut root = parse_xml(xml, "ppt/slides/slide1.xml", &mut budget).unwrap();
        let prefixes = Prefixes::from_root(&mut root);
        let properties = parse_run_properties(Some(&root));
        (root, prefixes, properties)
    }

    #[test]
    fn an_unedited_gradient_fill_survives_a_rewrite() {
        let (mut element, prefixes, properties) = run_properties(
            br#"<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" lang="en"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst></a:gradFill></a:rPr>"#,
        );
        assert!(
            properties.color.is_some(),
            "the gradient should flatten to a colour"
        );
        let gradient = element.child("gradFill").unwrap().clone();
        apply_run_properties(&mut element, &properties, None, &prefixes);
        assert_eq!(element.child("gradFill"), Some(&gradient));
        let serialized = String::from_utf8(serialize_xml(&element)).unwrap();
        assert!(
            serialized.contains("gradFill"),
            "the authored gradient must be left alone"
        );
        assert!(
            !serialized.contains("solidFill"),
            "it must not be rewritten as a solid fill"
        );
    }

    #[test]
    fn an_edited_colour_replaces_the_gradient() {
        let (mut element, prefixes, mut properties) = run_properties(
            br#"<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" lang="en"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst></a:gradFill></a:rPr>"#,
        );
        properties.color = Some(ColorValue {
            rgb: Some("FF0000".to_owned()),
            ..ColorValue::default()
        });
        apply_run_properties(&mut element, &properties, None, &prefixes);
        let serialized = String::from_utf8(serialize_xml(&element)).unwrap();
        assert!(!serialized.contains("gradFill"));
        assert!(serialized.contains("solidFill"));
        assert_eq!(parse_run_properties(Some(&element)).color, properties.color);
    }

    #[test]
    fn the_writer_recognises_exactly_what_the_parser_reads() {
        for local in ["sp", "cxnSp", "pic", "graphicFrame", "grpSp"] {
            assert!(
                ShapeElements::WithConnectors.contains(local),
                "{local} is parsed as a shape but not recognised here"
            );
        }
        for local in ["txBody", "nvGrpSpPr", "extLst"] {
            assert!(
                !ShapeElements::WithConnectors.contains(local),
                "{local} is not a shape element"
            );
        }
    }

    #[test]
    fn the_pre_connector_set_leaves_out_only_connectors() {
        assert!(!ShapeElements::WithoutConnectors.contains("cxnSp"));
        for local in ["sp", "pic", "graphicFrame", "grpSp"] {
            assert!(ShapeElements::WithoutConnectors.contains(local));
        }
    }

    const LINKED_PARAGRAPH: &[u8] = br#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:r><a:rPr lang="en-US" dirty="0"/><a:t>See </a:t></a:r><a:r><a:rPr lang="en-US" strike="sngStrike"><a:hlinkClick r:id="rId2"/></a:rPr><a:t>the docs</a:t></a:r><a:r><a:rPr lang="en-US"/><a:t> today</a:t></a:r></a:p>"#;
    const LINK_PROPERTIES: &str =
        r#"<a:rPr lang="en-US" strike="sngStrike"><a:hlinkClick r:id="rId2"/></a:rPr>"#;
    const FIELD_PARAGRAPH: &[u8] = br#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:fld id="{A}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld><a:r><a:rPr lang="en-US"/><a:t> of </a:t></a:r><a:fld id="{B}" type="datetime1"><a:rPr lang="en-US"/><a:t>2024</a:t></a:fld></a:p>"#;

    fn rebuilt_paragraph(xml: &[u8], runs: &[(&str, RunProperties)]) -> String {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let mut root = parse_xml(xml, "ppt/slides/slide1.xml", &mut budget).unwrap();
        let prefixes = Prefixes::from_root(&mut root);
        let write = ParagraphWrite {
            source_index: Some(0),
            rebuild: true,
            properties_changed: false,
            alignment: None,
            level: 0,
            bullet: None,
            runs: runs
                .iter()
                .map(|(text, properties)| RunWrite {
                    text: (*text).to_owned(),
                    properties: properties.clone(),
                })
                .collect(),
        };
        let paragraph = build_paragraph(&write, Some(root), None, &prefixes);
        String::from_utf8(serialize_xml(&paragraph)).unwrap()
    }

    #[test]
    fn an_edit_spanning_several_runs_keeps_each_survivor_on_its_source_run() {
        let xml = rebuilt_paragraph(
            LINKED_PARAGRAPH,
            &[("Seethe doc now", RunProperties::default())],
        );
        assert!(
            xml.contains(r#"<a:r><a:rPr dirty="0" lang="en-US"/><a:t>See</a:t></a:r>"#),
            "{xml}"
        );
        assert!(
            xml.contains(&format!("<a:r>{LINK_PROPERTIES}<a:t>the doc</a:t></a:r>")),
            "{xml}"
        );
        assert!(
            xml.contains(r#"<a:r><a:rPr lang="en-US"/><a:t> now</a:t></a:r>"#),
            "{xml}"
        );
        assert_eq!(xml.matches("<a:r>").count(), 3, "{xml}");
    }

    #[test]
    fn a_line_break_inside_the_span_keeps_the_runs_around_it_aligned() {
        let xml = rebuilt_paragraph(
            br#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:r><a:rPr lang="en-US"/><a:t>First</a:t></a:r><a:br/><a:r><a:rPr><a:hlinkClick r:id="rId2"/></a:rPr><a:t>Second</a:t></a:r></a:p>"#,
            &[("Firs\nSecond!", RunProperties::default())],
        );
        assert!(
            xml.contains(
                r#"<a:r><a:rPr lang="en-US"/><a:t>Firs</a:t></a:r><a:br/><a:r><a:rPr><a:hlinkClick r:id="rId2"/></a:rPr><a:t>Second!</a:t></a:r>"#
            ),
            "{xml}"
        );
    }

    #[test]
    fn a_field_survives_an_edit_around_it_and_degrades_once_its_text_changes() {
        let xml = rebuilt_paragraph(
            FIELD_PARAGRAPH,
            &[("1 out of 2025", RunProperties::default())],
        );
        assert!(
            xml.contains(
                r#"<a:fld id="{A}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld><a:r><a:rPr lang="en-US"/><a:t> out of </a:t></a:r><a:r><a:rPr lang="en-US"/><a:t>2025</a:t></a:r>"#
            ),
            "{xml}"
        );
        assert!(!xml.contains("datetime1"), "{xml}");

        let xml = rebuilt_paragraph(FIELD_PARAGRAPH, &[("1X of 2024", RunProperties::default())]);
        assert!(
            xml.contains(
                r#"<a:fld id="{A}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld><a:r><a:rPr lang="en-US"/><a:t>X</a:t></a:r><a:r><a:rPr lang="en-US"/><a:t> of </a:t></a:r><a:fld id="{B}" type="datetime1"><a:rPr lang="en-US"/><a:t>2024</a:t></a:fld>"#
            ),
            "{xml}"
        );
    }

    #[test]
    fn text_typed_at_the_end_of_a_link_stays_inside_it() {
        let xml = rebuilt_paragraph(
            LINKED_PARAGRAPH,
            &[("See the docsX today", RunProperties::default())],
        );
        assert!(
            xml.contains(&format!("<a:r>{LINK_PROPERTIES}<a:t>the docsX</a:t></a:r>")),
            "{xml}"
        );

        let bold = RunProperties {
            bold: Some(true),
            ..RunProperties::default()
        };
        let xml = rebuilt_paragraph(
            LINKED_PARAGRAPH,
            &[
                ("See ", RunProperties::default()),
                ("the docs", RunProperties::default()),
                ("X", bold),
                (" today", RunProperties::default()),
            ],
        );
        assert!(
            xml.contains(&format!(
                r#"<a:r>{LINK_PROPERTIES}<a:t>the docs</a:t></a:r><a:r><a:rPr b="1" lang="en-US" strike="sngStrike"><a:hlinkClick r:id="rId2"/></a:rPr><a:t>X</a:t></a:r><a:r><a:rPr lang="en-US"/><a:t> today</a:t></a:r>"#
            )),
            "{xml}"
        );
    }

    #[test]
    fn deleting_a_linked_run_drops_its_link() {
        let xml = rebuilt_paragraph(
            LINKED_PARAGRAPH,
            &[("See  today", RunProperties::default())],
        );
        assert!(!xml.contains("hlinkClick"), "{xml}");
        assert!(
            xml.contains(
                r#"<a:r><a:rPr dirty="0" lang="en-US"/><a:t>See </a:t></a:r><a:r><a:rPr lang="en-US"/><a:t> today</a:t></a:r>"#
            ),
            "{xml}"
        );
    }

    #[test]
    fn a_plain_duplicate_at_the_front_does_not_steal_the_field() {
        let xml = rebuilt_paragraph(
            br#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:rPr lang="en-US"/><a:t>1</a:t></a:r><a:fld id="{A}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld></a:p>"#,
            &[("1", RunProperties::default())],
        );
        assert!(
            xml.contains(
                r#"<a:fld id="{A}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld>"#
            ),
            "{xml}"
        );
        assert_eq!(xml.matches("<a:r>").count(), 0, "{xml}");
    }

    #[test]
    fn a_plain_duplicate_at_the_back_does_not_steal_the_field() {
        let xml = rebuilt_paragraph(
            br#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:rPr lang="en-US"/><a:t>X</a:t></a:r><a:fld id="{A}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld><a:r><a:rPr lang="en-US"/><a:t>1</a:t></a:r></a:p>"#,
            &[
                ("X", RunProperties::default()),
                ("1", RunProperties::default()),
            ],
        );
        assert!(
            xml.contains(
                r#"<a:fld id="{A}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld>"#
            ),
            "{xml}"
        );
        assert_eq!(xml.matches("<a:r>").count(), 1, "{xml}");
        assert!(
            !xml.contains(r#"<a:r><a:rPr lang="en-US"/><a:t>1</a:t></a:r>"#),
            "{xml}"
        );
    }

    #[test]
    fn adjustments_without_preset_geometry_error_instead_of_no_opting() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let part = "ppt/slides/slide1.xml";
        let mut root = parse_xml(
            br#"<p:sld><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="custom"/></p:nvSpPr><p:spPr><a:custGeom><a:pathLst><a:path w="10" h="10"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:close/></a:path></a:pathLst></a:custGeom></p:spPr></p:sp></p:spTree></p:cSld></p:sld>"#,
            part,
            &mut budget,
        )
        .unwrap();

        let package = PptxPackage::default();
        let mut replacements = HashMap::new();
        let mut new_parts = Vec::new();
        let mut sink = PartSink {
            package: &package,
            replacements: &mut replacements,
            new_parts: &mut new_parts,
        };
        let error = patch_slide(
            &mut root,
            &[ShapeWrite::Patch {
                source_index: 0,
                patch: Box::new(ShapePatch {
                    adjust_values: Some(BTreeMap::from([("adj".to_owned(), 0.25)])),
                    ..ShapePatch::default()
                }),
            }],
            None,
            part,
            ShapeElements::WithConnectors,
            &mut sink,
            &mut budget,
        )
        .unwrap_err();
        assert!(
            matches!(error, crate::PptxError::Write { ref message, .. } if message.contains("preset geometry")),
            "{error:?}"
        );
    }
}
