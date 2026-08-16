//! Baseline-diff write-back: the live CRDT state is compared against a
//! freshly seeded copy of the source package and only differences are written.

use std::collections::HashMap;

use ooxml_drawingml::{ColorValue, ShapeFill};
use pptx_parse::{
    Bullet, DeckWrite, InheritedTransform, ParagraphWrite, Placeholder, PptxPackage, RunProperties,
    RunWrite, ShapeAdd, ShapeNode, ShapePatch, ShapeTransform, ShapeWrite, SlideLayout,
    SlideMaster, SlideWrite, TextTarget, TextWrite,
};

use crate::deck::{seed_doc, snapshot_doc};
use crate::{
    BOOTSTRAP_CLIENT_ID, DeckSession, DeckSnapshot, EditError, EditResult, ParagraphSnapshot,
    ShapeKind, ShapeSnapshot, SlideSnapshot, StorySnapshot, TextRunSnapshot, doc_with_client_id,
};

/// The parsed parts a slide's shapes may inherit geometry from.
struct SlideContext<'a> {
    layout: Option<&'a SlideLayout>,
    master: Option<&'a SlideMaster>,
}

impl<'a> SlideContext<'a> {
    fn new(package: &'a PptxPackage, snapshot: &SlideSnapshot) -> Self {
        let layout = snapshot
            .layout_part_path
            .as_deref()
            .and_then(|path| {
                package
                    .layouts
                    .iter()
                    .find(|layout| layout.part_path == path)
            })
            .or_else(|| package.layouts.first());
        let master = layout
            .and_then(|layout| layout.master_part_path.as_deref())
            .and_then(|path| {
                package
                    .masters
                    .iter()
                    .find(|master| master.part_path == path)
            })
            .or_else(|| {
                layout.and_then(|layout| {
                    package.masters.iter().find(|master| {
                        master
                            .layout_part_paths
                            .iter()
                            .any(|path| path == &layout.part_path)
                    })
                })
            })
            .or_else(|| package.masters.first());
        Self { layout, master }
    }
}

impl DeckSession {
    /// Serializes the deck with all edits applied. Untouched slides keep their
    /// exact source part bytes; edited slides are patched at the XML level.
    pub fn save(&self) -> EditResult<Vec<u8>> {
        if !self.package.has_parts() {
            return Err(EditError::Write(
                "this session carries no source file bytes; open it from the \
                 file, or from the update plus the source file, to save"
                    .to_owned(),
            ));
        }
        let current = self.snapshot()?;
        let baseline = baseline_snapshot(&self.package)?;
        if current == baseline {
            return pptx_parse::write_pptx(&self.package)
                .map_err(|error| EditError::Write(error.to_string()));
        }
        let deck = deck_write(&current, &baseline, &self.package)?;
        pptx_parse::write_pptx_with_edits(&self.package, &deck)
            .map_err(|error| EditError::Write(error.to_string()))
    }
}

fn baseline_snapshot(package: &PptxPackage) -> EditResult<DeckSnapshot> {
    let doc = doc_with_client_id(BOOTSTRAP_CLIENT_ID);
    seed_doc(&doc, package, "")?;
    snapshot_doc(&doc, package)
}

fn deck_write(
    current: &DeckSnapshot,
    baseline: &DeckSnapshot,
    package: &PptxPackage,
) -> EditResult<DeckWrite> {
    let baseline_slides: HashMap<&str, &SlideSnapshot> = baseline
        .slides
        .iter()
        .map(|slide| (slide.id.as_str(), slide))
        .collect();
    let mut slides = Vec::with_capacity(current.slides.len());
    for slide in &current.slides {
        let write = match baseline_slides.get(slide.id.as_str()) {
            Some(base) if *base == slide => SlideWrite::Keep {
                part_path: source_part_path(slide)?,
            },
            Some(base) => SlideWrite::Patch {
                part_path: source_part_path(slide)?,
                shapes: shape_writes(
                    &slide.shapes,
                    &base.shapes,
                    &SlideContext::new(package, slide),
                )?,
            },
            None => SlideWrite::Add {
                name: slide.name.clone(),
                layout_part_path: slide.layout_part_path.clone(),
                shapes: slide
                    .shapes
                    .iter()
                    .map(shape_add)
                    .collect::<EditResult<Vec<_>>>()?,
            },
        };
        slides.push(write);
    }
    Ok(DeckWrite { slides })
}

fn source_part_path(slide: &SlideSnapshot) -> EditResult<String> {
    slide
        .source_part_path
        .clone()
        .ok_or_else(|| EditError::Write(format!("slide {} has no source part", slide.id)))
}

fn shape_writes(
    current: &[ShapeSnapshot],
    baseline: &[ShapeSnapshot],
    context: &SlideContext<'_>,
) -> EditResult<Vec<ShapeWrite>> {
    let baseline_shapes: HashMap<&str, &ShapeSnapshot> = baseline
        .iter()
        .map(|shape| (shape.id.as_str(), shape))
        .collect();
    let mut writes = Vec::with_capacity(current.len());
    for shape in current {
        let write = match baseline_shapes.get(shape.id.as_str()) {
            Some(base) if *base == shape => ShapeWrite::Keep {
                source_index: source_index(&shape.id)?,
            },
            Some(base) => ShapeWrite::Patch {
                source_index: source_index(&shape.id)?,
                patch: Box::new(shape_patch(shape, base, context)?),
            },
            None => ShapeWrite::Add(Box::new(shape_add(shape)?)),
        };
        writes.push(write);
    }
    Ok(writes)
}

fn source_index(shape_id: &str) -> EditResult<usize> {
    shape_id
        .rfind(":shape:")
        .map(|position| &shape_id[position + ":shape:".len()..])
        .and_then(|path| path.rsplit('.').next())
        .and_then(|segment| segment.parse().ok())
        .ok_or_else(|| EditError::Write(format!("shape {shape_id} has no source index")))
}

fn shape_patch(
    shape: &ShapeSnapshot,
    base: &ShapeSnapshot,
    context: &SlideContext<'_>,
) -> EditResult<ShapePatch> {
    let mut patch = ShapePatch::default();
    let moved = (shape.x, shape.y) != (base.x, base.y);
    let resized = (shape.width, shape.height) != (base.width, base.height);
    if moved || resized {
        patch.offset = moved.then_some((shape.x, shape.y));
        patch.extent = resized.then_some((shape.width, shape.height));
        let source_inherited = base.width <= 0 || base.height <= 0;
        patch.inherited = source_inherited.then(|| {
            inherited_transform(shape, context)
                .map(|transform| InheritedTransform {
                    x: transform.x,
                    y: transform.y,
                    width: transform.width,
                    height: transform.height,
                    rotation_deg: transform.rotation_deg,
                    flip_horizontal: transform.flip_h,
                    flip_vertical: transform.flip_v,
                })
                .unwrap_or(InheritedTransform {
                    x: shape.x,
                    y: shape.y,
                    width: shape.width,
                    height: shape.height,
                    rotation_deg: shape.rotation_deg,
                    flip_horizontal: shape.flip_h,
                    flip_vertical: shape.flip_v,
                })
        });
    }
    if shape.fill != base.fill {
        patch.fill = Some(
            shape
                .fill
                .clone()
                .unwrap_or_else(|| ShapeFill::named("none")),
        );
    }
    if shape.outline != base.outline {
        patch.outline = Some(shape.outline.clone().unwrap_or_default());
    }
    if shape.adjust_values != base.adjust_values {
        patch.adjust_values = Some(shape.adjust_values.clone());
    }
    let baseline_stories: HashMap<&str, &StorySnapshot> = base
        .text_stories
        .iter()
        .map(|story| (story.id.as_str(), story))
        .collect();
    for story in &shape.text_stories {
        let base_story = baseline_stories.get(story.id.as_str()).copied();
        if base_story == Some(story) {
            continue;
        }
        patch.texts.push(TextWrite {
            target: text_target(&story.id, &shape.id)?,
            paragraphs: paragraph_writes(story, base_story)?,
        });
    }
    if shape.children != base.children {
        patch.children = shape_writes(&shape.children, &base.children, context)?;
    }
    Ok(patch)
}

/// The transform a placeholder inherits: the layout's matching placeholder,
/// then the master's. The shape's own parsed node cannot contribute — a
/// positive extent there would already be in the snapshot.
fn inherited_transform<'a>(
    shape: &ShapeSnapshot,
    context: &SlideContext<'a>,
) -> Option<&'a ShapeTransform> {
    let layout = shape.placeholder.as_ref().and_then(|placeholder| {
        context
            .layout
            .and_then(|layout| find_placeholder(&layout.shapes, placeholder))
    });
    let master = shape.placeholder.as_ref().and_then(|placeholder| {
        context
            .master
            .and_then(|master| find_placeholder(&master.shapes, placeholder))
    });
    [layout, master]
        .into_iter()
        .flatten()
        .map(node_transform)
        .find(|transform| transform.width > 0 && transform.height > 0)
}

fn find_placeholder<'a>(nodes: &'a [ShapeNode], target: &Placeholder) -> Option<&'a ShapeNode> {
    for node in nodes {
        if node_placeholder(node).is_some_and(|value| placeholders_match(value, target)) {
            return Some(node);
        }
        if let ShapeNode::Group(group) = node
            && let Some(found) = find_placeholder(&group.children, target)
        {
            return Some(found);
        }
    }
    None
}

fn placeholders_match(left: &Placeholder, right: &Placeholder) -> bool {
    match (left.index, right.index) {
        (Some(left), Some(right)) => left == right,
        _ => {
            normalize_placeholder_type(left.placeholder_type.as_deref())
                == normalize_placeholder_type(right.placeholder_type.as_deref())
        }
    }
}

fn normalize_placeholder_type(value: Option<&str>) -> &str {
    match value.unwrap_or("body") {
        "ctrTitle" => "title",
        "obj" => "body",
        value => value,
    }
}

fn node_placeholder(node: &ShapeNode) -> Option<&Placeholder> {
    node_base(node).placeholder.as_ref()
}

fn node_transform(node: &ShapeNode) -> &ShapeTransform {
    &node_base(node).transform
}

fn node_base(node: &ShapeNode) -> &pptx_parse::ShapeBase {
    match node {
        ShapeNode::Shape(shape) => &shape.base,
        ShapeNode::Picture(shape) => &shape.base,
        ShapeNode::GraphicFrame(shape) => &shape.base,
        ShapeNode::Group(shape) => &shape.base,
    }
}

fn text_target(story_id: &str, shape_id: &str) -> EditResult<TextTarget> {
    let suffix = story_id
        .strip_prefix("story:")
        .and_then(|rest| rest.strip_prefix(shape_id))
        .and_then(|rest| rest.strip_prefix(':'))
        .ok_or_else(|| EditError::Write(format!("story {story_id} does not match its shape")))?;
    if let Some(cell) = suffix.strip_prefix("table:") {
        let (row, cell) = cell
            .split_once(':')
            .and_then(|(row, cell)| Some((row.parse().ok()?, cell.parse().ok()?)))
            .ok_or_else(|| EditError::Write(format!("story {story_id} has no table address")))?;
        return Ok(TextTarget::TableCell { row, cell });
    }
    Ok(TextTarget::Body)
}

fn paragraph_writes(
    story: &StorySnapshot,
    baseline: Option<&StorySnapshot>,
) -> EditResult<Vec<ParagraphWrite>> {
    let baseline_paragraphs: HashMap<&str, &ParagraphSnapshot> = baseline
        .into_iter()
        .flat_map(|story| &story.paragraphs)
        .map(|paragraph| (paragraph.id.as_str(), paragraph))
        .collect();
    let prefix = format!("para:{}:", story.id);
    let mut writes = Vec::with_capacity(story.paragraphs.len());
    for paragraph in &story.paragraphs {
        let source_index = paragraph
            .id
            .strip_prefix(&prefix)
            .and_then(|index| index.parse::<usize>().ok());
        let base = baseline_paragraphs.get(paragraph.id.as_str()).copied();
        if base == Some(paragraph) && source_index.is_some() {
            writes.push(ParagraphWrite {
                source_index,
                rebuild: false,
                properties_changed: false,
                alignment: None,
                level: 0,
                bullet: None,
                runs: Vec::new(),
            });
            continue;
        }
        let properties_changed = match base {
            Some(base) => {
                base.alignment != paragraph.alignment
                    || base.level != paragraph.level
                    || base.bullet_json != paragraph.bullet_json
            }
            None => true,
        };
        let bullet = paragraph
            .bullet_json
            .as_deref()
            .map(serde_json::from_str::<Bullet>)
            .transpose()
            .map_err(|error| EditError::Json(error.to_string()))?;
        writes.push(ParagraphWrite {
            source_index,
            rebuild: true,
            properties_changed,
            alignment: paragraph.alignment.clone(),
            level: paragraph.level,
            bullet,
            runs: paragraph.runs.iter().map(run_write).collect(),
        });
    }
    Ok(writes)
}

fn run_write(run: &TextRunSnapshot) -> RunWrite {
    RunWrite {
        text: run.text.clone(),
        properties: RunProperties {
            font_size_pt: run.style.font_size_pt,
            bold: run.style.bold,
            italic: run.style.italic,
            underline: run.style.underline.clone(),
            font_family: run.style.font_family.clone(),
            color: run.style.color.as_deref().map(color_from_hex),
            language: None,
            hyperlink_relationship_id: None,
        },
    }
}

fn color_from_hex(color: &str) -> ColorValue {
    ColorValue {
        rgb: Some(
            color
                .strip_prefix('#')
                .unwrap_or(color)
                .to_ascii_uppercase(),
        ),
        ..ColorValue::default()
    }
}

fn shape_add(shape: &ShapeSnapshot) -> EditResult<ShapeAdd> {
    if shape.kind != ShapeKind::Shape {
        return Err(EditError::Write(format!(
            "shape {} cannot be written as a new shape",
            shape.id
        )));
    }
    let paragraphs = shape
        .text_stories
        .first()
        .map(|story| paragraph_writes(story, None))
        .transpose()?;
    Ok(ShapeAdd {
        name: shape.name.clone(),
        geometry: shape.geometry.clone(),
        x: shape.x,
        y: shape.y,
        width: shape.width,
        height: shape.height,
        adjust_values: shape.adjust_values.clone(),
        fill: shape.fill.clone(),
        outline: shape.outline.clone(),
        paragraphs,
    })
}
