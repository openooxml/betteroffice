//! Font-independent paragraph semantics shared by rendering and export: the layout and master a
//! slide inherits from, the paragraph-property cascade, and list markers.

use std::cell::{Cell, RefCell};
use std::collections::{HashMap, HashSet};

use ooxml_drawingml::{
    ColorValue, ResolvedCellStyle, TableCellPosition, TableStyleFlags, Theme,
    resolve_color_value_to_hex_with_theme, resolve_theme_font_ref,
};
use pptx_parse::{
    Bullet, GraphicFrameData, ParagraphProperties, Placeholder, PptxPackage, Presentation,
    RunProperties, ShapeNode, Slide, SlideLayout, SlideMaster, TextBody, TextParagraph, TextRun,
    builtin_table_style,
};
use yrs::Transact;

use crate::deck::{map_string, required_map, shape_base, shape_parts, slide_ref};
use crate::{
    DeckSession, EditError, EditResult, ParagraphSnapshot, SHAPES, ShapeSnapshot, SlideSnapshot,
    StorySnapshot, TextRunSnapshot, TextStyle,
};

impl DeckSession {
    pub(crate) fn inherited_story_strikes(&self, story_id: &str) -> EditResult<Vec<bool>> {
        fn find_shape<'a>(
            shapes: &'a [ShapeSnapshot],
            story_id: &str,
        ) -> Option<&'a ShapeSnapshot> {
            shapes.iter().find_map(|shape| {
                if shape.text_stories.iter().any(|story| story.id == story_id) {
                    Some(shape)
                } else {
                    find_shape(&shape.children, story_id)
                }
            })
        }
        for slide in &self.snapshot()?.slides {
            if let Some(shape) = find_shape(&slide.shapes, story_id) {
                let story = shape
                    .text_stories
                    .iter()
                    .find(|story| story.id == story_id)
                    .unwrap();
                return Ok(inherited_strikes(self.package(), slide, shape, story));
            }
        }
        Err(EditError::StoryNotFound(story_id.to_owned()))
    }

    pub fn story_formatting(
        &self,
        slide_id: &str,
        shape_id: &str,
        story_id: &str,
    ) -> EditResult<StorySnapshot> {
        let txn = self.doc.transact();
        let slide = slide_ref(&txn, slide_id)?;
        let source_path = map_string(&slide, &txn, "sourcePartPath");
        let layout_path = map_string(&slide, &txn, "layoutPartPath");
        let package = self.package();
        let theme =
            pptx_parse::slide_theme(package, source_path.as_deref(), layout_path.as_deref());
        let parents =
            SlideParents::resolve(package, source_path.as_deref(), layout_path.as_deref());
        let shape = shape_parts(&required_map(&txn, SHAPES)?, &txn, shape_id, Some(&theme))?;
        if !shape.story_ids.iter().any(|id| id == story_id) {
            return Err(EditError::StoryNotFound(story_id.to_owned()));
        }
        let node = parents
            .slide
            .and_then(|slide| find_source_node(&slide.shapes, shape.snapshot.source_id));
        let placeholder = shape.snapshot.placeholder.as_ref();
        let mut cascade = ParagraphCascade {
            primary: node.and_then(node_text),
            layout: parents
                .layout
                .and_then(|layout| placeholder_text(&layout.shapes, placeholder)),
            master: parents
                .master
                .and_then(|master| placeholder_text(&master.shapes, placeholder)),
            master_slide: parents.master,
            default_style: &package.presentation.default_text_style,
            default_paragraph: package.presentation.default_text_paragraph.as_deref(),
            placeholder,
            style_color: match node {
                Some(ShapeNode::Shape(shape)) => shape
                    .style
                    .as_ref()
                    .and_then(|style| style.font_color.as_ref()),
                _ => None,
            },
        };
        let cell = cell_formatting(package, parents.master, node, &shape.snapshot, story_id);
        if let Some((text, inherited)) = &cell {
            cascade = ParagraphCascade {
                primary: Some(text),
                master: Some(inherited),
                ..ParagraphCascade::default()
            };
        }
        let mut story = self.story(story_id)?;
        let story_id = story.id.clone();
        for (index, paragraph) in story.paragraphs.iter_mut().enumerate() {
            let mut properties = cascade.properties(index, paragraph.level, None);
            let strike = paragraph_strike(&cascade, &story_id, paragraph);
            if properties.default_run.is_some() || strike.is_some() {
                properties
                    .default_run
                    .get_or_insert_with(RunProperties::default)
                    .strike = strike;
            }
            paragraph.alignment = paragraph.alignment.clone().or(properties.alignment);
            if paragraph.runs.is_empty() {
                let end = [cascade.primary, cascade.layout, cascade.master]
                    .into_iter()
                    .flatten()
                    .find_map(|body| body.paragraphs.get(index)?.end_properties.as_ref());
                if let Some(end) = end {
                    merge_run_properties(
                        properties
                            .default_run
                            .get_or_insert_with(RunProperties::default),
                        end,
                    );
                }
                paragraph.runs.push(TextRunSnapshot {
                    text: String::new(),
                    style: TextStyle::default(),
                });
            }
            for run in &mut paragraph.runs {
                run.style =
                    inherited_text_style(&run.style, properties.default_run.as_ref(), &theme);
            }
        }
        Ok(story)
    }
}

pub(crate) fn inherited_strikes(
    package: &PptxPackage,
    slide: &SlideSnapshot,
    shape: &ShapeSnapshot,
    story: &StorySnapshot,
) -> Vec<bool> {
    let parents = SlideParents::resolve(
        package,
        slide.source_part_path.as_deref(),
        slide.layout_part_path.as_deref(),
    );
    let node = parents
        .slide
        .and_then(|slide| find_source_node(&slide.shapes, shape.source_id));
    let cell = cell_formatting(package, parents.master, node, shape, &story.id);
    let cascade = if let Some((text, inherited)) = &cell {
        ParagraphCascade {
            primary: Some(text),
            master: Some(inherited),
            ..ParagraphCascade::default()
        }
    } else {
        shape_cascade(
            &parents,
            &package.presentation,
            shape.placeholder.as_ref(),
            node.and_then(node_text),
        )
    };
    story
        .paragraphs
        .iter()
        .map(|paragraph| {
            paragraph_strike(&cascade, &story.id, paragraph)
                .is_some_and(|strike| matches!(strike.as_str(), "sngStrike" | "dblStrike"))
        })
        .collect()
}

fn paragraph_strike(
    cascade: &ParagraphCascade<'_>,
    story_id: &str,
    paragraph: &ParagraphSnapshot,
) -> Option<String> {
    let source = paragraph.source_paragraphs.first();
    let xml = match source {
        Some(source) => source.xml.as_deref(),
        None => paragraph.preserved_xml.as_deref(),
    };
    let source_index = source
        .map(|source| source.id.as_str())
        .unwrap_or(&paragraph.id)
        .strip_prefix(&format!("para:{story_id}:"))
        .and_then(|index| index.parse::<usize>().ok());
    let own = if let Some(xml) = xml {
        pptx_parse::text_paragraph_properties(xml).ok()
    } else if source.is_none() {
        cascade
            .primary
            .and_then(|body| source_index.and_then(|index| body.paragraphs.get(index)))
            .map(|paragraph| paragraph.properties.clone())
    } else {
        None
    };
    let primary = TextBody {
        list_style: cascade
            .primary
            .map(|body| body.list_style.clone())
            .unwrap_or_default(),
        default_list_style: cascade
            .primary
            .and_then(|body| body.default_list_style.clone()),
        ..TextBody::default()
    };
    let inherited = ParagraphCascade {
        primary: Some(&primary),
        ..*cascade
    }
    .properties(source_index.unwrap_or(usize::MAX), paragraph.level, None)
    .default_run
    .and_then(|properties| properties.strike);
    own.and_then(|properties| properties.default_run)
        .and_then(|properties| properties.strike)
        .or(inherited)
}

fn inherited_text_style(
    direct: &TextStyle,
    fallback: Option<&RunProperties>,
    theme: &Theme,
) -> TextStyle {
    let fallback = fallback.cloned().unwrap_or_default();
    let family = direct
        .font_family
        .as_deref()
        .or(fallback.font_family.as_deref())
        .unwrap_or("+mn-lt");
    TextStyle {
        bold: Some(direct.bold.or(fallback.bold).unwrap_or(false)),
        italic: Some(direct.italic.or(fallback.italic).unwrap_or(false)),
        font_size_pt: Some(
            direct
                .font_size_pt
                .or(fallback.font_size_pt)
                .unwrap_or(18.0),
        ),
        kern_pt: direct.kern_pt.or(fallback.kern_pt),
        font_family: Some(if family.starts_with('+') {
            resolve_theme_font_ref(Some(theme), family)
        } else {
            family.to_owned()
        }),
        color: direct
            .color
            .clone()
            .or_else(|| resolve_color_value_to_hex_with_theme(fallback.color.as_ref(), Some(theme)))
            .or_else(|| Some("#000000".to_owned())),
        underline: Some(
            direct
                .underline
                .clone()
                .or(fallback.underline)
                .unwrap_or_else(|| "none".to_owned()),
        ),
        strike: direct.strike.clone().or(fallback.strike),
        spacing_pt: direct.spacing_pt.or(fallback.spacing_pt),
        baseline_pct: direct.baseline_pct.or(fallback.baseline_pct),
        caps: direct.caps.or(fallback.caps),
    }
}

fn find_source_node(nodes: &[ShapeNode], id: u32) -> Option<&ShapeNode> {
    for node in nodes {
        if node.id() == id {
            return Some(node);
        }
        if let ShapeNode::Group(group) = node
            && let Some(found) = find_source_node(&group.children, id)
        {
            return Some(found);
        }
    }
    None
}

fn node_text(node: &ShapeNode) -> Option<&TextBody> {
    match node {
        ShapeNode::Shape(shape) => shape.text.as_ref(),
        _ => None,
    }
}

fn placeholder_text<'a>(
    nodes: &'a [ShapeNode],
    placeholder: Option<&Placeholder>,
) -> Option<&'a TextBody> {
    placeholder
        .and_then(|target| find_placeholder(nodes, target))
        .and_then(node_text)
}

pub(crate) fn cell_formatting<'a>(
    package: &'a PptxPackage,
    master: Option<&SlideMaster>,
    node: Option<&'a ShapeNode>,
    shape: &crate::ShapeSnapshot,
    story_id: &str,
) -> Option<(&'a TextBody, TextBody)> {
    let ShapeNode::GraphicFrame(frame) = node? else {
        return None;
    };
    let GraphicFrameData::Table(table) = &frame.data else {
        return None;
    };
    let [row, column] = crate::table::source_cell(shape, story_id)?;
    let cell = table.rows.get(row)?.cells.get(column)?;
    let style = table
        .properties
        .style_id
        .as_deref()
        .and_then(|id| {
            package
                .table_styles
                .style(Some(id))
                .or_else(|| builtin_table_style(id))
        })
        .map(|style| {
            style.resolve_cell(
                TableStyleFlags {
                    first_row: table.properties.first_row,
                    last_row: table.properties.last_row,
                    first_column: table.properties.first_col,
                    last_column: table.properties.last_col,
                    band_row: table.properties.band_row,
                },
                TableCellPosition {
                    row,
                    column,
                    row_count: table.rows.len(),
                    column_count: table.grid.len().max(
                        table
                            .rows
                            .iter()
                            .map(|row| row.cells.len())
                            .max()
                            .unwrap_or_default(),
                    ),
                },
            )
        })
        .unwrap_or_default();
    let master_levels = master
        .map(|master| package.master_other_style_levels(master))
        .unwrap_or([false; 9]);
    let text_levels = table_text_levels(master, &master_levels, &package.presentation);
    Some((&cell.text, table_style_text_body(&style, &text_levels)))
}

/// The paragraph levels table text inherits outside the table style: each level the master's
/// `otherStyle` authors, else `p:defaultTextStyle` over its `a:defPPr`.
pub fn table_text_levels(
    master: Option<&SlideMaster>,
    master_levels: &[bool; 9],
    presentation: &Presentation,
) -> Vec<ParagraphProperties> {
    master_levels
        .iter()
        .enumerate()
        .map(|(index, present)| {
            let mut level = ParagraphProperties::default();
            if *present {
                if let Some(source) = master.and_then(|master| master.text_styles.other.get(index))
                {
                    merge_paragraph_properties(&mut level, source);
                }
            } else {
                if let Some(source) = presentation.default_text_paragraph.as_deref() {
                    merge_paragraph_properties(&mut level, source);
                }
                if let Some(source) = presentation
                    .default_text_style
                    .get(index)
                    .or_else(|| presentation.default_text_style.first())
                {
                    merge_paragraph_properties(&mut level, source);
                }
            }
            level
        })
        .collect()
}

/// Table text defaults from the master and table style, shaped as a body the cascade inherits
/// from.
pub fn table_style_text_body(
    style: &ResolvedCellStyle,
    text_levels: &[ParagraphProperties],
) -> TextBody {
    let properties = ParagraphProperties {
        default_run: Some(RunProperties {
            bold: style.bold,
            italic: style.italic,
            color: style.color.clone(),
            ..RunProperties::default()
        }),
        ..ParagraphProperties::default()
    };
    let mut list_style = text_levels.to_vec();
    for level in &mut list_style {
        merge_paragraph_properties(level, &properties);
    }
    TextBody {
        default_list_style: Some(Box::new(properties)),
        list_style,
        ..TextBody::default()
    }
}

/// The largest number a list marker shows: the largest `startAt` plus one per paragraph a story
/// renders.
pub const MAX_LIST_NUMBER: u32 = 32_767 + 20_000;

/// The source slide, layout and master a slide snapshot inherits from.
#[derive(Clone, Copy, Debug, Default)]
pub struct SlideParents<'a> {
    pub slide: Option<&'a Slide>,
    pub layout: Option<&'a SlideLayout>,
    pub master: Option<&'a SlideMaster>,
}

impl<'a> SlideParents<'a> {
    /// Resolves a slide's parts from its snapshot's source and layout part paths, falling back to
    /// the package's first layout and master.
    pub fn resolve(
        package: &'a PptxPackage,
        source_part_path: Option<&str>,
        layout_part_path: Option<&str>,
    ) -> Self {
        let slide = source_part_path
            .and_then(|path| package.slides.iter().find(|slide| slide.part_path == path));
        let layout_path =
            layout_part_path.or_else(|| slide.and_then(|slide| slide.layout_part_path.as_deref()));
        let layout = layout_path
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
        Self {
            slide,
            layout,
            master,
        }
    }
}

const MAX_LINK_WORK: usize = 65_536;

/// Conservatively detects text that can depend on deck order.
pub fn has_internal_slide_links(package: &PptxPackage, slide: &SlideSnapshot) -> bool {
    let parents = SlideParents::resolve(
        package,
        slide.source_part_path.as_deref(),
        slide.layout_part_path.as_deref(),
    );
    let links = SlideLinkDependencies::new(package);
    if let Some(source) = parents.slide
        && slide
            .shapes
            .iter()
            .any(|shape| links.snapshot_shape(shape, source))
    {
        return true;
    }
    let show_layout = parents.slide.is_none_or(|slide| slide.show_master_shapes);
    let show_master = show_layout
        && parents
            .layout
            .is_none_or(|layout| layout.show_master_shapes);
    parents
        .layout
        .filter(|_| show_layout)
        .is_some_and(|layout| {
            layout
                .shapes
                .iter()
                .any(|node| links.inherited_node(node, &layout.part_path))
        })
        || parents
            .master
            .filter(|_| show_master)
            .is_some_and(|master| {
                master
                    .shapes
                    .iter()
                    .any(|node| links.inherited_node(node, &master.part_path))
            })
}

struct SlideLinkDependencies<'a> {
    package: &'a PptxPackage,
    work: Cell<usize>,
    exhausted: Cell<bool>,
    parts: RefCell<HashMap<usize, bool>>,
    drawings: RefCell<HashSet<&'a str>>,
}

impl<'a> SlideLinkDependencies<'a> {
    fn new(package: &'a PptxPackage) -> Self {
        Self {
            package,
            work: Cell::new(0),
            exhausted: Cell::new(false),
            parts: RefCell::new(HashMap::new()),
            drawings: RefCell::new(HashSet::new()),
        }
    }

    fn charge_work(&self, work: usize) -> bool {
        if self.exhausted.get() {
            return false;
        }
        if work > MAX_LINK_WORK - self.work.get() {
            self.exhausted.set(true);
            return false;
        }
        self.work.set(self.work.get() + work);
        true
    }

    fn part_has_slide_relationship(&self, part: &str) -> Option<bool> {
        if self.exhausted.get() {
            return None;
        }
        let Some(relationships) = self.package.relationships.get(part) else {
            return Some(false);
        };
        let mut parts = self.parts.borrow_mut();
        let key = relationships.as_ptr() as usize;
        if let Some(has_slide) = parts.get(&key) {
            return Some(*has_slide);
        }
        if !self.charge_work(relationships.len()) {
            return None;
        }
        let has_slide = relationships.iter().any(|link| {
            link.target_mode == pptx_parse::TargetMode::Internal && link.has_type("/slide")
        });
        parts.insert(key, has_slide);
        Some(has_slide)
    }

    fn run(&self, run: &TextRun, has_slide_relationship: bool) -> bool {
        !self.charge_work(1)
            || run.properties.hyperlink_action.as_deref() == Some("ppaction://hlinksldjump")
            || (has_slide_relationship && run.properties.hyperlink_relationship_id.is_some())
    }

    fn body(&self, body: &TextBody, part: &str) -> bool {
        let Some(has_slide_relationship) = self.part_has_slide_relationship(part) else {
            return true;
        };
        body.paragraphs.iter().any(|paragraph| {
            paragraph
                .runs
                .iter()
                .any(|run| self.run(run, has_slide_relationship))
        })
    }

    fn graphic(&self, graphic: &GraphicFrameData, part: &str, depth: usize) -> bool {
        match graphic {
            GraphicFrameData::Table(table) => table
                .rows
                .iter()
                .flat_map(|row| &row.cells)
                .any(|cell| !cell.merged && self.body(&cell.text, part)),
            GraphicFrameData::Diagram {
                drawing_part_path: Some(path),
                ..
            } => self.drawing(path, depth),
            _ => false,
        }
    }

    fn drawing(&self, path: &str, depth: usize) -> bool {
        if self.exhausted.get() {
            return true;
        }
        if self.drawings.borrow().contains(path) {
            return false;
        }
        if !self.charge_work(1) {
            return true;
        }
        for drawing in &self.package.diagram_drawings {
            if !self.charge_work(1) {
                return true;
            }
            if drawing.part_path == path {
                self.drawings.borrow_mut().insert(&drawing.part_path);
                return drawing
                    .shapes
                    .iter()
                    .any(|node| self.parsed_node(node, &drawing.part_path, depth + 1));
            }
        }
        false
    }

    fn inherited_node(&self, node: &ShapeNode, part: &str) -> bool {
        !self.charge_work(1)
            || (shape_base(node).placeholder.is_none() && self.parsed_node_contents(node, part, 0))
    }

    fn parsed_node(&self, node: &ShapeNode, part: &str, depth: usize) -> bool {
        !self.charge_work(1) || self.parsed_node_contents(node, part, depth)
    }

    fn parsed_node_contents(&self, node: &ShapeNode, part: &str, depth: usize) -> bool {
        if shape_base(node).hidden {
            return false;
        }
        if depth >= 128 {
            return true;
        }
        match node {
            ShapeNode::Shape(shape) => shape
                .text
                .as_ref()
                .is_some_and(|body| self.body(body, part)),
            ShapeNode::Group(group) => group
                .children
                .iter()
                .any(|node| self.parsed_node(node, part, depth + 1)),
            ShapeNode::GraphicFrame(frame) => self.graphic(&frame.data, part, depth),
            ShapeNode::Picture(_) => false,
        }
    }

    fn snapshot_shape(&self, shape: &ShapeSnapshot, source: &Slide) -> bool {
        if !self.charge_work(1) {
            return true;
        }
        if shape.hidden {
            return false;
        }
        if shape.kind == crate::ShapeKind::Group {
            return shape
                .children
                .iter()
                .any(|shape| self.snapshot_shape(shape, source));
        }
        if shape.kind == crate::ShapeKind::GraphicFrame {
            return shape
                .graphic
                .as_ref()
                .is_some_and(|graphic| self.graphic(graphic, &source.part_path, 0));
        }
        if shape.text_stories.is_empty() || shape.source_id == 0 {
            return false;
        }
        let node = self.source_node(&source.shapes, shape.source_id);
        self.exhausted.get()
            || node
                .and_then(node_text)
                .is_some_and(|body| self.body(body, &source.part_path))
    }

    fn source_node<'b>(&self, nodes: &'b [ShapeNode], id: u32) -> Option<&'b ShapeNode> {
        for node in nodes {
            if !self.charge_work(1) {
                return None;
            }
            if shape_base(node).id == id {
                return Some(node);
            }
            if let ShapeNode::Group(group) = node {
                let found = self.source_node(&group.children, id);
                if found.is_some() || self.exhausted.get() {
                    return found;
                }
            }
        }
        None
    }
}

/// The first shape in `nodes`, groups included, whose placeholder matches `target`.
pub fn find_placeholder<'a>(nodes: &'a [ShapeNode], target: &Placeholder) -> Option<&'a ShapeNode> {
    for node in nodes {
        if shape_base(node)
            .placeholder
            .as_ref()
            .is_some_and(|value| value.matches(target))
        {
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

/// A placeholder type with its synonyms folded: `ctrTitle` is `title`, `obj` and an absent type
/// are `body`.
pub fn normalize_placeholder_type(value: Option<&str>) -> &str {
    match value.unwrap_or("body") {
        "ctrTitle" => "title",
        "obj" => "body",
        value => value,
    }
}

pub(crate) fn is_title_placeholder(placeholder: Option<&Placeholder>) -> bool {
    placeholder.is_some_and(|value| {
        normalize_placeholder_type(value.placeholder_type.as_deref()) == "title"
    })
}

pub(crate) fn slide_hidden(package: &PptxPackage, source_path: Option<&str>) -> Option<bool> {
    match source_path {
        None => Some(false),
        Some(path) => package
            .slides
            .iter()
            .find(|slide| slide.part_path == path)
            .and_then(|slide| slide.hidden),
    }
}

pub(crate) fn shape_cascade<'a>(
    parents: &SlideParents<'a>,
    presentation: &'a pptx_parse::Presentation,
    placeholder: Option<&'a Placeholder>,
    primary: Option<&'a TextBody>,
) -> ParagraphCascade<'a> {
    let inherited = |nodes: &'a [ShapeNode]| {
        placeholder
            .and_then(|value| find_placeholder(nodes, value))
            .and_then(node_text)
    };
    ParagraphCascade {
        primary,
        layout: parents.layout.and_then(|layout| inherited(&layout.shapes)),
        master: parents.master.and_then(|master| inherited(&master.shapes)),
        master_slide: parents.master,
        default_style: &presentation.default_text_style,
        default_paragraph: presentation.default_text_paragraph.as_deref(),
        placeholder,
        style_color: None,
    }
}

pub(crate) fn effective_list(
    counters: &mut ListCounters,
    properties: &ParagraphProperties,
    level: u32,
    has_text: bool,
) -> Option<crate::structured::ExportList> {
    use crate::structured::ExportList;
    match counters.next(properties.bullet.as_ref(), level, has_text)? {
        ListMarker::Character(character) => Some(ExportList::Bullet {
            character,
            font: match &properties.bullet_font {
                Some(pptx_parse::BulletFont::Typeface(font)) => Some(font.clone()),
                _ => None,
            },
        }),
        ListMarker::Number { value, marker } => {
            let (scheme, start_at) = match &properties.bullet {
                Some(Bullet::AutoNumber {
                    scheme, start_at, ..
                }) => (scheme.clone(), *start_at),
                _ => (String::new(), 1),
            };
            Some(ExportList::Number {
                scheme,
                start_at,
                value,
                marker,
            })
        }
    }
}

/// The source paragraph a seeded paragraph id names.
pub fn source_paragraph<'b>(
    body: &'b TextBody,
    story_id: &str,
    paragraph_id: &str,
) -> Option<&'b TextParagraph> {
    let index = paragraph_id
        .strip_prefix("para:")?
        .strip_prefix(story_id)?
        .strip_prefix(':')?
        .parse::<usize>()
        .ok()?;
    body.paragraphs.get(index)
}

/// The leading and trailing bytes `source` and `target` share, compared by character, the
/// trailing ones within what the leading ones leave.
pub fn common_ends(source: &str, target: &str) -> (usize, usize) {
    let prefix: usize = source
        .chars()
        .zip(target.chars())
        .take_while(|(left, right)| left == right)
        .map(|(value, _)| value.len_utf8())
        .sum();
    let suffix = source[prefix..]
        .chars()
        .rev()
        .zip(target[prefix..].chars().rev())
        .take_while(|(left, right)| left == right)
        .map(|(value, _)| value.len_utf8())
        .sum();
    (prefix, suffix)
}

/// A source paragraph's field results as byte spans of `text`, its current text, empty results
/// included. `None` when a field left the unchanged leading or trailing text.
pub fn field_spans(
    source: &TextParagraph,
    text: &str,
) -> Option<Vec<(std::ops::Range<usize>, Option<String>)>> {
    field_spans_with_runs(source, text).map(|spans| {
        spans
            .into_iter()
            .map(|(span, field_type, _)| (span, field_type))
            .collect()
    })
}

type FieldSpan<'a> = (std::ops::Range<usize>, Option<String>, &'a TextRun);

fn field_spans_with_runs<'a>(source: &'a TextParagraph, text: &str) -> Option<Vec<FieldSpan<'a>>> {
    let mut seeded = String::new();
    let mut fields = Vec::new();
    for run in &source.runs {
        let start = seeded.len();
        seeded.push_str(&run.text);
        if run.field_id.is_some() || run.field_type.is_some() {
            fields.push((start..seeded.len(), run.field_type.clone(), run));
        }
    }
    let (prefix, suffix) = common_ends(&seeded, text);
    let (tail, moved_tail) = (seeded.len() - suffix, text.len() - suffix);
    fields
        .into_iter()
        .map(|(span, field_type, run)| {
            if span.end <= prefix {
                Some((span, field_type, run))
            } else if span.start >= tail {
                let start = moved_tail + (span.start - tail);
                Some((start..start + span.len(), field_type, run))
            } else {
                None
            }
        })
        .collect()
}

/// `live` with each located `slidenum` field result replaced by `slide_number`.
pub fn resolved_field_text(source: &TextParagraph, live: &str, slide_number: i64) -> String {
    let Some(spans) = field_spans(source, live) else {
        return live.to_owned();
    };
    let mut resolved = String::new();
    let mut last = 0;
    for (span, field_type) in &spans {
        if field_type.as_deref() == Some("slidenum") {
            resolved.push_str(&live[last..span.start]);
            resolved.push_str(&slide_number.to_string());
            last = span.end;
        }
    }
    resolved.push_str(&live[last..]);
    resolved
}

/// Per live run, the replacement for a run a located `slidenum` field overlaps, or `None` to
/// keep the live run.
pub fn slide_number_run_texts(
    source: &TextParagraph,
    live_runs: &[&str],
    slide_number: i64,
) -> Vec<Option<String>> {
    let mut replacements = vec![None; live_runs.len()];
    let live: String = live_runs.iter().copied().collect();
    let Some(spans) = field_spans(source, &live) else {
        return replacements;
    };
    let mut boundaries = Vec::with_capacity(live_runs.len() + 1);
    boundaries.push(0);
    for run in live_runs {
        boundaries.push(boundaries.last().unwrap() + run.len());
    }
    let mut edits: Vec<Vec<(usize, usize, String)>> = vec![Vec::new(); live_runs.len()];
    for (span, field_type) in &spans {
        if field_type.as_deref() != Some("slidenum") {
            continue;
        }
        if span.is_empty() {
            if let Some(index) = live_runs
                .iter()
                .enumerate()
                .find(|(index, _)| {
                    boundaries[*index] < span.start && span.start < boundaries[index + 1]
                })
                .map(|(index, _)| index)
            {
                let at = span.start - boundaries[index];
                edits[index].push((at, at, slide_number.to_string()));
            }
            continue;
        }
        let mut first = true;
        for (index, bounds) in boundaries.windows(2).enumerate() {
            let start = bounds[0].max(span.start);
            let end = bounds[1].min(span.end);
            if start >= end {
                continue;
            }
            let replacement = if first {
                slide_number.to_string()
            } else {
                String::new()
            };
            first = false;
            edits[index].push((start - bounds[0], end - bounds[0], replacement));
        }
    }
    for (index, mut run_edits) in edits.into_iter().enumerate() {
        if run_edits.is_empty() {
            continue;
        }
        run_edits.sort_by_key(|(start, end, _)| (*start, *end));
        let original = live_runs[index];
        let mut text = String::with_capacity(original.len());
        let mut cursor = 0;
        for (start, end, replacement) in run_edits {
            if start < cursor {
                continue;
            }
            text.push_str(&original[cursor..start]);
            text.push_str(&replacement);
            cursor = end;
        }
        text.push_str(&original[cursor..]);
        if text != original {
            replacements[index] = Some(text);
        }
    }
    replacements
}

/// The live byte offsets, at run boundaries, of located empty `slidenum` fields that lie at a
/// run boundary or in a paragraph with no runs, with each field's source run.
pub fn slide_number_insertions<'a>(
    source: &'a TextParagraph,
    live_runs: &[&str],
) -> Vec<(usize, &'a TextRun)> {
    let live: String = live_runs.iter().copied().collect();
    let Some(spans) = field_spans_with_runs(source, &live) else {
        return Vec::new();
    };
    let mut boundaries = Vec::with_capacity(live_runs.len() + 1);
    boundaries.push(0);
    for run in live_runs {
        boundaries.push(boundaries.last().unwrap() + run.len());
    }
    let mut insertions = Vec::new();
    for (span, field_type, run) in &spans {
        if field_type.as_deref() != Some("slidenum") || !span.is_empty() {
            continue;
        }
        if live_runs
            .iter()
            .enumerate()
            .any(|(index, _)| boundaries[index] < span.start && span.start < boundaries[index + 1])
        {
            continue;
        }
        if boundaries.contains(&span.start) {
            insertions.push((span.start, *run));
        }
    }
    insertions
}

/// The text bodies a shape's paragraphs inherit from, most specific first.
#[derive(Clone, Copy, Debug, Default)]
pub struct ParagraphCascade<'a> {
    /// The shape's own source text body.
    pub primary: Option<&'a TextBody>,
    /// The matching layout placeholder's text body.
    pub layout: Option<&'a TextBody>,
    /// The matching master placeholder's text body.
    pub master: Option<&'a TextBody>,
    /// The master whose text styles the placeholder type selects.
    pub master_slide: Option<&'a SlideMaster>,
    /// `p:defaultTextStyle`, which outranks the master's `otherStyle` for text outside
    /// placeholders.
    pub default_style: &'a [ParagraphProperties],
    /// `p:defaultTextStyle/a:defPPr`, under every level of `default_style`.
    pub default_paragraph: Option<&'a ParagraphProperties>,
    pub placeholder: Option<&'a Placeholder>,
    /// `p:style/a:fontRef` colour.
    pub style_color: Option<&'a ColorValue>,
}

impl ParagraphCascade<'_> {
    /// The effective properties of paragraph `index` at `level`. An `authored` numbered bullet
    /// replaces the inherited bullet.
    pub fn properties(
        &self,
        index: usize,
        level: u32,
        authored: Option<&Bullet>,
    ) -> ParagraphProperties {
        let mut properties = self
            .master_slide
            .and_then(|master| master_style(master, self.placeholder, level))
            .cloned()
            .unwrap_or_default();
        if self.placeholder.is_none() {
            let level_style = self
                .default_style
                .get(level as usize)
                .or_else(|| self.default_style.first());
            for source in self.default_paragraph.into_iter().chain(level_style) {
                merge_paragraph_properties(&mut properties, source);
            }
        }
        if let Some(color) = self.style_color {
            properties
                .default_run
                .get_or_insert_with(RunProperties::default)
                .color = Some(color.clone());
        }
        for body in [self.master, self.layout, self.primary]
            .into_iter()
            .flatten()
        {
            if let Some(source) = &body.default_list_style {
                merge_paragraph_properties(&mut properties, source);
            }
            if let Some(source) = body.list_style.get(level as usize) {
                merge_paragraph_properties(&mut properties, source);
            }
            if let Some(source) = body
                .paragraphs
                .get(index)
                .or_else(|| body.paragraphs.get(level as usize))
                .map(|paragraph| &paragraph.properties)
            {
                merge_paragraph_properties(&mut properties, source);
            }
        }
        if let Some(Bullet::AutoNumber { restart, .. }) = &mut properties.bullet {
            *restart = self
                .primary
                .and_then(|body| body.paragraphs.get(index))
                .is_some_and(|paragraph| {
                    matches!(
                        paragraph.properties.bullet,
                        Some(Bullet::AutoNumber { restart: true, .. })
                            | Some(Bullet::AutoNumber { start_at: 2.., .. })
                    )
                });
        }
        if let Some(authored @ Bullet::AutoNumber { .. }) = authored {
            properties.bullet = Some(authored.clone());
            if let Some(Bullet::AutoNumber {
                restart, start_at, ..
            }) = &mut properties.bullet
            {
                *restart |= *start_at != 1;
            }
        }
        properties
    }
}

fn master_style<'a>(
    master: &'a SlideMaster,
    placeholder: Option<&Placeholder>,
    level: u32,
) -> Option<&'a ParagraphProperties> {
    let styles = match placeholder {
        Some(placeholder) => {
            match normalize_placeholder_type(placeholder.placeholder_type.as_deref()) {
                "title" => &master.text_styles.title,
                "body" | "subTitle" => &master.text_styles.body,
                _ => &master.text_styles.other,
            }
        }
        None => &master.text_styles.other,
    };
    styles.get(level as usize).or_else(|| styles.first())
}

/// Overlays every property `source` declares onto `target`.
pub fn merge_paragraph_properties(target: &mut ParagraphProperties, source: &ParagraphProperties) {
    if source.alignment.is_some() {
        target.alignment.clone_from(&source.alignment);
    }
    if source.margin_left.is_some() {
        target.margin_left = source.margin_left;
    }
    if source.margin_right.is_some() {
        target.margin_right = source.margin_right;
    }
    if source.indent.is_some() {
        target.indent = source.indent;
    }
    if source.bullet.is_some() {
        target.bullet.clone_from(&source.bullet);
    }
    if source.line_spacing.is_some() {
        target.line_spacing = source.line_spacing;
    }
    if source.space_before.is_some() {
        target.space_before = source.space_before;
    }
    if source.space_after.is_some() {
        target.space_after = source.space_after;
    }
    if source.bullet_font.is_some() {
        target.bullet_font.clone_from(&source.bullet_font);
    }
    if source.bullet_color.is_some() {
        target.bullet_color.clone_from(&source.bullet_color);
    }
    if source.bullet_size.is_some() {
        target.bullet_size.clone_from(&source.bullet_size);
    }
    if source.default_tab_size.is_some() {
        target.default_tab_size = source.default_tab_size;
    }
    if source.tab_stops.is_some() {
        target.tab_stops.clone_from(&source.tab_stops);
    }
    if source.rtl.is_some() {
        target.rtl = source.rtl;
    }
    if let Some(source) = &source.default_run {
        let target = target
            .default_run
            .get_or_insert_with(RunProperties::default);
        merge_run_properties(target, source);
    }
}

fn merge_run_properties(target: &mut RunProperties, source: &RunProperties) {
    if source.font_size_pt.is_some() {
        target.font_size_pt = source.font_size_pt;
    }
    if source.bold.is_some() {
        target.bold = source.bold;
    }
    if source.italic.is_some() {
        target.italic = source.italic;
    }
    if source.underline.is_some() {
        target.underline.clone_from(&source.underline);
    }
    if source.strike.is_some() {
        target.strike.clone_from(&source.strike);
    }
    if source.font_family.is_some() {
        target.font_family.clone_from(&source.font_family);
    }
    if source.color.is_some() {
        target.color.clone_from(&source.color);
    }
    if source.language.is_some() {
        target.language.clone_from(&source.language);
    }
    if source.spacing_pt.is_some() {
        target.spacing_pt = source.spacing_pt;
    }
    if source.baseline_pct.is_some() {
        target.baseline_pct = source.baseline_pct;
    }
    if source.kern_pt.is_some() {
        target.kern_pt = source.kern_pt;
    }
    if source.caps.is_some() {
        target.caps = source.caps;
    }
}

/// A paragraph's list marker, before any symbol-font substitution.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ListMarker {
    /// A `buChar` bullet, as authored.
    Character(String),
    /// A `buAutoNum` number; `marker` is `None` for a scheme [`format_autonum`] cannot format.
    Number { value: u32, marker: Option<String> },
}

impl ListMarker {
    /// The marker to draw: an unformattable scheme falls back to `"{value}."`.
    pub fn text(&self) -> String {
        match self {
            Self::Character(value) => value.clone(),
            Self::Number { value, marker } => marker.clone().unwrap_or_else(|| format!("{value}.")),
        }
    }
}

/// Per-level `a:buAutoNum` state of one story: the number last shown and the `startAt` the run
/// was seeded from.
#[derive(Clone, Debug, Default)]
pub struct ListCounters {
    numbers: [u32; 9],
    starts: [u32; 9],
}

impl ListCounters {
    /// The marker of the story's next paragraph. A paragraph without text is spacing, not an
    /// item: PowerPoint neither marks it nor counts it towards the next number.
    pub fn next(
        &mut self,
        bullet: Option<&Bullet>,
        level: u32,
        has_text: bool,
    ) -> Option<ListMarker> {
        if !has_text {
            return None;
        }
        let level = (level as usize).min(self.numbers.len() - 1);
        self.numbers[level + 1..].fill(0);
        self.starts[level + 1..].fill(0);
        match bullet {
            Some(Bullet::AutoNumber {
                scheme,
                start_at,
                restart,
            }) => {
                let start = (*start_at).clamp(1, 32_767);
                // PowerPoint writes the list's `startAt` on every one of its paragraphs, so
                // repeating the seed continues the run; only a different declared start opens a
                // new list.
                self.numbers[level] = match self.numbers[level] {
                    0 => start,
                    _ if *restart && self.starts[level] != start => start,
                    current => current.saturating_add(1),
                };
                if self.numbers[level] == start {
                    self.starts[level] = start;
                }
                let value = self.numbers[level].clamp(1, MAX_LIST_NUMBER);
                Some(ListMarker::Number {
                    value,
                    marker: format_autonum(value, scheme),
                })
            }
            _ => {
                self.numbers[level] = 0;
                self.starts[level] = 0;
                match bullet {
                    Some(Bullet::Character { value }) if !value.trim().is_empty() => {
                        Some(ListMarker::Character(value.clone()))
                    }
                    _ => None,
                }
            }
        }
    }
}

/// Formats a Latin, Roman or decimal `ST_TextAutonumberScheme` marker; `None` for any other
/// scheme.
pub fn format_autonum(value: u32, scheme: &str) -> Option<String> {
    let value = value.clamp(1, MAX_LIST_NUMBER);
    let (numeral, suffix) = ["ParenBoth", "ParenR", "Period", "Plain"]
        .into_iter()
        .find_map(|suffix| Some((scheme.strip_suffix(suffix)?, suffix)))?;
    let body = match numeral {
        "alphaLc" => format_alpha(value, false),
        "alphaUc" => format_alpha(value, true),
        "romanLc" => format_roman(value, false),
        "romanUc" => format_roman(value, true),
        "arabic" => value.to_string(),
        _ => return None,
    };
    Some(match suffix {
        "ParenBoth" => format!("({body})"),
        "ParenR" => format!("{body})"),
        "Plain" => body,
        _ => format!("{body}."),
    })
}

fn format_alpha(value: u32, upper: bool) -> String {
    let base = if upper { b'A' } else { b'a' };
    let mut value = value.max(1);
    let mut out = Vec::new();
    while value > 0 {
        let index = (value - 1) % 26;
        out.push(base + index as u8);
        value = (value - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).unwrap_or_default()
}

fn format_roman(value: u32, upper: bool) -> String {
    const NUMERALS: [(u32, &str); 13] = [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ];
    let mut value = value.max(1);
    let mut out = String::new();
    for (amount, numeral) in NUMERALS {
        while value >= amount {
            out.push_str(numeral);
            value -= amount;
        }
    }
    if upper { out.to_uppercase() } else { out }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn dependency_base(id: u32) -> pptx_parse::ShapeBase {
        pptx_parse::ShapeBase {
            id,
            name: format!("Shape {id}"),
            title: None,
            description: None,
            decorative: false,
            hidden: false,
            placeholder: None,
            transform: pptx_parse::ShapeTransform::default(),
        }
    }

    fn dependency_frame(id: u32, path: &str) -> ShapeNode {
        ShapeNode::GraphicFrame(pptx_parse::GraphicFrame {
            base: dependency_base(id),
            data: GraphicFrameData::Diagram {
                relationship_ids: Vec::new(),
                drawing_part_path: Some(path.to_owned()),
            },
        })
    }

    fn dependency_leaf(linked: bool) -> ShapeNode {
        ShapeNode::Shape(pptx_parse::Shape {
            base: dependency_base(3),
            geometry: "rect".to_owned(),
            has_preset_geometry: true,
            paths: Vec::new(),
            style: None,
            adjust_values: Default::default(),
            use_bg_fill: false,
            fill: None,
            picture_fill: None,
            outline: None,
            effects: None,
            text: linked.then(|| TextBody {
                paragraphs: vec![TextParagraph {
                    runs: vec![TextRun {
                        text: "Jump".to_owned(),
                        properties: RunProperties {
                            hyperlink_action: Some("ppaction://hlinksldjump".to_owned()),
                            ..RunProperties::default()
                        },
                        field_id: None,
                        field_type: None,
                        line_break: false,
                    }],
                    ..TextParagraph::default()
                }],
                ..TextBody::default()
            }),
        })
    }

    #[test]
    fn doubling_diagram_dependencies_visit_each_part_once() {
        let mut package = PptxPackage::default();
        let paths: Vec<_> = (0..32)
            .map(|index| format!("ppt/diagrams/drawing{index}.xml"))
            .collect();
        for index in 0..paths.len() {
            package.diagram_drawings.push(pptx_parse::DiagramDrawing {
                part_path: paths[index].clone(),
                synthesis_extent: None,
                shapes: if index == 0 {
                    vec![dependency_leaf(false)]
                } else {
                    vec![
                        dependency_frame(1, &paths[index - 1]),
                        dependency_frame(2, &paths[index - 1]),
                    ]
                },
            });
        }
        let root = dependency_frame(100, paths.last().unwrap());
        {
            let links = SlideLinkDependencies::new(&package);
            assert!(!links.parsed_node(&root, "ppt/slides/slide1.xml", 0));
            assert_eq!(links.drawings.borrow().len(), 32);
            assert!(links.work.get() < 1_024);
            assert!(!links.exhausted.get());
        }
        package.diagram_drawings[0].shapes[0] = dependency_leaf(true);
        let links = SlideLinkDependencies::new(&package);
        assert!(links.parsed_node(&root, "ppt/slides/slide1.xml", 0));
        assert!(!links.exhausted.get());
    }

    #[test]
    fn cyclic_diagram_dependencies_keep_links_beyond_back_edges() {
        for count in [1, 2] {
            let mut package = PptxPackage::default();
            for index in 0..count {
                package.diagram_drawings.push(pptx_parse::DiagramDrawing {
                    part_path: format!("ppt/diagrams/drawing{index}.xml"),
                    synthesis_extent: None,
                    shapes: vec![
                        dependency_frame(
                            1,
                            &format!("ppt/diagrams/drawing{}.xml", (index + 1) % count),
                        ),
                        dependency_leaf(false),
                    ],
                });
            }
            let root = dependency_frame(100, "ppt/diagrams/drawing0.xml");
            {
                let links = SlideLinkDependencies::new(&package);
                assert!(!links.parsed_node(&root, "ppt/slides/slide1.xml", 0));
                assert_eq!(links.drawings.borrow().len(), count);
                assert!(!links.exhausted.get());
            }
            package.diagram_drawings[count - 1].shapes[1] = dependency_leaf(true);
            let links = SlideLinkDependencies::new(&package);
            assert!(links.parsed_node(&root, "ppt/slides/slide1.xml", 0));
            assert!(!links.exhausted.get());
        }
    }

    #[test]
    fn structural_dependency_work_cap_counts_textless_diagrams_and_groups() {
        for diagram in [false, true] {
            let mut package = PptxPackage::default();
            let children = vec![dependency_leaf(false); MAX_LINK_WORK + 1];
            let root = if diagram {
                package.diagram_drawings.push(pptx_parse::DiagramDrawing {
                    part_path: "ppt/diagrams/drawing1.xml".to_owned(),
                    synthesis_extent: None,
                    shapes: children,
                });
                dependency_frame(100, "ppt/diagrams/drawing1.xml")
            } else {
                ShapeNode::Group(pptx_parse::GroupShape {
                    base: dependency_base(100),
                    children,
                })
            };
            let links = SlideLinkDependencies::new(&package);
            assert!(links.parsed_node(&root, "ppt/slides/slide1.xml", 0));
            assert_eq!(links.work.get(), MAX_LINK_WORK);
            assert!(links.exhausted.get());
        }
    }

    #[test]
    fn slide_link_dependencies_scan_runs_and_relationships_conservatively() {
        const COUNT: usize = 20_000;
        let part = "ppt/slides/slide1.xml";
        let mut package = PptxPackage::default();
        let run = |relationship_id| TextRun {
            text: "x".to_owned(),
            properties: RunProperties {
                hyperlink_relationship_id: relationship_id,
                ..RunProperties::default()
            },
            field_id: None,
            field_type: None,
            line_break: false,
        };
        let mut body = TextBody {
            paragraphs: vec![TextParagraph {
                runs: (0..COUNT)
                    .map(|index| run(Some(format!("missing{index}"))))
                    .collect(),
                ..TextParagraph::default()
            }],
            ..TextBody::default()
        };
        package.relationships.insert(
            part.to_owned(),
            (0..COUNT)
                .map(|index| pptx_parse::Relationship {
                    id: format!("present{index}"),
                    relationship_type: "/slide".to_owned(),
                    target: "https://example.test".to_owned(),
                    target_mode: pptx_parse::TargetMode::External,
                    resolved_target: None,
                })
                .collect(),
        );
        {
            let links = SlideLinkDependencies::new(&package);
            assert!(!links.body(&body, part));
            assert_eq!(links.work.get(), 2 * COUNT);
            assert!(!links.exhausted.get());
        }
        let relationship = &mut package.relationships.get_mut(part).unwrap()[0];
        relationship.target_mode = pptx_parse::TargetMode::Internal;
        relationship.target = "slide1.xml".to_owned();
        relationship.resolved_target = Some(part.to_owned());
        {
            let links = SlideLinkDependencies::new(&package);
            assert!(links.body(&body, part));
            assert_eq!(links.work.get(), COUNT + 1);
            assert!(!links.exhausted.get());
        }
        body.paragraphs[0].runs[0]
            .properties
            .hyperlink_relationship_id = Some("present0".to_owned());
        {
            let links = SlideLinkDependencies::new(&package);
            assert!(links.body(&body, part));
        }
        package.relationships.clear();
        {
            let properties = &mut body.paragraphs[0].runs[0].properties;
            properties.hyperlink_relationship_id = None;
            properties.hyperlink_action = Some("ppaction://hlinksldjump".to_owned());
        }
        {
            let links = SlideLinkDependencies::new(&package);
            assert!(links.body(&body, part));
            assert_eq!(links.work.get(), 1);
        }
        body.paragraphs[0].runs[0].properties.hyperlink_action = None;
        {
            let links = SlideLinkDependencies::new(&package);
            assert!(!links.body(&body, part));
            assert_eq!(links.work.get(), COUNT);
            assert!(!links.exhausted.get());
        }
        package.relationships.insert(
            part.to_owned(),
            vec![
                pptx_parse::Relationship {
                    id: "unused".to_owned(),
                    relationship_type: "/hyperlink".to_owned(),
                    target: "https://example.test".to_owned(),
                    target_mode: pptx_parse::TargetMode::External,
                    resolved_target: None,
                };
                100_000
            ],
        );
        {
            let links = SlideLinkDependencies::new(&package);
            assert!(links.body(&body, part));
            assert_eq!(links.work.get(), 0);
            assert!(links.exhausted.get());
            assert!(links.parts.borrow().is_empty());
        }
    }

    #[test]
    fn slide_link_dependency_work_cap_returns_the_conservative_answer() {
        let mut package = PptxPackage::default();
        let mut body = TextBody {
            paragraphs: vec![TextParagraph {
                runs: vec![
                    TextRun {
                        text: "x".to_owned(),
                        properties: RunProperties::default(),
                        field_id: None,
                        field_type: None,
                        line_break: false,
                    };
                    MAX_LINK_WORK
                ],
                ..TextParagraph::default()
            }],
            ..TextBody::default()
        };
        {
            let links = SlideLinkDependencies::new(&package);
            assert!(!links.body(&body, "ppt/slides/slide1.xml"));
            assert_eq!(links.work.get(), MAX_LINK_WORK);
            assert!(links.body(&body, "ppt/slides/slide1.xml"));
            assert!(links.exhausted.get());
            assert_eq!(links.work.get(), MAX_LINK_WORK);
        }
        package.relationships.insert(
            "ppt/slides/slide1.xml".to_owned(),
            vec![
                pptx_parse::Relationship {
                    id: "unused".to_owned(),
                    relationship_type: "/hyperlink".to_owned(),
                    target: "https://example.test".to_owned(),
                    target_mode: pptx_parse::TargetMode::External,
                    resolved_target: None,
                };
                MAX_LINK_WORK + 1
            ],
        );
        body.paragraphs[0].runs.truncate(1);
        let links = SlideLinkDependencies::new(&package);
        assert!(links.body(&body, "ppt/slides/slide1.xml"));
        assert_eq!(links.work.get(), 0);
        assert!(links.parts.borrow().is_empty());
    }

    fn marker(counters: &mut ListCounters, bullet: Option<&Bullet>, level: u32) -> Option<String> {
        counters
            .next(bullet, level, true)
            .map(|marker| marker.text())
    }

    fn number(scheme: &str, start_at: u32, restart: bool) -> Bullet {
        Bullet::AutoNumber {
            scheme: scheme.to_owned(),
            start_at,
            restart,
        }
    }

    fn character(value: &str) -> Bullet {
        Bullet::Character {
            value: value.to_owned(),
        }
    }

    #[test]
    fn autonumbering_counts_per_level_and_resumes_across_other_levels() {
        let mut counters = ListCounters::default();
        let arabic = number("arabicPeriod", 1, false);
        let dash = character("-");
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("1.")
        );
        assert_eq!(marker(&mut counters, Some(&dash), 1).as_deref(), Some("-"));
        assert_eq!(marker(&mut counters, Some(&dash), 1).as_deref(), Some("-"));
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("2.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("3.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 1).as_deref(),
            Some("1.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 1).as_deref(),
            Some("2.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("4.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 1).as_deref(),
            Some("1.")
        );
        assert_eq!(marker(&mut counters, Some(&Bullet::None), 0), None);
    }

    #[test]
    fn inherited_autonumber_start_at_applies_only_to_the_first_item() {
        let mut counters = ListCounters::default();
        let seven = number("arabicPeriod", 7, false);
        assert_eq!(
            marker(&mut counters, Some(&seven), 0).as_deref(),
            Some("7.")
        );
        assert_eq!(
            marker(&mut counters, Some(&seven), 0).as_deref(),
            Some("8.")
        );
        let mut counters = ListCounters::default();
        let last_start = number("arabicPeriod", 32_767, false);
        assert_eq!(
            marker(&mut counters, Some(&last_start), 0).as_deref(),
            Some("32767.")
        );
        assert_eq!(
            marker(&mut counters, Some(&last_start), 0).as_deref(),
            Some("32768.")
        );
    }

    #[test]
    fn autonumber_schemes_format_their_numeral_and_suffix() {
        assert_eq!(format_autonum(4, "arabicPeriod").as_deref(), Some("4."));
        assert_eq!(format_autonum(4, "arabicParenR").as_deref(), Some("4)"));
        assert_eq!(format_autonum(4, "arabicParenBoth").as_deref(), Some("(4)"));
        assert_eq!(format_autonum(4, "arabicPlain").as_deref(), Some("4"));
        assert_eq!(format_autonum(1, "alphaLcParenR").as_deref(), Some("a)"));
        assert_eq!(format_autonum(27, "alphaUcPeriod").as_deref(), Some("AA."));
        assert_eq!(format_autonum(9, "romanLcPeriod").as_deref(), Some("ix."));
        assert_eq!(
            format_autonum(2024, "romanUcPeriod").as_deref(),
            Some("MMXXIV.")
        );
        for scheme in ["somethingElse", "somethingElsePlain", "arabicDbPeriod"] {
            assert_eq!(format_autonum(3, scheme), None);
        }
    }

    #[test]
    fn unformattable_schemes_fall_back_to_a_decimal_marker() {
        let mut counters = ListCounters::default();
        let circled = number("circleNumDbPlain", 3, false);
        assert_eq!(
            counters.next(Some(&circled), 0, true),
            Some(ListMarker::Number {
                value: 3,
                marker: None
            })
        );
        assert_eq!(
            marker(&mut counters, Some(&circled), 0).as_deref(),
            Some("4.")
        );
    }

    #[test]
    fn paragraphs_without_text_neither_mark_nor_count() {
        let mut counters = ListCounters::default();
        let arabic = number("arabicPeriod", 1, false);
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("1.")
        );
        assert_eq!(counters.next(Some(&arabic), 0, false), None);
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("2.")
        );
    }

    #[test]
    fn autonumber_sequences_restart_after_plain_paragraphs_and_explicit_starts() {
        let mut counters = ListCounters::default();
        let arabic = number("arabicPeriod", 1, false);
        let restart = number("arabicPeriod", 7, true);
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("1.")
        );
        assert_eq!(
            marker(&mut counters, Some(&restart), 0).as_deref(),
            Some("7.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("8.")
        );
        assert_eq!(marker(&mut counters, None, 0), None);
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("1.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 1).as_deref(),
            Some("1.")
        );
        assert_eq!(
            marker(&mut counters, Some(&character("•")), 0).as_deref(),
            Some("•")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 1).as_deref(),
            Some("1.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("1.")
        );
    }

    /// `pptarena-018-original` slide 11 declares `startAt="4"` on all four of its paragraphs and
    /// PowerPoint renders 4, 5, 6, 7; `pptarena-034-original` slide 11 declares `startAt="1"` on
    /// all five and PowerPoint renders a) through e).
    #[test]
    fn a_repeated_declared_start_continues_the_list() {
        let mut counters = ListCounters::default();
        let arabic = number("arabicPeriod", 4, true);
        let alpha = number("alphaLcParenR", 1, true);
        assert_eq!(
            (0..4)
                .filter_map(|_| marker(&mut counters, Some(&arabic), 0))
                .collect::<Vec<_>>(),
            ["4.", "5.", "6.", "7."]
        );
        let mut counters = ListCounters::default();
        assert_eq!(
            (0..5)
                .filter_map(|_| marker(&mut counters, Some(&alpha), 0))
                .collect::<Vec<_>>(),
            ["a)", "b)", "c)", "d)", "e)"]
        );
    }

    #[test]
    fn autonumber_roman_markers_bound_untrusted_start_values() {
        assert_eq!(format_autonum(0, "arabicPeriod").as_deref(), Some("1."));
        let largest = format_autonum(u32::MAX, "romanUcPeriod").unwrap();
        assert_eq!(largest.len(), 61);
        assert_eq!(largest, format!("{}DCCLXVII.", "M".repeat(52)));
    }

    #[test]
    fn text_outside_placeholders_takes_the_presentation_default_style() {
        let sized = |size| ParagraphProperties {
            default_run: Some(RunProperties {
                font_size_pt: Some(size),
                ..RunProperties::default()
            }),
            ..ParagraphProperties::default()
        };
        let levels = [sized(27.0), sized(20.0)];
        let under = ParagraphProperties {
            bullet: Some(character("-")),
            ..sized(12.0)
        };
        let body = Placeholder {
            placeholder_type: Some("body".to_owned()),
            index: Some(1),
            orientation: None,
            size: None,
        };
        let cascade = |placeholder| ParagraphCascade {
            default_style: &levels,
            default_paragraph: Some(&under),
            placeholder,
            ..ParagraphCascade::default()
        };
        let size = |properties: ParagraphProperties| properties.default_run.unwrap().font_size_pt;
        let plain = cascade(None).properties(0, 1, None);
        assert_eq!(plain.bullet, Some(character("-")));
        assert_eq!(size(plain), Some(20.0));
        assert_eq!(size(cascade(None).properties(0, 4, None)), Some(27.0));
        let placed = cascade(Some(&body)).properties(0, 1, None);
        assert_eq!((placed.bullet, placed.default_run), (None, None));
    }
}
