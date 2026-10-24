//! Font-independent paragraph semantics shared by rendering and export: the layout and master a
//! slide inherits from, the paragraph-property cascade, and list markers.

use ooxml_drawingml::{
    ColorValue, TableCellPosition, TableStyleFlags, Theme, resolve_color_value_to_hex_with_theme,
    resolve_theme_font_ref,
};
use pptx_parse::{
    Bullet, GraphicFrameData, ParagraphProperties, Placeholder, PptxPackage, RunProperties,
    ShapeNode, Slide, SlideLayout, SlideMaster, TextBody, builtin_table_style,
};
use yrs::Transact;

use crate::deck::{map_string, required_map, shape_base, shape_parts, slide_ref};
use crate::{
    DeckSession, EditError, EditResult, SHAPES, StorySnapshot, TextRunSnapshot, TextStyle,
};

impl DeckSession {
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
        let cell = cell_formatting(package, node, &shape.snapshot, story_id);
        if let Some((text, inherited)) = &cell {
            cascade = ParagraphCascade {
                primary: Some(text),
                master: Some(inherited),
                ..ParagraphCascade::default()
            };
        }
        let mut story = self.story(story_id)?;
        for (index, paragraph) in story.paragraphs.iter_mut().enumerate() {
            let mut properties = cascade.properties(index, paragraph.level, None);
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

fn cell_formatting<'a>(
    package: &'a PptxPackage,
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
    Some((
        &cell.text,
        TextBody {
            default_list_style: Some(Box::new(ParagraphProperties {
                default_run: Some(RunProperties {
                    bold: style.bold,
                    italic: style.italic,
                    color: style.color,
                    ..RunProperties::default()
                }),
                ..ParagraphProperties::default()
            })),
            ..TextBody::default()
        },
    ))
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
